"""
Gmailラベル分けエージェント (実行環境版)

入力: message_id, message_text, message_timestamp, sender, subject
処理: 内容からカテゴリを判定し、Gmail にラベルを付与する

ラベルは2階層。

  親ラベル   受信箱を6分類する。数を増やさない（増やすと分類が揺れる）
  子ラベル   親のうち一部だけ、差出組織ごとに切る
             例: 「取引/<組織名>」「要対応/<組織名>」

子ラベルは事前に用意しない。分類のたびに GMAIL_LIST_LABELS のキャッシュを引き、
無ければ GMAIL_CREATE_LABEL でネスト名のまま作る。
組織は増え続けるので、人が先にラベルを掘る運用にすると必ず追いつかなくなる。
"""
import asyncio
import json
from agent_sdk import (
    llm_call,
    GMAIL_LIST_LABELS,
    GMAIL_CREATE_LABEL,
    GMAIL_ADD_LABEL_TO_EMAIL,
)


# ============================================================
# ラベル定義
# ============================================================
# 値は Gmail 側で発行されるラベルID。利用環境のIDに差し替えて使う
# （GMAIL_LIST_LABELS の戻りに入っている）。
LABELS = {
    "要対応":   "<LABEL_ID_ACTION>",     # 自分が動かないと止まるもの
    "予定":     "<LABEL_ID_SCHEDULE>",   # 日時が決まる・決まっているもの
    "取引":     "<LABEL_ID_BILLING>",    # お金が動くもの
    "配信":     "<LABEL_ID_NEWSLETTER>", # 読み物。読まなくても困らない
    "通知":     "<LABEL_ID_SYSTEM>",     # 機械が出した記録
    "その他":   "<LABEL_ID_OTHER>",
}

# 子ラベル（差出組織ごと）を切る親カテゴリ。
# 「あとで組織単位で掘り返したくなるもの」だけに限る。
# 配信・通知まで組織で割るとラベルが爆発して、探すコストがかえって上がる。
ORG_SUBLABEL_PARENTS = {"要対応", "取引"}

# Gmail のシステムラベル。ここに触れると受信箱の状態が壊れるので付与対象から外す
SYSTEM_LABELS = {
    "INBOX", "SENT", "DRAFT", "SPAM", "TRASH", "UNREAD", "STARRED", "IMPORTANT",
    "CATEGORY_PERSONAL", "CATEGORY_SOCIAL", "CATEGORY_PROMOTIONS",
    "CATEGORY_UPDATES", "CATEGORY_FORUMS",
}


# ============================================================
# ラベルキャッシュ（display name -> label_id）
# ============================================================
_LABEL_CACHE: dict[str, str] = {}


async def refresh_label_cache() -> None:
    """GMAIL_LIST_LABELS を呼んで {name: id} のキャッシュを作る"""
    global _LABEL_CACHE
    res = await GMAIL_LIST_LABELS()
    # 連携基盤のレスポンス形を吸収
    data = res.get("data", res) if isinstance(res, dict) else res
    labels = (
        data.get("labels")
        or data.get("response_data", {}).get("labels")
        or []
    )
    _LABEL_CACHE = {lb["name"]: lb["id"] for lb in labels if "name" in lb and "id" in lb}


async def get_or_create_label_id(label_name: str) -> str:
    """ラベル名から ID を返す。なければ作成する（ネスト '取引/<組織名>' 対応）"""
    if not _LABEL_CACHE:
        await refresh_label_cache()
    if label_name in _LABEL_CACHE:
        return _LABEL_CACHE[label_name]

    created = await GMAIL_CREATE_LABEL(label_name=label_name)
    data = created.get("data", created) if isinstance(created, dict) else created
    new_id = data.get("id") or data.get("labelId") or data.get("response_data", {}).get("id")
    if not new_id:
        # 同名ラベルが並行して作られた場合はここに来る。作成失敗とは限らないので
        # 一覧を取り直して存在を確認してから落とす
        await refresh_label_cache()
        if label_name in _LABEL_CACHE:
            return _LABEL_CACHE[label_name]
        raise RuntimeError(f"label 作成失敗: {label_name}, raw={created}")

    _LABEL_CACHE[label_name] = new_id
    return new_id


# ============================================================
# Step1: LLM でカテゴリ判定 + 対象カテゴリなら差出組織を抽出
# ============================================================
CLASSIFY_PROMPT = """あなたはメール分類器です。以下のメールを判定基準に従って分類してください。

# 判定基準（優先順位: 要対応 > 予定 > 取引 > 配信 > 通知 > その他）

## 1. 要対応
自分が何か返さないと先に進まないもの。
- キーワード: ご確認ください / ご返信 / 承認 / 依頼 / 提出 / 期限 / 締切 / 要対応 / お願いします
- 判断の軸: 「読むだけで終わるか、自分の行動が要るか」。要るなら要対応

## 2. 予定
日時が決まる、または決まっているもの。
- キーワード: 日程調整 / 会議 / 打ち合わせ / 招待 / 予約確認 / リマインド / 開催のご案内
- カレンダー招待（.ics 添付、招待の承諾依頼）を含む

## 3. 取引
お金が動くもの。
- キーワード: 請求 / 支払い / 入金 / 出金 / 引き落とし / 残高 / 明細 / 振込 / 領収書 / 契約更新 / 決済

## 4. 配信
定期的に届く読み物。読まなくても実害が無いもの。
- キーワード: メルマガ / ニュースレター / キャンペーン / セール / 割引 / クーポン / 期間限定 / 配信停止

## 5. 通知
機械が出した記録。人が書いていないもの。
- 送信元: noreply@ / no-reply@ / notifications@ / alert@
- キーワード: ログイン通知 / パスワード変更 / バックアップ完了 / ステータス変更 / 自動返信

## 6. その他
上記1〜5のいずれにも該当しない場合。

# ルール
- 件名で候補を絞り、本文で確定する
- 複数該当する場合は、本文と最も直接的に一致する 1 つだけを選ぶ
- 必ず 1 つ選択する（該当なしは不可）
- 迷ったら優先順位の上位を取る。取りこぼすより、多めに拾って人が下げるほうが安い

# 差出組織の抽出
category が「要対応」「取引」のときだけ、送信者ドメイン・署名・件名から
**差出組織名を 1 語で** 抽出し "org" フィールドに入れる。
- 表記は一般的な正式表記に正規化する（法人格の有無・全角半角・略称の揺れをそろえる）
- ドメインが組織を表さない汎用メール（フリーメール等）で、本文からも判別できない場合は
  推測せず空文字 "" にする
- 上記2カテゴリ以外のときは必ず空文字 "" にする

# 入力メール
- 送信者: {sender}
- 件名: {subject}
- タイムスタンプ: {timestamp}
- 本文:
{body}
"""


async def classify_email(message: dict) -> dict:
    prompt = CLASSIFY_PROMPT.format(
        sender=message.get("sender", ""),
        subject=message.get("subject", ""),
        timestamp=message.get("message_timestamp", ""),
        body=message.get("message_text", ""),
    )
    schema = {
        "type": "object",
        "properties": {
            "category": {
                "type": "string",
                "enum": list(LABELS.keys()),
            },
            "org": {
                "type": "string",
                "description": "差出組織名。要対応・取引 以外、または判別できない場合は空文字",
            },
            "reason": {
                "type": "string",
                "description": "そのカテゴリにした理由。本文中の根拠を1文で",
            },
        },
        "required": ["category", "org", "reason"],
    }
    res = await llm_call(prompt, schema=schema)
    return res["data"]


# ============================================================
# Step2: ラベル付与
# ============================================================
async def apply_labels(message_id: str, label_ids: list[str]) -> dict:
    return await GMAIL_ADD_LABEL_TO_EMAIL(
        message_id=message_id,
        add_label_ids=label_ids,
    )


# ============================================================
# メインフロー
# ============================================================
async def label_email(message: dict) -> dict:
    # Step1: 分類
    classified = await classify_email(message)
    category = classified["category"]
    org = (classified.get("org") or "").strip()
    reason = classified["reason"]

    # 付与するラベル ID を組み立てる
    label_ids: list[str] = [LABELS[category]]
    child_label_name = None

    if category in ORG_SUBLABEL_PARENTS and org:
        # ネスト形式: 「取引/<組織名>」
        child_label_name = f"{category}/{org}"
        child_id = await get_or_create_label_id(child_label_name)
        label_ids.append(child_id)

    # システムラベルが紛れていないことを確認してから投げる。
    # 受信箱の状態（既読・重要度・受信トレイ在席）を書き換えてしまうと戻せない
    for lid in label_ids:
        assert lid not in SYSTEM_LABELS, f"システムラベルは付与しない: {lid}"

    # Step2: ラベル付与
    api_result = await apply_labels(message["message_id"], label_ids)

    return {
        "message_id": message["message_id"],
        "subject": message.get("subject", ""),
        "category": category,
        "org": org,
        "child_label_name": child_label_name,
        "label_ids": label_ids,
        "reason": reason,
        "api_result": api_result,
    }


# ============================================================
# 動作確認サンプル
# ============================================================
async def main():
    samples = [
        {
            "message_id": "msg_action",
            "sender": "sales@supplier.example.jp",
            "subject": "【ご確認】見積書送付の件（9/20 まで）",
            "message_timestamp": "2026-09-12T09:00:00+09:00",
            "message_text": "お見積書を添付いたします。9/20 までにご返信をお願いいたします。",
        },
        {
            "message_id": "msg_schedule",
            "sender": "pm@partner.example.jp",
            "subject": "定例MTG 日程調整のお願い",
            "message_timestamp": "2026-09-12T10:00:00+09:00",
            "message_text": "来週の定例について、候補日を3つお送りします。ご都合をお知らせください。",
        },
        {
            "message_id": "msg_billing",
            "sender": "noreply@card-x.example.jp",
            "subject": "9月分ご請求金額のお知らせ",
            "message_timestamp": "2026-09-05T10:00:00+09:00",
            "message_text": "今月のご請求金額は 35,210 円です。お引き落とし日は 9/27 です。",
        },
        {
            "message_id": "msg_newsletter",
            "sender": "news@media.example.com",
            "subject": "【週刊】今週の注目記事5本",
            "message_timestamp": "2026-09-12T07:00:00+09:00",
            "message_text": "今週の記事をまとめてお届けします。配信停止はこちら。",
        },
        {
            "message_id": "msg_system",
            "sender": "no-reply@storage.example.com",
            "subject": "バックアップが完了しました",
            "message_timestamp": "2026-09-12T03:00:00+09:00",
            "message_text": "2026-09-12 03:00 のバックアップが正常に完了しました。",
        },
    ]

    results = []
    for m in samples:
        r = await label_email(m)
        results.append(r)
        print(json.dumps(r, ensure_ascii=False, indent=2))
    return results


# 実行方法:
#   - code_execute（IPython）: このファイルを読み込んだ上で `await main()` を呼ぶ
#   - bash_execute（python script.py）: 下の __main__ ブロックが asyncio.run(main()) を実行
if __name__ == "__main__":
    asyncio.run(main())
