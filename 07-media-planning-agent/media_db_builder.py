"""
メディアDB構築AI 実行環境版 v4・50列（統合コール＋抜粋バッチ方式）

OneDriveトリガーで追加された媒体資料PDFから50項目を抽出し、マスタExcelに
媒体数ぶんの行を追記する。複合資料（1PDFに複数媒体）に対応。

処理の流れ:
  - コール1（全ページ添付）: 媒体数カウント＋先頭THRESHOLD媒体の抽出＋残り媒体のページ
    リスト＋共通ページの特定 を1回で行う（THRESHOLD以下の資料はこれだけで完結）。
  - コール2以降（THRESHOLD超の資料のみ）: 残りをTHRESHOLD媒体ずつ、抜粋PDF（対象媒体＋
    共通ページのみ）で並列抽出。ページ番号はAIに抜粋の通し順で書かせコードが原本番号へ変換。
  - 律速はサーバ側120秒タイムアウト。THRESHOLDはその余白から決めた上限。

属性（役職/職種/業種）はAIが values に%を直接返し、コードは範囲・欠損・桁の検査のみ行う。
従業員規模はAIが「区分名=値」文字列で返し、コードが3列へ範囲変換（跨り区分は重なり幅の
大きい方へ全額）。リードサービスは「付けるか」をAIが媒体別に判断し、文言はコードが多数決で統一。
コードで直せない値は "-"（空欄）にして続行し、AI修復ループは採らない。

途中再開: 実行状態を tmp/media_db_ckpt.json に段階保存。失敗・タイムアウト時に同じ引数で
再実行すると、済んだAIコールをスキップし未追記の行だけ追記する（重複追記は起きない）。
成功時のsummaryは tmp/media_db_summary.txt にも保存（stdout消失時の復旧用）。
デバッグ: python-hunter の実行トレースを tmp/trace_media_db.log に記録。
テスト時は WRITE_TO_EXCEL = False にする。

code_execute での実行例（トリガーJSONの値を渡す）:

    from ptc_media_db_build_50col import run_media_db_build
    report = await run_media_db_build(
        item_id="<payload.item.item_id>",
        file_name="<payload.item.name>",
        web_url="<payload.item.web_url>",
    )
    print(report["summary"])
"""

from __future__ import annotations

import asyncio
import json
import logging
import os
import re
import shutil
import time
import unicodedata
from collections import Counter
from pathlib import Path

logger = logging.getLogger(__name__)

# ============================================================
# 定数
# ============================================================

# スクリプト版数。summaryの末尾に出す＝RAGに載っているのが最新版かを実行結果から即判定できる。
# 中身を変えたらここも上げる（RAG反映漏れの検知用）
SCRIPT_VERSION = "2026-07-08j (属性コンパクト/規模パース強化/価格オプション除外/PV捏造ガード/結合資料優先/チェックポイント)"

WRITE_TO_EXCEL = True  # 本番運用: マスタExcelに媒体数ぶんの行を追記する（テスト時は False に戻す）

THRESHOLD = 12  # コール1で抽出する媒体数の上限（2026-07-08ユーザー指定で15→12。余白を厚く。
                # 属性はコンパクト方式に戻したので出力は軽いが、保守的に12を維持）

PDF_LOCAL_PATH = "tmp/media_source.pdf"  # トリガーは1ファイルずつ処理されるため固定パスで安全
EXCERPT_PATH_FMT = "tmp/excerpt_batch_{}.pdf"  # バッチ用抜粋PDF（バッチ番号ごと）
TRACE_LOG_PATH = "tmp/trace_media_db.log"  # hunter の実行トレースログ
ENABLE_TRACE = True  # False でトレース無効化（本処理には影響しない）

# 書込先Excel（固定値。接続と別人のドライブのため drive_id 必須。PDF取得も同じドライブ）
DRIVE_ID = "b!gwkc2lR-LUakOlaQjqZ9tZLiTQSVEyJCt3YptP4PtyL31KRxZhpgSLiJdUohS1tC"
EXCEL_ITEM_ID = "<ONEDRIVE_ITEM_ID>"
EXCEL_TABLE_ID = "{4EE4840D-B94B-4E19-8787-962976CC7678}"  # 波カッコ込み。エラー時は "Table1" でも可

# PDF読解用モデル（PDFをページ画像込みで読めるGeminiを指定する）
# gemini-3.5-flash に戻す（2026-07-08ユーザー決定。属性転記テストと同一モデルで精度保証を揃える）
# フォールバックも同一モデル＝失敗時は同モデルで1回リトライする
EXTRACT_MODEL = "gemini/gemini-3.5-flash"
FALLBACK_MODEL = "gemini/gemini-3.5-flash"

MAX_DL_ATTEMPTS = 3          # DL失敗は2回まで再試行（計3回）
MAX_EXCEL_ATTEMPTS = 3       # Excel追記の最大試行回数
MAX_PDF_BYTES = 64 * 1024 * 1024  # llm_call の添付サイズ上限
EXPECTED_COLS = 50
EXCEL_BATCH = 1              # Excel追記の並列数。1=直列（2026-07-08確定）。
                             # 16並列で21行を書いたら同一ファイルへの同時書き込みをGraph APIが
                             # 401で拒否（11行成功/5行失敗）。同一Excelへの並列書き込みは不可

# 系統別カラム名
A_COLUMNS: list[str] = [
    "メディア運営会社", "メディア名称", "メディアカテゴリ", "資料対象期間",
    "月間PV数", "月間UU数", "会員数",
]
AGE_COLUMNS: list[str] = ["年代_20代", "年代_30代", "年代_40代", "年代_50代", "年代_60代以上"]
ROLE_COLUMNS: list[str] = ["役職_経営者役員", "役職_部長", "役職_課長", "役職_係長主任", "役職_一般社員"]
JOB_COLUMNS: list[str] = [
    "職種_営業", "職種_情シス", "職種_企画", "職種_技術", "職種_製造",
    "職種_人事総務", "職種_経理財務", "職種_マーケ",
]
INDUSTRY_COLUMNS: list[str] = [
    "業種_IT", "業種_製造", "業種_金融", "業種_サービス", "業種_建設不動産",
    "業種_医療", "業種_官公庁教育", "業種_その他",
]
D_COLUMNS: list[str] = ["対応ファネル", "主要メニューフォーマット", "獲得リードの質", "外部システム連携"]
E_COLUMNS: list[str] = ["最低出稿金額", "CPL目安", "PV保証_最小"]

# 正規カラム名（最終値dictのキー全量。メタ・規模3列はコード側が確定する）
CANONICAL_COLUMNS: list[str] = (
    A_COLUMNS + AGE_COLUMNS + ROLE_COLUMNS + JOB_COLUMNS + INDUSTRY_COLUMNS
    + ["従業員規模"] + D_COLUMNS + E_COLUMNS + ["メディアの独自価値"]
)

# 属性（役職・職種・業種）はAIが%を直接 values に返すコンパクト方式（2026-07-08 再決定）。
# graphs転記方式はコール1の出力を膨らませ大資料でタイムアウトしたため撤回。分類ルールは
# プロンプトに残す（食品→その他 等）が、精密さは求めない（食品がどの業種でも合計が合えばよい）。
# 従業員規模だけは values に「区分名=値」文字列で受け取り、3列への変換はコードが範囲計算で行う。
AI_VALUE_COLUMNS: list[str] = list(CANONICAL_COLUMNS)

# マスタExcelの50列ヘッダー（実物の Sheet1 で確認済みの並び）
SIZE_COLUMNS: list[str] = [
    "規模_エンタープライズ(1000名以上)", "規模_SMB(100-999名)", "規模_スタートアップ零細(100名未満)",
]
PAYLOAD_HEADERS: list[str] = (
    ["メディアNo"] + A_COLUMNS
    + AGE_COLUMNS + ROLE_COLUMNS + JOB_COLUMNS + INDUSTRY_COLUMNS
    + SIZE_COLUMNS + D_COLUMNS + E_COLUMNS
    + ["PDFファイルURL", "データ更新日", "参照ページ一覧", "メディアの独自価値", "item_id", "file_name"]
)

# 数値として payload に書く列（%値・万単位値）
PERCENT_COLUMNS: set[str] = set(AGE_COLUMNS + ROLE_COLUMNS + JOB_COLUMNS + INDUSTRY_COLUMNS)
NUMERIC_COLUMNS: set[str] = PERCENT_COLUMNS | {"月間PV数", "月間UU数", "会員数", "最低出稿金額", "PV保証_最小"}

# 欠損表記の揺れを "-" に正規化する変換表（"–"は16媒体テストで実際に観測された揺れ）
DASH_TRANS = str.maketrans({"–": "-", "—": "-", "―": "-", "_": "-"})

# 数値カラムに紛れる単位・修飾語（「20万人以上」「28.5288万」「10万円」等）を除去する正規表現。
# 数値部だけ残して float 判定に通す。名称等の文字列カラムには適用しない（数値7カラム限定）。
_NUM_UNIT_RE = re.compile(r"(万|人|名|社|件|通|円|PV|UU|UB|imp|%|％|以上|以下|約|強|弱|程度|前後|\s)")

# メディアカテゴリの許容5択（これ以外は _finalize_media でコードが機械的に寄せる）
_CATEGORY_CHOICES = {"IT・通信", "製造・メーカー", "ビジネス総合", "経営・マネジメント", "業界専門誌"}

# リードサービスの正規トークン（ハイブリッド方式のコード統一で使用）
LEAD_Q_TOKENS: list[str] = ["名刺情報のみ", "役職・部門あり", "課題アンケート付与可"]
LEAD_I_TOKENS: list[str] = ["CSV納品", "MAツールAPI連携可"]

# 主要メニューフォーマットの正規8種と表記揺れヒント（規格外語をDBに残さない）
MENU_TOKENS: list[str] = ["ディスプレイ広告", "タイアップ", "メール広告", "ホワイトペーパー",
                          "セミナー", "製品掲載", "製品レビュー", "リードジェネレーション"]
_MENU_HINTS: dict[str, tuple] = {
    "ディスプレイ広告": ("バナー", "純広"),
    "タイアップ": ("記事広告",),
    "メール広告": ("メルマガ", "メール"),
    "ホワイトペーパー": ("WP",),
    "セミナー": ("ウェビナー", "イベント"),
    "リードジェネレーション": ("リードジェン", "リード獲得"),
}
_CPL_RANGES = ("〜5,000円", "5,000〜10,000円", "10,000〜20,000円", "20,000円〜")

# 従業員規模3列の帯（従業員規模の「区分名=値」をどの列に入れるかの重なり幅判定に使用。
# 跨り区分は重なり幅の大きい方に全額入れる＝比率按分はしない）
_SIZE_BUCKETS: list[tuple[float, float]] = [(1000.0, float("inf")), (100.0, 1000.0), (1.0, 100.0)]
_SIZE_BUCKET_COLS: list[str] = ["エンタープライズ", "SMB", "スタートアップ零細"]
_SIZE_INF_CAP = 10000.0  # 「○○以上」区分の幅計算用の仮上限

# ============================================================
# llm_call 構造化出力スキーマ（v4: カラム名を固定キーとして埋め込む）
# 属性（役職/職種/業種）はAIが values に%を直接返すコンパクト方式。従業員規模は「区分名=値」文字列
# ============================================================

_MEDIA_ITEM: dict = {
    "type": "object",
    "properties": {
        "name": {"type": "string"},
        "page_start": {"type": "integer"},
        "page_end": {"type": "integer"},
        "values": {
            "type": "object",
            "properties": {c: {"type": "string"} for c in AI_VALUE_COLUMNS},
            # values の中は意図的に required にしない（無い値を無理に埋めさせない＝捏造防止）
        },
        "age_labels": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {"label": {"type": "string"}, "value": {"type": "number"}},
                "required": ["label", "value"],
            },
        },
        "price_candidates": {"type": "string"},
    },
    "required": ["name", "page_start", "page_end", "values"],
}

CALL1_SCHEMA: dict = {
    "type": "object",
    "properties": {
        "media_count": {"type": "integer"},
        "extracted": {"type": "array", "items": _MEDIA_ITEM},
        "remaining": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "name": {"type": "string"},
                    "pages": {"type": "array", "items": {"type": "integer"}},
                },
                "required": ["name", "pages"],
            },
        },
        "common_pages": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {"page": {"type": "integer"}, "topic": {"type": "string"}},
                "required": ["page", "topic"],
            },
        },
    },
    "required": ["media_count", "extracted", "remaining", "common_pages"],
}

BATCH_SCHEMA: dict = {
    "type": "object",
    "properties": {"media": {"type": "array", "items": _MEDIA_ITEM}},
    "required": ["media"],
}

# ============================================================
# プロンプト（照合テストで蓄積した本番ルール全文。値のみ形式に適合）
# ============================================================

VALUE_RULES = """
## 手順
1. 資料全体から「属性セクション」（年代/役職/職種/業種/従業員規模/会員数のグラフ）と
   「共通ページ」（リードサービス/会社概要/共通の広告メニュー）の場所を先に特定する。
   属性は媒体の個別ページの外（資料末尾など）にあることが多い。
2. 各媒体の個別ページ＋共通ページから各カラムの値を埋める。

## 媒体全体の値だけを使う（次のものは媒体全体ではないので採用しない）
- フォーラム/チャンネル/サブサイト別のPV（全体186万なのにモビリティ35万を拾わない）
- イベント/セミナー参加者アンケートの属性・登録者数
- 管理画面/レポートの「画面イメージ・サンプル」スクショ内のデモ数値
- nが極端に小さい創刊時アンケート（媒体全体の会員グラフがあればそちらを優先）

## 記載がない値は "-"（推測・外部知識での補完は禁止＝捏造が最悪のミス）。
概数と精密値が両方あれば精密値（「約28万人」と「285,288名」→ 28.5288）。

## 結合資料（複数の元資料を1つのPDFにまとめたもの）での優先順位【全カラム共通】
1つのPDFに複数の元資料（媒体ガイド／リード資料／共通の会員プロフィール／末尾の全媒体紹介
カタログ 等）が結合されていることがある。同じ媒体の同じ項目が複数箇所に別の値で出てきたら、
次の順で上位にある値を採る（PV・会員数・期間・価格などすべてのカラムに適用）:
  ① その媒体自身の詳細資料（媒体ガイド・メディアシート本体）＝最優先
  ② 全社共通の会員プロフィール・読者データ専用ページ（共通の属性・共通会員基盤はここが正）
  ③ 末尾の全媒体紹介カタログ・他資料（リード資料等）内での言及＝最も弱い（要約・古い値の恐れ）
- 上位に値があれば必ずそれを採る。③カタログにしか載っていない媒体は③の値でよい。
- 同じ資料・同じ順位の中で実績日が違う2値があるときは、**媒体規模として正式に記載された
  スペック値**を優先する（「月間PV ◯◯万」等）。本文中のプロモ表現（「◯◯突破しました」）は
  たとえ日付が新しくても採らない。
- 食い違いを理由に "-" にはしない（必ず上位の値を採る）。

## values のカラム（キーはこの正規名のみ）
- """ + " / ".join(AI_VALUE_COLUMNS) + """
※**値のある（記載がある）カラムだけを返す。記載が無い/該当しないカラムはキーごと省略してよい**
  （出力を減らすため。省略したカラムはコード側が "-"・「該当なし」で自動補完する）。
  無い値を "-" や "該当なし" と明記して返す必要はない。ただし値を捏造して埋めるのは厳禁。
※役職(5)/職種(8)/業種(8)は各カラムに%を直接入れる。従業員規模は「区分名=値」の文字列で返す（下記ルール）

## カラム別
- メディア運営会社: 会社概要・©表記の正式社名
- メディアカテゴリ: 掲載内容で5択から1つ推論。定義:
  IT・通信=IT/デジタル技術が主題 / 製造・メーカー=製造業・ものづくりが主題 /
  経営・マネジメント=経営層向けの経営課題が主題 / 業界専門誌=特定業界・職域に読者を絞る
  （IT業界に絞る場合はIT・通信を優先）/ ビジネス総合=領域横断・一般ビジネス
- 資料対象期間: 「YYYY年M月」。優先順位=**発行年月＞対象期間の表記＞本文の基準日注記**
  （「2025年10-12月版・2025年9月発行」→2025年9月。対象期間しか無い範囲表記は開始月。
  月が特定できなければ「YYYY年」まで。年も無ければ本文の基準日注記から）。
  同一資料の全媒体で同じ値になる
- 月間PV数/月間UU数: 媒体全体のサマリー値のみ（万単位の数字）。
  **万への換算は÷10,000**。例: 26,400pv→2.64 / 133,000UB→13.3 / 「約230万PV」→230
  （÷1,000にしない。換算ミスは最悪の事故）。
  **「◯◯PV」「◯◯UB/UU」と明記された実数が無ければ必ず "-"。次の数字は絶対にPV/UUにしない**:
  想定imp・配信imp / 会員数・メール会員数・登録者数 / **リード獲得件数・「実施例」「◯◯件」等の
  提案プランのサンプル件数** / 料金・掲載枠数・部数。これらをPV/UUに転用・換算しない
  （imp課金モデルやリードジェネ資料はPV/UU記載が無いことが多い＝その場合は"-"。
  実害例: 日経クロステックで「約200件（リード実施例）」をPV=200に、imp/会員をUUに捏造した）。
  複数箇所で食い違う場合は 実績＞想定、全体＞内訳 を優先し、元資料をまたぐ食い違いは
  上の【結合資料での優先順位】に従う（詳細資料＞共通ページ＞付録カタログ。ただし
  どこにも実数が無ければ"-"）。
  UUとして可=ユニーク系（UU/UB/ユニークユーザー/ユニークビジター）。
  セッション・訪問回数・DAUはUUではない（それしか無ければ"-"）
- 会員数: 媒体の会員基盤（万単位）。メルマガ配信数(通)/メール配信可能数/調査パネル数/
  SNSフォロワー/問合せ数/掲載製品数/グループ累計は会員数ではない。
  **リードサービス等のサービス説明文中に出てくる会員規模（「約100万人のID会員向け」等）も
  媒体の会員数にしない**（会員プロフィール・媒体データ専用ページの値のみ使う）。
  共通会員基盤の値は、**共通ページにその媒体名が列挙されている媒体だけ**に入れる
  （名前が無い媒体は"-"。姉妹媒体で同値になるのは正当）。会員かどうか迷ったら"-"
- 年代: 正規5区分(20/30/40/50/60代以上)に一致すれば入れる。違えば年代_を全て"-"にして
  age_labels に原文の全区分 [{label,value}] を出す（コードが変換する。自分で合算しない）
- 役職(5)/職種(8)/業種(8): 読者・会員構成の%（0〜100スケール）を各カラムに直接入れる。
  グラフの各区分を下のカラムに割り当て、同じカラムに入る区分は**合算**して%を入れる。
  会員全体のグラフを使う（全体が無ければ内訳を使い、その旨は気にしなくてよい）。
  ベンダー/ユーザー等の企業属性は業種に使わない。**カラムに当てはまらない区分は捨ててよい**
  （職種・役職の「その他」「学生」「契約社員」「派遣社員」等は入れる列が無いので無視。
  結果その軸の合計が100%未満になってよい）。割り当てルール:
  ・業種(IT/製造/金融/サービス/建設不動産/医療/官公庁教育/その他): 8番目の「その他」が受け皿
    なので業種は必ず合計100%になる。製造=資料が製造/メーカーと呼ぶ区分（食品・素材・化学等の
    個別業種名で迷ったら「その他」でよい＝製造かその他かは厳密でなくてよい）。
    卸/小売/商社/流通/人材/飲食/宿泊/コンサル/会計/士業/広告/出版/専門サービス→サービス。
    運輸/物流/電力/ガス/エネルギー/農林水産/鉱業→その他。情報処理/情報通信/ソフトウェア/通信→IT
  ・職種(営業/情シス/企画/技術/製造/人事総務/経理財務/マーケ): 生産/製造/品質/研究開発/設計→製造、
    情報システム/SE/ITエンジニア/DX推進→情シス、上記以外の技術職→技術、販売/セールス→営業、
    経営/役員/経営企画→企画、広報/宣伝/調査→マーケ、財務/経理→経理財務、総務/人事/事務→人事総務。
    複合ラベルは主たる語で判断（現場職の語を含めば製造）
  ・役職(経営者役員/部長/課長/係長主任/一般社員): 本部長/事業部長/工場長→部長、マネージャー→課長、
    リーダー→係長主任、複合「一般社員・その他」→一般社員
- 従業員規模: values の「従業員規模」に「区分名=値;区分名=値;…」の**文字列**で返す
  （区分名は原文のまま。例「1〜299人=39;300〜999人=18;1,000〜9,999人=25;10,000人以上=18」）。
  3列(エンタープライズ/SMB/零細)への変換はコードが範囲計算で行う（自分で分類・合算しない）
- 選択肢カラム（主要メニューフォーマット/獲得リードの質/外部システム連携）:
  記載・適用があるものだけ返す（無ければキーごと省略可。コードが「該当なし」を補う）。
  **下の選択肢の文言を一字一句そのまま使う**:
  ・主要メニューフォーマット: **次の8種だけ**を;区切りで列挙（個別メニュー名・独自の種別名は書かない）:
    ディスプレイ広告 / タイアップ / メール広告 / ホワイトペーパー / セミナー /
    製品掲載 / 製品レビュー / リードジェネレーション
    集約の原則: イベント協賛→セミナー、記事広告・動画タイアップ→タイアップ。
    広告出稿でないもの（調査・制作代行・運用代行）は含めない
  ・獲得リードの質: 名刺情報のみ / 役職・部門あり / 課題アンケート付与可（名刺情報のみは排他）。
    **「役職・部門あり」は納品されるリード項目に役職/部署が含まれる場合のみ**
  ・外部システム連携: CSV納品 / MAツールAPI連携可（連携可は明示記載がある時のみ）
- 対応ファネル: 記入不要（"-"でよい。メニューからコードが導出する）
- CPL目安: リード単価の最安を**円の数字だけ**で（例: 8000）。明示単価が無ければ「該当なし」
  （レンジへの変換はコードが行う）
- 全社共通リードサービスは媒体ごとに実施可否を判断（純広告・コンシューマ媒体→該当なし、
  別会社運営→該当なし。同じ基準を全媒体に適用）
- 最低出稿金額: まず price_candidates に価格候補を安い順に最大6件、
  「名前=金額[種別]」を;区切りで列挙する。金額は万円の数字。種別は次の4つ:
  [固]=1回の発注で必ず支払う固定額（掲載料含む。成果報酬型でも固定の掲載料は[固]）
  [初]=初期費用・最低利用料・最低配信価格 [成果]=成果報酬の単価部分 [単]=imp/クリック/通あたり単価
  例: メール号外=15[固];タイアップ=120[固];初期費用=3[初];レクタングル=1.5円/imp[単]
  （メール広告は表の全枠 ヘッダー/センター等 を比較して安い枠も含める）
  **候補に入れないもの（絶対）**: 「オプション」「オプション:」と書かれた項目は金額に関わらず
  price_candidates に一切入れない。再誘導・ブースト・クリップ・属性/企業名指定(ABM)・
  X投稿・パンフレット・初校/念校の修正費など、既存メニューに付随する追加料金も入れない。
  全社共通リードサービスの料金も入れない。
  候補は**それ単体を発注できる固定費メニュー**だけ（例: タイアップ本体・バナー・メール号外）。
  ※オプションを固定費として拾うと最安がオプション額になり誤る（実害例あり）
  金額は**税込を優先**（税抜/税込が併記なら税込。税込が無ければ税抜のまま。換算はしない）
  最低出稿金額 カラムには[固]の最安（無ければ[初]の最安）を書く。[単][成果]は候補外。0は書かない
- PV保証_最小: 最低出稿金額で採用したプラン自身のPV保証のみ（万PV）
- メディアの独自価値: 100字以内の要約

## 属性（役職・職種・業種・従業員規模）の対象の選び方
- 会員全体・読者全体のグラフを使う（全体が無ければ内訳でよい）。
- 対象外: 管理画面/レポートのスクショ内デモ数値・イベント/セミナー参加者アンケート・発行部数等
- 年代は values の年代_5列 / age_labels で返す（属性%とは別扱い）
"""

CALL1_PROMPT = f"""あなたはBtoBメディアの媒体資料PDFからデータを抽出する専門家です。
添付PDFを冒頭・中盤・末尾まで全ページ読み、次の4つを1回で返してください。

1. media_count: 資料に収録されている媒体の総数
   - 専用の紹介ページが1枚でもあれば1媒体（表紙・ロゴ一覧に名前だけの媒体は数えない）
   - 同一サイト内の分野別セクションは1媒体にまとめる
     （例: 日経クロステックActiveの「IT」「製造」「建設」→ Active 1件）
   - 別サイト・別ドメインとして運営されている場合のみ別媒体に分ける
     （例: イプロスものづくり/都市まちづくり/医薬食品技術 → 3件）
2. extracted: **掲載順で最初の{THRESHOLD}媒体**のカラム値
   （media_countが{THRESHOLD}以下なら全媒体。{THRESHOLD + 1}媒体目以降はここに含めない）
3. remaining: {THRESHOLD + 1}媒体目以降の「媒体名＋そのデータが載っているページ番号のリスト」
   （値の抽出はしない。media_countが{THRESHOLD}以下なら空配列で返す）
4. common_pages: 特定の媒体に属さず全媒体に効くページ（page と topic を返す）
   - 全社共通のリードサービス / 会社概要・コピーライト / 共通の広告メニュー・入稿規定
   - **読者・会員の属性データ（年代・役職・職種・業種・従業員規模・会員数）が
     媒体の個別ページの外にまとまっている場合、そのページも必ず含める**
     （例: 資料末尾の「会員プロフィール」「読者属性」セクション）
   - extracted の各媒体の値は、媒体個別ページだけでなくこの共通ページも根拠にしてよい
     （例: 資料末尾の会員属性を全媒体・親媒体の属性として採用する場合）

※ページ番号はすべてPDFの物理的な通し順（1枚目=p.1）。ページ内に印字された
  スライド番号ではない。
""" + VALUE_RULES


def _batch_prompt(media_names: list[str], n_pages: int, total_pages: int) -> str:
    """バッチ抽出コールのプロンプトを組み立てる（抜粋PDF添付・通し番号方式）"""
    names = "\n".join(f"- {n}" for n in media_names)
    return (
        "あなたはBtoBメディアの媒体資料PDFからデータを抽出する専門家です。\n"
        f"次の{len(media_names)}媒体それぞれについてカラムの値を抽出してください。\n"
        "指定された媒体以外は返さないでください。\n\n"
        f"## 抽出対象媒体\n{names}\n\n"
        "## 添付PDFのページ番号について（重要）\n"
        f"- 添付は原本（全{total_pages}ページ）から{n_pages}ページだけを抜粋したPDFです。\n"
        "- page_start / page_end は**添付PDFの通し順**で書くこと\n"
        "  （1枚目=1、2枚目=2、…）。原本のページ番号を推測して書かない。\n"
        "  ページ内に印字されたスライド番号も使わない。\n"
        "- 抜粋には全媒体共通のページ（会社概要・共通リードサービス等）も含まれている。\n"
        "  運営会社・リードサービス・会員属性の判断にはその共通ページも根拠にしてよい。\n"
        + VALUE_RULES
    )

# ============================================================
# 内部ヘルパー（抜粋PDFのページ対応）
# ============================================================


def _ensure_pypdf() -> None:
    """pypdf が無い環境ではサブプロセスのpipで導入する（ページ数取得・抜粋PDF生成に必須）

    プロセス内で pip._internal を呼ぶ方式は禁止（以降の全importにpipのDEPRECATION警告が
    連鎖し、stdout 80KB上限をノイズが食い潰して本来の結果が消える。本番で実害を観測済み）。
    """
    try:
        import pypdf  # noqa: F401
    except ImportError:
        import subprocess
        import sys
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "--quiet", "pypdf"],
            capture_output=True, timeout=120,
        )
        import pypdf  # noqa: F401  # 導入失敗ならここでImportError→呼び出し元に伝播


def _build_excerpt_pdf(src_pdf: str, pages: list[int], out_path: str) -> dict[int, int]:
    """原本PDFから指定ページだけの抜粋PDFを作り、通し番号→原本番号の対応表を返す

    pages は原本の物理ページ番号（1始まり。ページ内に印字された番号ではない）。
    重複除去・昇順化・範囲外の除外を行い、抜粋PDF内の通し番号 n 枚目が
    原本の何ページかを {n: 原本ページ} で返す。
    例: pages=[1,2,3,4,5,10] → 抜粋は6枚、対応表 {1:1, ..., 5:5, 6:10}
    """
    from pypdf import PdfReader, PdfWriter

    reader = PdfReader(src_pdf)
    total = len(reader.pages)
    valid = sorted({p for p in pages if 1 <= p <= total})
    if not valid:
        raise Exception(f"抜粋対象ページがありません（指定: {pages}, 原本: {total}ページ）")
    writer = PdfWriter()
    for p in valid:
        writer.add_page(reader.pages[p - 1])
    with open(out_path, "wb") as f:
        writer.write(f)
    return {i: p for i, p in enumerate(valid, start=1)}


def _to_original_page(n, page_map: dict[int, int]):
    """LLMが返した抜粋PDFの通し番号を原本のページ番号へ変換する

    LLMは常に抜粋の通し順（1枚目=1）で返す約束（_batch_prompt の指示）なので、
    変換は対応表を引くだけ。例: 抜粋が原本[1,3,6,12,17,18]なら 4 → 12。
    対応表に無い番号（範囲外・非整数）は変換せずそのまま返す。
    """
    if isinstance(n, int) and n in page_map:
        return page_map[n]
    return n

# ============================================================
# 内部ヘルパー（実行トレースロギング）
# ============================================================


def _start_trace():
    """python-hunter で本モジュールの関数呼び出し/リターンをトレースログに記録する

    デバッグ用のロギング。行単位トレースは膨大になるため kind を call/return に絞る。
    hunter が無い環境では pip install を試み、それでも使えなければ黙ってスキップして
    本処理は続行する（ロギングのために処理を止めない）。
    戻り値は _stop_trace に渡すログストリーム（無効時は None）。
    """
    if not ENABLE_TRACE:
        return None
    try:
        try:
            import hunter
        except ImportError:
            # pip はサブプロセスで隔離実行する。プロセス内で pip._internal を呼ぶと
            # 以降の全 import に pip の DEPRECATION 警告が出て stdout の80KB上限を
            # ノイズが食い潰し、本来の結果・エラーが切り落とされる（本番で実害を観測済み）
            import subprocess
            import sys
            subprocess.run(
                [sys.executable, "-m", "pip", "install", "--quiet", "hunter"],
                capture_output=True, timeout=60,
            )
            import hunter
    except Exception as e:
        logger.warning(f"hunter が利用できないためトレースなしで続行: {e}")
        return None
    try:
        Path("tmp").mkdir(exist_ok=True)
        stream = open(TRACE_LOG_PATH, "w", encoding="utf-8")
        hunter.trace(
            module_startswith="ptc_media_db_build",
            kind_in=("call", "return"),
            action=hunter.CallPrinter(stream=stream),
        )
        return stream
    except Exception as e:
        logger.warning(f"hunter トレースを開始できませんでした（続行）: {e}")
        return None


def _stop_trace(stream) -> None:
    """トレースを停止してログストリームを閉じる（開始していなければ何もしない）"""
    if stream is None:
        return
    try:
        import hunter

        hunter.stop()
        stream.close()
        print(f"[trace] 実行トレース: {TRACE_LOG_PATH}")
    except Exception as e:
        logger.warning(f"hunter トレースの停止に失敗: {e}")

# ============================================================
# 内部ヘルパー（外部ツール・LLM呼び出し）
# ============================================================


async def _composio_call(tool_name: str, args: dict) -> dict:
    """Composioツールを呼び出す（agent_sdk直接import → tool_bridge の順で試す）"""
    import agent_sdk

    fn = getattr(agent_sdk, tool_name, None)
    if fn is not None:
        return await fn(**args)
    from agent_sdk import tool_bridge

    return await tool_bridge(tool_name, args)


async def _download_pdf(item_id: str, file_name: str) -> None:
    """OneDriveから対象PDFを取得して tmp/ に実体化する（2回まで再試行）

    サイズ超過はリトライで解決しないため、取得成功後にループの外で判定する。
    """
    import requests

    Path("tmp").mkdir(exist_ok=True)
    last_error = ""
    for attempt in range(1, MAX_DL_ATTEMPTS + 1):
        try:
            dl = await _composio_call(
                "ONE_DRIVE_DOWNLOAD_FILE",
                {"item_id": item_id, "file_name": file_name, "drive_id": DRIVE_ID},
            )
            if not dl.get("successful"):
                last_error = str(dl.get("error", "unknown"))
                logger.warning(f"PDFダウンロード失敗 ({attempt}/{MAX_DL_ATTEMPTS}): {last_error}")
                continue
            s3url = dl["data"]["content"]["s3url"]
            if s3url.startswith("http"):
                # 大きいファイルでもメモリに全載せしないようストリーミング書き込み
                with requests.get(s3url, allow_redirects=True, stream=True) as r:
                    r.raise_for_status()
                    with open(PDF_LOCAL_PATH, "wb") as f:
                        for chunk in r.iter_content(chunk_size=1024 * 1024):
                            f.write(chunk)
            else:
                shutil.copy(s3url, PDF_LOCAL_PATH)
            break  # 取得成功
        except Exception as e:
            last_error = str(e)
            logger.warning(f"PDFダウンロード例外 ({attempt}/{MAX_DL_ATTEMPTS}): {e}")
    else:
        raise Exception(
            f"PDFのダウンロードに{MAX_DL_ATTEMPTS}回失敗しました。ユーザーに報告してください。"
            f" item_id={item_id} file_name={file_name} 最終エラー: {last_error}"
        )

    size = Path(PDF_LOCAL_PATH).stat().st_size
    if size > MAX_PDF_BYTES:
        raise Exception(
            f"PDFが {size // (1024 * 1024)}MB あり llm_call の添付上限(64MB)を超えています。"
            "実行環境では処理できないため、このファイルは手動対応してください。"
        )
    print(f"[Step1] PDF取得完了: {file_name} ({size // 1024}KB) → {PDF_LOCAL_PATH}")


async def _llm_structured(prompt: str, *, schema: dict, context: str,
                          attach_path: str = "") -> dict:
    """PDF添付で llm_call の構造化出力を取得する（失敗時は同モデルで1回再試行）

    attach_path を指定するとそのファイルを添付する（既定は原本PDF）。
    バッチ抽出で抜粋PDFを渡すために使う。
    """
    from agent_sdk import ToolCallError, llm_call

    last_error = ""
    for model in (EXTRACT_MODEL, FALLBACK_MODEL):
        t0 = time.time()
        try:
            res = await llm_call(
                prompt=prompt, model=model,
                file_paths=[attach_path or PDF_LOCAL_PATH], schema=schema,
            )
            elapsed = time.time() - t0
            if isinstance(res, dict) and "error" in res:
                # 空文字のerrorはサーバ側120秒タイムアウト（出力超過）の典型パターン
                last_error = str(res["error"]) or "(空エラー=サーバ側タイムアウトの疑い)"
                logger.warning(f"{context}: llm_call失敗 {elapsed:.1f}秒 ({model}): {last_error}")
                continue
            if isinstance(res.get("data"), dict):
                usage = res.get("usage") or {}
                print(f"[llm] {context}: {elapsed:.1f}秒 "
                      f"in={usage.get('input_tokens', '?')} out={usage.get('output_tokens', '?')}")
                return res["data"]
            last_error = "構造化出力が不正"
            logger.warning(f"{context}: 構造化出力が不正 ({model})")
        except ToolCallError as e:
            last_error = str(e)
            logger.warning(f"{context}: ToolCallError {time.time() - t0:.1f}秒 ({model}): {e}")
    if "inline limit" in last_error:
        raise Exception(
            f"{context}: PDFがこの環境の llm_call 添付上限を超えています（{last_error}）。"
            "PDFを圧縮するか手動対応してください。"
        )
    raise Exception(f"{context}: LLM呼び出しが全モデルで失敗しました（最終エラー: {last_error}）")

# ============================================================
# Step2〜3: 統合コール＋抜粋バッチ抽出
# ============================================================


def _norm_name(s) -> str:
    """媒体名の照合用正規化（全角半角・大小文字・空白を吸収）"""
    return unicodedata.normalize("NFKC", str(s)).lower().replace(" ", "")


async def _call1() -> dict:
    """コール1（統合コール）: 媒体数＋先頭THRESHOLD媒体の抽出＋残りページリスト＋共通ページ"""
    data = await _llm_structured(CALL1_PROMPT, schema=CALL1_SCHEMA, context="コール1(統合)")
    extracted = [m for m in data.get("extracted", []) if isinstance(m, dict) and m.get("name")]
    remaining = [r for r in data.get("remaining", []) if isinstance(r, dict) and r.get("name")]
    common_pages = [c for c in data.get("common_pages", []) if isinstance(c, dict)]
    media_count = data.get("media_count", 0)
    if not extracted:
        raise Exception("収録媒体が1件も抽出されませんでした（媒体資料でない可能性）")
    print(f"[Step2] media_count={media_count} / 抽出済み={len(extracted)}件 / "
          f"残り={len(remaining)}件 / 共通ページ={len(common_pages)}件")
    # 宣言された総数と実物が矛盾しても remaining の実物を正として動く
    if media_count > len(extracted) and not remaining:
        logger.warning(f"media_count={media_count} なのに remaining が空（実物を正として続行）")
    return {"extracted": extracted, "remaining": remaining,
            "common_pages": common_pages, "media_count": media_count}


async def _extract_one_batch(idx: int, batch: list[dict], common_page_nums: list[int],
                             total_pages: int) -> list[dict]:
    """1バッチ分の媒体を抜粋PDFで抽出し、ページ番号を原本番号へ変換して返す"""
    pages = sorted({p for rm in batch for p in rm.get("pages", []) if isinstance(p, int)}
                   | set(common_page_nums))
    pages = [p for p in pages if 1 <= p <= total_pages]
    excerpt = EXCERPT_PATH_FMT.format(idx)
    page_map = _build_excerpt_pdf(PDF_LOCAL_PATH, pages, excerpt)
    prompt = _batch_prompt([rm["name"] for rm in batch], len(page_map), total_pages)
    try:
        data = await _llm_structured(prompt, schema=BATCH_SCHEMA,
                                     context=f"バッチ{idx}({len(batch)}媒体)",
                                     attach_path=excerpt)
    except Exception as e:
        logger.warning(f"バッチ{idx}が失敗: {e}")
        return []
    rows = [m for m in data.get("media", []) if isinstance(m, dict) and m.get("name")]
    # ページ番号を抜粋の通し順 → 原本番号へ変換（コードが100%行う）
    for m in rows:
        m["page_start"] = _to_original_page(m.get("page_start"), page_map)
        m["page_end"] = _to_original_page(m.get("page_end"), page_map)
    return rows


async def _extract_batches(remaining: list[dict], common_pages: list[dict],
                           total_pages: int) -> tuple[list[dict], list[str]]:
    """残り媒体をTHRESHOLD件ずつのバッチで並列抽出する

    返却媒体数が指定より少ないバッチは、欠けた媒体だけの小バッチで1回だけ再試行する
    （出力溢れ・タイムアウト対策。それでも欠けた媒体は「抽出失敗」として報告）。
    """
    common_page_nums = [c.get("page") for c in common_pages if isinstance(c.get("page"), int)]
    batches = [remaining[i:i + THRESHOLD] for i in range(0, len(remaining), THRESHOLD)]
    print(f"[Step3] バッチ抽出: {len(remaining)}媒体を{len(batches)}バッチで並列実行")
    results = await asyncio.gather(
        *(_extract_one_batch(i + 1, b, common_page_nums, total_pages)
          for i, b in enumerate(batches))
    )
    rows: list[dict] = [m for batch_rows in results for m in batch_rows]

    # 欠け媒体の検出 → 小バッチで1回だけ再試行
    got = {_norm_name(m["name"]) for m in rows}
    missing = [rm for rm in remaining if _norm_name(rm["name"]) not in got]
    if missing:
        print(f"[Step3] 未返却{len(missing)}媒体を小バッチで再試行: "
              + " / ".join(rm["name"] for rm in missing))
        retry_rows = await _extract_one_batch(0, missing, common_page_nums, total_pages)
        rows += retry_rows
        got |= {_norm_name(m["name"]) for m in retry_rows}
    failed = [rm["name"] for rm in remaining if _norm_name(rm["name"]) not in got]
    return rows, failed

# ============================================================
# Step4: コード清書（正規化・リード文言統一。追加のAIコールなし）
# ============================================================


def _normalize_ages(age_labels: list[dict]) -> tuple[dict[str, float], list[str]]:
    """PDFの年代区分を正規5区分へ変換（各区分を上限年齢の年代に割当）"""
    buckets = {"20代": 0.0, "30代": 0.0, "40代": 0.0, "50代": 0.0, "60代以上": 0.0}
    dropped: list[str] = []
    for it in age_labels:
        label = str(it.get("label", ""))
        val = float(it.get("value", 0))
        nums = [int(n) for n in re.findall(r"\d+", label)]
        if not nums:
            dropped.append(f"{label}={val}")
            continue
        open_end = len(nums) == 1 and any(k in label for k in ("以上", "〜", "～", "+", "over"))
        if open_end:
            upper = 999
        elif "代" in label and len(nums) == 1:
            upper = nums[0] + 9   # 「30代」→ 上限39
        else:
            upper = max(nums)     # 「25-34」→ 上限34
        decade = 60 if upper >= 60 else (upper // 10) * 10
        if 20 <= decade <= 50:
            buckets[f"{decade}代"] += val
        elif decade >= 60:
            buckets["60代以上"] += val
        else:
            dropped.append(f"{label}={val}")  # 10代などは正規区分の対象外
    return buckets, dropped


def _fmt_pct(v: float) -> str:
    s = f"{v:.1f}"
    return s[:-2] if s.endswith(".0") else s


def _canon_choice(value: str, tokens: list[str]) -> str:
    """選択肢カラムの値を正規トークンに分解し、正規の順序・「;」区切りで組み直す

    「A/B」「B;A」のような区切り・順序の揺れを吸収する。トークンが1つも見つからなければ
    「該当なし」。「名刺情報のみ」は排他選択肢なので他の属性があれば外す。
    """
    s = str(value)
    found = [t for t in tokens if t in s]
    if not found:
        return "該当なし"
    if "名刺情報のみ" in found and len(found) > 1:
        found.remove("名刺情報のみ")
    return ";".join(t for t in tokens if t in found)


_PRICE_KIND_RE = re.compile(r"[\[［](固|初|成果|単)[\]］]")


def _select_min_price(candidates: str, ai_value: str, media_name: str) -> str:
    """price_candidates から優先順位でコードが最低出稿金額を確定する

    優先順位: [固]固定掲載費の最安 → 無ければ[初]初期費用・最低利用料の最安。
    [単]配信単価・[成果]成果報酬は除外。種別マーカーが無い候補はテキストから推定する。
    金額が1万(万円)を超える場合は生値(円)混入とみなし÷10000で補正。
    パース可能な候補がゼロのときだけAIの値を維持する（従来動作へのフォールバック）。
    AIの選んだ値と食い違えばコードの選択が勝つ＝候補に無い値（捏造）も採用されない。
    """
    fixed: list[float] = []
    initial: list[float] = []
    for part in re.split(r"[;；]", str(candidates or "")):
        part = part.strip()
        if not part:
            continue
        kind_m = _PRICE_KIND_RE.search(part)
        if kind_m:
            kind = kind_m.group(1)
        elif re.search(r"円/|/imp|/クリック|/通|単価|imps", part, re.IGNORECASE):
            kind = "単"
        elif re.search(r"初期|最低利用|最低配信", part):
            kind = "初"
        elif "成果" in part or "CPL" in part.upper():
            kind = "成果"
        else:
            kind = "固"
        m_amt = re.search(r"=\s*¥?([0-9][0-9,.]*)", part)
        if not m_amt:
            continue
        try:
            amt = float(m_amt.group(1).replace(",", ""))
        except ValueError:
            continue
        while amt > 10000:  # 生値(円)混入 → 万円へ
            amt /= 10000
        if amt <= 0:
            continue
        if kind == "固":
            fixed.append(amt)
        elif kind == "初":
            initial.append(amt)
    pool = fixed or initial
    if not pool:
        return ai_value
    best = min(pool)
    best_s = str(int(best)) if float(best).is_integer() else str(round(best, 4))
    if ai_value not in ("", "-") and ai_value != best_s:
        logger.warning(f"{media_name}: 最低出稿金額をコード選択で補正 {ai_value} → {best_s}"
                       f"（候補: 固定{len(fixed)}件/初期{len(initial)}件）")
    return best_s


def _canon_menu(value: str) -> str:
    """主要メニューフォーマットを正規8種トークンに集約する（表記揺れヒント込み）"""
    sv = str(value)
    found = [t for t in MENU_TOKENS
             if t in sv or any(h in sv for h in _MENU_HINTS.get(t, ()))]
    return ";".join(found) if found else "該当なし"


def _cpl_to_range(value: str) -> str:
    """CPL単価（円の数値）を選択肢レンジへ変換する（下限以上・上限未満）

    既にレンジ表記ならそのまま。数値が100未満なら万円表記とみなし×10000。
    数値として読めなければ「該当なし」。
    """
    v = str(value).strip()
    if v in _CPL_RANGES or v == "該当なし":
        return v
    stripped = _NUM_UNIT_RE.sub("", v).replace(",", "").replace("@", "").replace("＠", "").replace("~", "").replace("〜", "")
    try:
        n = float(stripped)
    except ValueError:
        return "該当なし"
    if n <= 0:
        return "該当なし"
    if n < 100:      # 万円で返された場合（1.5 = 15,000円）
        n *= 10000
    if n < 5000:
        return "〜5,000円"
    if n < 10000:
        return "5,000〜10,000円"
    if n < 20000:
        return "10,000〜20,000円"
    return "20,000円〜"


# ============================================================
# 従業員規模「区分名=値;…」→ 3列(エンタープライズ/SMB/零細) のコード変換
# ============================================================


def _size_band_col(lo: float, hi: float) -> str:
    """人数帯 (lo, hi) を重なり幅が最大の規模カラム名へ割り当てる"""
    hi_capped = max(lo * 2, _SIZE_INF_CAP) if hi == float("inf") else min(hi, _SIZE_INF_CAP)
    best_i, best_ov = 0, -1.0
    for i, (b_lo, b_hi) in enumerate(_SIZE_BUCKETS):
        ov = max(0.0, min(hi_capped, b_hi) - max(lo, b_lo))
        if ov > best_ov:
            best_i, best_ov = i, ov
    return _SIZE_BUCKET_COLS[best_i]


def _apportion_company_size(value_text: str) -> list[float] | None:
    """従業員規模の「区分名=値;…」文字列を3列の%へ変換する

    各区分を人数レンジに解釈し、重なり幅が最大の1列に全額入れて合算する
    （比率按分はしない＝跨り区分は重なり幅の大きい方へ全額。2026-07-08決定）。
    人数レンジとして解釈できない区分（「不明」「上場企業」「個人事業主」等の非人数ラベル）は
    その区分だけスキップして残りを合算する（結果その軸は<100%になってよい＝ユーザー決定）。
    1件も解釈できなければ None（規模3列は空欄）。
    戻り値は [エンタープライズ, SMB, スタートアップ零細] の順の%リスト。
    """
    if not value_text or value_text == "-":
        return None
    totals = [0.0, 0.0, 0.0]
    parsed = False
    skipped: list[str] = []
    for part in str(value_text).split(";"):
        part = part.strip()
        if not part:
            continue
        name, sep, val_text = part.partition("=")
        rng = _parse_range(name) if sep else None
        try:
            value = float(val_text.replace("%", "").replace(",", "").strip()) if sep else None
        except ValueError:
            value = None
        if rng is None or value is None:
            skipped.append(part)  # 非人数ラベル（不明・上場企業 等）は捨てて続行
            continue
        totals[_SIZE_BUCKET_COLS.index(_size_band_col(rng[0], rng[1]))] += value
        parsed = True
    if skipped:
        logger.info(f"従業員規模: 人数でない区分をスキップ {skipped}")
    return [round(t, 1) for t in totals] if parsed else None


def _finalize_media(m: dict) -> dict[str, str]:
    """1媒体分の抽出結果を最終値dictへ清書する（欠損統一・年代変換・数値カラム検査）"""
    vals = m.get("values") or {}
    final: dict[str, str] = {}
    for c in CANONICAL_COLUMNS:
        raw = str(vals.get(c, "-")).strip()
        # 値全体が欠損記号（"–"・"_"・空 等）のときだけ "-" に正規化する。
        # 文字列の一部置換はしない（「PC USER_」→「PC USER-」のような名称破壊を防ぐ）
        final[c] = "-" if raw.translate(DASH_TRANS) in ("", "-") else raw

    # 数値カラム: 単位付き表記（「20万人以上」「28.5288万」「10万円」「285,288名」等）から
    # 数値部だけを取り出す。単位を付けて返すモデルがあり、素のfloat判定だと値ごと欠落するため
    # 機械的に除去する。除去後もfloatにできなければ（「該当なし」等）"-" にする。
    for c in NUMERIC_COLUMNS:
        v = final[c]
        if v != "-":
            stripped = _NUM_UNIT_RE.sub("", v).replace(",", "")
            try:
                float(stripped)
                final[c] = stripped
            except ValueError:
                final[c] = "-"

    # 万単位カラムの桁補正: LLMが÷10000を忘れて生値（PV=3000000等）を入れた場合を機械的に直す。
    # 実在の万単位値の上限を超えたら生値混入とみなし÷10000する。閾値は誤爆しないよう保守的に:
    #   PV/UU … 10億(=万単位10万)超はありえない（ねとらぼ4.1億=41000は温存）
    #   会員数/最低出稿金額/PV保証 … 1億(=万単位1万)超はありえない
    _RAW_CAP = {"月間PV数": 100000, "月間UU数": 100000,
                "会員数": 10000, "最低出稿金額": 10000, "PV保証_最小": 10000}
    for c, cap in _RAW_CAP.items():
        v = final.get(c, "-")
        if v != "-":
            try:
                num = float(v.replace(",", ""))
            except ValueError:
                continue
            while num > cap:  # 桁が合うまで÷10000（通常1回。生値混入を万単位へ）
                num /= 10000
                final[c] = str(int(num) if num.is_integer() else round(num, 4))
                logger.warning(f"{final.get('メディア名称', '?')}: {c} 生値混入 {v} → {final[c]} に桁補正")

    # %カラム: 0〜100の範囲外は "-"（桁ミスをDBに入れない）
    for c in PERCENT_COLUMNS:
        v = final[c]
        if v != "-":
            try:
                if not 0 <= float(v.replace(",", "")) <= 100:
                    logger.warning(f"{final.get('メディア名称', '?')}: {c}={v} が%範囲外 → 空欄化")
                    final[c] = "-"
            except ValueError:
                final[c] = "-"

    # メディアカテゴリ: 5択外（「ビジネスIT」等のハルシネーション）を機械的に5択へ寄せる
    if final.get("メディアカテゴリ", "-") not in _CATEGORY_CHOICES:
        cat = final.get("メディアカテゴリ", "")
        if "製造" in cat or "メーカー" in cat:
            final["メディアカテゴリ"] = "製造・メーカー"
        elif "経営" in cat or "マネジメント" in cat:
            final["メディアカテゴリ"] = "経営・マネジメント"
        elif "専門" in cat or "業界" in cat:
            final["メディアカテゴリ"] = "業界専門誌"
        else:
            final["メディアカテゴリ"] = "ビジネス総合"  # IT系・金融系・不明は総合に寄せる

    # D系（選択肢カラム）: 区切りを ";" に正規化し、"-" は「該当なし」に寄せる
    for c in D_COLUMNS:
        v = final[c]
        if v == "-":
            final[c] = "該当なし"
        else:
            final[c] = ";".join(t.strip() for t in re.split(r"[;,／/、]", v) if t.strip())

    # 主要メニューフォーマット: 正規8種に集約（規格外語・個別メニュー名をDBに残さない）
    final["主要メニューフォーマット"] = _canon_menu(final["主要メニューフォーマット"])

    # 対応ファネル: メニューから決定論で導出（AIの判断は使わない＝揺れゼロ）
    menu_set = set(final["主要メニューフォーマット"].split(";"))
    funnel = []
    if menu_set & {"ディスプレイ広告", "タイアップ", "メール広告"}:
        funnel += ["認知・ブランディング", "興味喚起・理解促進"]
    if menu_set & {"ホワイトペーパー", "セミナー", "製品掲載", "製品レビュー", "リードジェネレーション"}:
        funnel.append("比較検討・リード獲得")
    final["対応ファネル"] = ";".join(funnel) if funnel else "該当なし"

    # CPL目安: 円の数値をレンジへ変換（境界は下限以上・上限未満）
    final["CPL目安"] = _cpl_to_range(final["CPL目安"])

    # 最低出稿金額: price_candidates からコードが優先順位で確定する
    # （[固]最安→[初]最安。imp/クリック単価・成果報酬は除外。候補に無い値＝捏造は採用されない。
    # 選択したプランが変わった場合、PV保証_最小は元プラン基準のままの可能性がある点は許容）
    final["最低出稿金額"] = _select_min_price(
        m.get("price_candidates", ""), final.get("最低出稿金額", "-"),
        str(final.get("メディア名称", "?")))

    # 年代: 原文区分が正規5区分と違う場合は age_labels から変換
    if m.get("age_labels"):
        buckets, dropped = _normalize_ages(m["age_labels"])
        if any(v > 0 for v in buckets.values()):
            for k, v in buckets.items():
                final[f"年代_{k}"] = _fmt_pct(v) if v > 0 else "-"
            if dropped:
                logger.info(f"{final.get('メディア名称', '?')}: 年代の対象外区分 {dropped}")
    return final


def _unify_lead_wording(finals: dict[str, dict[str, str]]) -> tuple[str, list[str]]:
    """リードサービスの文言をコードで統一する（ハイブリッド方式の後半）

    「付ける/付けない」のAI判断は変えず、「付ける」と判断された主運営会社の媒体の
    文言（獲得リードの質・外部システム連携）を正規トークン化＋多数決で全員同じに揃える。
    コール1とバッチの境界を跨いだ統一もここで一括して行う。
    戻り値: (判定サマリ行, 媒体別noteのリスト)
    """
    notes: list[str] = []
    ops = Counter(v.get("メディア運営会社", "") for v in finals.values()
                  if v.get("メディア運営会社", "-") not in ("", "-"))
    main_op = ops.most_common(1)[0][0] if ops else ""

    # 正規トークン化（区切り・順序の揺れを先に吸収）
    for name, v in finals.items():
        v["獲得リードの質"] = _canon_choice(v.get("獲得リードの質", ""), LEAD_Q_TOKENS)
        v["外部システム連携"] = _canon_choice(v.get("外部システム連携", ""), LEAD_I_TOKENS)

    applied = [name for name, v in finals.items()
               if v.get("メディア運営会社") == main_op
               and (v["獲得リードの質"] != "該当なし" or v["外部システム連携"] != "該当なし")]
    mode_q = mode_i = ""
    if applied:
        q_counts = Counter(finals[n]["獲得リードの質"] for n in applied
                           if finals[n]["獲得リードの質"] != "該当なし")
        i_counts = Counter(finals[n]["外部システム連携"] for n in applied
                           if finals[n]["外部システム連携"] != "該当なし")
        mode_q = q_counts.most_common(1)[0][0] if q_counts else ""
        mode_i = i_counts.most_common(1)[0][0] if i_counts else ""
        for n in applied:
            changed = []
            if mode_q and finals[n]["獲得リードの質"] != mode_q:
                finals[n]["獲得リードの質"] = mode_q
                changed.append("質")
            if mode_i and finals[n]["外部システム連携"] != mode_i:
                finals[n]["外部システム連携"] = mode_i
                changed.append("連携")
            if changed:
                notes.append(f"{n}: リード{'・'.join(changed)}の文言を多数決で統一")

    # 別会社ガード: 主運営会社と異なる媒体に共通リードサービスは適用しない（コードで強制）。
    # プロンプトでも指示しているがLLMが従わない実例（発注ナビ）があったため決定的に倒す。
    # 複数媒体の資料でのみ発動（単一媒体資料は自社=主運営会社なので影響なし）
    if len(finals) > 1:
        for name, v in finals.items():
            if v.get("メディア運営会社") not in ("", "-", main_op) and (
                    v["獲得リードの質"] != "該当なし" or v["外部システム連携"] != "該当なし"):
                v["獲得リードの質"] = "該当なし"
                v["外部システム連携"] = "該当なし"
                notes.append(f"{name}: 別会社（{v.get('メディア運営会社')}）のため"
                             "共通リードサービス対象外 → 該当なしへ強制")

    n_none = sum(1 for v in finals.values()
                 if v["獲得リードの質"] == "該当なし" and v["外部システム連携"] == "該当なし")
    summary = (f"[リードサービス] 主運営会社={main_op or '不明'} / 適用{len(applied)}媒体"
               f"（統一文言: 質={mode_q or '-'} 連携={mode_i or '-'}） / 該当なし{n_none}媒体")
    return summary, notes

# ============================================================
# Step5: 50列行の組み立て
# ============================================================


def _parse_range(label: str) -> tuple[float, float] | None:
    """「18-24歳」「65歳以上」「10人未満」「1000-5000人」等を (下限, 上限) の半開区間に解釈する

    NFKC正規化してから判定する（全角チルダ～(U+FF5E)・全角数字を吸収）。波ダッシュ〜(U+301C)は
    NFKCで変換されないため正規表現の文字クラスに明示的に含める。
    """
    text = unicodedata.normalize("NFKC", str(label))
    text = re.sub(r"[歳才人名\s]", "", text.replace(",", ""))
    m = re.fullmatch(r"(\d+)[-〜~](\d+)", text)
    if m:
        return float(m.group(1)), float(m.group(2)) + 1  # 「18-24」は24を含むため+1
    m = re.fullmatch(r"(\d+)[-〜~](\d+)未満", text)
    if m:
        return float(m.group(1)), float(m.group(2))  # 「500-1000未満」は上限を含まない
    m = re.fullmatch(r"(\d+)以上", text)
    if m:
        return float(m.group(1)), float("inf")
    m = re.fullmatch(r"(\d+)(未満|以下)", text)
    if m:
        hi = float(m.group(1)) + (1 if m.group(2) == "以下" else 0)  # 「50以下」は50を含む
        return 1.0, hi
    m = re.fullmatch(r"(\d+)代", text)
    if m:
        return float(m.group(1)), float(m.group(1)) + 10
    return None


def _compact_pages(nums: list[int]) -> str:
    """ページ番号リストを「p2-p6;p28」のような圧縮表記にする"""
    nums = sorted(set(n for n in nums if isinstance(n, int)))
    if not nums:
        return ""
    parts: list[str] = []
    start = prev = nums[0]
    for n in nums[1:] + [None]:
        if n is not None and n == prev + 1:
            prev = n
            continue
        parts.append(f"p{start}" if start == prev else f"p{start}-p{prev}")
        if n is not None:
            start = prev = n
    return ";".join(parts)


def _pages_text(m: dict, common_text: str = "") -> str:
    """媒体のページ範囲＋共通ページ（属性・会社概要等の根拠）を参照ページ一覧のセル値にする"""
    ps, pe = m.get("page_start"), m.get("page_end")
    if not isinstance(ps, int):
        own = "-"
    elif isinstance(pe, int) and pe > ps:
        own = f"p{ps}-p{pe}"
    else:
        own = f"p{ps}"
    if common_text:
        return f"{own};共通{common_text}"
    return own


def _build_row(final: dict[str, str], pages_text: str, *, web_url: str, item_id: str,
               file_name: str) -> list:
    """1媒体分の最終値dictから50列の行を組み立てる（空欄は "-"）"""
    from datetime import datetime, timedelta, timezone

    def cell(col: str):
        value = final.get(col, "")
        if not value or value == "-":
            return "-"
        if col in NUMERIC_COLUMNS:
            try:
                number = float(value.replace(",", ""))
                return int(number) if number.is_integer() else number
            except ValueError:
                return "-"
        return value

    today_jst = datetime.now(timezone(timedelta(hours=9))).strftime("%Y-%m-%d")

    # 従業員規模の文字列（「1〜299人=39;…」）を3列へコード変換（重なり幅の大きい方へ全額）
    size_text = final.get("従業員規模", "")
    size_values = _apportion_company_size(size_text)
    if size_text not in ("", "-") and size_values is None:
        logger.warning(f"{final.get('メディア名称', '?')}: 従業員規模がパースできず変換不可: "
                       f"{str(size_text)[:80]}")

    row: list = []
    for header in PAYLOAD_HEADERS:
        if header == "メディアNo":
            row.append("")  # 空のまま（Excel側で採番想定）
        elif header in SIZE_COLUMNS:
            if size_values:
                n = size_values[SIZE_COLUMNS.index(header)]
                row.append(int(n) if float(n).is_integer() else n)
            else:
                row.append("-")
        elif header == "PDFファイルURL":
            row.append(web_url or "-")
        elif header == "データ更新日":
            row.append(today_jst)
        elif header == "参照ページ一覧":
            row.append(pages_text or "-")
        elif header == "item_id":
            row.append(item_id or "-")
        elif header == "file_name":
            row.append(file_name or "-")
        else:
            row.append(cell(header))
    if len(row) != EXPECTED_COLS:
        raise Exception(f"行の列数が{len(row)}列になりました（期待: {EXPECTED_COLS}列）")
    return row

# ============================================================
# Step6: Excel追記
# ============================================================


def _classify_excel_error(res: dict, error_text: str) -> str:
    """Excel系ツールのエラーを4分類する（mismatch / transient / auth / fatal）"""
    text = error_text.lower()
    if "column count mismatch" in text:
        return "mismatch"
    if res.get("auth_refresh_required") or any(
        p in text for p in ("unauthorized", "forbidden", "invalid authentication",
                            "token", "expired", "401", "403")
    ):
        return "auth"
    http_code = res.get("mercury_last_http_status_code")
    if isinstance(http_code, int) and (http_code == 429 or http_code >= 500):
        return "transient"
    if res.get("error_class"):  # Composioのサーバーエラー分類（メディア選定AIの実測仕様）
        return "transient"
    if any(
        p in text for p in ("429", "rate limit", "throttl", "timeout", "timed out",
                            "temporarily", "service unavailable", "internal server error",
                            "bad gateway", "502", "503", "504", "connection")
    ):
        return "transient"
    return "fatal"


async def _append_one_row(values_row: list) -> None:
    """EXCEL_ADD_TABLE_ROW でマスタExcelに1行追記する（4分類エラー処理・3回まで）"""
    from agent_sdk import ToolCallError

    last_error = ""
    kind = "fatal"
    for attempt in range(1, MAX_EXCEL_ATTEMPTS + 1):
        try:
            res = await _composio_call(
                "EXCEL_ADD_TABLE_ROW",
                {
                    "drive_id": DRIVE_ID,  # /me接続と別人のドライブのため必須
                    "item_id": EXCEL_ITEM_ID,
                    "table_id": EXCEL_TABLE_ID,
                    "values": [values_row],
                    # index は省略＝末尾に追記。worksheet というパラメータは存在しない
                },
            )
        except ToolCallError as e:
            res = {}
            last_error = f"ToolCallError: {e}"
            kind = "transient"
        else:
            if res.get("successful"):
                return
            last_error = str(res.get("error", "unknown"))
            kind = _classify_excel_error(res, last_error)
        logger.warning(f"Excel追記失敗 ({attempt}/{MAX_EXCEL_ATTEMPTS}) [{kind}]: {last_error[:300]}")
        # auth(401)も再試行対象にする: 並列書き込み競合が401で返る実例があり（2026-07-08）、
        # 本物の認証切れなら3回失敗して止まるだけ（無駄は+2試行のみ）
        if attempt == MAX_EXCEL_ATTEMPTS or kind == "fatal":
            break
        if kind in ("transient", "auth"):
            await asyncio.sleep(2 ** attempt)  # 2秒 → 4秒
    raise Exception(f"Excel追記失敗 [{kind}]: {last_error}")


# 未追記行のセッション内保持（code_execute は実行をまたいで変数が残るため、
# ファイルに書き出さなくても同一セッション内なら retry_excel_append で復旧できる）
_pending_rows: list[list] = []


async def _append_rows_to_excel(rows: list[list], *, ckpt: dict | None = None,
                                base_done: int = 0) -> None:
    """全媒体の行を EXCEL_BATCH 行ずつ直列で追記する

    失敗行があっても投入済みの行は追記され、未追記分（失敗行＋未投入行）を
    _pending_rows に残して例外を出す。ckpt を渡すと1行追記するたびに
    「追記済み行数」をチェックポイントへ保存する（再実行時はその続きから追記され、
    重複追記が起きない）。
    """
    global _pending_rows
    remaining = list(rows)
    done_total = 0
    while remaining:
        chunk = remaining[:EXCEL_BATCH]
        rest = remaining[len(chunk):]
        results = await asyncio.gather(
            *(_append_one_row(row) for row in chunk), return_exceptions=True,
        )
        failed = [(chunk[i], r) for i, r in enumerate(results) if isinstance(r, Exception)]
        done_total += len(chunk) - len(failed)
        # 未追記 = このバッチの失敗行 + まだ投げていない行（復旧・再送用にメモリ保持）
        remaining = [row for row, _ in failed] + rest
        _pending_rows = remaining
        if ckpt is not None:
            ckpt["appended"] = base_done + done_total
            ckpt["stage"] = f"Excel追記 {base_done + done_total}行"
            _save_ckpt(ckpt)
        print(f"[Excel] 追記 {base_done + done_total}/{base_done + len(rows)} 行完了")
        if failed:
            errors = " / ".join(str(e)[:150] for _, e in failed[:3])
            raise Exception(
                f"❌ Excel追記失敗（{len(failed)}行 / 全{base_done + len(rows)}行。"
                f"追記済み{base_done + done_total}行）: {errors}\n"
                "- 同じコードをもう1回実行すれば、チェックポイントから未追記分のみ追記されます\n"
                "  （抽出のやり直し・重複追記は起きません）。"
            )


async def retry_excel_append() -> dict:
    """Excel追記だけを再実行する（旧復旧手段。run_media_db_build の再実行が正になったため
    セッション内メモリに未追記行が残っている場合の予備として残す）"""
    if not _pending_rows:
        return {"summary": "未追記の行はありません（追記済み、または同一セッションでの実行歴なし）"}
    rows = list(_pending_rows)
    await _append_rows_to_excel(rows)
    _clear_ckpt()  # 全行追記済み＝再実行での重複追記を防ぐ
    summary = f"✅ Excel追記の再実行完了（{len(rows)}行）"
    print(summary)
    return {"summary": summary}

# ============================================================
# チェックポイント（タイムアウト・失敗からの途中再開用）
# ============================================================
# LLMの産出物（コール1・バッチ結果）とExcel追記済み行数だけを保存する。
# 清書・行組み立ては決定論のコードなので保存せず、再開時に再計算する。
# 成功時に削除。tmp/ は同一セッション内でのみ持続する前提（セッションを跨いだ再開は保証しない）。

CKPT_PATH = "tmp/media_db_ckpt.json"       # 直近実行のチェックポイント（item_idで照合）
SUMMARY_PATH = "tmp/media_db_summary.txt"  # stdoutが消えた場合の復旧用summary


def _load_ckpt(item_id: str, file_name: str) -> dict:
    """同一ファイルの実行途中チェックポイントがあれば読み込む（無ければ空dict）"""
    try:
        with open(CKPT_PATH, encoding="utf-8") as f:
            ck = json.load(f)
        if ck.get("item_id") == item_id and ck.get("file_name") == file_name:
            return ck
    except Exception:
        pass
    return {}


def _save_ckpt(ck: dict) -> None:
    """チェックポイントを保存する（保存の失敗で本処理は止めない）"""
    try:
        os.makedirs("tmp", exist_ok=True)
        with open(CKPT_PATH, "w", encoding="utf-8") as f:
            json.dump(ck, f, ensure_ascii=False)
    except Exception as e:
        logger.warning(f"チェックポイント保存に失敗（処理は続行）: {e}")


def _clear_ckpt() -> None:
    """チェックポイントを削除する（成功時・再実行不要になった時）"""
    try:
        os.remove(CKPT_PATH)
    except OSError:
        pass

# ============================================================
# 公開関数（フルフロー）
# ============================================================


async def run_media_db_build(item_id: str, file_name: str, web_url: str) -> dict:
    """メディアDB構築のフルフロー（v4・統合コール＋抜粋バッチ方式）

    PDF取得 → 統合コール（媒体数判定＋先頭THRESHOLD媒体の抽出＋ページ構成）
    → 16媒体以上なら抜粋バッチで残りを並列抽出 → コード清書（欠損統一・年代変換・
    リード文言統一）→ 50列行の組み立て → Excel追記（最大16行並列）。
    実行中は hunter の関数トレースを TRACE_LOG_PATH に記録する。

    Args:
        item_id: トリガーJSONの payload.item.item_id
        file_name: トリガーJSONの payload.item.name
        web_url: トリガーJSONの payload.item.web_url（PDFファイルURLとして各行に記録）

    Returns:
        完了報告 dict（summary / media_names / failed / rows_count）
    """
    trace_stream = _start_trace()
    try:
        return await _run_build(item_id, file_name, web_url)
    finally:
        _stop_trace(trace_stream)


async def _run_build(item_id: str, file_name: str, web_url: str) -> dict:
    """フルフローの本体（run_media_db_build がトレース付きで呼ぶ）"""
    t_start = time.time()

    # 途中再開: 同一ファイルのチェックポイントがあれば済んだ段階をスキップする
    ck = _load_ckpt(item_id, file_name)
    if ck:
        print(f"[Resume] チェックポイント検出（{ck.get('stage', '?')}まで完了済み）→ 続きから再開")
    else:
        ck = {"item_id": item_id, "file_name": file_name}

    # Step1: PDF取得（再開時も取り直す＝一時ファイルが消えていても動く）
    _ensure_pypdf()  # ページ数取得・抜粋PDF生成に必須。無い環境ではsubprocess pipで導入
    await _download_pdf(item_id, file_name)
    from pypdf import PdfReader
    total_pages = len(PdfReader(PDF_LOCAL_PATH).pages)

    # Step2: 統合コール（媒体数判定＋先頭THRESHOLD媒体の抽出＋残りページリスト＋共通ページ）
    if ck.get("call1"):
        c1 = ck["call1"]
        print("[Step2] チェックポイントから再利用（コール1スキップ）")
    else:
        c1 = await _call1()
        ck["call1"] = c1
        ck["stage"] = "コール1完了"
        _save_ckpt(ck)
    media_rows: list[dict] = list(c1["extracted"])
    failed: list[str] = []

    # Step3: 残り媒体の抜粋バッチ抽出（remaining の実物が空かどうかで分岐）
    if c1["remaining"]:
        if "batch_rows" in ck:
            batch_rows, batch_failed = ck["batch_rows"], ck.get("batch_failed", [])
            print("[Step3] チェックポイントから再利用（バッチ抽出スキップ）")
        else:
            batch_rows, batch_failed = await _extract_batches(
                c1["remaining"], c1["common_pages"], total_pages)
            ck["batch_rows"] = batch_rows
            ck["batch_failed"] = batch_failed
            ck["stage"] = "バッチ抽出完了"
            _save_ckpt(ck)
        media_rows += batch_rows
        failed += batch_failed

    if not media_rows:
        raise Exception("全媒体の抽出に失敗しました")

    # 同名媒体の重複排除（コール1とバッチの重複・LLMの二重出力対策。先勝ち）
    _seen: set[str] = set()
    _deduped: list[dict] = []
    for m in media_rows:
        key = _norm_name(m.get("name"))
        if key in _seen:
            logger.warning(f"重複媒体をスキップ: {m.get('name')}")
            continue
        _seen.add(key)
        _deduped.append(m)
    media_rows = _deduped

    # Step4: コード清書（媒体ごとの正規化 → リード文言の全体統一）
    finals: dict[str, dict[str, str]] = {}
    for m in media_rows:
        final = _finalize_media(m)
        if not final.get("メディア名称") or final["メディア名称"] == "-":
            final["メディア名称"] = str(m.get("name", "-"))
        finals[str(m.get("name"))] = final
    lead_summary, lead_notes = _unify_lead_wording(finals)
    print(lead_summary)

    # Step5: 50列行の組み立て（参照ページには媒体のページ＋共通ページを併記）
    common_text = _compact_pages([c.get("page") for c in c1["common_pages"]])
    rows: list[list] = []
    for m in media_rows:
        name = str(m.get("name"))
        if name in finals:
            rows.append(_build_row(finals[name], _pages_text(m, common_text), web_url=web_url,
                                   item_id=item_id, file_name=file_name))

    # Step6: Excel追記（再開時はチェックポイントの追記済み行数の続きから）
    if WRITE_TO_EXCEL:
        row_names = [str(m.get("name")) for m in media_rows if str(m.get("name")) in finals]
        appended = int(ck.get("appended", 0))
        if appended and ck.get("row_names") != row_names:
            raise Exception(
                "チェックポイントの行構成が現在の抽出結果と一致しません。"
                f"tmp/{CKPT_PATH.split('/')[-1]} を削除してから再実行してください"
                "（そのまま続けると重複・欠落の恐れがあるため停止）"
            )
        if appended:
            print(f"[Excel] 追記済み{appended}行をスキップして続きから追記")
        ck["row_names"] = row_names
        _save_ckpt(ck)
        await _append_rows_to_excel(rows[appended:], ckpt=ck, base_done=appended)

    # 完了報告（媒体ごとの主要値1行）
    media_lines: list[str] = []
    for name, final in finals.items():
        filled = sum(1 for c in CANONICAL_COLUMNS if final.get(c, "-") not in ("", "-", "該当なし"))
        parts = [f"PV={final.get('月間PV数', '-')}", f"UU={final.get('月間UU数', '-')}",
                 f"会員={final.get('会員数', '-')}", f"最低出稿={final.get('最低出稿金額', '-')}"]
        media_lines.append(f"  - {name}（実値{filled}件 | {' '.join(parts)}）")
    write_line = (f"Excelに{len(rows)}行追記済み" if WRITE_TO_EXCEL
                  else "⚠️ テストモードのためExcel未書き込み")
    elapsed_total = time.time() - t_start
    summary = (
        f"✅ メディアDB構築v4 完了（{len(rows)}媒体 / {write_line} / 全体{elapsed_total:.0f}秒）\n"
        f"- 入力PDF: {file_name}（{total_pages}ページ / 総媒体数申告={c1['media_count']}）\n"
        f"- {lead_summary}\n"
        "- 媒体別:\n" + "\n".join(media_lines) + "\n"
        + (f"- ❌ 抽出失敗媒体: {' / '.join(failed)}\n" if failed else "")
        + ("".join(f"- ※{n}\n" for n in lead_notes) if lead_notes else "")
        + f"- 実行トレース: {TRACE_LOG_PATH}\n"
        + f"- スクリプト版数: {SCRIPT_VERSION}"
    )
    # 成功: チェックポイントを消し、summaryをファイルにも残す（stdout消失時の復旧用。
    # 万一summaryが表示されなくても file_read で取り出せる＝ログからの再構築を不要にする）
    _clear_ckpt()
    try:
        os.makedirs("tmp", exist_ok=True)
        with open(SUMMARY_PATH, "w", encoding="utf-8") as f:
            f.write(summary)
    except Exception as e:
        logger.warning(f"summaryのファイル保存に失敗（処理結果には影響なし）: {e}")
    print(summary)
    return {
        "summary": summary,
        "media_names": list(finals),
        "failed": failed,
        "rows_count": len(rows),
    }
