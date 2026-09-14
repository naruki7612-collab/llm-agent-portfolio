import asyncio
import os
import re
import json
from agent_sdk import (
    OUTLOOK_GET_MESSAGE,
    OUTLOOK_GET_MASTER_CATEGORIES,
    OUTLOOK_BATCH_UPDATE_MESSAGES,
    OUTLOOK_CREATE_USER_MASTER_CATEGORY,
    OUTLOOK_LIST_OUTLOOK_ATTACHMENTS,
    OUTLOOK_DOWNLOAD_OUTLOOK_ATTACHMENT,
    llm_call,
)

DEFAULT_COLOR_CATEGORIES = {
    "赤の分類", "青の分類", "緑の分類", "黄色の分類", "オレンジの分類", "紫の分類",
    "Red category", "Blue category", "Green category",
    "Yellow category", "Orange category", "Purple category",
}

DEFAULT_LABELS = [
    {"displayName": "要返信",         "color": "preset0"},
    {"displayName": "通知",           "color": "preset1"},
    {"displayName": "マーケティング", "color": "preset2"},
    {"displayName": "返信待ち",       "color": "preset3"},
    {"displayName": "FYI",            "color": "preset4"},
]

async def _ensure_default_categories():
    """必要なカテゴリが存在しない場合、作成する。"""
    master = await OUTLOOK_GET_MASTER_CATEGORIES()
    existing = [c.get("displayName") for c in master.get("data", {}).get("value", [])]
    
    for label in DEFAULT_LABELS:
        if label["displayName"] not in existing:
            try:
                await OUTLOOK_CREATE_USER_MASTER_CATEGORY(
                    user_id="me", display_name=label["displayName"], color=label["color"]
                )
            except Exception:
                pass

def _extract_body_text(message: dict) -> str:
    message_data = message.get("data", message)
    body = message_data.get("body", "")
    if isinstance(body, dict):
        content = body.get("content", "")
        if body.get("contentType", "").lower() == "html":
            content = re.sub(r"<[^>]+>", " ", content)
            content = re.sub(r"\s+", " ", content).strip()
        return content
    return str(body)

async def _judge_with_llm(message: dict, candidate_names: list[str]) -> dict:
    message_data = message.get("data", message)
    subject = message_data.get("subject", "")
    body = _extract_body_text(message)
    sender = message_data.get("from", {})
    sender_str = sender.get("emailAddress", {}).get("address", "不明") if isinstance(sender, dict) else str(sender)
    candidates_str = "\n".join(f"- {n}" for n in candidate_names)

    prompt = f"""以下のメールをもとに判定してください。
- 差出人: {sender_str}
- 件名: {subject}
- 本文: {body}

---
判定1: 候補カテゴリから最も合致するものを一つ選択（どれにも合致しない場合は「なし」）
候補: {candidates_str}

判定2: メール本文に返答・確認・承認・回答・日程調整などを求める記述がある場合は「要返信」、ない場合は「返信不要」と判定
"""
    schema = {
        "type": "object",
        "properties": {
            "category": {"type": "string", "enum": candidate_names + ["なし"]},
            "reply": {"type": "string", "enum": ["要返信", "返信不要"]},
        },
        "required": ["category", "reply"],
    }
    result = await llm_call(prompt=prompt, schema=schema)
    data = result.get("data", {}) if isinstance(result, dict) else {}
    return {"category": data.get("category", "なし"), "reply": data.get("reply", "返信不要")}

async def _get_attachment_urls(message_id: str) -> list[dict]:
    """添付ファイルの一覧を取得し、各ファイルの S3URL を返す。"""
    print("\n[Step 5] 添付ファイル S3URL を取得中...")

    try:
        result = await OUTLOOK_LIST_OUTLOOK_ATTACHMENTS(message_id=message_id)
        attachments = result.get("data", {}).get("value", [])
    except Exception as e:
        print(f"  ❌ 添付ファイル一覧取得失敗: {e}")
        return []

    if not attachments:
        print("  → 添付ファイルなし")
        return []

    print(f"  ✅ {len(attachments)}個の添付ファイルを検出")

    attachment_urls = []
    for attachment in attachments:
        attachment_id = attachment.get("id")
        filename = attachment.get("name", "file")
        ext = filename.split(".")[-1] if "." in filename else "bin"

        print(f"  - {filename}")
        s3_url = None

        try:
            dl_result = await OUTLOOK_DOWNLOAD_OUTLOOK_ATTACHMENT(
                message_id=message_id,
                attachment_id=attachment_id,
                file_name=filename
            )
            if isinstance(dl_result, dict):
                if "data" in dl_result and "file" in dl_result.get("data", {}):
                    s3_url = dl_result["data"]["file"].get("s3url")
                elif "data" in dl_result:
                    s3_url = dl_result["data"].get("s3url")
                else:
                    s3_url = dl_result.get("s3url")

            print(f"      {'✅ S3URL取得完了' if s3_url else '⚠️  S3URL取得失敗'}")
        except Exception as e:
            print(f"      ❌ エラー: {e}")

        attachment_urls.append({
            "filename": filename,
            "s3_url": s3_url,
            "ext": ext,
            "attachment_id": attachment_id,
            "message_id": message_id
        })

    return attachment_urls


async def categorize_message(message_id: str) -> str:
    """
    メール分類処理
    
    1. カテゴリの整合性確保
    2. メッセージ取得
    3. マスターカテゴリ取得
    4. LLM判定（カテゴリ & 返信要否）
    5. メッセージにカテゴリを適用
    6. 返信が必要な場合、mail_info.json と attachment_urls.json を出力
    """
    
    # Step 1: カテゴリの整合性確保
    await _ensure_default_categories()

    # Step 2: メッセージ取得
    print(f"[Step 1] メッセージ取得中...")
    message = await OUTLOOK_GET_MESSAGE(message_id=message_id)
    message_data = message.get("data", message)
    subject = message_data.get("subject", "")
    print(f"  ✅ 取得完了: {subject}")
    
    # Step 3: マスターカテゴリ取得
    print(f"\n[Step 2] カテゴリ取得中...")
    master = await OUTLOOK_GET_MASTER_CATEGORIES()
    categories: list[dict] = master.get("data", {}).get("value", [])
    available_names = [c.get("displayName", "") for c in categories if c.get("displayName")]
    candidate_names = [n for n in available_names if n not in DEFAULT_COLOR_CATEGORIES]
    
    if not candidate_names:
        print(f"  → 候補カテゴリなし: 更新しない")
        return "更新しない"
    
    print(f"  ✅ {len(candidate_names)}個の候補を取得: {candidate_names}")

    # Step 4: LLM判定
    print(f"\n[Step 3] LLM判定中...")
    judgment = await _judge_with_llm(message, candidate_names)
    selected = judgment["category"]
    needs_reply = judgment["reply"] == "要返信"
    print(f"  ✅ カテゴリ: {selected}, 返信: {judgment['reply']}")

    # 候補外 or 「なし」は更新しない
    if selected == "なし" or selected not in candidate_names:
        print(f"  → 「なし」: 更新しない")
        return "更新しない"

    # Step 5: カテゴリ適用
    print(f"\n[Step 4] カテゴリ適用中...")
    await OUTLOOK_BATCH_UPDATE_MESSAGES(
        updates=[{"message_id": message_id, "patch": {"categories": [selected]}}]
    )
    print(f"  ✅ 「{selected}」を適用しました")

    # Step 6: 返信が必要な場合、メール情報と添付ファイル S3URL を出力
    if needs_reply:
        print(f"\n[Step 5] 返信ドラフト作成の準備中...")

        os.makedirs("tmp", exist_ok=True)

        # メール情報を JSON で保存
        mail_info = {
            "message_id": message_id,
            "subject": message_data.get("subject", ""),
            "from": message_data.get("from", {}),
            "to_recipients": message_data.get("toRecipients", []),
            "cc_recipients": message_data.get("ccRecipients", []),
            "body": _extract_body_text(message),
            "received_datetime": message_data.get("receivedDateTime", ""),
        }

        with open("tmp/mail_info.json", "w", encoding="utf-8") as f:
            json.dump(mail_info, f, ensure_ascii=False, indent=2)

        print(f"  ✅ メール情報を保存: tmp/mail_info.json")

        # 添付ファイル S3URL を取得して保存
        attachment_urls = await _get_attachment_urls(message_id)

        with open("tmp/attachment_urls.json", "w", encoding="utf-8") as f:
            json.dump({
                "attachments": attachment_urls,
                "count": len(attachment_urls),
                "status": "completed" if attachment_urls else "no_attachments"
            }, f, ensure_ascii=False, indent=2)

        print(f"  ✅ S3URL を保存: tmp/attachment_urls.json")
        print(f"\n⚠️  実行環境環境を抜けて以下を実行してください:")
        print(f"    python step2_extract_attachments.py")
        print(f"\n実行後、ドラフト作成を再開します。")

        return "分類情報の更新が完了していて返信が必要"
    
    return "分類情報の更新が完了"

if __name__ == "__main__":
    import asyncio
    async def run_standalone():
        message_id = input("message_id を入力してください: ").strip()
        if message_id:
            result = await categorize_message(message_id)
            print(result)
    asyncio.run(run_standalone())
