import json
import os
from agent_sdk import (
    OUTLOOK_CREATE_DRAFT_REPLY,
    llm_call,
)

async def main():
    print(f"=== メール返信ドラフト作成 ===\n")
    
    # ファイルの確認
    if not os.path.exists("tmp/mail_info.json"):
        print("❌ tmp/mail_info.json が見つかりません")
        return
    
    if not os.path.exists("tmp/attachment_contents.json"):
        print("❌ tmp/attachment_contents.json が見つかりません")
        print("先に prepare_attachments.py を実行してください")
        return
    
    # メール情報を読み込む
    print("[Step 1] メール情報を読み込み中...")
    with open("tmp/mail_info.json", "r", encoding="utf-8") as f:
        mail_info = json.load(f)
    
    message_id = mail_info.get("message_id")
    subject = mail_info.get("subject", "")
    body = mail_info.get("body", "")
    from_info = mail_info.get("from", {})
    sender_name = from_info.get("emailAddress", {}).get("name", "不明") if isinstance(from_info, dict) else "不明"
    cc_recipients = mail_info.get("cc_recipients", [])
    
    print(f"  ✅ メール情報取得: 件名 = {subject}")
    
    # 添付ファイル内容を読み込む
    print("\n[Step 2] 添付ファイル内容を読み込み中...")
    with open("tmp/attachment_contents.json", "r", encoding="utf-8") as f:
        attachment_data = json.load(f)
    
    attachments = attachment_data.get("attachments", [])
    print(f"  ✅ {len(attachments)}個のファイルを読み込み")
    
    # 添付ファイルの内容をまとめる
    attachment_context = ""
    if attachments:
        attachment_context = "\n\n【添付ファイル内容】\n"
        for item in attachments:
            filename = item.get("filename", "不明")
            content = item.get("content", "")
            attachment_context += f"ファイル『{filename}』:\n{content}\n\n"
    
    # 返信ドラフト作成プロンプト
    print("\n[Step 3] LLM でドラフト生成中...")
    draft_prompt = f"""メール返信ドラフトを作成してください。

【元のメール情報】
- 差出人: {sender_name}
- 件名: {subject}
- 本文:
{body}
{attachment_context}

【ドラフト作成要件】
- ビジネスメールとして適切な敬語・丁寧語を使用
- 挨拶・用件・締めの3段落に分ける
- 原文の意図と異なる内容・約束・確約は含めない
- 段落は改行で分ける

以下のような形式を参考にしてください：

お世話になっております。

××の件、ご連絡いただきありがとうございます。
〇月〇日には確実にお届けいただけるとのこと、承知いたしました。
早急なご対応、誠にありがとうございます。

よろしくお願いいたします。

【返信ドラフト本文】
"""

    schema = {
        "type": "object",
        "properties": {
            "draft": {
                "type": "string",
                "description": "返信ドラフトの本文"
            }
        },
        "required": ["draft"]
    }
    
    result = await llm_call(prompt=draft_prompt, schema=schema)
    draft_body = result.get("data", {}).get("draft", "") if isinstance(result, dict) else ""
    
    print(f"  ✅ ドラフト生成完了")
    print(f"\n【生成されたドラフト】")
    print(f"{draft_body[:500]}...")
    
    # Outlook に下書きを保存
    print(f"\n[Step 4] Outlook に保存中...")
    cc_addresses = [cc.get("emailAddress", {}).get("address") for cc in cc_recipients if isinstance(cc, dict)]
    
    try:
        await OUTLOOK_CREATE_DRAFT_REPLY(
            message_id=message_id,
            comment=draft_body,
            cc_recipients=cc_addresses if cc_addresses else []
        )
        print(f"  ✅ 下書きを保存しました（CC: {len(cc_addresses)}件）")
        print(f"\n✅ 完了")
        
        # 後処理: JSON ファイルを削除
        try:
            os.remove("tmp/mail_info.json")
            os.remove("tmp/attachment_contents.json")
        except:
            pass
        
    except Exception as e:
        print(f"  ❌ 保存失敗: {e}")

if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
