"""
カレンダー登録エージェント (実行環境版)

入力: ユーザーが自由記述した予定テキスト（複数可）+ 今日の日付
処理: LLMで予定を構造化 → GOOGLECALENDAR_CREATE_EVENT で一括登録
出力: 登録結果サマリー
"""

import asyncio
import json
import logging
from datetime import datetime, timedelta

from agent_sdk import GOOGLECALENDAR_CREATE_EVENT, llm_call

logger = logging.getLogger(__name__)

# --- 設定 ---
DEFAULT_TIMEZONE = "Asia/Tokyo"


# ============================================================
# LLMによる予定抽出
# ============================================================
_EXTRACT_PROMPT = """以下のテキストに含まれるカレンダー予定をすべて抽出してください。

【今日の日付】{today}（JST / Asia/Tokyo）

【変換ルール】
- start_datetime は必ず YYYY-MM-DDTHH:MM:SS 形式の絶対日時に変換する
- 「明日」「来週月曜」などは今日の日付を基準に絶対日付へ変換する
- end_datetime は必ず設定すること。終了時刻が明示されていない場合は start_datetime + 1時間 で算出する
- attendees にはメールアドレス形式（@を含む）のものだけ含める
- 時刻が完全に不明（「いつか」「未定」など）で絶対日時に変換できない場合は skippable=true

【テキスト】
{text}"""

_EXTRACT_SCHEMA = {
    "type": "object",
    "required": ["events"],
    "properties": {
        "events": {
            "type": "array",
            "items": {
                "type": "object",
                "required": ["summary", "start_datetime", "end_datetime", "skippable"],
                "properties": {
                    "summary":        {"type": "string", "description": "予定のタイトル"},
                    "start_datetime": {"type": "string", "description": "開始日時 YYYY-MM-DDTHH:MM:SS"},
                    "end_datetime":   {"type": "string", "description": "終了日時 YYYY-MM-DDTHH:MM:SS。不明な場合は start + 1時間"},
                    "description":    {"type": "string", "description": "詳細・メモ"},
                    "location":       {"type": "string", "description": "場所"},
                    "attendees": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "参加者のメールアドレスリスト",
                    },
                    "skippable": {
                        "type": "boolean",
                        "description": "日時が特定できず登録不可な場合 true",
                    },
                },
            },
        }
    },
}


async def _extract_events(text: str, today: str) -> list[dict]:
    """テキストから予定リストをLLMで抽出する。"""
    prompt = _EXTRACT_PROMPT.format(text=text, today=today)
    res = await llm_call(prompt=prompt, schema=_EXTRACT_SCHEMA)
    events: list[dict] = res["data"]["events"]
    logger.info(f"LLMが {len(events)} 件の予定を抽出しました")
    return events


# ============================================================
# カレンダー登録
# ============================================================
def _resolve_end_datetime(start: str, end: str) -> str:
    """end が空・不正な場合は start + 1時間で補完する。"""
    if end and end.strip():
        return end.strip()
    dt = datetime.strptime(start, "%Y-%m-%dT%H:%M:%S")
    return (dt + timedelta(hours=1)).strftime("%Y-%m-%dT%H:%M:%S")


def _build_params(event: dict) -> dict:
    """抽出済みイベントdict をAPIパラメータに変換する。"""
    params: dict = {
        "summary":        event["summary"],
        "start_datetime": event["start_datetime"],
        "end_datetime":   _resolve_end_datetime(event["start_datetime"], event.get("end_datetime", "")),
        "timezone":       DEFAULT_TIMEZONE,
    }
    for key in ("description", "location"):
        if event.get(key):
            params[key] = event[key]
    if event.get("attendees"):
        params["attendees"] = event["attendees"]
    return params


async def _register_one(event: dict, idx: int) -> dict:
    """単一イベントをカレンダーに登録し、結果dictを返す。"""
    summary = event.get("summary", f"予定{idx + 1}")

    # 日時が特定できない予定はスキップ
    if event.get("skippable"):
        logger.info(f"[{idx + 1}] スキップ（日時不明）: {summary}")
        return {"success": False, "skipped": True, "summary": summary, "error": "日時が特定できませんでした"}

    params = _build_params(event)
    try:
        res = await GOOGLECALENDAR_CREATE_EVENT(**params)
        if res.get("successful"):
            url = res.get("data", {}).get("display_url", "")
            logger.info(f"[{idx + 1}] 登録成功: {summary}")
            return {"success": True, "skipped": False, "summary": summary, "url": url, "params": params}
        error = res.get("error") or "不明なエラー"
        logger.warning(f"[{idx + 1}] 登録失敗: {summary} – {error}")
        return {"success": False, "skipped": False, "summary": summary, "error": error, "params": params}
    except Exception as e:
        logger.warning(f"[{idx + 1}] 例外: {summary} – {e}")
        return {"success": False, "skipped": False, "summary": summary, "error": str(e), "params": params}


# ============================================================
# メインフロー
# ============================================================
async def process_message(text: str, today: str) -> dict:
    """
    ユーザーが送った自然言語テキストから予定を一括登録する。

    Args:
        text:  予定を記述した自然言語テキスト（複数可）
        today: 今日の日付（YYYY-MM-DD形式）

    Returns:
        {
            "results":  list[dict],  # 各予定の登録結果
            "summary":  str,         # ユーザーへの返答文
        }
    """
    events = await _extract_events(text, today)
    if not events:
        return {"results": [], "summary": "予定が見つかりませんでした。もう少し具体的に書いてみてください。"}

    logger.info(f"{len(events)} 件を並列登録します...")
    tasks = [_register_one(ev, i) for i, ev in enumerate(events)]
    results: list[dict] = list(await asyncio.gather(*tasks))

    # --- 返答文の組み立て ---
    lines = []
    succeeded = [r for r in results if r["success"]]
    failed    = [r for r in results if not r["success"] and not r.get("skipped")]
    skipped   = [r for r in results if r.get("skipped")]

    if succeeded:
        lines.append(f"✅ {len(succeeded)} 件登録しました")
        for r in succeeded:
            lines.append(f"  ・{r['summary']}　{r.get('url', '')}")
    if skipped:
        lines.append(f"⚠️ {len(skipped)} 件は日時が特定できず登録できませんでした")
        for r in skipped:
            lines.append(f"  ・{r['summary']}")
    if failed:
        lines.append(f"❌ {len(failed)} 件でエラーが発生しました")
        for r in failed:
            lines.append(f"  ・{r['summary']}：{r.get('error', '')}")

    return {"results": results, "summary": "\n".join(lines)}


# ============================================================
# 動作確認サンプル
# ============================================================
async def main():
    sample_text = """
    明日14時から15時 田中さんと進捗MTG（場所：会議室B）
    6/20 10:00〜11:30 週次定例 参加者: tanaka@example.com, yamada@example.com
    来週月曜 15時 歯医者
    """
    today = "2026-06-16"
    result = await process_message(sample_text, today)
    print(result["summary"])
    print("\n--- 詳細 ---")
    print(json.dumps(result["results"], ensure_ascii=False, indent=2))


if __name__ == "__main__":
    asyncio.run(main())
