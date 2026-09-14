"""slack_recorder_notion_v5 — Slack投稿を統合Notion DBに記録する（1DB移行版）

v4 からの変更点:
- 保存先を統合DB（<RAW_DB_ID>）に変更。
  プロパティは subject / parents / timestamp / sender / name / content / project /
  source / links / to / cc-bcc の11個に統一（slack は to / cc-bcc / subject 未使用）。
- 添付ファイルは共有ドライブ「project」内の「slack」フォルダに保存し、
  1ファイルにつき source=drive のファイル行を1行 upsert する
  （添付付きメッセージはメッセージ行＋ファイル行の計2行以上になる）。
- 保存前に共有ドライブ「project」全体（「slack」フォルダに限らない）を対象に
  同名ファイルの重複チェック（GOOGLEDRIVE_FIND_FILE）。
  同名かつ同サイズ → 保存スキップ・既存リンクを links に記録。
  同名かつサイズ不一致（更新版）→ ファイル名に日時を付けて別名保存。
- メッセージ内の Google Drive URL も links に記録しつつ、ファイル行（要約あり）を
  upsert する。google_drive_recorder_v4 の定期スキャンは createdTime が直近
  LOOKBACK_HOURS 以内のファイルしか対象にせず、以前から存在するファイルが
  共有された場合は永久に拾われないため、v4 に要約を任せず v5 側でその場で
  ダウンロード・要約する（2026-07-14 ユーザー確認：URL/添付ファイル行の
  content には最初から要約を入れる想定だった）。
- Miro / Canva / Notion の URL は従来どおり要約を生成して links に入れる。
- 添付ファイル・メッセージ内Drive URLどちらも、markitdown/pypdf/fitz による
  テキスト抽出・要約を行い、ファイル行の content に要約JSONを入れる
  （2026-07-12 ユーザー確認、fitz込みでフル実装）。
"""

import json
import asyncio
import re
import os
import unicodedata
import urllib.parse
from datetime import datetime, timezone, timedelta

import pypdf
from markitdown import MarkItDown
from agent_sdk import (
    slack_list_all_users as SLACK_LIST_ALL_USERS,
    slack_list_conversations as SLACK_LIST_CONVERSATIONS,
    SLACK_DOWNLOAD_SLACK_FILE,
    NOTION_UPSERT_ROW_DATABASE,
    NOTION_FETCH_ALL_BLOCK_CONTENTS,
    GOOGLEDRIVE_GET_FILE_METADATA,
    GOOGLEDRIVE_FIND_FILE,
    GOOGLEDRIVE_UPLOAD_FILE,
    GOOGLEDRIVE_DOWNLOAD_FILE,
    MIRO_GET_BOARD_ITEMS,
    CANVA_LIST_DESIGN_PAGES_WITH_PAGINATION,
    llm_call,
)

# 統合DB（raw 4DB を1本化したもの）
DB_ID = "<RAW_DB_ID>"

# 共有ドライブ「project」内の「slack」フォルダ（添付ファイルの保存先）
SLACK_DRIVE_FOLDER_ID = "<DRIVE_FOLDER_ID_2>"

# 共有ドライブ「project」自体のID（重複チェックをドライブ全体に対して行うため）。
# アップロードAPIのレスポンスに含まれる driveId/teamDriveId と一致することを確認済み。
PROJECT_SHARED_DRIVE_ID = "0AINJmYrIeBf9Uk9PVA"

# Slack形式 <URL|text> とプレーンURLの両方にマッチ
_URL_RE = re.compile(r'<(https?://[^|>\s]+)(?:\|[^>]*)?>|(?<![<|])(https?://[^\s<>|]+)')
_GDRIVE_FILE_ID_RE = re.compile(r'/d/([a-zA-Z0-9_-]+)|[?&]id=([a-zA-Z0-9_-]+)')
_NOTION_UUID_RE = re.compile(
    r'([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})'
    r'|([a-f0-9]{32})'
)

# llm_callに渡せる最大ファイル数（Canva/添付ファイル画像化で使用）
LLM_MAX_FILES = 20

# LLMに渡すテキストの最大文字数
MAX_TEXT_LENGTH = 8000

# markitdown が失敗しても素読み（UTF-8）でフォールバックできるテキスト系形式
_PLAIN_TEXT_EXTS = frozenset({
    ".html", ".htm", ".txt", ".md", ".csv", ".tsv", ".json", ".xml", ".log", ".yaml", ".yml",
})

# fitz では開けず、markitdown でテキスト化する形式（Office + html/テキスト系）
_MARKITDOWN_EXTS = frozenset({".docx", ".xlsx", ".pptx"}) | _PLAIN_TEXT_EXTS

# stdout に出す要約プレビューの文字数
SUMMARY_PREVIEW_CHARS = 200

# Notion rich_text / title の1テキスト上限
NOTION_TEXT_LIMIT = 2000

# 初回運用は不具合が出やすい前提でデバッグログを常時ONにしておく。
# 落ち着いたら False にして良い。
DEBUG = True


def _dbg(label: str, **kv) -> None:
    """デバッグログ。長い値は切り詰めて1行にまとめて出す。"""
    if not DEBUG:
        return
    parts = []
    for k, v in kv.items():
        s = v if isinstance(v, str) else repr(v)
        if len(s) > 500:
            s = s[:500] + f"...(計{len(s)}文字)"
        parts.append(f"{k}={s}")
    print(f"[DEBUG] {label}: " + " / ".join(parts))

# 要約スキーマ（Miro / Canva 用）
FILE_SUMMARY_SCHEMA = {
    "type": "object",
    "properties": {
        "doc_type":          {"type": "string", "description": "ドキュメントの種類（LLM判定）"},
        "one_line_summary":  {"type": "string", "description": "1行要約"},
        "subject":           {"type": "string", "description": "主題"},
        "key_points":        {
            "type": "array",
            "items": {"type": "string"},
            "description": "主要な重要ポイントの箇条書き"
        },
        "conclusion":        {"type": "string", "description": "結論"},
        "important_figures": {"type": "string", "description": "重要な数値や定量データ"},
        "notes":             {"type": "string", "description": "備考"},
    },
    "required": [
        "doc_type", "one_line_summary",
        "subject", "key_points", "conclusion", "important_figures", "notes"
    ],
}


def _extract_urls(text: str) -> list[str]:
    """Slack形式 <URL|text> とプレーンURLを重複なく抽出する。"""
    seen: set[str] = set()
    urls: list[str] = []
    for m in _URL_RE.finditer(text):
        url = m.group(1) or m.group(2)
        if url and url not in seen:
            seen.add(url)
            urls.append(url)
    return urls


def _gdrive_file_id(url: str) -> str | None:
    m = _GDRIVE_FILE_ID_RE.search(url)
    return (m.group(1) or m.group(2)) if m else None


def _notion_page_id(url: str) -> str | None:
    m = _NOTION_UUID_RE.search(url)
    if not m:
        return None
    if m.group(1):
        return m.group(1)
    raw = m.group(2)
    return f"{raw[:8]}-{raw[8:12]}-{raw[12:16]}-{raw[16:20]}-{raw[20:]}"


# ─────────────────────────────────────────────
# Notion プロパティビルダー（統合DBの11プロパティ用）
# ─────────────────────────────────────────────

def _rt(value: str) -> dict:
    return {"rich_text": [{"text": {"content": (value or "")[:NOTION_TEXT_LIMIT]}}]}


def _title(value: str) -> dict:
    return {"title": [{"text": {"content": (value or "")[:NOTION_TEXT_LIMIT]}}]}


def _date(iso: str) -> dict:
    return {"date": {"start": iso}}


def _select(name: str) -> dict:
    return {"select": {"name": name}}


def _file_links_value(file_name: str, url: str) -> str:
    """ファイル行の links 値（＝upsert の match キー）。

    slack 側と drive 側（google_drive_recorder_v4）で同一フォーマットにすることで、
    同じファイルがどちらの経路で登録されても links の完全一致 match により
    同一行に upsert される（重複行防止）。

    URLはクエリ文字列（?以降）を除いて正規化する。Google Drive/DocsのwebViewLinkは
    ouid等のクエリパラメータが取得元によって変わり得るため、そのままだと同じファイルでも
    毎回違うmatchキーになり重複行が量産される（実運用で確認済み）。
    """
    canonical_url = url.split("?", 1)[0]
    return f"{file_name}\n{canonical_url}"[:NOTION_TEXT_LIMIT]


# ─────────────────────────────────────────────
# 要約（Miro / Canva / Notion URL 用に維持）
# ─────────────────────────────────────────────

async def _summarize_image_paths(
    img_paths: list[str],
    file_name: str,
) -> dict | None:
    """ページ画像リストを llm_call でスキーマ要約する共通処理。"""
    prompt = (
        f"以下はファイル「{file_name}」の各ページ画像です。"
        "内容を読み取り、指定のJSONスキーマに従って日本語で要約してください。"
    )
    try:
        result = await llm_call(prompt, file_paths=img_paths, schema=FILE_SUMMARY_SCHEMA)
        _dbg("画像要約 llm_call 結果", file_name=file_name, result_type=type(result).__name__)
        return result.get("data") if isinstance(result, dict) else None
    except Exception as e:
        _dbg("画像要約 llm_call 例外", file_name=file_name, error=repr(e))
        return None


async def _summarize_text_content(
    text: str,
    file_name: str,
) -> dict | None:
    """抽出済みテキストを llm_call でスキーマ要約する共通処理。"""
    combined = text[:MAX_TEXT_LENGTH]

    prompt = (
        f"以下はファイル「{file_name}」のテキスト内容です。"
        "指定のJSONスキーマに従って日本語で要約してください。"
        f"\n\n{combined}"
    )
    try:
        result = await llm_call(prompt, schema=FILE_SUMMARY_SCHEMA)
        _dbg("テキスト要約 llm_call 結果", file_name=file_name, result_type=type(result).__name__)
        return result.get("data") if isinstance(result, dict) else None
    except Exception as e:
        _dbg("テキスト要約 llm_call 例外", file_name=file_name, error=repr(e))
        return None


def _format_summary(summary: dict) -> str:
    """要約dictをNotionのrich_textに入れる文字列に変換する。"""
    if not summary:
        return ""
    return json.dumps(summary, ensure_ascii=False, indent=2)


# ─────────────────────────────────────────────
# 添付ファイル要約（Slack「slack」フォルダ保存分。
# google_drive_recorder_v4 が対象外にするため他に要約する経路が無い）
# ─────────────────────────────────────────────

def _merge_file_summaries(summaries: list[dict], total_batches: int) -> dict:
    """複数バッチ（llm_call 1回あたりの画像枚数上限ごと）の要約を1つに統合する。

    ページを切り捨てず全ページ分の情報を読むため、長い文書は LLM_MAX_FILES 枚ずつの
    バッチに分けて別々に要約し、ここで1つのdictにまとめる。情報を落とさないよう
    各フィールドはバッチごとの内容を連結する（要約を要約し直すような追加のllm_call
    は行わない＝コストを増やさない）。
    """
    if len(summaries) == 1:
        return summaries[0]

    def _join(key: str) -> str:
        parts = [s.get(key, "") for s in summaries if s.get(key)]
        return "\n---\n".join(parts)

    merged_key_points: list[str] = []
    for i, s in enumerate(summaries, 1):
        for kp in s.get("key_points") or []:
            merged_key_points.append(f"[{i}/{total_batches}バッチ目] {kp}")

    return {
        "url":               summaries[0].get("url", ""),
        "file_name":         summaries[0].get("file_name", ""),
        "doc_type":          summaries[0].get("doc_type", ""),
        "one_line_summary":  _join("one_line_summary"),
        "subject":           _join("subject"),
        "key_points":        merged_key_points,
        "conclusion":        _join("conclusion"),
        "important_figures": _join("important_figures"),
        "notes": _join("notes") + f"\n（{total_batches}バッチに分割して全ページを要約した文書）",
    }


async def _summarize_file_with_images(
    file_path: str,
    file_name: str,
) -> dict | None:
    """
    fitz で全ページを画像化 → llm_call(画像) で要約。

    llm_call 1回に渡せる画像枚数は LLM_MAX_FILES 枚までという技術的な上限があるため、
    ページ数がそれを超える文書は for ループでバッチに分けて複数回 llm_call し、
    結果を _merge_file_summaries で統合する（ページを切り捨てず全ページ読む）。
    失敗時は None を返す。
    """
    try:
        import fitz as _fitz
    except ImportError as e:
        _dbg("fitz import失敗", file_name=file_name, error=repr(e))
        return None

    try:
        doc = _fitz.open(file_path)
        total_pages = len(doc)

        img_paths: list[str] = []
        os.makedirs("tmp/pages", exist_ok=True)
        for i in range(total_pages):
            page = doc[i]
            mat = _fitz.Matrix(1.5, 1.5)  # 解像度係数
            pix = page.get_pixmap(matrix=mat)
            img_path = f"tmp/pages/{os.path.basename(file_path)}_page{i+1}.png"
            pix.save(img_path)
            img_paths.append(img_path)
        doc.close()
    except Exception as e:
        _dbg("fitz画像化 例外", file_name=file_name, error=repr(e))
        return None

    if not img_paths:
        return None

    batches = [img_paths[i:i + LLM_MAX_FILES] for i in range(0, len(img_paths), LLM_MAX_FILES)]
    _dbg("画像要約バッチ分割", file_name=file_name, total_pages=len(img_paths), batch_count=len(batches))

    summaries: list[dict] = []
    for batch in batches:
        s = await _summarize_image_paths(batch, file_name)
        if s:
            summaries.append(s)
        else:
            _dbg("画像要約バッチ失敗", file_name=file_name, batch_size=len(batch))

    if not summaries:
        return None
    return _merge_file_summaries(summaries, len(batches))


async def _summarize_file_with_text(
    file_path: str,
    file_name: str,
) -> dict | None:
    """
    fitz でテキスト抽出 → llm_call(テキスト) で要約。
    テキストが取れない場合は None を返す。
    """
    try:
        import fitz as _fitz
    except ImportError as e:
        _dbg("fitz import失敗", file_name=file_name, error=repr(e))
        return None

    try:
        doc = _fitz.open(file_path)
        page_texts: list[str] = []
        for page in doc:
            text = page.get_text()
            if text.strip():
                page_texts.append(text)
        doc.close()
    except Exception as e:
        _dbg("fitzテキスト抽出 例外", file_name=file_name, error=repr(e))
        return None

    if not page_texts:
        return None

    combined = "\n---\n".join(page_texts)
    return await _summarize_text_content(combined, file_name)


def _extract_with_pypdf(path: str) -> str:
    reader = pypdf.PdfReader(path)
    pages = [page.extract_text() or "" for page in reader.pages]
    return "\n".join(pages)


def _extract_with_docx(path: str) -> str:
    import docx
    doc = docx.Document(path)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())


def _convert_sync(local_path: str, name_ext: str) -> str:
    """markitdown → pypdf/docx フォールバックの同期本体（asyncio.to_thread から呼ばれる）。"""
    md = MarkItDown()
    text = ""
    try:
        result = md.convert(local_path)
        text = result.text_content or ""
    except Exception as md_err:
        text = ""
        if name_ext == ".docx":
            text = _extract_with_docx(local_path)
            if not text.strip():
                raise md_err

    if not text.strip() and name_ext == ".pdf":
        text = _extract_with_pypdf(local_path)

    return text


async def _summarize_file_with_markitdown(
    file_path: str,
    file_name: str,
) -> dict | None:
    """
    markitdown（→ docx/pypdf フォールバック）でテキスト抽出 → llm_call(テキスト) で要約。
    html/テキスト系は markitdown で取れなかった場合に素読み（UTF-8）でフォールバックする。
    テキストが取れない場合は None を返す。
    """
    name_ext = os.path.splitext(file_name.lower())[1]
    try:
        text = await asyncio.to_thread(_convert_sync, file_path, name_ext)
    except Exception as e:
        _dbg("markitdown抽出 例外", file_name=file_name, error=repr(e))
        text = ""

    if not text.strip() and name_ext in _PLAIN_TEXT_EXTS:
        try:
            with open(file_path, "r", encoding="utf-8", errors="replace") as f:
                text = f.read()
        except Exception as e:
            _dbg("素読みフォールバック失敗", file_name=file_name, error=repr(e))

    if not text.strip():
        return None

    return await _summarize_text_content(text, file_name)


async def _summarize_file_with_pypdf(
    file_path: str,
    file_name: str,
) -> dict | None:
    """
    pypdf でテキスト抽出 → llm_call(テキスト) で要約。
    テキストが取れない場合（スキャンPDF等）は None を返す。
    """
    try:
        text = await asyncio.to_thread(_extract_with_pypdf, file_path)
    except Exception as e:
        _dbg("pypdf抽出 例外", file_name=file_name, error=repr(e))
        return None

    if not text.strip():
        return None

    return await _summarize_text_content(text, file_name)


async def _summarize_attached_file(
    file_path: str,
    file_name: str,
) -> dict | None:
    """
    添付ファイルを要約する。ページ数上限は設けず、全ページを読む
    （_summarize_file_with_images 内でバッチ分割して全ページ処理する）。

    docx/xlsx/pptx と html/テキスト系は fitz で開けないため markitdown 経路のみ：
      markitdown テキスト抽出（テキスト系は素読みフォールバック付き）→ llm_call（テキスト）
    PDF は pypdf テキスト抽出を先に試み、無理なら fitz 画像化（全ページ）にフォールバック：
      1. pypdf テキスト抽出 → llm_call（テキスト）
      2. フォールバック: 全ページ画像化 → llm_call（画像、バッチ分割）
    それ以外（画像等）は以下の順で試みる：
    1. 全ページ画像化 → llm_call（画像、バッチ分割）
    2. フォールバック: fitz テキスト抽出 → llm_call（テキスト）
    3. 全て失敗 → None
    """
    name_ext = os.path.splitext(file_name.lower())[1]
    _dbg("添付ファイル要約開始", file_name=file_name, name_ext=name_ext)

    # STEP 0: Office 形式・html/テキスト系は markitdown でテキスト化
    #（fitz はこれらの形式を開けないため、失敗時のフォールバックはない）
    if name_ext in _MARKITDOWN_EXTS:
        return await _summarize_file_with_markitdown(file_path, file_name)

    # PDF: pypdf テキスト抽出 → 無理なら fitz 画像化（全ページ） → llm_call
    if name_ext == ".pdf":
        result = await _summarize_file_with_pypdf(file_path, file_name)
        if result:
            return result
        return await _summarize_file_with_images(file_path, file_name)

    # STEP 1: 全ページ画像（バッチ分割） → llm_call
    result = await _summarize_file_with_images(file_path, file_name)
    if result:
        return result

    # STEP 2: テキスト抽出 → llm_call
    return await _summarize_file_with_text(file_path, file_name)


async def _summarize_notion(url: str) -> str | None:
    """Notion ページのブロックを取得し、llm_call で要約する。"""
    page_id = _notion_page_id(url)
    if not page_id:
        _dbg("Notion URL page_id抽出失敗", url=url)
        return None
    try:
        blocks_res = await NOTION_FETCH_ALL_BLOCK_CONTENTS(block_id=page_id)
    except Exception as e:
        _dbg("NOTION_FETCH_ALL_BLOCK_CONTENTS 例外", page_id=page_id, error=repr(e))
        return None
    if not blocks_res:
        _dbg("NOTION_FETCH_ALL_BLOCK_CONTENTS 空レスポンス", page_id=page_id)
        return None

    blocks = blocks_res if isinstance(blocks_res, list) else (
        (blocks_res.get("data") or blocks_res).get("results", [])
    )
    texts: list[str] = []
    for block in blocks:
        block_type = block.get("type")
        if not block_type:
            continue
        rich_texts = block.get(block_type, {}).get("rich_text", [])
        text = "".join(t.get("plain_text", "") for t in rich_texts)
        if text.strip():
            texts.append(text)

    if not texts:
        return None

    combined = "\n".join(texts)[:6000]
    try:
        raw = await llm_call(
            f"以下はNotionページの内容です。タイトルと要旨を日本語で1〜3文に要約してください。\n\n{combined}"
        )
        if isinstance(raw, dict):
            raw = raw.get("content") or raw.get("text") or str(raw)
        return f"[Notion] {raw}" if raw else None
    except Exception as e:
        _dbg("Notion要約 llm_call 例外", page_id=page_id, error=repr(e))
        return None


def _strip_html(html: str) -> str:
    """Miro アイテムの content から簡易的に HTML タグを除去する。"""
    if not html:
        return ""
    text = re.sub(r"<[^>]+>", "", html)
    text = text.replace("&#xff08;", "(").replace("&#xff09;", ")")
    text = text.replace("&#xff1a;", ":").replace("&#xff1f;", "?")
    text = text.replace("&nbsp;", " ")
    text = re.sub(r"​", "", text)  # zero-width space
    return text.strip()


def _tool_res_ok(res) -> bool:
    """ツールレスポンスの成功判定（successful / successfull の表記ゆれ両対応）。"""
    if not isinstance(res, dict):
        return False
    return bool(res.get("successful") or res.get("successfull"))


async def _summarize_miro(url: str) -> str | None:
    """Miro ボードの全アイテムを取得し、テキスト要素を抽出して llm_call で要約する。"""
    m = re.search(r"/board/([^/?#]+)", url)
    if not m:
        return None
    board_id = urllib.parse.unquote(m.group(1))

    texts: list[str] = []
    cursor = ""
    while True:
        kwargs: dict = {"board_id": board_id, "limit": 50}
        if cursor:
            kwargs["cursor"] = cursor
        try:
            res = await MIRO_GET_BOARD_ITEMS(**kwargs)
        except Exception as e:
            _dbg("MIRO_GET_BOARD_ITEMS 例外", board_id=board_id, error=repr(e))
            return None
        if not _tool_res_ok(res):
            _dbg("MIRO_GET_BOARD_ITEMS 失敗", board_id=board_id, error=str(res.get("error") or res))
            return None
        data = res.get("data") or {}
        for item in (data.get("data") or []):
            content_html = (item.get("data") or {}).get("content")
            if content_html:
                text = _strip_html(content_html)
                if text:
                    texts.append(text)
        cursor = data.get("cursor") or ""
        if not cursor:
            break

    if not texts:
        return None

    combined = "\n".join(f"- {t}" for t in texts)
    summary = await _summarize_text_content(combined, f"Miroボード({board_id})")
    if summary:
        return _format_summary(summary)
    return None


def _resolve_canva_design_id(canva_url: str) -> str | None:
    """Canva URL から design_id を抽出する。canva.link 等の短縮URLはリダイレクトを解決する。"""
    m = re.search(r"/design/([A-Za-z0-9_-]+)", canva_url)
    if m:
        return m.group(1)

    import requests as _req
    headers = {"User-Agent": "Mozilla/5.0 (compatible; SummaryBot/1.0)"}
    final_url = ""
    try:
        resp = _req.head(canva_url, allow_redirects=True, timeout=15, headers=headers)
        final_url = resp.url or ""
    except Exception:
        pass
    m = re.search(r"/design/([A-Za-z0-9_-]+)", final_url)
    if m:
        return m.group(1)

    # HEAD がブロックされるサイト向けの GET フォールバック
    try:
        resp = _req.get(canva_url, allow_redirects=True, timeout=15, headers=headers, stream=True)
        final_url = resp.url or ""
        resp.close()
    except Exception:
        return None
    m = re.search(r"/design/([A-Za-z0-9_-]+)", final_url)
    return m.group(1) if m else None


async def _summarize_canva(url: str) -> str | None:
    """Canva デザインの各ページ画像を1ページずつDLし、llm_call(画像) で要約する。"""
    design_id = _resolve_canva_design_id(url)
    if not design_id:
        _dbg("Canva design_id抽出失敗", url=url)
        return None

    try:
        res = await CANVA_LIST_DESIGN_PAGES_WITH_PAGINATION(
            designId=design_id, limit=200, offset=1
        )
    except Exception as e:
        _dbg("CANVA_LIST_DESIGN_PAGES 例外", design_id=design_id, error=repr(e))
        return None
    if not _tool_res_ok(res):
        _dbg("CANVA_LIST_DESIGN_PAGES 失敗", design_id=design_id, error=str(res.get("error") or res))
        return None

    items = (res.get("data") or {}).get("items") or []

    import requests as _req
    headers = {"User-Agent": "Mozilla/5.0 (compatible; SummaryBot/1.0)"}
    os.makedirs("tmp/canva_pages", exist_ok=True)
    img_paths: list[str] = []
    for i, item in enumerate(items, 1):
        if len(img_paths) >= LLM_MAX_FILES:
            break
        thumb_url = (item.get("thumbnail") or {}).get("url")
        if not thumb_url:
            continue
        page_no = item.get("page_number") or item.get("index") or i
        img_path = f"tmp/canva_pages/{design_id}_page{page_no}.png"
        try:
            resp = _req.get(thumb_url, timeout=30, headers=headers)
            resp.raise_for_status()
            with open(img_path, "wb") as f:
                f.write(resp.content)
            img_paths.append(img_path)
        except Exception:
            continue

    if not img_paths:
        return None

    summary = await _summarize_image_paths(img_paths, f"Canvaデザイン({design_id})")
    if summary:
        return _format_summary(summary)
    return None


# ─────────────────────────────────────────────
# ツール呼び出し（リトライ付き）
# ─────────────────────────────────────────────

TEAM = "T03PCARC859"
SERVER_ERROR_WAIT = 3
SERVER_ERROR_RETRIES = 5


async def _tool(fn, label: str = "tool") -> dict:
    for attempt in range(SERVER_ERROR_RETRIES + 1):
        try:
            result = await fn()
        except Exception as e:
            _dbg(f"{label} 呼び出し例外", attempt=f"{attempt + 1}/{SERVER_ERROR_RETRIES + 1}", error=repr(e))
            if attempt < SERVER_ERROR_RETRIES:
                await asyncio.sleep(SERVER_ERROR_WAIT)
                continue
            raise
        # successful/successfull の表記ゆれ両対応（Composio側で混在するため）。
        # ここを片方だけ見ると、実際は成功している呼び出しを誤って「失敗」と
        # ログ・リトライ判定してしまう（実運用で確認済みの表記ゆれ）。
        ok = (result or {}).get("successful") or (result or {}).get("successfull")
        if ok:
            return result
        error = str((result or {}).get("error") or "")
        _dbg(
            f"{label} 失敗レスポンス",
            attempt=f"{attempt + 1}/{SERVER_ERROR_RETRIES + 1}",
            error=error,
            result_keys=list((result or {}).keys()),
        )
        if ("500" in error or "Internal Server Error" in error) and attempt < SERVER_ERROR_RETRIES:
            await asyncio.sleep(SERVER_ERROR_WAIT)
        else:
            return result
    return result

JST = timezone(timedelta(hours=9))


def _normalize_pjt(name: str) -> str:
    return unicodedata.normalize("NFC", (name or "")).replace("pjt_ ", "pjt_")


# ─────────────────────────────────────────────
# 統合DBへの upsert
# ─────────────────────────────────────────────

def _rt_match(property_name: str, value: str) -> dict:
    """rich_text型プロパティの完全一致matchフィルタ。"""
    return {"property": property_name, "equals": value}


def _date_match(property_name: str, iso_value: str) -> dict:
    """date型プロパティの完全一致matchフィルタ。

    Notion APIのフィルタはプロパティ型ごとに形が変わり、date型は
    {"property": p, "date": {"equals": v}} が必要（rich_text用の
    {"property": p, "equals": v} を date型プロパティに送ると
    「database property date does not match filter text」で失敗する。
    実運用の初回実行で実際にこのエラーを確認済み）。
    """
    return {"property": property_name, "date": {"equals": iso_value}}


def _title_match(property_name: str, value: str) -> dict:
    """title型プロパティの完全一致matchフィルタ。"""
    return {"property": property_name, "title": {"equals": value}}


def _upsert_ok(res: dict) -> tuple[bool, str]:
    """Notion UPSERT の実際の成否を判定する。

    トップレベルの successful/successfull は Composio ツール呼び出し自体が
    実行できたか（＝APIと通信できたか）を示すだけで、実際に行の作成/更新が
    成功したかどうかは data.error_count / data.results[].ok に入っている。
    ここを見ずに top-level フラグだけで判定すると、プロパティ不一致等の
    バリデーションエラーで書き込みが失敗していても「成功」と誤検知する
    （実運用の初回実行でこの誤検知を確認済み・要修正済み）。
    """
    if not isinstance(res, dict):
        return False, "レスポンスがdictでない"
    top_ok = bool(res.get("successful") or res.get("successfull"))
    if not top_ok:
        return False, str(res.get("error") or res)
    data = res.get("data") or {}
    results = data.get("results")
    if results is None:
        # results を返さない実装もあり得るため、その場合はトップレベルの成否を信じる
        return True, ""
    bad = [r for r in results if not r.get("ok")]
    if bad:
        detail = "; ".join(
            str(r.get("error_message") or r.get("error_code") or r) for r in bad
        )
        return False, detail
    return True, ""


async def _upsert_row(properties: dict, match: dict) -> bool:
    """統合DBへ1行 upsert する。match と一致する行があれば update、なければ create。"""
    _dbg(
        "upsert 送信内容",
        database_id=DB_ID,
        match=match,
        properties=json.dumps(properties, ensure_ascii=False),
    )
    res = await _tool(lambda: NOTION_UPSERT_ROW_DATABASE(
        database_id=DB_ID,
        items=[{
            "match":  match,
            "create": {"properties": properties},
            "update": {"properties": properties},
        }],
    ), label="NOTION_UPSERT_ROW_DATABASE")
    ok, err_detail = _upsert_ok(res)
    _dbg(
        "upsert 結果",
        match=match,
        successful=ok,
        error=err_detail,
        raw_result=json.dumps(res, ensure_ascii=False, default=str),
    )
    if not ok:
        print(f"[ERROR] upsert 失敗（Notionに書き込まれていません）: match={match} / {err_detail}")
    return ok


async def _upsert_drive_file_row(
    *,
    file_name: str,
    drive_url: str,
    ts_iso: str,
    sender_email: str,
    sender_name: str,
    project: str,
    parents: str,
    summary: dict | None = None,
) -> bool:
    """Drive に実体があるファイルの行（source=drive）を upsert する。

    match キーは links（"ファイル名\\nURL"）。同じファイルが slack 経由と
    drive 情報集約エージェントの両方で登録されても同一行に収束する。

    summary を渡した場合は content に要約JSONを入れる（Slack添付ファイル用。
    google_drive_recorder_v4 が「slack」フォルダを対象外にしているため、
    要約はここで生成しないと永久にファイル名のみになってしまう）。
    summary が無ければ従来通り content=ファイル名（メッセージ内Drive URL用。
    こちらは v4 の定期スキャンが同じ links キーで後から要約入りに上書きする想定）。
    """
    links_value = _file_links_value(file_name, drive_url)
    content_value = _format_summary(summary) if summary else file_name
    properties = {
        "content":   _title(content_value),
        "subject":   _rt(file_name),
        "parents":   _rt(parents),
        "timestamp": _date(ts_iso),
        "sender":    _rt(sender_email),
        "name":      _rt(sender_name),
        "project":   _rt(project),
        "source":    _select("drive"),
        "links":     _rt(links_value),
    }
    ok = await _upsert_row(properties, _rt_match("links", links_value))
    if ok:
        print(f"[OK] ファイル行 upsert: {file_name}" + ("（要約あり）" if summary else ""))
    else:
        print(f"[WARN] ファイル行 upsert 失敗: {file_name}")
    return ok


# ─────────────────────────────────────────────
# Google Drive アップロード（添付ファイルの保存）
# ─────────────────────────────────────────────

_LINK_KEYS = ("webViewLink", "display_url", "webContentLink", "web_view_link")


def _walk(obj):
    """ネストした dict/list を深さ優先で全ノード走査する。"""
    if isinstance(obj, dict):
        yield obj
        for v in obj.values():
            yield from _walk(v)
    elif isinstance(obj, list):
        for v in obj:
            yield from _walk(v)


def _extract_drive_link(res: dict) -> str:
    """UPLOAD_FILE レスポンスから閲覧URLを取り出す（入れ子が変わっても拾えるよう全走査）。

    実運用で確認済み: レスポンスに webViewLink 等のリンクキーが一切含まれず
    id/kind("drive#file")/name のみが返るケースがある。その場合は
    _drive_view_url（重複チェック側）と同じロジックで id から URL を組み立てる
    フォールバックが無いと、アップロード自体は成功しているのにファイル行が
    upsert されずデータが消える（実運用で発生確認済み）。
    """
    nodes = list(_walk(res))
    for node in nodes:
        for key in _LINK_KEYS:
            link = node.get(key)
            if isinstance(link, str) and link:
                return link
    for node in nodes:
        for v in node.values():
            if isinstance(v, str) and "drive.google.com" in v:
                return v
    for node in nodes:
        if node.get("kind") == "drive#file" and isinstance(node.get("id"), str) and node["id"]:
            return f"https://drive.google.com/file/d/{node['id']}/view"
    return ""


def _versioned_name(file_name: str, ts_iso: str) -> str:
    """同名で内容が異なるファイル（更新版）用に、日時サフィックス付きの別名を作る。"""
    stem, ext = os.path.splitext(file_name)
    stamp = ts_iso.replace(":", "").replace("-", "")[:15]  # 20260712T103000
    return f"{stem}_{stamp}{ext}"


async def _find_existing_in_project_drive(file_name: str) -> dict | None:
    """共有ドライブ「project」全体（フォルダを問わず）で同名ファイルを検索する。

    「slack」フォルダ内だけに限定すると、project内の別フォルダ（例えば
    google_drive_recorder_v4 が既に集約済みの通常フォルダ）に同名ファイルが
    既にある場合の重複を検出できない。そのため検索対象は特定フォルダではなく
    共有ドライブ「project」全体にする（driveId + corpora=drive で指定）。
    見つからなければ None。
    """
    safe = file_name.replace("\\", "\\\\").replace("'", "\\'")
    q = f"name = '{safe}' and trashed = false"
    res = await _tool(lambda: GOOGLEDRIVE_FIND_FILE(
        q=q, corpora="drive", driveId=PROJECT_SHARED_DRIVE_ID,
    ), label="GOOGLEDRIVE_FIND_FILE(重複チェック)")
    if not res.get("successful"):
        _dbg("重複チェック失敗", query=q, error=str(res.get("error") or res))
        return None
    files = (res.get("data") or {}).get("files") or []
    _dbg("重複チェック結果", file_name=file_name, query=q, hit_count=len(files),
         first_hit=files[0] if files else None)
    return files[0] if files else None


def _drive_view_url(file: dict) -> str:
    """FIND_FILE の結果から閲覧URLを得る。webViewLink が無ければ fileId から組み立てる。"""
    url = file.get("webViewLink") or file.get("web_view_link") or ""
    if not url and file.get("id"):
        url = f"https://drive.google.com/file/d/{file['id']}/view"
    return url


async def _save_attachment_to_drive(
    local_path: str,
    file_name: str,
    mimetype: str,
    slack_size: int | None,
    ts_iso: str,
) -> tuple[str, str] | None:
    """添付ファイルを「slack」フォルダに保存し、(保存名, DriveURL) を返す。

    重複チェックは共有ドライブ「project」全体が対象（「slack」フォルダ限定ではない）:
    - 同名ファイルなし → そのままアップロード
    - 同名あり・サイズ一致（またはサイズ不明）→ アップロードせず既存リンクを返す
    - 同名あり・サイズ不一致（更新版）→ 日時サフィックス付き別名でアップロード
    失敗時は None。
    """
    upload_name = file_name
    existing = await _find_existing_in_project_drive(file_name)
    if existing:
        try:
            existing_size = int(existing.get("size") or 0)
        except (TypeError, ValueError):
            existing_size = 0
        _dbg(
            "サイズ比較（重複/更新版判定）",
            file_name=file_name,
            slack_size=slack_size,
            existing_size=existing_size,
            existing_file_id=existing.get("id"),
        )
        if not slack_size or not existing_size or slack_size == existing_size:
            url = _drive_view_url(existing)
            if url:
                print(f"[DUP] 同名ファイルが既に存在するため保存スキップ: {file_name}")
                return file_name, url
            _dbg("既存ファイルのURLが取得できず新規アップロードにフォールバック", file_name=file_name,
                 existing=existing)
        else:
            upload_name = _versioned_name(file_name, ts_iso)
            print(f"[VER] 同名・内容違いのため別名で保存: {upload_name}")

    _dbg(
        "Drive アップロード実行",
        upload_name=upload_name,
        mimetype=mimetype,
        local_path=local_path,
        local_path_exists=os.path.exists(local_path),
        folder_id=SLACK_DRIVE_FOLDER_ID,
    )
    res = await _tool(lambda: GOOGLEDRIVE_UPLOAD_FILE(
        file_to_upload={
            "name":     upload_name,
            "mimetype": mimetype or "application/octet-stream",
            "s3key":    local_path,
        },
        folder_to_upload_to=SLACK_DRIVE_FOLDER_ID,
    ), label="GOOGLEDRIVE_UPLOAD_FILE")
    _dbg("Drive アップロードレスポンス全体", upload_name=upload_name,
         raw_result=json.dumps(res, ensure_ascii=False, default=str))
    if not (res.get("successful") or res.get("successfull")):
        print(f"[WARN] Drive アップロード失敗（5MB超の可能性あり）: {upload_name} / "
              f"{str(res.get('error', ''))[:200]}")
        return None

    url = _extract_drive_link(res)
    if not url:
        # アップロード自体は成功しているのにURLが取れないと、このファイルの行は
        # 二度とDBに登録されない（実運用でwebViewLink無し・id/kindのみのレスポンスに
        # 遭遇し、データが消えることを確認済み）。ERRORとして明示する。
        print(f"[ERROR] アップロード成功したがURLを抽出できずファイルが失われます: {upload_name} "
              f"/ raw={json.dumps(res, ensure_ascii=False, default=str)[:300]}")
        return None
    print(f"[UP] Drive 保存完了: {upload_name} → {url}")
    return upload_name, url


# ─────────────────────────────────────────────
# メッセージ内 URL の処理
# ─────────────────────────────────────────────

async def _summarize_gdrive_url_file(file_id: str, file_name: str) -> dict | None:
    """メッセージ内で共有された Drive ファイルをダウンロードして要約する。

    Slack添付ファイルと同じ _summarize_attached_file パイプライン（markitdown/pypdf/fitz）
    を再利用する。google_drive_recorder_v4 の定期スキャンは createdTime が直近
    LOOKBACK_HOURS 以内のファイルしか対象にしないため、以前から存在するファイルが
    共有された場合は v4 が永久に拾わない。そのためここで直接要約する
    （content をファイル名のみにしない。2026-07-14 ユーザー確認：URL/添付ファイル行の
    content には最初から要約を入れる想定だった）。
    失敗時は None（呼び出し元は content をファイル名のみにフォールバックする）。
    """
    try:
        dl_res = await GOOGLEDRIVE_DOWNLOAD_FILE(fileId=file_id)
    except Exception as e:
        _dbg("Drive URLファイル ダウンロード例外", file_id=file_id, error=repr(e))
        return None
    if not isinstance(dl_res, dict) or not (dl_res.get("successful") or dl_res.get("successfull")):
        _dbg("Drive URLファイル ダウンロード失敗", file_id=file_id,
             error=str((dl_res or {}).get("error") if isinstance(dl_res, dict) else dl_res))
        return None

    data = dl_res.get("data") or {}
    content = data.get("downloaded_file_content") or {}
    local_path = content.get("s3url") or ""
    actual_name = content.get("name") or file_name

    if not local_path or not os.path.exists(local_path):
        _dbg("Drive URLファイル ローカルパスが見つからない", file_id=file_id, local_path=local_path)
        return None

    return await _summarize_attached_file(local_path, actual_name)


async def _record_gdrive_url(
    url: str,
    *,
    ts_iso: str,
    sender_email: str,
    sender_name: str,
    project: str,
) -> tuple[str, str] | None:
    """メッセージ内で共有された Drive URL のファイル行（要約あり）を upsert する。

    links の完全一致 match により、既に登録済み（drive エージェント経由含む）なら
    重複行は作られず既存行が update されるだけ。メタデータが取れない場合は何もしない。

    戻り値は (file_name, web_url)。メッセージ行の links にもファイル名を含めるため、
    呼び出し元（_process_urls）がこれを使って links エントリを組み立てる。
    """
    file_id = _gdrive_file_id(url)
    if not file_id:
        _dbg("Drive URL file_id抽出失敗", url=url)
        return None
    try:
        meta_res = await GOOGLEDRIVE_GET_FILE_METADATA(
            fileId=file_id, fields="id,name,webViewLink,trashed"
        )
    except Exception as e:
        _dbg("GOOGLEDRIVE_GET_FILE_METADATA 例外", file_id=file_id, error=repr(e))
        return None
    if not isinstance(meta_res, dict) or meta_res.get("successful") is False:
        _dbg("GOOGLEDRIVE_GET_FILE_METADATA 失敗", file_id=file_id,
             error=str((meta_res or {}).get("error") if isinstance(meta_res, dict) else meta_res))
        return None
    meta = meta_res.get("data") or meta_res
    if meta.get("trashed"):
        _dbg("Drive URLファイルはゴミ箱のためスキップ", file_id=file_id)
        return None
    file_name = meta.get("name") or ""
    if not file_name:
        _dbg("Drive URLメタデータにnameが無い", file_id=file_id, meta=meta)
        return None
    web_url = meta.get("webViewLink") or url
    _dbg("Drive URLファイル行 upsert 対象", file_id=file_id, file_name=file_name, web_url=web_url)

    summary = await _summarize_gdrive_url_file(file_id, file_name)
    _dbg("Drive URLファイル要約結果", file_id=file_id, file_name=file_name, summarized=bool(summary))

    await _upsert_drive_file_row(
        file_name=file_name,
        drive_url=web_url,
        ts_iso=ts_iso,
        sender_email=sender_email,
        sender_name=sender_name,
        project=project,
        parents="",
        summary=summary,
    )
    return file_name, web_url


async def _process_urls(
    content_text: str,
    *,
    ts_iso: str,
    sender_email: str,
    sender_name: str,
    project: str,
) -> list[str]:
    """メッセージ内URLを処理し、links プロパティに入れるパーツのリストを返す。

    - Google Drive/Docs URL: ファイル行を upsert し、links にはファイル名+URL を入れる
    - Miro / Canva / Notion: 要約を生成し「URL + 要約」を links に入れる
    - それ以外（slack.com/files 含む）: URL をそのまま links に入れる
    """
    parts: list[str] = []
    urls = _extract_urls(content_text)
    _dbg("メッセージ内URL抽出", count=len(urls), urls=urls)
    for url in urls:
        if "docs.google.com" in url or "drive.google.com" in url:
            _dbg("URL振り分け", url=url, handler="gdrive")
            resolved = await _record_gdrive_url(
                url,
                ts_iso=ts_iso,
                sender_email=sender_email,
                sender_name=sender_name,
                project=project,
            )
            if resolved:
                file_name, web_url = resolved
                parts.append(_file_links_value(file_name, web_url))
            else:
                parts.append(url)
        elif any(d in url for d in ("notion.site", "notion.so", "notion.com", "app.notion.com")):
            _dbg("URL振り分け", url=url, handler="notion")
            s = await _summarize_notion(url)
            parts.append(f"{url}\n{s}" if s else url)
        elif "miro.com" in url:
            _dbg("URL振り分け", url=url, handler="miro")
            s = await _summarize_miro(url)
            parts.append(f"{url}\n{s}" if s else url)
        elif "canva.com" in url or "canva.link" in url:
            _dbg("URL振り分け", url=url, handler="canva")
            s = await _summarize_canva(url)
            parts.append(f"{url}\n{s}" if s else url)
        else:
            _dbg("URL振り分け", url=url, handler="そのまま(links格納のみ)")
            parts.append(url)
    return parts


# ─────────────────────────────────────────────
# メイン処理
# ─────────────────────────────────────────────

async def fetch_channel_map() -> dict[str, str]:
    """ワークスペース内の全チャンネル（channel_id → 生の名前）を取得する。

    record_slack_post が毎回内部で呼ぶほか、slack_backfill 等の一括処理側が
    事前に1回だけ呼んで record_slack_post に渡すことで、メッセージ件数分の
    重複したSLACK_LIST_CONVERSATIONS呼び出し（レート制限の主因）を避けられる。
    """
    channel_map: dict[str, str] = {}
    cursor = ""
    while True:
        res = await _tool(lambda cursor=cursor: SLACK_LIST_CONVERSATIONS(
            limit=200,
            types="public_channel,private_channel",
            team_id=TEAM,
            cursor=cursor,
        ))
        if not res.get("successful"):
            break
        data = res.get("data") or {}
        for ch in (data.get("channels") or []):
            channel_map[ch["id"]] = ch.get("name", ch["id"])
        cursor = (data.get("response_metadata") or {}).get("next_cursor", "")
        if not cursor:
            break
    return channel_map


async def fetch_user_maps() -> tuple[dict[str, str], dict[str, str]]:
    """(user_map, email_map) を取得する。fetch_channel_map と同じ理由で外部から再利用可能にする。"""
    user_map:  dict[str, str] = {}
    email_map: dict[str, str] = {}
    cursor = ""
    while True:
        res = await _tool(lambda cursor=cursor: SLACK_LIST_ALL_USERS(limit=200, cursor=cursor))
        if not res.get("successful"):
            break
        data = res.get("data") or {}
        for m in (data.get("members") or []):
            if m.get("deleted") or m.get("is_bot"):
                continue
            uid  = m.get("id")
            prof = m.get("profile", {})
            user_map[uid] = (
                prof.get("display_name")
                or prof.get("real_name")
                or m.get("name")
                or uid
            )
            email_map[uid] = prof.get("email") or ""
        cursor = (data.get("response_metadata") or {}).get("next_cursor", "")
        if not cursor:
            break
    return user_map, email_map


async def record_slack_post(
    payload: dict,
    *,
    channel_map: dict[str, str] | None = None,
    user_map: dict[str, str] | None = None,
    email_map: dict[str, str] | None = None,
) -> None:
    """1件のSlack投稿を統合DBに記録する。

    channel_map/user_map/email_map を渡すと、その場でのSlack API取得を省略する
    （slack_backfill のように大量メッセージを処理する場合、1件ずつ毎回取得すると
    レート制限に即座に引っかかるため、呼び出し側で1回だけ取得して使い回す）。
    リアルタイム1件処理（__main__ 経由）では渡さず、これまで通り毎回取得する。
    """
    _dbg("payload受信", channel=payload.get("channel"), user=payload.get("user"),
         ts=payload.get("ts"), thread_ts=payload.get("thread_ts"),
         file_count=len(payload.get("files") or []))

    # STEP 1: チャンネル名解決 → pjt_ 以外はスキップ
    channel_id = payload.get("channel")
    if channel_map is None:
        channel_map = await fetch_channel_map()

    channel_name = _normalize_pjt(channel_map.get(channel_id, channel_id))
    _dbg("チャンネル解決", channel_id=channel_id, channel_name=channel_name,
         channel_map_size=len(channel_map))
    if not channel_name.startswith("pjt_"):
        print(f"[SKIP] pjt_チャンネルでないため保存しません: {channel_name}")
        return

    # STEP 2: 投稿情報の整形
    ts_f = float(payload.get("ts", "0"))
    sent_at = datetime.fromtimestamp(ts_f, tz=JST).strftime("%Y-%m-%dT%H:%M:%S+09:00")
    _dbg("sent_at算出", raw_ts=payload.get("ts"), sent_at=sent_at)

    file_records = []
    for f in (payload.get("files") or []):
        # file_access=check_file_info のイベントは name/mimetype を含まない。
        # その場合は id + filetype から "F0XXXX.pdf" 形式の名前を組み立てる
        # （実名はDL後に fc.name で上書きされる）。
        name = f.get("name") or f.get("title")
        if not name:
            name = f.get("id") or "unknown"
            if f.get("filetype"):
                name = f"{name}.{f['filetype']}"
        file_records.append({
            "id":        f.get("id"),
            "name":      name,
            "mimetype":  f.get("mimetype") or "",
            "size":      f.get("size"),
            "drive_url": "",   # Drive 保存後の閲覧URL
            "saved":     False,
        })

    record: dict = {
        "sender":  payload.get("user"),
        "content": payload.get("text") or "",
        "sent_at": sent_at,
        "channel": channel_name,
        "files":   file_records,
    }
    # parents: スレッド返信時は親メッセージの ts（timestamp と同じ ISO 形式）
    if payload.get("thread_ts"):
        parent_ts = float(payload["thread_ts"])
        record["parents"] = datetime.fromtimestamp(parent_ts, tz=JST).strftime("%Y-%m-%dT%H:%M:%S+09:00")

    # STEP 3: ユーザーマップ取得（sender=メール / name=表示名 に使う）
    if user_map is None or email_map is None:
        user_map, email_map = await fetch_user_maps()

    sender_uid   = record["sender"]
    sender_name  = user_map.get(sender_uid, sender_uid) or ""
    sender_email = email_map.get(sender_uid, "") or ""
    _dbg("送信者解決", sender_uid=sender_uid, sender_name=sender_name,
         sender_email=sender_email, email_found=bool(sender_email))
    if not sender_email:
        _dbg("警告: 送信者メールが空（senderプロパティが空欄になる）", sender_uid=sender_uid)
    record["content"] = re.sub(
        r"<@(U[A-Z0-9]+)>",
        lambda m: f"@{user_map.get(m.group(1), m.group(1))}",
        record["content"],
    )

    # STEP 4: メッセージ内 URL の処理（Drive URL はファイル行 upsert / Miro等は要約）
    links_parts = await _process_urls(
        record["content"],
        ts_iso=sent_at,
        sender_email=sender_email,
        sender_name=sender_name,
        project=channel_name,
    )

    # STEP 5: 添付ファイルのダウンロード → Drive「slack」フォルダへ保存 → ファイル行 upsert
    if file_records:
        import requests as _req, shutil as _shutil, os as _os
        _os.makedirs("tmp", exist_ok=True)

        for fr in file_records:
            _dbg("添付ファイル処理開始", id=fr["id"], name=fr["name"], mimetype=fr["mimetype"],
                 size=fr.get("size"))
            try:
                dl_res = await SLACK_DOWNLOAD_SLACK_FILE(file=fr["id"])
                if not dl_res.get("successful"):
                    _dbg("SLACK_DOWNLOAD_SLACK_FILE 失敗", id=fr["id"], name=fr["name"],
                         error=str(dl_res.get("error") or dl_res))
                    continue

                fc = (dl_res.get("data") or {}).get("file_content") or {}
                s3url = fc.get("s3url")
                if not s3url:
                    _dbg("SLACK_DOWNLOAD_SLACK_FILE レスポンスにs3urlが無い", id=fr["id"], data=fc)
                    continue
                fname = fc.get("name") or fr["name"]
                fr["mimetype"] = fc.get("mimetype", fr["mimetype"])
                fr["name"] = fname

                tmp_path = f"tmp/{fr['id']}_{fname}"
                if s3url.startswith("http"):
                    resp = _req.get(s3url, timeout=60)
                    resp.raise_for_status()
                    with open(tmp_path, "wb") as _f:
                        _f.write(resp.content)
                else:
                    # ツールブリッジはファイルを sandbox 内に保存し、
                    # s3url フィールドにそのローカルパスを返す
                    if not _os.path.exists(s3url):
                        _dbg("s3url がローカルパスとして存在しない", id=fr["id"], s3url=s3url)
                        continue
                    _shutil.copy(s3url, tmp_path)

                try:
                    slack_size = int(fr.get("size") or 0) or os.path.getsize(tmp_path)
                except OSError:
                    slack_size = int(fr.get("size") or 0)
                _dbg("ダウンロード完了", id=fr["id"], fname=fname, tmp_path=tmp_path,
                     slack_size=slack_size)

                saved = await _save_attachment_to_drive(
                    local_path=tmp_path,
                    file_name=fname,
                    mimetype=fr["mimetype"],
                    slack_size=slack_size,
                    ts_iso=sent_at,
                )
                if not saved:
                    _dbg("Drive保存失敗のためこのファイルはスキップ", id=fr["id"], name=fname)
                    continue

                saved_name, drive_url = saved
                fr["name"] = saved_name
                fr["drive_url"] = drive_url
                fr["saved"] = True

                # 添付ファイルの中身を要約する。google_drive_recorder_v4 は「slack」
                # フォルダを対象外にしているため、ここで要約しないと永久にファイル名
                # のみのままになる（tmp_path はDrive保存の成否に関わらずローカルに
                # 残っているのでダウンロードし直さず使える）。
                summary = await _summarize_attached_file(
                    file_path=tmp_path, file_name=saved_name,
                )
                _dbg("添付ファイル要約結果", id=fr["id"], name=saved_name, summarized=bool(summary))

                # ファイル行 upsert（source=drive / parents=slack 固定 / 要約あり）
                await _upsert_drive_file_row(
                    file_name=saved_name,
                    drive_url=drive_url,
                    ts_iso=sent_at,
                    sender_email=sender_email,
                    sender_name=sender_name,
                    project=channel_name,
                    parents="slack",
                    summary=summary,
                )

                # メッセージ行の links にも保存先を記載
                links_parts.append(_file_links_value(saved_name, drive_url))

            except Exception as e:
                import traceback
                print(f"[WARN] 添付ファイル処理失敗: {fr['name']} / {e}")
                traceback.print_exc()

    # STEP 6: メッセージ行の upsert
    # content: 本文＋ファイル名（仕様）
    content_value = record["content"]
    file_names = [fr["name"] for fr in file_records if fr.get("name")]
    if file_names:
        content_value = f"{content_value}\n[添付ファイル] {', '.join(file_names)}".strip()

    links_text = "\n\n".join(links_parts)
    _dbg("メッセージ行 確定内容", content_value=content_value, links_count=len(links_parts))

    for part in links_parts:
        print(f"[links] {part[:SUMMARY_PREVIEW_CHARS]}")

    properties = {
        "content":   _title(content_value),
        "timestamp": _date(sent_at),
        "sender":    _rt(sender_email),
        "name":      _rt(sender_name),
        "parents":   _rt(record.get("parents") or ""),
        "project":   _rt(channel_name),
        "source":    _select("slack"),
        "links":     _rt(links_text),
    }

    # match は timestamp ではなく content（title）にする。
    # ファイル行とメッセージ行は同じ Slack メッセージ由来だと timestamp が完全に一致するため、
    # timestamp を match キーにするとメッセージ行の upsert がファイル行のページに誤って
    # マッチしてしまい、subject 等は古いまま content/source/links だけ上書きされて
    # データが壊れる（実運用で確認済み）。content はメッセージ行とファイル行で
    # 構造的に一致しないため、この衝突が起きない。
    ok = await _upsert_row(properties, _title_match("content", content_value[:NOTION_TEXT_LIMIT]))
    if not ok:
        raise RuntimeError(f"Notion 保存失敗: timestamp={sent_at}")

    print(
        f"[OK] Notion 保存完了: timestamp={sent_at} / project={channel_name}"
        f" / links {len(links_parts)}件 / 添付 {sum(1 for fr in file_records if fr.get('saved'))}件保存"
    )


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")

    # file_create は sandbox の tmp/ 配下に書き出すため tmp/payload.json を優先。
    # ローカルデバッグ用にカレント直下もフォールバックで見る。
    payload_path = next(
        (p for p in ("tmp/payload.json", "payload.json") if os.path.exists(p)),
        "tmp/payload.json",
    )
    with open(payload_path, "r", encoding="utf-8") as f:
        outer = json.load(f)
        payload = outer.get("payload") or outer

    try:
        loop = asyncio.get_running_loop()
        loop.run_until_complete(record_slack_post(payload))
    except RuntimeError:
        asyncio.run(record_slack_post(payload))
