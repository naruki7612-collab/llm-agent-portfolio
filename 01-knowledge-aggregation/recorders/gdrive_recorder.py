"""google_drive_recorder_v4 — Drive 保存ファイルを統合Notion DBに記録する（1DB移行版）

v3 からの変更点:
- 保存先を統合DB（<RAW_DB_ID>）に変更。
  プロパティは subject / parents / timestamp / sender / name / content / project /
  source / links / to / cc-bcc の11個に統一（drive は to / cc-bcc 未使用）。
- content(title) にはファイルの要約を入れる（要約できない場合はファイル名）。
  旧 summary プロパティは廃止。
- parents には直親から pjt_ フォルダまでの先祖フォルダ名チェーンを入れる。
- 共有ドライブ「project」内の「slack」フォルダ配下のファイルは集約しない
  （slack 情報集約エージェント側で登録済みのため重複を防ぐ）。
- upsert の match キーを sent_at から links（"ファイル名\\nURL"）に変更。
  slack 側（slack_recorder_notion_v5）と同一フォーマットのため、同じファイルが
  どちらの経路で登録されても同一行に収束する。
"""

import asyncio
import json
import os
import re
import unicodedata
from datetime import datetime, timezone, timedelta

import pypdf
from markitdown import MarkItDown
from agent_sdk import (
    GOOGLEDRIVE_GET_FILE_METADATA,
    GOOGLEDRIVE_FIND_FILE,
    GOOGLEDRIVE_DOWNLOAD_FILE,
    NOTION_UPSERT_ROW_DATABASE,
    llm_call,
)

JST = timezone(timedelta(hours=9))
DB_ID = "<RAW_DB_ID>"  # 統合DB（raw 4DB を1本化したもの）

# 共有ドライブ「project」内の「slack」フォルダ。
# この配下のファイルは slack 情報集約エージェントが登録するため、ここでは集約しない。
SLACK_FOLDER_ID = "<DRIVE_FOLDER_ID_2>"

RETRY_COUNT = 3       # API 呼び出しのリトライ回数（130秒待ちのタイムアウトが何度も重なると
                      # SOFT_DEADLINE_SECONDS を超えて処理全体が打ち切られるため、少なめに抑える）
RETRY_DELAY = 3.0     # リトライ間隔（秒）
FILE_SEMAPHORE = 16   # ファイル並列処理の最大同時実行数（実行環境ブリッジの並列スロット数と同じ＝これが上限）
MAX_TEXT_LENGTH = 8000  # LLMに渡すテキストの最大文字数
SOFT_DEADLINE_SECONDS = 780  # bash_execute/code_execute の同期実行ハードキャップ（900秒、AWS側で
                              # 調整不可）に強制終了される前に自ら打ち切り、正常にログを吐いて終了する
LOOKBACK_HOURS = 7    # 定期実行の間隔（6時間）に対して1時間分オーバーラップさせる。
                      # SOFT_DEADLINE で今回処理しきれなかったファイルも次回の実行で対象に含まれる
                      # （Notion UPSERT は links 一致で重複防止済みのため再処理しても安全）

NOTION_TEXT_LIMIT = 2000  # Notion rich_text / title の1テキスト上限

# markitdown/pypdf では絶対にテキスト化できない形式。ダウンロード自体を試みずに
# 要約をスキップし、無駄なAPI呼び出しと待ち時間を削る（速度最適化）。
# Google Docs/Sheets/Slides のネイティブ形式（application/vnd.google-apps.*）は
# エクスポートすればテキスト化できるため、ここには含めない。
_NON_TEXT_MIME_PREFIXES = ("image/", "video/", "audio/")
_NON_TEXT_MIME_EXACT = frozenset({
    "application/zip",
    "application/x-zip-compressed",
    "application/x-rar-compressed",
    "application/x-7z-compressed",
    "application/x-tar",
    "application/gzip",
})


def _is_non_text_mime(mime_type: str) -> bool:
    return mime_type.startswith(_NON_TEXT_MIME_PREFIXES) or mime_type in _NON_TEXT_MIME_EXACT


SUMMARY_PROMPT = """\
あなたは社内文書要約アシスタントです。
以下の文書を読み、必ずJSON形式のみで返してください。前後の説明は不要です。

出力フォーマット：
{{
  "doc_type": "議事録／提案書／仕様書／契約書／分析資料 など",
  "one_line_summary": "1〜2文で文書の核心",
  "subject": "何について書かれているか（対象・テーマ）",
  "key_points": ["重要ポイント1", "重要ポイント2", "重要ポイント3"],
  "conclusion": "結論・まとめ・推奨事項・決定事項",
  "important_figures": "重要な数値・日付・固有名詞（金額、期限、当事者名など）",
  "notes": "注意点・制約・前提条件など特記事項"
}}

文書内容：
{text}
"""


def _normalize_pjt(name: str) -> str:
    """NFC正規化 + 'pjt_' 直後の半角スペースを吸収。"""
    return unicodedata.normalize("NFC", (name or "")).replace("pjt_ ", "pjt_")


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

    slack 側（slack_recorder_notion_v5）と同一フォーマットにすることで、
    同じファイルがどちらの経路で登録されても links の完全一致 match により
    同一行に upsert される（重複行防止）。

    URLはクエリ文字列（?以降）を除いて正規化する。Google Drive/DocsのwebViewLinkは
    ouid等のクエリパラメータが取得元によって変わり得るため、そのままだと同じファイルでも
    毎回違うmatchキーになり重複行が量産される（実運用で確認済み。slack_recorder_notion_v5と
    同じ正規化をしないと、v4/v5どちらの経路で登録されたかで links が食い違い、
    「同一行に収束する」という設計が成立しなくなる）。
    """
    canonical_url = url.split("?", 1)[0]
    return f"{file_name}\n{canonical_url}"[:NOTION_TEXT_LIMIT]


def _upsert_ok(res: dict) -> tuple[bool, str]:
    """Notion UPSERT の実際の成否を判定する。

    トップレベルの successful/successfull は Composio ツール呼び出し自体が
    実行できたか（＝APIと通信できたか）を示すだけで、実際に行の作成/更新が
    成功したかどうかは data.error_count / data.results[].ok に入っている。
    ここを見ずに top-level フラグだけで判定すると、プロパティ不一致等の
    バリデーションエラーで書き込みが失敗していても「成功」と誤検知する
    （slack_recorder_notion_v5 の実運用で実際に確認済みの不具合と同種）。
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


async def call_with_retry(tool, *, label: str | None = None, validate=None, **kwargs):
    """
    agent_sdk のツール呼び出しをリトライ付きで実行する。
    - 例外（タイムアウト等）と successful=False をリトライ対象にする。
    - validate を渡すと、successful=True でも validate(res) が False の場合はリトライする。
    - リトライを使い切ったら None を返す。
    """
    name = label or getattr(tool, "__name__", "tool")
    last_err = None
    for attempt in range(1, RETRY_COUNT + 1):
        try:
            res = await tool(**kwargs)
        except Exception as e:
            last_err = e
        else:
            # successful/successfull の表記ゆれ両対応（Composio側で混在するため）。
            # 片方だけ見ると実際は失敗している呼び出しを成功と誤判定しかねない。
            if isinstance(res, dict) and (
                res.get("successful") is False or res.get("successfull") is False
            ):
                last_err = res.get("error")
            elif validate is not None and not validate(res):
                last_err = "不完全なレスポンス（期待した要素が見つからない）"
            else:
                return res
        if attempt < RETRY_COUNT:
            await asyncio.sleep(RETRY_DELAY)

    print(f"[RETRY] {name} を {RETRY_COUNT} 回試行しましたが失敗しました: {last_err}")
    return None


async def resolve_ancestors(
    start_folder_id: str,
    cache: dict[str, tuple[str | None, list[str], bool]],
) -> tuple[str | None, list[str], bool]:
    """先祖フォルダを辿り、(pjt_名 | None, 先祖名チェーン, slackフォルダ配下か) を返す。

    - 先祖名チェーンは直親 → … → pjt_ フォルダの順（pjt_ 自身を含む）。
    - 途中で SLACK_FOLDER_ID に到達したら slack フォルダ配下と判定する（集約対象外）。
    - ``cache`` は1回の実行（handle_scheduled_run）内で共有する folder_id → 結果 の辞書。
      同じフォルダに複数ファイルがある実運用では、2件目以降は同じ階層をAPIで
      辿り直さずに即座に解決でき、API呼び出し数を大きく削減できる。
    """
    if start_folder_id in cache:
        return cache[start_folder_id]

    visited: set[str] = set()
    chain_ids: list[str] = []
    chain_names: list[str] = []
    current_id: str | None = start_folder_id
    pjt_name: str | None = None
    under_slack = False

    while current_id and current_id not in visited:
        if current_id == SLACK_FOLDER_ID:
            chain_names.append("slack")
            under_slack = True
            break
        if current_id in cache:
            c_pjt, c_chain, c_slack = cache[current_id]
            pjt_name = c_pjt
            chain_names.extend(c_chain)
            under_slack = c_slack
            break
        visited.add(current_id)
        chain_ids.append(current_id)
        res = await call_with_retry(
            GOOGLEDRIVE_GET_FILE_METADATA,
            label="GET_FILE_METADATA(ancestor)",
            fileId=current_id,
            fields="id,name,parents",
        )
        if not res:
            break
        data = res.get("data") or res
        name = data.get("name") or ""
        chain_names.append(name)
        if name.startswith("pjt_"):
            pjt_name = _normalize_pjt(name)
            break
        parents = data.get("parents") or []
        current_id = parents[0] if parents else None

    # 辿った各フォルダに「そこから見た結果」をキャッシュする
    for i, fid in enumerate(chain_ids):
        cache[fid] = (pjt_name, chain_names[i:], under_slack)
    result = (pjt_name, chain_names, under_slack)
    cache[start_folder_id] = result
    return result


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


async def extract_text(file_id: str, file_name: str) -> str | None:
    """ファイルをダウンロードして markitdown でテキストを抽出する。
    PDF は pypdf フォールバック、docx は python-docx フォールバックあり。
    変換自体は同期(ブロッキング)処理のため asyncio.to_thread でスレッドに逃がし、
    他ファイルの並列処理がイベントループごと止まらないようにする。"""
    dl_res = await call_with_retry(
        GOOGLEDRIVE_DOWNLOAD_FILE,
        label=f"DOWNLOAD_FILE({file_name})",
        fileId=file_id,
    )
    if not dl_res:
        print(f"[WARN] ダウンロード失敗: {file_name}")
        return None

    data = dl_res.get("data") or {}
    content = data.get("downloaded_file_content") or {}
    local_path = content.get("s3url") or ""
    actual_name = content.get("name") or file_name

    if not local_path or not os.path.exists(local_path):
        print(f"[WARN] ローカルパスが見つかりません: {file_name}")
        return None

    name_ext = os.path.splitext(actual_name.lower())[1]
    try:
        text = await asyncio.to_thread(_convert_sync, local_path, name_ext)
    except Exception as e:
        print(f"[WARN] テキスト抽出失敗 ({file_name}): {e}")
        return None

    if not text.strip():
        print(f"[SKIP] テキストが空（スキャン画像等）: {file_name}")
        return None

    return text


async def generate_summary(text: str) -> str | None:
    """テキストからJSON要約を生成する。パース失敗時はテキストをそのまま返す。"""
    truncated = text[:MAX_TEXT_LENGTH]
    prompt = SUMMARY_PROMPT.format(text=truncated)

    try:
        raw = await llm_call(prompt)
    except Exception as e:
        print(f"[WARN] llm_call 失敗: {e}")
        return None

    # llm_call が dict を返す場合に文字列を取り出す
    if isinstance(raw, dict):
        raw = (
            raw.get("content")
            or raw.get("text")
            or (raw.get("data") or {}).get("content")
            or str(raw)
        )

    try:
        json_match = re.search(r'\{.*\}', raw, re.DOTALL)
        if json_match:
            summary_json = json.loads(json_match.group())
            return json.dumps(summary_json, ensure_ascii=False, indent=2)
    except (json.JSONDecodeError, AttributeError):
        pass

    return raw  # パース失敗時はテキストそのまま


async def process_file(
    file_id: str,
    ancestor_cache: dict[str, tuple[str | None, list[str], bool]],
) -> None:
    """ファイルを処理して統合DBに保存する。"""
    # ── STEP 1: メタデータ取得 ──────────────────────────────────────────────
    meta_res = await call_with_retry(
        GOOGLEDRIVE_GET_FILE_METADATA,
        label="GET_FILE_METADATA",
        fileId=file_id,
        fields="id,name,mimeType,webViewLink,modifiedTime,lastModifyingUser,parents,trashed",
    )
    if not meta_res:
        print(f"[SKIP] メタデータ取得失敗: {file_id}")
        return
    meta = meta_res.get("data") or meta_res

    if meta.get("trashed"):
        print(f"[SKIP] ゴミ箱のファイル: {meta.get('name')}")
        return

    # ── STEP 2: 先祖フォルダを辿って pjt_ フィルター + slack フォルダ除外 ────
    parent_folder_id = (meta.get("parents") or [""])[0]
    if not parent_folder_id:
        print(f"[SKIP] 親フォルダが不明: {meta.get('name')}")
        return

    pjt_name, ancestor_chain, under_slack = await resolve_ancestors(
        parent_folder_id, ancestor_cache
    )
    if under_slack:
        print(f"[SKIP] slack フォルダ配下（slack エージェントで登録済み）: {meta.get('name')}")
        return
    if not pjt_name:
        print(f"[SKIP] pjt_ フォルダ外のファイル: {meta.get('name')}")
        return

    # ── STEP 3: 保存者・日時・先祖フォルダ名の整形 ─────────────────────────
    last_modifier  = meta.get("lastModifyingUser") or {}
    uploader       = last_modifier.get("displayName") or "不明"
    uploader_email = last_modifier.get("emailAddress") or ""

    modified_time_str = meta.get("modifiedTime") or ""
    try:
        modified_dt = datetime.fromisoformat(modified_time_str.replace("Z", "+00:00"))
        saved_at = modified_dt.astimezone(JST).strftime("%Y-%m-%dT%H:%M:%S+09:00")
    except Exception:
        saved_at = modified_time_str

    file_name = meta.get("name") or ""
    url       = meta.get("webViewLink") or ""
    # parents: 直親 → … → pjt_ の先祖フォルダ名チェーン
    parents_value = " / ".join(ancestor_chain)

    # ── STEP 4: テキスト抽出 + LLM 要約生成（content 用） ──────────────────
    # 画像/動画/音声/zip等は markitdown で絶対にテキスト化できないため、
    # ダウンロード自体を試みずスキップする（速度最適化）。
    mime_type = meta.get("mimeType") or ""
    summary = None
    if _is_non_text_mime(mime_type):
        print(f"[SKIP-EXTRACT] 非テキスト形式のためDL/要約をスキップ: {file_name} ({mime_type})")
    else:
        text = await extract_text(file_id, file_name)
        if text:
            summary = await generate_summary(text)
            if summary:
                print(f"[SUMMARY] 要約生成完了: {file_name}")
            else:
                print(f"[WARN] 要約生成失敗: {file_name}")

    # ── STEP 5: 統合DBへ UPSERT 保存 ────────────────────────────────────────
    links_value = _file_links_value(file_name, url)
    properties: dict = {
        "content":   _title(summary or file_name),
        "subject":   _rt(file_name),
        "parents":   _rt(parents_value),
        "timestamp": _date(saved_at),
        "sender":    _rt(uploader_email),
        "name":      _rt(uploader),
        "project":   _rt(pjt_name),
        "source":    _select("drive"),
        "links":     _rt(links_value),
    }

    res = await call_with_retry(
        NOTION_UPSERT_ROW_DATABASE,
        label="UPSERT",
        database_id=DB_ID,
        items=[{
            "match":  {"property": "links", "equals": links_value},
            "create": {"properties": properties},
            "update": {"properties": properties},
        }],
    )
    if not res:
        raise RuntimeError(f"[ERROR] Notion 保存失敗（リトライ後も失敗）: {file_name} ({saved_at})")
    ok, err_detail = _upsert_ok(res)
    if not ok:
        raise RuntimeError(f"[ERROR] Notion 保存失敗: {file_name} ({saved_at}) / {err_detail}")
    print(f"[OK] {file_name} → {pjt_name} ({saved_at})")


async def handle_scheduled_run(hours: int = LOOKBACK_HOURS) -> None:
    """
    定期実行エントリーポイント。
    過去 hours 時間以内に作成されたファイルを取得して処理する。
    """
    since = datetime.now(timezone.utc) - timedelta(hours=hours)
    since_str = since.strftime("%Y-%m-%dT%H:%M:%SZ")
    print(f"[SCHEDULED] 過去{hours}時間以内のファイルを取得: since={since_str}")

    all_file_ids: list[str] = []
    page_token = ""
    while True:
        kwargs: dict = {
            "query": (
                f"createdTime > '{since_str}'"
                " and mimeType != 'application/vnd.google-apps.folder'"
                " and trashed = false"
            ),
        }
        if page_token:
            kwargs["page_token"] = page_token

        res = await call_with_retry(GOOGLEDRIVE_FIND_FILE, label="FIND_FILE", **kwargs)
        if not res:
            print("[SCHEDULED] ファイル検索に失敗したため中断します。")
            break
        res_data = res.get("data") or res
        files = res_data.get("files") or []
        all_file_ids.extend(f["id"] for f in files if f.get("id"))
        print(f"[SCHEDULED] {len(files)} 件取得（累計 {len(all_file_ids)} 件）")

        page_token = res_data.get("nextPageToken") or ""
        if not page_token:
            break

    sem = asyncio.Semaphore(FILE_SEMAPHORE)
    completed: list[str] = []
    # 同じフォルダに複数ファイルがある実運用を想定した run 内共有キャッシュ
    # （resolve_ancestors 参照）。
    ancestor_cache: dict[str, tuple[str | None, list[str], bool]] = {}

    async def process_with_sem(file_id: str) -> None:
        async with sem:
            await process_file(file_id, ancestor_cache)
            completed.append(file_id)

    try:
        await asyncio.wait_for(
            asyncio.gather(*[process_with_sem(fid) for fid in all_file_ids], return_exceptions=True),
            timeout=SOFT_DEADLINE_SECONDS,
        )
    except asyncio.TimeoutError:
        remaining = len(all_file_ids) - len(completed)
        print(
            f"[SCHEDULED] 時間切れのため打ち切り: 完了 {len(completed)} 件 / 未処理 {remaining} 件"
            "（未処理分は LOOKBACK_HOURS のオーバーラップにより次回の定期実行で再度対象になります）"
        )

    print(f"[SCHEDULED] 処理結果    : 成功 {len(completed)} 件 / 合計 {len(all_file_ids)} 件")


# code_execute（IPython カーネル）からの実行方法:
#   from rag.google_drive_recorder_v4 import handle_scheduled_run
#   await handle_scheduled_run()
# カーネルは既にイベントループ上で動くため asyncio.run() は使えない。
# 下の __main__ ガードは import 時には発火しない（ローカル実行・デバッグ用）。
if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    asyncio.run(handle_scheduled_run())
