import asyncio
import json
import os
import random
import re
from datetime import datetime, timezone, timedelta
from email.header import decode_header, make_header
from email.utils import getaddresses
from agent_sdk import (
    NOTION_UPSERT_ROW_DATABASE,
    NOTION_QUERY_DATABASE_WITH_FILTER,
    GMAIL_FETCH_MESSAGE_BY_MESSAGE_ID,
    GMAIL_GET_ATTACHMENT,
    GOOGLEDRIVE_FIND_FILE,
    GOOGLEDRIVE_UPLOAD_FILE,
    llm_call,
)

JST = timezone(timedelta(hours=9))

# raw DB ID（固定。全ソース統合DB「社内情報集約PJ > raw」）
RAW_DB_ID         = "<RAW_DB_ID>"
# ドメイン情報DB ID（固定）
DOMAIN_INFO_DB_ID = "<DOMAIN_MAP_DB_ID>"
DOMAIN_PROP       = "domain"
# 人物辞書DB ID（meta > 人。登場人物の自動登録先。空にすると人物登録をスキップ）
PEOPLE_DB_ID = "<PEOPLE_DB_ID>"

# raw DB のプロパティ名（Notion側の列名が変わったらここだけ直す）
PROP_CONTENT    = "content"      # title: 本文＋[添付ファイル] 名前（メール行）／要約JSON（ファイル行）
PROP_SUBJECT    = "subject"      # rich_text: 件名
PROP_TIMESTAMP  = "timestamp"    # date: 送受信日時（JST）
PROP_SENDER     = "sender"       # rich_text: 送信者アドレス
PROP_NAME       = "name"         # rich_text: 送受信者の表示名
PROP_PROJECT    = "project"      # rich_text: プロジェクト名
PROP_SOURCE     = "source"       # select: 情報ソース
PROP_LINKS      = "links"        # rich_text: 添付の「ファイル名\n正規化URL」（ファイル行では match キー）
PROP_TO         = "to"           # rich_text: 宛先アドレス
PROP_CC_BCC     = "cc_bcc"       # rich_text: CC/BCCアドレス
PROP_PARENTS    = "parents"      # rich_text: スレッドID
SOURCE_GMAIL    = "gmail"        # source select の選択肢名（メール行）
SOURCE_DRIVE    = "drive"        # source select の選択肢名（添付ファイル行）

# 共有ドライブ「project」のID（添付の重複チェック用。slack_recorder_v5 と同じ値）
PROJECT_SHARED_DRIVE_ID = "0AINJmYrIeBf9Uk9PVA"
# 添付ファイルの保存先「gmail」フォルダID（プロジェクト > gmail。URLの /folders/ 以降）
GMAIL_DRIVE_FOLDER_ID = "<DRIVE_FOLDER_ID>"

_MAX_RETRIES         = 4
_INITIAL_BACKOFF_SEC = 0.5


class NotionTransientError(Exception):
    """一時的エラー（リトライ上限超過）。"""


class NotionPermanentError(Exception):
    """リトライしても回復しないエラー（許可外ツール・予算超過・認証エラー）。"""


# リトライしても回復しないエラーの識別文言（許可外ツール・予算超過）
_NON_RETRYABLE_MARKERS = ("Tool not allowed", "Budget exceeded")


def _is_non_retryable(err_text: str) -> bool:
    """リトライしても成功し得ないエラーか判定する。"""
    return any(m in err_text for m in _NON_RETRYABLE_MARKERS)


def _is_error_response(response) -> tuple[bool, str | None]:
    """Notionレスポンスがエラーかどうかを判定する。"""
    if not isinstance(response, dict):
        return False, None
    if response.get("successful") is False or response.get("successfull") is False:
        return True, str(response.get("error") or response.get("error_class") or "unknown")
    if response.get("error") and "results" not in response and "data" not in response:
        return True, f"{response.get('error_class') or 'error'}: {response.get('error')}"
    data = response.get("data")
    if isinstance(data, dict):
        if data.get("error") and "results" not in data:
            return True, f"{data.get('error_class') or 'error'}: {data.get('error')}"
    return False, None


async def _call_with_retry(tool, **kwargs):
    """一時的エラーに対して指数バックオフでリトライする。上限超過で例外を送出。"""
    backoff = _INITIAL_BACKOFF_SEC
    last_err: str | Exception | None = None
    for attempt in range(_MAX_RETRIES):
        try:
            response = await tool(**kwargs)
        except Exception as e:
            if _is_non_retryable(str(e)):
                raise NotionPermanentError(f"リトライ不能エラー: {e}") from e
            last_err = e
            if attempt == _MAX_RETRIES - 1:
                break
            await asyncio.sleep(backoff + random.uniform(0, 1))
            backoff *= 2
            continue

        is_err, err_msg = _is_error_response(response)
        if not is_err:
            return response

        if response.get("error_type") == "authentication" or _is_non_retryable(err_msg or ""):
            raise NotionPermanentError(f"リトライ不能エラー: {err_msg}")

        last_err = err_msg
        if attempt == _MAX_RETRIES - 1:
            break
        await asyncio.sleep(backoff + random.uniform(0, 1))
        backoff *= 2

    raise NotionTransientError(
        f"Notion呼び出しが{_MAX_RETRIES}回失敗: tool={getattr(tool, '__name__', tool)}, "
        f"last_error={last_err}"
    )


def _to_jst_str(value) -> str:
    """任意の日時表現を JST の ISO 8601 形式（例: 2026-07-09T13:53:01+09:00）で返す。

    chat/drive DB と日付表記を統一し、検索側 _parse_iso_aware で読める形にする。

    ISO8601 / メールの Date ヘッダ(RFC822) / epoch秒・ミリ秒 / 一般的な
    日付フォーマットを解釈する。空入力のみ "" を返し、それ以外は必ず解釈する。
    """
    if value is None or value == "":
        return ""
    if isinstance(value, datetime):
        dt = value
    elif isinstance(value, (int, float)):
        ts = float(value)
        if ts > 1e12:  # ミリ秒とみなす
            ts /= 1000.0
        dt = datetime.fromtimestamp(ts, tz=timezone.utc)
    else:
        s = str(value).strip()
        dt = None
        if s.isdigit():
            ts = float(s)
            if ts > 1e12:
                ts /= 1000.0
            dt = datetime.fromtimestamp(ts, tz=timezone.utc)
        if dt is None:
            try:
                dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
            except ValueError:
                dt = None
        if dt is None:
            try:
                from email.utils import parsedate_to_datetime
                dt = parsedate_to_datetime(s)
            except (TypeError, ValueError):
                dt = None
        if dt is None:
            for fmt in ("%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S",
                        "%Y/%m/%d %H:%M", "%Y-%m-%d %H:%M",
                        "%Y/%m/%d", "%Y-%m-%d"):
                try:
                    dt = datetime.strptime(s, fmt)
                    break
                except ValueError:
                    continue
        if dt is None:
            raise ValueError(f"日時を解釈できません: {value!r}")
    if dt.tzinfo is None:
        # タイムゾーン無しの日時は JST とみなす
        dt = dt.replace(tzinfo=JST)
    return dt.astimezone(JST).isoformat(timespec="seconds")


def _page_title(page: dict) -> str:
    props = page.get("properties") or {}
    for _, v in props.items():
        if (v or {}).get("type") == "title":
            parts = v.get("title") or []
            return "".join((p.get("plain_text") or "") for p in parts)
    return ""


def _extract_email(s: str) -> str:
    m = re.search(r"[\w.+-]+@[\w-]+\.[\w.-]+", s or "")
    return m.group(0).lower() if m else (s or "").strip().lower()


def _extract_emails(s: str) -> str:
    """文字列中の全メールアドレスを抽出し、カンマ区切りで返す（CC/BCC は複数のため）。"""
    found = re.findall(r"[\w.+-]+@[\w-]+\.[\w.-]+", s or "")
    seen: dict[str, None] = {}
    for a in found:
        seen.setdefault(a.lower(), None)
    return ", ".join(seen.keys())


def _extract_names(*header_strs: str) -> str:
    """From/To ヘッダから表示名を抽出し、カンマ区切りで返す（name プロパティ用）。

    送信者と受信者のヘッダを両方渡す想定。MIMEエンコードされた表示名はデコードし、
    表示名が無いアドレスはメールアドレスをそのまま名前として使う。
    """
    seen: dict[str, None] = {}
    for s in header_strs:
        if not s:
            continue
        for name, addr in getaddresses([s]):
            if name:
                # 表示名はアドレス分解後に個別デコードする
                # （ヘッダ全体を先にデコードすると生UTF-8混在時に文字化けする）
                try:
                    name = str(make_header(decode_header(name)))
                except Exception:
                    pass
            label = (name or "").strip() or (addr or "").strip()
            if label:
                seen.setdefault(label, None)
    return ", ".join(seen.keys())


def _get_header(payload_inner: dict, name: str) -> str:
    """Gmail payload の headers から指定ヘッダ値を取り出す。"""
    headers = (payload_inner or {}).get("headers", []) or []
    name_l = name.lower()
    for h in headers:
        if (h.get("name") or "").lower() == name_l:
            return h.get("value") or ""
    return ""


def _normalize_payload(gmail_payload: dict) -> dict:
    """トリガーの入れ子を吸収し、メール本体(sender等)を持つ dict を返す。

    トリガー全体（payload 配下にメール本体）を渡しても、メール本体そのものを
    渡しても、同じメール本体 dict を返す。これにより CC/BCC の取りこぼしを防ぐ。
    """
    p = gmail_payload or {}
    inner = p.get("payload")
    if isinstance(inner, dict) and any(
        k in inner for k in ("sender", "message_id", "messageId", "message_text")
    ):
        return inner
    return p


def _rt(value: str) -> list[dict]:
    """rich_text/title 用の配列を作る。Notion の 2000字上限に合わせてチャンク分割する。"""
    s = value or ""
    if not s:
        return []
    return [{"text": {"content": s[i:i + 2000]}} for i in range(0, len(s), 2000)]


def _domain_of(email: str) -> str:
    return email.split("@", 1)[1].lower() if "@" in email else email


# 添付要約に使うモデル（PDF読解は Gemini の明示指定が必要）
_SUMMARY_MODEL = "gemini/gemini-3.5-flash"

# ファイル行の content に入れる要約のスキーマ（slack_recorder_v5 ベース。
# url / file_name は links・subject 列と重複するため含めない＝要約のみ）
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
        "doc_type", "one_line_summary", "subject",
        "key_points", "conclusion", "important_figures", "notes"
    ],
}


def _canonical_url(url: str) -> str:
    """URLのクエリ文字列を除いて正規化する（match キーの揺れ防止）。"""
    return (url or "").split("?", 1)[0]


def _file_links_value(file_name: str, url: str) -> str:
    """ファイル行の links 値（＝upsert の match キー）。

    slack_recorder_v5 / drive エージェントと同一フォーマット（"ファイル名\\n正規化URL"）。
    同じファイルがどの経路で登録されても links の完全一致 match により
    同一行に収束する（重複行防止）。
    """
    return f"{file_name}\n{_canonical_url(url)}"[:2000]


def _versioned_name(file_name: str, ts_iso: str) -> str:
    """同名で内容が異なるファイル（更新版）用に、日時サフィックス付きの別名を作る。"""
    stem, ext = os.path.splitext(file_name)
    stamp = ts_iso.replace(":", "").replace("-", "")[:15]  # 例: 20260712T103000
    return f"{stem}_{stamp}{ext}"


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
    """UPLOAD_FILE レスポンスから閲覧URLを取り出す（slack_recorder_v5 と同じ全走査）。

    レスポンスにリンクキーが無く id/kind のみ返るケースが実運用で確認されているため、
    最後は id からURLを組み立てるフォールバックを持つ。
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


def _drive_view_url(file: dict) -> str:
    """FIND_FILE の結果から閲覧URLを得る。webViewLink が無ければ fileId から組み立てる。"""
    url = file.get("webViewLink") or file.get("web_view_link") or ""
    if not url and file.get("id"):
        url = f"https://drive.google.com/file/d/{file['id']}/view"
    return url


def _upsert_failed_detail(result) -> str | None:
    """UPSERT応答の中身まで見て、行の作成/更新が実際に失敗していないか調べる。

    トップレベルの successful はツール呼び出し自体の成否しか示さず、実際の
    行書き込みの成否は data.results[].ok に入る（slack_recorder_v5 の実運用で
    誤検知を確認済み）。失敗があれば詳細文字列、無ければ None を返す。
    """
    if not isinstance(result, dict):
        return "レスポンスがdictでない"
    data = result.get("data") or {}
    results = data.get("results") if isinstance(data, dict) else None
    if results is None:
        return None
    bad = [r for r in results if isinstance(r, dict) and not r.get("ok")]
    if bad:
        return "; ".join(str(r.get("error_message") or r.get("error_code") or r) for r in bad)
    return None


def _upsert_page_info(result) -> dict:
    """UPSERT応答から作成/更新されたページ情報を取り出す。

    ページの id / url はトップレベルの data ではなく data.results[0] に入る
    （2026-07-14 実走で確認。data.get("url") では常に None になる）。
    """
    if not isinstance(result, dict):
        return {}
    data = result.get("data") or {}
    results = data.get("results") if isinstance(data, dict) else None
    if isinstance(results, list) and results and isinstance(results[0], dict):
        return results[0]
    return {}


# ===== 人物辞書DB（people）自動更新 =====
# ※ PEOPLE_DB_ID はファイル冒頭のDB ID定数セクションにある

# 人物辞書DBのプロパティ名
PROP_P_NAME    = "name"     # title: 代表表示名（初回観測時に固定）
PROP_P_EMAIL   = "email"    # rich_text: 主キー（小文字正規化）
PROP_P_ALIASES = "aliases"  # rich_text: 観測した表記ゆれの改行区切り蓄積
PROP_P_COMPANY = "company"  # rich_text: メールの@より後ろの文字列そのまま（判断なしの事実。
                            # 会社名・プロジェクトへの意味づけは手動のドメイン情報DBが担う）

# 人物として登録しない機械アドレス（部分一致）
_PEOPLE_EMAIL_BLOCKLIST = (
    "noreply", "no-reply", "donotreply", "mailer-daemon", "notifications@",
)

# バッチ実行中の人物収集バッファ（書き込みは run() の最後にまとめて行う。
# 16並列の保存中に都度書くと同一人物のマージが競合するため）
_people_buffer: list[dict] = []


def _valid_person_email(email: str) -> str:
    """人物辞書に登録可能なメールなら小文字正規化して返す。不可なら空文字。"""
    e = (email or "").strip().lower()
    if not e or "@" not in e:
        return ""
    if any(b in e for b in _PEOPLE_EMAIL_BLOCKLIST):
        return ""
    return e


def _extract_name_email_pairs(*header_strs: str) -> list[dict]:
    """From/To ヘッダから {name, email} のペアを抽出する（人物辞書の登録用）。

    _extract_names と同じ解析だが、名前⇄メールの対応を保ったまま返す
    （この対応情報は保存の瞬間にしか存在しないため、ここで確保する）。
    """
    pairs: list[dict] = []
    for s in header_strs:
        if not s:
            continue
        for name, addr in getaddresses([s]):
            if name:
                try:
                    name = str(make_header(decode_header(name)))
                except Exception:
                    pass
            if addr:
                pairs.append({"name": (name or "").strip(), "email": addr.strip()})
    return pairs


async def _upsert_people(persons: list[dict]) -> None:
    """登場人物を人物辞書DBに upsert する（name / email / aliases / company の4列）。

    aliases は set マージで増えるだけ＝冪等（再実行しても壊れず、競合しても
    次の登場で自己修復する）。失敗は WARN のみで続行し、本体の保存処理を
    絶対に巻き込まない。PEOPLE_DB_ID 未設定なら何もしない。
    """
    if not PEOPLE_DB_ID:
        return
    # email で集約（同一実行内の重複をまとめ、書き込みを人数分に抑える）
    by_email: dict[str, set] = {}
    for p in persons or []:
        email = _valid_person_email(p.get("email", ""))
        if not email:
            continue
        names = by_email.setdefault(email, set())
        name = (p.get("name") or "").strip()
        if name and name.lower() != email:
            names.add(name)
            # 複合表示名（"担当A / Tomomi Mashiko" 等）は分割した各部分も蓄積する
            # （検索側が contains 検索の起点に使うため、短い単位の方がヒットしやすい）
            for part in re.split(r"[/／|｜]", name):
                part = part.strip()
                if part and part.lower() != email:
                    names.add(part)
    for email, names in by_email.items():
        try:
            resp = await _call_with_retry(
                NOTION_QUERY_DATABASE_WITH_FILTER,
                database_id=PEOPLE_DB_ID,
                filter={"property": PROP_P_EMAIL, "rich_text": {"equals": email}},
                page_size=1,
            )
            data = resp.get("data", resp) or {}
            rows = data.get("results") or []
            if rows:
                props = rows[0].get("properties") or {}
                cur_name = "".join(
                    t.get("plain_text", "")
                    for t in ((props.get(PROP_P_NAME) or {}).get("title") or [])
                )
                cur_aliases = "".join(
                    t.get("plain_text", "")
                    for t in ((props.get(PROP_P_ALIASES) or {}).get("rich_text") or [])
                )
                alias_set = {a.strip() for a in cur_aliases.split("\n") if a.strip()} | names
                display = cur_name or (sorted(names)[0] if names else email)
            else:
                alias_set = set(names)
                display = sorted(names)[0] if names else email
            props_out = {
                PROP_P_NAME:    {"title":     _rt(display)},
                PROP_P_EMAIL:   {"rich_text": _rt(email)},
                PROP_P_ALIASES: {"rich_text": _rt("\n".join(sorted(alias_set)))},
                PROP_P_COMPANY: {"rich_text": _rt(email.split("@", 1)[1])},
            }
            result = await _call_with_retry(
                NOTION_UPSERT_ROW_DATABASE,
                database_id=PEOPLE_DB_ID,
                items=[{
                    "match":  {"property": PROP_P_EMAIL, "equals": email},
                    "create": {"properties": props_out},
                    "update": {"properties": props_out},
                }],
            )
            err = _upsert_failed_detail(result)
            if err:
                print(f"[WARN] 人物辞書の更新失敗: {email} / {err}")
        except Exception as e:
            print(f"[WARN] 人物辞書の更新失敗: {email} / {str(e)[:120]}")


async def _find_existing_in_project_drive(file_name: str) -> dict | None:
    """共有ドライブ「project」全体で同名ファイルを検索する。見つからなければ None。

    「gmail」フォルダ内だけに限定すると、project内の別フォルダに同名ファイルが
    既にある場合の重複を検出できないため、検索対象はドライブ全体にする。
    """
    safe = file_name.replace("\\", "\\\\").replace("'", "\\'")
    q = f"name = '{safe}' and trashed = false"
    try:
        res = await _call_with_retry(
            GOOGLEDRIVE_FIND_FILE, q=q, corpora="drive", driveId=PROJECT_SHARED_DRIVE_ID,
        )
    except Exception as e:
        print(f"[WARN] 重複チェック失敗: {file_name} / {str(e)[:120]}")
        return None
    files = ((res.get("data") or res) or {}).get("files") or []
    return files[0] if files else None


async def _save_attachment_to_drive(local_path: str, file_name: str, mimetype: str,
                                    size: int, ts_iso: str) -> tuple[str, str] | None:
    """添付を「gmail」フォルダに保存し、(保存名, DriveURL) を返す。失敗時は None。

    重複チェックは共有ドライブ「project」全体（slack_recorder_v5 と同じ）:
    - 同名なし → そのままアップロード
    - 同名・サイズ一致（またはサイズ不明）→ 保存スキップ・既存リンクを再利用
    - 同名・サイズ不一致（更新版）→ 日時サフィックス付き別名でアップロード
    """
    upload_name = file_name
    existing = await _find_existing_in_project_drive(file_name)
    if existing:
        try:
            existing_size = int(existing.get("size") or 0)
        except (TypeError, ValueError):
            existing_size = 0
        if not size or not existing_size or size == existing_size:
            url = _drive_view_url(existing)
            if url:
                print(f"[DUP] 同名ファイルが既に存在するため保存スキップ: {file_name}")
                return file_name, url
        else:
            upload_name = _versioned_name(file_name, ts_iso)
            print(f"[VER] 同名・内容違いのため別名で保存: {upload_name}")

    if not GMAIL_DRIVE_FOLDER_ID:
        print("[WARN] GMAIL_DRIVE_FOLDER_ID が未設定のため添付を保存できません")
        return None
    try:
        res = await _call_with_retry(
            GOOGLEDRIVE_UPLOAD_FILE,
            file_to_upload={
                "name":     upload_name,
                "mimetype": mimetype or "application/octet-stream",
                "s3key":    local_path,
            },
            folder_to_upload_to=GMAIL_DRIVE_FOLDER_ID,
        )
    except Exception as e:
        print(f"[WARN] Drive アップロード失敗: {upload_name} / {str(e)[:200]}")
        return None
    url = _extract_drive_link(res)
    if not url:
        print(f"[WARN] アップロード成功したがURLを抽出できません: {upload_name}")
        return None
    return upload_name, url


# テキスト抽出経由で LLM に渡す最大文字数（slack_recorder_v5 と同値）
_MAX_TEXT_LENGTH = 8000


def _extract_pptx_text(path: str) -> str:
    """pptx をテキスト化する（llm_call が pptx 添付に非対応のため markitdown 経由）。"""
    from markitdown import MarkItDown
    result = MarkItDown().convert(path)
    return result.text_content or ""


async def _summarize_attachment_file(local_path: str, file_name: str) -> dict | None:
    """添付1件を llm_call（スキーマ指定）で要約JSONにする。

    content に入れるのは要約のみ（URL・ファイル名は links / subject 列にあるため
    JSONには含めない）。pptx は llm_call の添付に非対応（fail closed）のため、
    markitdown でテキスト抽出してからテキストとして要約する（slack_recorder_v5 と同じ方式）。
    """
    prompt = (
        f"添付ファイル「{file_name}」の内容を読み取り、指定のJSONスキーマに従って"
        f"日本語で要約してください。"
    )
    kwargs: dict = {"model": _SUMMARY_MODEL, "schema": FILE_SUMMARY_SCHEMA}
    if file_name.lower().endswith(".pptx"):
        try:
            text = await asyncio.to_thread(_extract_pptx_text, local_path)
        except Exception as e:
            print(f"[WARN] pptx テキスト抽出失敗: {file_name} / {str(e)[:120]}")
            return None
        if not text.strip():
            print(f"[WARN] pptx からテキストを抽出できず要約スキップ: {file_name}")
            return None
        kwargs["prompt"] = f"{prompt}\n\n【抽出テキスト】\n{text[:_MAX_TEXT_LENGTH]}"
    else:
        kwargs["prompt"] = prompt
        kwargs["file_paths"] = [local_path]
    try:
        res = await llm_call(**kwargs)
        data = res.get("data") if isinstance(res, dict) else None
        return data if isinstance(data, dict) else None
    except Exception as e:
        print(f"[WARN] 添付要約失敗: {file_name} / {str(e)[:120]}")
        return None


async def _upsert_drive_file_row(*, file_name: str, drive_url: str, ts_iso: str,
                                 sender: str, names: str, project: str,
                                 summary: dict | None) -> bool:
    """添付ファイルの行（source=drive）を raw DB に upsert する。

    match キーは links（"ファイル名\\n正規化URL"）。slack/drive エージェントと
    同一書式のため、同じファイルがどの経路で登録されても同一行に収束する。
    content には要約JSON（要約できない形式はファイル名）を入れる。
    """
    links_value = _file_links_value(file_name, drive_url)
    content_value = json.dumps(summary, ensure_ascii=False, indent=2) if summary else file_name
    props = {
        PROP_CONTENT: {"title":     _rt(content_value)},
        PROP_SUBJECT: {"rich_text": _rt(file_name)},
        PROP_PARENTS: {"rich_text": _rt("gmail")},
        PROP_SENDER:  {"rich_text": _rt(sender)},
        PROP_NAME:    {"rich_text": _rt(names)},
        PROP_PROJECT: {"rich_text": _rt(project)},
        PROP_SOURCE:  {"select":    {"name": SOURCE_DRIVE}},
        PROP_LINKS:   {"rich_text": _rt(links_value)},
    }
    if ts_iso:
        props[PROP_TIMESTAMP] = {"date": {"start": ts_iso}}
    try:
        result = await _call_with_retry(
            NOTION_UPSERT_ROW_DATABASE,
            database_id=RAW_DB_ID,
            items=[{
                "match":  {"property": PROP_LINKS, "equals": links_value},
                "create": {"properties": props},
                "update": {"properties": props},
            }],
        )
    except Exception as e:
        print(f"[WARN] ファイル行 upsert 失敗: {file_name} / {str(e)[:120]}")
        return False
    err = _upsert_failed_detail(result)
    if err:
        print(f"[WARN] ファイル行 upsert がNotion側で失敗: {file_name} / {err}")
        return False
    return True


# これ未満のサイズの画像添付は署名・HTMLメールのインライン画像とみなして除外する
_MIN_IMAGE_BYTES = 20_000

# 並列バッチでの二重アップロード防止（同名チェック→アップロードを直列化する）
_drive_save_lock = asyncio.Lock()


async def _process_attachments(message_id: str, attachments: list[dict], *,
                               ts_iso: str, sender: str, names: str,
                               project: str) -> tuple[list[str], list[str]]:
    """添付を処理し、(添付ファイル名リスト, links用エントリリスト) を返す。

    各添付について slack_recorder_v5 と同じ流れを踏む:
    1. GMAIL_GET_ATTACHMENT でダウンロード
    2. 共有ドライブ「project」全体の重複チェック → 「gmail」フォルダへ実体保存
    3. llm_call（スキーマ）で要約し、source=drive のファイル行を upsert
    失敗した添付はスキップして続行する（メール本体の保存は止めない）。
    """
    file_names: list[str] = []
    links_parts: list[str] = []
    for att in attachments:
        name   = att.get("filename") or ""
        att_id = att.get("attachmentId") or att.get("attachment_id") or ""
        mime   = (att.get("mimeType") or "").lower()
        if not name or not att_id:
            continue
        file_names.append(name)
        try:
            resp = await _call_with_retry(
                GMAIL_GET_ATTACHMENT,
                message_id=message_id, attachment_id=att_id, file_name=name,
            )
            data = resp.get("data", resp) or {}
            s3url = ((data.get("file") or {}).get("s3url")) or ""
        except Exception as e:
            print(f"[WARN] 添付ダウンロード失敗: {name} / {str(e)[:120]}")
            continue
        local_path = s3url
        if s3url.startswith("http"):
            # s3url がURL形式の場合はローカルへ落としてから使う
            import requests
            os.makedirs("tmp", exist_ok=True)
            local_path = f"tmp/att_{message_id}_{name}"
            try:
                r = requests.get(s3url, timeout=60)
                r.raise_for_status()
                with open(local_path, "wb") as f:
                    f.write(r.content)
            except Exception as e:
                print(f"[WARN] 添付取得失敗: {name} / {str(e)[:120]}")
                continue
        elif not local_path:
            # s3url 欠落時は既定の保存先を推定する
            local_path = f"uploads/{name}"
        if not os.path.exists(local_path):
            print(f"[WARN] 添付のローカルパスが見つからない: {name} / {local_path}")
            continue
        try:
            size = int(att.get("size") or 0) or os.path.getsize(local_path)
        except OSError:
            size = int(att.get("size") or 0)

        if mime.startswith("image/") and size and size < _MIN_IMAGE_BYTES:
            # 署名・インライン画像とみなしスキップ（Drive/DBの汚染防止）
            file_names.pop()
            continue

        # 並列実行時に同名添付が同時に重複チェックを通過して
        # 二重アップロードされるのを防ぐ（過去情報バッチは16並列）
        async with _drive_save_lock:
            saved = await _save_attachment_to_drive(local_path, name, mime, size, ts_iso)
        if not saved:
            continue
        saved_name, drive_url = saved
        if saved_name != name:
            file_names[-1] = saved_name

        summary = await _summarize_attachment_file(local_path, saved_name)
        await _upsert_drive_file_row(
            file_name=saved_name, drive_url=drive_url, ts_iso=ts_iso,
            sender=sender, names=names, project=project, summary=summary,
        )
        links_parts.append(_file_links_value(saved_name, drive_url))
    return file_names, links_parts


async def _get_attachment_list(gp: dict, payload_inner: dict, message_id: str) -> list[dict]:
    """添付一覧（attachmentList）を取得する。payload に無ければ full 取得で補完する。"""
    for src in (gp, payload_inner):
        if isinstance(src, dict) and "attachmentList" in src:
            return [a for a in (src.get("attachmentList") or []) if isinstance(a, dict)]
    if not message_id:
        return []
    try:
        detail = await _call_with_retry(
            GMAIL_FETCH_MESSAGE_BY_MESSAGE_ID, message_id=message_id, format="full"
        )
        body = detail.get("data", detail) or {}
        return [a for a in (body.get("attachmentList") or []) if isinstance(a, dict)]
    except Exception:
        return []



def _build_raw_props(*, content: str, subject: str, project: str, sender: str,
                     names: str, recipient: str, cc_bcc: str, links: str,
                     timestamp_jst: str, thread_id: str) -> dict:
    """raw DB 用のプロパティ dict を組み立てる（Notion 生API形式）。"""
    props = {
        PROP_CONTENT:    {"title":     _rt(content)},
        PROP_SUBJECT:    {"rich_text": _rt(subject)},
        PROP_PROJECT:    {"rich_text": _rt(project)},
        PROP_SENDER:     {"rich_text": _rt(sender)},
        PROP_NAME:       {"rich_text": _rt(names)},
        PROP_TO:         {"rich_text": _rt(recipient)},
        PROP_CC_BCC:     {"rich_text": _rt(cc_bcc)},
        PROP_SOURCE:     {"select":    {"name": SOURCE_GMAIL}},
        PROP_LINKS:      {"rich_text": _rt(links)},
        PROP_PARENTS:    {"rich_text": _rt(thread_id)},
    }
    if timestamp_jst:
        # 空文字を date に入れるとエラーになるため、日時がある場合のみ付与する
        props[PROP_TIMESTAMP] = {"date": {"start": timestamp_jst}}
    return props


async def _get_project(domain: str) -> str | None:
    """ドメインでドメイン情報DBを絞り込み、プロジェクト名を返す。未登録ドメインなら None。

    タイトルは pjt_ prefix を剥がさずそのまま使う（chat/drive DB と表記統一）。
    """
    resp = await _call_with_retry(
        NOTION_QUERY_DATABASE_WITH_FILTER,
        database_id=DOMAIN_INFO_DB_ID,
        filter={"property": DOMAIN_PROP, "rich_text": {"equals": domain.lower()}},
        page_size=1,
    )
    data = resp.get("data", resp) or {}
    results = data.get("results") or []
    if not results:
        return None
    company_name = _page_title(results[0]).strip()
    return company_name or domain


async def save_chat_to_notion(gmail_payload: dict) -> dict:
    gp = _normalize_payload(gmail_payload)
    payload_inner = gp.get("payload") or {}
    message_id   = gp.get("messageId") or gp.get("message_id") or gp.get("id") or ""
    subject      = gp.get("subject", "")
    sender_raw   = gp.get("sender", "")
    recipient    = _extract_emails(gp.get("to", ""))  # 複数宛先を全て保存
    # name には送信者・受信者両方の表示名を入れる
    names        = _extract_names(sender_raw, gp.get("to", ""))
    received_at  = _to_jst_str(gp.get("message_timestamp", ""))
    message_body = gp.get("message_text", "")
    # スレッド管理用: 同じ threadId を持つ行＝同じ会話
    thread_id    = gp.get("threadId") or gp.get("thread_id") or ""

    # CC / BCC は1つのプロパティにまとめる（トップレベル優先、無ければ payload ヘッダ）
    cc  = gp.get("cc")  or gp.get("Cc")  or _get_header(payload_inner, "Cc")
    bcc = gp.get("bcc") or gp.get("Bcc") or _get_header(payload_inner, "Bcc")
    cc_bcc = _extract_emails(f"{cc or ''} {bcc or ''}")

    sender       = _extract_email(sender_raw)
    project = await _get_project(_domain_of(sender))
    if project is None:
        return {"status": "skipped",
                "reason": f"未登録ドメインのため保存しない: {_domain_of(sender)}",
                "message_id": message_id}

    if not message_id:
        return {"status": "skipped", "reason": "message_id が空のため添付取得・原文特定が不能", "message_id": ""}

    # 添付は slack_recorder_v5 と同じ扱い: Drive「gmail」フォルダに実体保存し、
    # source=drive のファイル行を別行で upsert する。メール行には
    # content=本文＋添付ファイル名 / links=「ファイル名\n正規化URL」を入れる
    attachment_list = await _get_attachment_list(gp, payload_inner, message_id)
    att_names, links_parts = await _process_attachments(
        message_id, attachment_list,
        ts_iso=received_at, sender=sender, names=names, project=project,
    )
    content = message_body or subject  # 本文が空のメールは件名を content に立てる
    if att_names:
        content = f"{content}\n[添付ファイル] {', '.join(att_names)}".strip()
    # 末尾に [mail:メールID] の行を付ける（重複判定キー。IDはメール毎に一意なので
    # 長文本文を filter に渡す必要がなく、本文が同一の別メールとも混ざらない）
    content = f"{content}\n[mail:{message_id}]".strip()

    props = _build_raw_props(
        content=content, subject=subject, project=project, sender=sender,
        names=names, recipient=recipient, cc_bcc=cc_bcc,
        links="\n\n".join(links_parts),
        timestamp_jst=received_at, thread_id=thread_id,
    )

    # メール行の重複防止は content 末尾に埋めた [mail:ID] 行の contains 一致
    # （再処理は既存行の上書きで済む。ファイル行は [mail:] を含まないため誤マッチしない）
    result = await _call_with_retry(
        NOTION_UPSERT_ROW_DATABASE,
        database_id=RAW_DB_ID,
        items=[{
            "match":  {"property": PROP_CONTENT,
                       "title": {"contains": f"[mail:{message_id}]"}},
            "create": {"properties": props},
            "update": {"properties": props},
        }],
    )
    err = _upsert_failed_detail(result)
    if err:
        return {"status": "error", "reason": f"Notion書き込み失敗: {err}",
                "message_id": message_id}
    # 人物辞書の自動更新（送受信者。失敗しても保存結果には影響しない）
    await _upsert_people(_extract_name_email_pairs(sender_raw, gp.get("to", "")))

    return {
        "status": "success",
        "id": _upsert_page_info(result).get("id"),
        "message_id": message_id,
    }
