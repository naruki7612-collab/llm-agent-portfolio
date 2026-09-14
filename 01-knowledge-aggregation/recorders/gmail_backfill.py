"""Gmail から各プロジェクトのメールを取得し、raw DB に保存する（バックフィル）。

★2026-07-25 改修（実測に基づく3点）:
  1. **hydrate（1通ずつの詳細取得）を廃止**。include_payload=True の一覧応答に
     本文（messageText）と全ヘッダ（payload.headers の To/From/Cc/Bcc/Date）が入っている
     ことを確認したため。取得が 1.3件/秒 → 5.7件/秒 に改善（旧コメントの前提が誤りだった）。
  2. **バッチ upsert**（UPSERT_BATCH=50件/回）。100件の保存が 88秒 → 50秒 に短縮。
  3. **時間予算＋チェックポイント方式**。code_execute の実行窓は約250〜300秒しかないため、
     1ページ（既定 PAGE_SIZE 件）ずつ「取得→保存→再開位置を保存」を繰り返し、
     予算が尽きたら自分で止まる。
     → 「→ 続きあり」が出たら**同じ呼び出しを「✅ 全完了」まで繰り返す**だけ。
  4. **応答サイズが大きすぎる場合はページサイズを自動で縮める**（2026-07-29 nanocosme社で
     実測。本文が重いメールが多い会社は既定件数でも一括で返せないことがある）。
     PAGE_SIZE の既定値も25件に下げてあり（本文が重い会社でも通りやすくするため）、
     それでも大きすぎれば _fetch_page_adaptive がさらに半分ずつ縮めて再試行する。

  再開位置は2段構え（tmp/ が消えても止まらない）:
    ① tmp/gmail_bf_*.json の pageToken（同一セッション内＝正確に続き）
    ② tmp/ が無い（セッション30分TTLで消えた）→ raw DB の**保存済み最古 timestamp** から
       `before:` を付けて再開（取得は新しい順なので最古＝取り込みのフロンティア）

- ドメイン情報DB（domain_map）からプロジェクト⇄ドメインを取得
- 会社ごとに Gmail を (from/to/cc/bcc:domain) で検索（★期間指定は無く全期間）
- 保存先は全ソース統合の raw DB（source=gmail）
- 添付は Drive「gmail」フォルダに実体保存し、source=drive のファイル行を別行で upsert
  （slack_recorder_v5 と同じ設計）
- 重複防止: メール行は content 末尾の [mail:ID] の contains 一致 / ファイル行は links 完全一致
  → 再実行は既存行の上書きで済む＝**タイムアウトしても損失なし**
- エントリポイント: run() / run(projects=["A社","B社"]) / run(projects=[...], budget_seconds=...)
"""
import asyncio
import hashlib
import json
import logging
import os
import random
import re
import time
from datetime import datetime, timezone, timedelta
from email.header import decode_header, make_header
from email.utils import getaddresses

from agent_sdk import (
    GMAIL_FETCH_EMAILS,
    GMAIL_GET_ATTACHMENT,
    GOOGLEDRIVE_FIND_FILE,
    GOOGLEDRIVE_UPLOAD_FILE,
    NOTION_UPSERT_ROW_DATABASE,
    NOTION_QUERY_DATABASE_WITH_FILTER,
    llm_call,
)

# ========== 設定 ==========

logger = logging.getLogger(__name__)

# ★2026-07-29: 既定を100→25に変更。100件だと本文が重い会社（nanocosme社で実測）で
#   「応答が大きすぎる」エラーになり、毎回100から縮小を試すのは無駄なため、
#   最初から通りやすいサイズで始める。それでも大きすぎる場合は _fetch_page_adaptive が
#   さらに自動で縮める（この値はその安全弁の初期値という位置づけ）。
PAGE_SIZE = 25

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

JST = timezone(timedelta(hours=9))

_MAX_RETRIES         = 4
_INITIAL_BACKOFF_SEC = 0.5
# 同時実行数。実行環境実行スロット上限=16（2026-07-12実測: Gmail 187件・Notion 100件ともエラー0）
_GMAIL_CONCURRENCY  = 16
_NOTION_CONCURRENCY = 16

# ===== 時間予算（code_execute の実行窓 約300秒に対する安全圏。2026-07-25 実測で確定） =====
# 実測: 250秒予算→実際314秒で成功 / 385秒相当・850秒予算→ReadTimeoutError。
# ★ただしタイムアウトしてもサンドボックス側の処理は進み、ページ単位のチェックポイントは残る
#   （850秒runで1100件保存済みを確認）。つまり「打ち切って途中から再実行」で損失は出ない。
TIME_BUDGET_SECONDS = 250
# 1ページの想定所要秒。100件時代の実測（fetch17＋save57＝74秒、マージン込み110秒）を
# 25件ぶんに比例配分した見積もり（固定オーバーヘッド分を考慮し切り上げ）。
# ★25件での実測値ではないため、実行後に read_log() の page_done の間隔を見て
#   必要なら調整すること。残り時間がこれ未満なら次のページに入らない
PAGE_COST_EST = 40
UPSERT_BATCH = 50         # 1回の upsert に詰める行数（実測: 100件の保存が88秒→50秒に短縮）
_UPSERT_CONCURRENCY = 2   # バッチ呼び出しの並列数（Notionのレート制限 約3req/s に配慮）

# リトライしても回復しないエラーの識別文言（許可外ツール・予算超過・
# 応答サイズ超過＝同じ引数のまま再試行しても必ず同じ結果になるもの）
_NON_RETRYABLE_MARKERS = ("Tool not allowed", "Budget exceeded", "payload is too large")


class ToolTransientError(Exception):
    """一時的エラー（リトライ上限超過）。"""


class ToolPermanentError(Exception):
    """同じ引数でのリトライでは回復しないエラー（許可外ツール・予算超過・
    認証エラー・応答サイズ超過など）。呼び出し元が引数を変えて再試行する想定。"""


# ===== デバッグ用 JSONL ログ =====
# ★ReadTimeoutError やセッション破棄が起きると stdout は丸ごと失われる。
#   このファイルはイベントごとに追記＋即フラッシュするので、落ちた後でも
#   「どこまで何をやったか」を read_log() で追跡できる（Zoom版と同じ仕組み）。
DEBUG_LOG_PATH = "tmp/logs/gmail_過去情報.jsonl"
DEBUG_LOG = True
_run_id = ""


def _log(event: str, msg: str = "", **fields) -> None:
    """1行=1イベントのJSONLを追記し、同時に人が読む行を print する。

    ログの書き込み失敗は握りつぶす（デバッグ用の仕組みが本体を止めては本末転倒）。
    """
    if msg:
        print(msg)
    if not DEBUG_LOG:
        return
    rec = {"ts": datetime.now(JST).isoformat(timespec="seconds"),
           "run": _run_id, "event": event, **fields}
    try:
        os.makedirs(os.path.dirname(DEBUG_LOG_PATH), exist_ok=True)
        with open(DEBUG_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
            f.flush()
    except Exception:
        pass


def read_log(last: int = 40, event: str | None = None) -> list[dict]:
    """JSONLログの末尾を読む（タイムアウト後の原因追跡用）。

    使い方: read_log(60) で直近60イベント / read_log(20, "company_error") で失敗だけ。
    """
    try:
        with open(DEBUG_LOG_PATH, encoding="utf-8") as f:
            lines = f.readlines()
    except OSError:
        return []
    out: list[dict] = []
    for line in lines:
        try:
            rec = json.loads(line)
        except Exception:
            continue
        if event and rec.get("event") != event:
            continue
        out.append(rec)
    return out[-last:]


def _is_non_retryable(err_text: str) -> bool:
    """リトライしても成功し得ないエラーか判定する。"""
    return any(m in err_text for m in _NON_RETRYABLE_MARKERS)


def _is_error_response(response) -> tuple[bool, str | None]:
    """ツール応答が失敗封筒かどうかを判定する。

    Composio の失敗は例外にならず {"successful": False, ...} の dict で
    素通しされるため、dict の中身を見て判定する必要がある。
    """
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
    """一時的エラーに対して指数バックオフでリトライする。

    許可外ツール・予算超過・認証エラーはリトライ不能として即時打ち切る。
    上限超過で ToolTransientError を送出。
    """
    backoff = _INITIAL_BACKOFF_SEC
    last_err: str | Exception | None = None
    for attempt in range(_MAX_RETRIES):
        try:
            response = await tool(**kwargs)
        except Exception as e:
            if _is_non_retryable(str(e)):
                raise ToolPermanentError(f"リトライ不能エラー: {e}") from e
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
            raise ToolPermanentError(f"リトライ不能エラー: {err_msg}")

        last_err = err_msg
        if attempt == _MAX_RETRIES - 1:
            break
        await asyncio.sleep(backoff + random.uniform(0, 1))
        backoff *= 2

    raise ToolTransientError(
        f"ツール呼び出しが{_MAX_RETRIES}回失敗: tool={getattr(tool, '__name__', tool)}, "
        f"last_error={last_err}"
    )


# ========== ヘルパー ==========

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
    # 重複を除きつつ順序を保持
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


# この実行で人物辞書に書き終えた email → 既に書いた表示名の集合。
# ★run の最後にまとめて書くと、途中でランタイムが落ちた（ReadTimeout/セッション破棄）
#   時にその実行で集めた人物が全部失われる。次回はチェックポイントの先から進むため、
#   その範囲の人物は二度と登録されない。そこで「ページごとに差分だけ書く」ためのメモ。
_people_written: dict[str, set[str]] = {}


def _valid_person_email(email: str) -> str:
    """人物辞書に登録可能なメールなら小文字正規化して返す。不可なら空文字。"""
    e = (email or "").strip().lower()
    if not e or "@" not in e:
        return ""
    if any(b in e for b in _PEOPLE_EMAIL_BLOCKLIST):
        return ""
    return e


async def _flush_people() -> None:
    """人物辞書バッファのうち「まだ書いていない分」だけを書き、バッファを空にする。

    ページ単位で呼ぶことで、途中でランタイムが落ちても人物を取りこぼさない。
    1社に登場する顔ぶれはほぼ同じなので、2ページ目以降の新規はごくわずかで済む
    （コストが先頭ページに集まるだけ）。新しい表記（別名）が出た人は再度書いて
    aliases に追記する。失敗は _upsert_people 側で WARN のみ＝本体を巻き込まない。
    """
    if not _people_buffer:
        return
    if not PEOPLE_DB_ID:
        _people_buffer.clear()
        return

    todo: list[dict] = []
    for p in _people_buffer:
        email = _valid_person_email(p.get("email", ""))
        if not email:
            continue
        name = (p.get("name") or "").strip()
        known = _people_written.get(email)
        if known is not None and (not name or name in known):
            continue  # 既に書いた人で、新しい表記も無い
        todo.append(p)
    _people_buffer.clear()
    if not todo:
        return

    await _upsert_people(todo)
    for p in todo:
        email = _valid_person_email(p.get("email", ""))
        if not email:
            continue
        names = _people_written.setdefault(email, set())
        name = (p.get("name") or "").strip()
        if name:
            names.add(name)


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


_PEOPLE_QUERY_CHUNK = 50   # OR条件・バッチupsertともにこの件数ずつまとめる
_PEOPLE_WRITE_CONCURRENCY = 2  # 書き込みチャンクの並列数（Notionのレート制限に配慮）


async def _upsert_people(persons: list[dict]) -> None:
    """登場人物を人物辞書DBに upsert する（name / email / aliases / company の4列）。

    ★2026-07-25 バッチ化: 以前は1人ずつ「読み取り→書き込み」を直列でやっており、
      新規人物が18人いれば最大36回の逐次Notion呼び出しになっていた（メール行を
      50件ずつバッチupsertした改善が人物辞書だけ未適用だったため）。
      読み取りは全員ぶんを1回のOR検索でまとめ、書き込みは _save_items と同じ
      items=[...] のバッチupsertにする（件数が多い場合は _PEOPLE_QUERY_CHUNK 件ずつ）。

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
    if not by_email:
        return
    emails = list(by_email.keys())

    # ---- 読み取り：対象メール全員ぶんを1回のOR検索でまとめて引く（chunkごと） ----
    existing: dict[str, dict] = {}
    for i in range(0, len(emails), _PEOPLE_QUERY_CHUNK):
        chunk_emails = emails[i:i + _PEOPLE_QUERY_CHUNK]
        cursor: str | None = None
        while True:
            kwargs: dict = {
                "database_id": PEOPLE_DB_ID,
                "filter": {"or": [{"property": PROP_P_EMAIL, "rich_text": {"equals": e}}
                                  for e in chunk_emails]},
                "page_size": 100,
            }
            if cursor:
                kwargs["start_cursor"] = cursor
            try:
                resp = await _call_with_retry(NOTION_QUERY_DATABASE_WITH_FILTER, **kwargs)
            except Exception as e:
                print(f"[WARN] 人物辞書の一括読み取りに失敗（{len(chunk_emails)}件）: {e}")
                break
            data = resp.get("data", resp) or {}
            for row in data.get("results") or []:
                props = row.get("properties") or {}
                row_email = "".join(
                    t.get("plain_text", "")
                    for t in ((props.get(PROP_P_EMAIL) or {}).get("rich_text") or [])
                ).strip().lower()
                if row_email:
                    existing[row_email] = props
            if not data.get("has_more"):
                break
            cursor = data.get("next_cursor")
            if not cursor:
                break

    # ---- 書き込み用の items を組み立てる（Python側でマージ） ----
    items: list[dict] = []
    for email, names in by_email.items():
        props = existing.get(email)
        if props:
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
        items.append({
            "match":  {"property": PROP_P_EMAIL, "equals": email},
            "create": {"properties": props_out},
            "update": {"properties": props_out},
        })

    # ---- バッチ書き込み（_save_items と同じ形：チャンク分割＋並列） ----
    chunks = [items[i:i + _PEOPLE_QUERY_CHUNK] for i in range(0, len(items), _PEOPLE_QUERY_CHUNK)]
    sem = asyncio.Semaphore(_PEOPLE_WRITE_CONCURRENCY)

    async def _write_chunk(chunk: list[dict]) -> None:
        async with sem:
            try:
                result = await _call_with_retry(
                    NOTION_UPSERT_ROW_DATABASE, database_id=PEOPLE_DB_ID, items=chunk
                )
            except Exception as e:
                print(f"[WARN] 人物辞書の一括更新に失敗（{len(chunk)}件）: {e}")
                return
            err = _upsert_failed_detail(result)
            if err:
                print(f"[WARN] 人物辞書の一括更新でNotion側の一部失敗: {err}")

    await asyncio.gather(*(_write_chunk(c) for c in chunks))


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


def _download_url_to_file(url: str, dest_path: str) -> None:
    """URLの内容をファイルに書き出す（同期・to_thread 経由で呼ぶこと）。

    requests は同期I/Oのため、async関数内で直接 await 無しに呼ぶと
    イベントループ全体をブロックする（実測: 5並列が7.5秒→1.5秒に短縮するのを確認）。
    """
    import requests
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    with open(dest_path, "wb") as f:
        f.write(r.content)


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
            os.makedirs("tmp", exist_ok=True)
            local_path = f"tmp/att_{message_id}_{name}"
            try:
                # ★requests は同期I/O。await 無しで直接呼ぶとイベントループを
                #   ブロックし、_build_item を並列gatherしている他の全メールの
                #   処理が止まる（実測で確認済み）。to_thread で別スレッドに逃がす。
                await asyncio.to_thread(_download_url_to_file, s3url, local_path)
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


# ========== Notionナビゲーション ==========

async def _query_all(database_id: str) -> list[dict]:
    """指定DBを全件クエリして返す。"""
    rows: list[dict] = []
    cursor: str | None = None
    while True:
        kwargs = {"database_id": database_id, "page_size": 100}
        if cursor:
            kwargs["start_cursor"] = cursor
        resp = await _call_with_retry(NOTION_QUERY_DATABASE_WITH_FILTER, **kwargs)
        data = resp.get("data", resp) or {}
        rows.extend(data.get("results") or [])
        if not data.get("has_more"):
            break
        cursor = data.get("next_cursor")
        if not cursor:
            break
    return rows


async def _load_projects() -> list[dict]:
    """ドメイン情報DBから [{project, domain}, ...] を返す。

    プロジェクト名は pjt_ prefix を剥がさずタイトルのまま使う（chat/drive DB と表記統一）。
    """
    projects: list[dict] = []
    for row in await _query_all(DOMAIN_INFO_DB_ID):
        company_name = _page_title(row).strip()
        props = row.get("properties") or {}
        domain_prop = props.get(DOMAIN_PROP) or {}
        domain = "".join(x.get("plain_text", "") for x in (domain_prop.get("rich_text") or []))
        domain = domain.strip().lstrip("@").lower()
        if company_name and domain:
            projects.append({"project": company_name, "domain": domain})
    return projects


# ========== Gmail取得部 ==========

def _pick(msg: dict) -> dict:
    payload_inner = msg.get("payload") or {}
    # CC / BCC はトップレベルフィールド優先、無ければ payload ヘッダから補完
    cc  = msg.get("cc")  or msg.get("Cc")  or _get_header(payload_inner, "Cc")
    bcc = msg.get("bcc") or msg.get("Bcc") or _get_header(payload_inner, "Bcc")
    return {
        "message_id":        msg.get("messageId") or msg.get("message_id") or msg.get("id") or "",
        "subject":           msg.get("subject") or "",
        "message_text":      msg.get("messageText") or msg.get("message_text") or "",
        "sender":            msg.get("sender") or "",
        "to":                msg.get("to") or "",
        "cc":                cc or "",
        "bcc":               bcc or "",
        "message_timestamp": msg.get("messageTimestamp") or msg.get("message_timestamp") or "",
        # スレッド管理用: 同じ threadId を持つ行＝同じ会話
        "thread_id":         msg.get("threadId") or msg.get("thread_id") or "",
        # 添付一覧（保存時に _process_attachments で要約とURL収集をする）
        "attachment_list":   [a for a in (msg.get("attachmentList") or []) if isinstance(a, dict)],
    }


# ★2026-07-25 hydrate（1通ずつ GMAIL_FETCH_MESSAGE_BY_MESSAGE_ID で詳細取得）は廃止した。
#   include_payload=True の一覧応答に messageText（本文）と payload.headers（To/From/Cc/Bcc/Date）が
#   すべて入っていることを実測で確認したため（旧コメント「一覧応答には headers が無い」は誤り）。
#   1通0.59秒 × 件数 かかっていた詳細取得が丸ごと不要になり、取得が 1.3件/秒 → 5.7件/秒 に改善。

# 下書き（未送信メール）を取り込むか。本番は必ず False にすること。
# ★下書きは実際に送受信していない=通信記録ではないが、Gmail検索には既定でヒットする
#   （実測: to:ドメイン で検索すると labelIds=["DRAFT"] の行がそのまま返ってくる）。
#   テストで作った下書きメールを使って動作検証したいときだけ True にする。
INCLUDE_DRAFTS = False

# 添付ファイルを処理するか（Drive保存＋LLM要約＋source=drive のファイル行作成）。
# ★False にするとメール本文だけを保存する。添付は1件ごとに
#   ダウンロード→Drive重複チェック→アップロード→LLM要約 と重いため、
#   まず本文だけを早く揃えたい場合は False にする（後から True で再実行すれば
#   [mail:ID] の冪等 upsert で同じ行に添付情報が足される）。
PROCESS_ATTACHMENTS = False


def _base_query(domain: str) -> str:
    """その会社のドメインが from/to/cc/bcc のどこかに現れるメールを対象にする。"""
    q = f"(from:{domain} OR to:{domain} OR cc:{domain} OR bcc:{domain})"
    if not INCLUDE_DRAFTS:
        q += " -in:drafts"
    return q


# 応答が大きすぎるエラーの識別文言（2026-07-29 nanocosme社で実測。本文が重いメールが
# 多い会社だと、include_payload=True で100件ぶんの応答が一括では返せないほど巨大になる）
_PAYLOAD_TOO_LARGE_MARKER = "payload is too large"
# ページサイズを縮める下限。ここまで縮めても失敗するなら別要因（1通が単体で巨大等）
_MIN_PAGE_SIZE = 5


def _is_payload_too_large(err_text: str) -> bool:
    return _PAYLOAD_TOO_LARGE_MARKER in (err_text or "").lower()


async def _fetch_page(query: str, page_token: str | None,
                      max_results: int = PAGE_SIZE) -> tuple[list[dict], str | None]:
    """Gmail を1ページ（最大 max_results 件）取得して (メール一覧, 次のトークン) を返す。

    ★include_payload=True で本文・全ヘッダまで取れるので追加の詳細取得は不要。
      1ページずつ取得→保存→チェックポイント、を繰り返す（全件取ってから保存はしない）。
    ★呼び出し側（_backfill_one_company）が「応答が大きすぎる」エラーを検知した場合、
      max_results を縮めて同じ page_token でこの関数を呼び直す想定
      （本文が重いメールが多い会社では既定件数でも一括では返せないことがあるため）。
    """
    kwargs: dict = {
        "query": query,
        "max_results": max_results,
        "verbose": True,
        "include_payload": True,
        "include_spam_trash": True,
    }
    if page_token:
        kwargs["page_token"] = page_token
    resp = await _call_with_retry(GMAIL_FETCH_EMAILS, **kwargs)
    body = resp.get("data") or resp
    messages = [_pick(m) for m in (body.get("messages") or [])]
    return messages, body.get("nextPageToken")


async def _fetch_page_adaptive(query: str, page_token: str | None,
                               state: dict) -> tuple[list[dict], str | None]:
    """_fetch_page を呼び、「応答が大きすぎる」エラーが出たらページサイズを半分に
    縮めて同じ場所（同じ page_token）から再試行する。

    ★2026-07-29 nanocosme社で実測: 本文が重いメールが多い会社は、100件ぶんを
      include_payload=True で一括取得すると応答サイズ上限を超えることがある。
      「件数を減らしてもらう」のではなく、ここで自動的に折り合いをつける。
    見つかったサイズは state["page_size"] に記録し、以降のページでも使い回す
    （毎ページ100から縮め直すのは無駄なため）。
    """
    size = state.get("page_size") or PAGE_SIZE
    while True:
        try:
            messages, next_token = await _fetch_page(query, page_token, max_results=size)
            if size != state.get("page_size"):
                state["page_size"] = size
                print(f"  ℹ 以降このページサイズ（{size}件）を使います")
            return messages, next_token
        except Exception as e:
            if not _is_payload_too_large(str(e)) or size <= _MIN_PAGE_SIZE:
                raise
            new_size = max(_MIN_PAGE_SIZE, size // 2)
            print(f"  ⚠ 応答が大きすぎたため件数を {size}→{new_size} 件に縮めて再試行します")
            size = new_size


# ========== Notion保存部 ==========

def _rt(value: str) -> list[dict]:
    """rich_text/title 用の配列を作る。Notion の 2000字上限に合わせてチャンク分割する。"""
    s = value or ""
    if not s:
        return []
    return [{"text": {"content": s[i:i + 2000]}} for i in range(0, len(s), 2000)]


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


async def _build_item(payload: dict, project: str) -> dict | None:
    """1通から upsert 用の item（match/create/update）を作る。保存はまだしない。

    ★添付があるメールだけ Drive 保存＋要約＋ファイル行 upsert をここで済ませる
      （links にファイルURLが必要なため、メール行の item を作る前に実行する）。
    message_id が無い、または例外が出たメールは None（呼び出し側でスキップ扱い）。
    """
    message_id = payload.get("message_id", "")
    if not message_id:
        # message_id が無いと添付取得・重複判定キーが作れないためスキップ
        return None
    try:
        subject      = payload.get("subject", "")
        sender_raw   = payload.get("sender", "")
        sender       = _extract_email(sender_raw)
        recipient    = _extract_emails(payload.get("to", ""))  # 複数宛先を全て保存
        # name には送信者・受信者両方の表示名を入れる
        names        = _extract_names(sender_raw, payload.get("to", ""))
        # CC / BCC は1つのプロパティにまとめる（両方のアドレスを重複除去して結合）
        cc_bcc       = _extract_emails(f"{payload.get('cc', '')} {payload.get('bcc', '')}")
        received_at  = _to_jst_str(payload.get("message_timestamp", ""))
        message_body = payload.get("message_text", "")
        thread_id    = payload.get("thread_id", "")

        # 添付は slack_recorder_v5 と同じ扱い: Drive「gmail」フォルダに実体保存し、
        # source=drive のファイル行を別行で upsert する。メール行には
        # content=本文＋添付ファイル名 / links=「ファイル名\n正規化URL」を入れる
        att_names: list[str] = []
        links_parts: list[str] = []
        if PROCESS_ATTACHMENTS and payload.get("attachment_list"):
            att_names, links_parts = await _process_attachments(
                message_id, payload["attachment_list"],
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
        # 人物辞書用に送受信者を収集（書き込みは run() の最後にまとめて行う）
        _people_buffer.extend(_extract_name_email_pairs(sender_raw, payload.get("to", "")))
        # メール行の重複防止は content 末尾に埋めた [mail:ID] 行の contains 一致
        # （再処理は既存行の上書きで済む。ファイル行は [mail:] を含まないため誤マッチしない）
        return {
            "match":  {"property": PROP_CONTENT, "title": {"contains": f"[mail:{message_id}]"}},
            "create": {"properties": props},
            "update": {"properties": props},
        }
    except Exception as e:
        print(f"  ⚠ item構築に失敗 (id={message_id}): {e}")
        return None


async def _save_items(items: list[dict], raw_db_id: str) -> int:
    """items をまとめて upsert する（成功件数を返す）。

    ★1件ずつ upsert すると Notion のレート制限（約3req/s）で 100件=88秒かかっていたのを、
      UPSERT_BATCH 件ずつまとめることで 100件=50秒に短縮（実測1.6倍速）。
    """
    if not items:
        return 0
    chunks = [items[i:i + UPSERT_BATCH] for i in range(0, len(items), UPSERT_BATCH)]
    sem = asyncio.Semaphore(_UPSERT_CONCURRENCY)

    async def _one(chunk: list[dict]) -> int:
        async with sem:
            try:
                result = await _call_with_retry(
                    NOTION_UPSERT_ROW_DATABASE, database_id=raw_db_id, items=chunk
                )
            except Exception as e:
                print(f"  ⚠ バッチ保存失敗（{len(chunk)}件）: {e}")
                return 0
            data = result.get("data") or {} if isinstance(result, dict) else {}
            results = data.get("results") if isinstance(data, dict) else None
            if isinstance(results, list):
                ng = [r for r in results if isinstance(r, dict) and not r.get("ok")]
                if ng:
                    print(f"  ⚠ 行の書き込み失敗 {len(ng)}件: {_upsert_failed_detail(result)}")
                return len(results) - len(ng)
            return len(chunk)

    return sum(await asyncio.gather(*(_one(c) for c in chunks)))


async def _oldest_saved_epoch(project: str) -> int | None:
    """raw DB に保存済みの、そのプロジェクトの最古 timestamp を epoch 秒で返す。

    ★tmp/ のしおりが消えた（セッション切替）ときの再開位置に使う。
      取得は新しい順なので「保存済みの最古」＝取り込みのフロンティア。
      そこより古い方を `before:` で続ければ、ほぼ続きから再開できる（境界の重複は冪等で吸収）。
    """
    try:
        resp = await _call_with_retry(
            NOTION_QUERY_DATABASE_WITH_FILTER,
            database_id=RAW_DB_ID,
            filter={"and": [
                {"property": PROP_PROJECT, "rich_text": {"equals": project}},
                {"property": PROP_SOURCE, "select": {"equals": SOURCE_GMAIL}},
            ]},
            sorts=[{"property": PROP_TIMESTAMP, "direction": "ascending"}],
            page_size=1,
        )
        data = resp.get("data", resp) if isinstance(resp, dict) else {}
        results = data.get("results") or []
        if not results:
            return None
        ts = (((results[0].get("properties") or {}).get(PROP_TIMESTAMP) or {})
              .get("date") or {}).get("start")
        if not ts:
            return None
        return int(datetime.fromisoformat(ts).timestamp())
    except Exception as e:
        print(f"  ⚠ 最古timestampの取得に失敗（最初から取得します）: {e}")
        return None


def _state_path(project: str, domain: str = "") -> str:
    """チェックポイントファイルのパス（日本語名でも安全な英数字に落とす）。

    ★組み込み hash() はプロセス毎に値が変わり再開位置を見失うため md5 を使う。
    ★2026-08-16: project だけでなく domain も鍵に含める。1社が複数ドメインを
      持つ場合（例: 中央コンピューター = ccsplus-hd.jp と ccs1981.jp）、
      domain_map に同じプロジェクト名で複数行を置く運用になるが、project だけを
      鍵にすると両者が同じチェックポイントを共有し、2つ目のドメインが
      1つ目の pageToken を引き継いで永久に取得されなくなる。
    """
    key = f"{project}|{domain}" if domain else project
    safe = re.sub(r"[^A-Za-z0-9]", "", project)[:20] or "pjt"
    digest = hashlib.md5(key.encode("utf-8")).hexdigest()[:8]
    return f"tmp/gmail_bf_{safe}_{digest}.json"


# ========== エントリポイント ==========

def _resolve_targets(all_projects: list[dict], projects: list[str] | None) -> list[dict]:
    """projects 指定（pjt_ prefix の有無は自動吸収）を domain_map の行に解決する。"""
    if not projects:
        return all_projects
    wanted: set[str] = set()
    for name in projects:
        n = (name or "").strip()
        if n:
            wanted.add(n)
            wanted.add(n if n.startswith("pjt_") else f"pjt_{n}")
    targets = [p for p in all_projects if p["project"] in wanted]
    found = {p["project"] for p in targets}
    for name in projects:
        n = (name or "").strip()
        if n and not ({n, f"pjt_{n}"} & found):
            print(f"⚠ domain_map に見つからないプロジェクト: {name}")
    print(f"プロジェクト絞り込み: 全{len(all_projects)}件中 {len(targets)}件を対象")
    return targets


async def _resume_state(project: str, domain: str, *,
                        multi_domain: bool = False) -> tuple[dict, str]:
    """チェックポイントを読み、無ければ Notion から再開位置を導出する（2段構え）。

    1. tmp/ に state があればそれを使う（同一セッション内＝pageToken で正確に続き）
    2. state が無い（セッションが切り替わって tmp/ が消えた）→ raw DB に保存済みの
       最古 timestamp を見て `before:` を付けたクエリで再開する（取得は新しい順なので
       「保存済みの最古」が取り込みのフロンティア）。境界の重複は [mail:ID] 冪等で吸収。

    ★multi_domain=True（同じプロジェクト名が domain_map に複数行ある会社）のときは
      2 を使わない。raw DB の project 列にはドメインの区別が無いため、先に取り込んだ
      別ドメインの最古 timestamp を掴んでしまい、こちらのドメインの新しいメールが
      まるごとスキップされる。最初から取り直す（[mail:ID] 冪等なので重複は出ない）。
    """
    path = _state_path(project, domain)
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as f:
                state = json.load(f)
            return state, f"tmp再開（token={'あり' if state.get('page_token') else 'なし'}）"
        except Exception as e:
            print(f"  ⚠ state 読み込み失敗（最初から）: {e}")
    if multi_domain:
        return ({"query": _base_query(domain), "page_token": None, "saved": 0, "done": False},
                "最初から（複数ドメインの会社のため durable 再開は使わない）")
    epoch = await _oldest_saved_epoch(project)
    if epoch:
        query = f"{_base_query(domain)} before:{epoch + 1}"
        return ({"query": query, "page_token": None, "saved": 0, "done": False},
                f"Notion由来のdurable再開（before:{epoch + 1}）")
    return ({"query": _base_query(domain), "page_token": None, "saved": 0, "done": False},
            "最初から（保存済みなし）")


def _write_state(project: str, state: dict, domain: str = "") -> None:
    """チェックポイントを書く（★1ページ分を保存し終えた後に呼ぶこと）。

    ★一時ファイルに書いてから os.replace で置き換える（原子的な書き込み）。
      書き込み中にランタイムが落ちても、既存の state ファイルは壊れずに残る
      （直接同名ファイルに書くと、書きかけの不完全なJSONが残るリスクがある）。
    """
    path = _state_path(project, domain)
    try:
        os.makedirs("tmp", exist_ok=True)
        tmp_path = f"{path}.tmp"
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False)
        os.replace(tmp_path, path)
    except Exception as e:
        print(f"  ⚠ state 書き込み失敗（再開位置が保存されません）: {e}")


async def _backfill_one_company(project: str, domain: str, *, deadline: float,
                                multi_domain: bool = False) -> dict:
    """1社ぶんを時間予算内で取り込む（1ページ取得→保存→チェックポイント の繰り返し）。"""
    state, how = await _resume_state(project, domain, multi_domain=multi_domain)
    if state.get("done"):
        _log("company_already_done", f"  ✅ この会社は完了済み（累計{state.get('saved', 0)}件）",
             project=project, saved_total=state.get("saved", 0))
        return {"project": project, "domain": domain, "done": True,
                "saved_total": state.get("saved", 0), "saved_this_run": 0, "pages": 0}

    query = state.get("query") or _base_query(domain)
    token = state.get("page_token")
    _log("company_start", f"  再開: {how}\n  query: {query}",
         project=project, domain=domain, resume=how, query=query)

    saved_run, pages, skipped = 0, 0, 0
    # ★残り時間が1ページ分を下回ったら次のページに入らない（予算超過後の走り切りを見込む）
    while deadline - time.monotonic() > PAGE_COST_EST:
        try:
            messages, next_token = await _fetch_page_adaptive(query, token, state)
        except Exception as e:
            _log("company_fetch_error", f"  Gmail取得エラー: {e}",
                 project=project, error=str(e))
            return {"project": project, "domain": domain, "done": False,
                    "saved_total": state.get("saved", 0), "saved_this_run": saved_run,
                    "pages": pages, "error": str(e)}

        # 添付ありメールは Drive 保存＋要約＋ファイル行 upsert も item 構築の中で行う
        items = [it for it in await asyncio.gather(
            *(_build_item(m, project) for m in messages)) if it]
        skipped += len(messages) - len(items)
        ok = await _save_items(items, RAW_DB_ID)
        saved_run += ok

        token = next_token
        state["query"] = query
        state["page_token"] = token
        state["saved"] = state.get("saved", 0) + ok
        state["done"] = not token
        pages += 1
        _write_state(project, state, domain)   # ★保存し終えてからチェックポイント
        # ★人物辞書もページ単位で差分を書く（run末尾まで溜めるとランタイム死で全部失うため）
        people_before = len(_people_written)
        await _flush_people()
        new_people = len(_people_written) - people_before
        _log("page_done",
             f"  page{pages}: 取得{len(messages)} / 保存{ok} / 累計{state['saved']} "
             f"/ 人物+{new_people} / 残り{deadline - time.monotonic():.0f}秒 "
             f"/ token={'あり' if token else '★なし（完了）★'}",
             project=project, page=pages, fetched=len(messages), saved=ok,
             saved_total=state["saved"], new_people=new_people,
             remaining=round(deadline - time.monotonic(), 1), has_next=bool(token))
        if state["done"]:
            break

    _log("company_end", project=project, done=bool(state.get("done")),
         saved_total=state.get("saved", 0), saved_this_run=saved_run,
         pages=pages, skipped=skipped)
    return {"project": project, "domain": domain, "done": bool(state.get("done")),
            "saved_total": state.get("saved", 0), "saved_this_run": saved_run,
            "pages": pages, "skipped": skipped}


async def run(projects: list[str] | None = None, *,
              budget_seconds: float = TIME_BUDGET_SECONDS) -> dict:
    """domain_map のプロジェクトごとに Gmail を取得して raw DB に保存する。

    Args:
        projects: プロジェクト名のリスト（pjt_ prefix の有無どちらでも可）。
                  省略時は domain_map の全プロジェクト。
        budget_seconds: 1回の実行の打ち切り目標（既定 TIME_BUDGET_SECONDS=250秒）。

    ★1回で終わらない設計。出力に「→ 続きあり」と出たら同じ呼び出しを繰り返すこと
      （「✅ 全完了」まで）。再開位置は tmp/ の pageToken、tmp/ が消えていれば
      raw DB の最古 timestamp から導出するため、セッションが切れても続きから進む。
    ★期間指定は無く、その会社のドメインが絡むメールを全期間取得する。
    """
    global _run_id
    t0 = time.monotonic()
    deadline = t0 + budget_seconds
    _run_id = "gbf-" + datetime.now(JST).strftime("%m%d-%H%M%S")
    _log("run_start", f"raw DB: {RAW_DB_ID} / 時間予算: {budget_seconds:.0f}秒",
         projects=projects, budget=budget_seconds)
    _people_buffer.clear()
    _people_written.clear()   # 実行ごとにリセット（前回書いた人も今回また確認する）

    targets = _resolve_targets(await _load_projects(), projects)
    print(f"プロジェクト数: {len(targets)} 件")

    # 同じプロジェクト名が複数行ある会社（＝複数ドメインを持つ会社）を把握する。
    # その会社は durable 再開（保存済み最古 timestamp からの before: 継続）を使えない
    # ため、_backfill_one_company に伝える（詳細は _resume_state の docstring）。
    _dom_count: dict[str, int] = {}
    for p in targets:
        _dom_count[p["project"]] = _dom_count.get(p["project"], 0) + 1
    multi = {k for k, v in _dom_count.items() if v > 1}
    if multi:
        print(f"複数ドメインの会社: {', '.join(sorted(multi))}")

    results: list[dict] = []
    for pjt in targets:
        if deadline - time.monotonic() <= PAGE_COST_EST:
            print(f"\n（時間予算に達したため残り{len(targets) - len(results)}社は次回に回します）")
            break
        print(f"\n--- {pjt['project']} ({pjt['domain']}) ---")
        try:
            results.append(await _backfill_one_company(
                pjt["project"], pjt["domain"], deadline=deadline,
                multi_domain=pjt["project"] in multi))
        except Exception as e:
            # ★_backfill_one_company は Gmail取得の失敗しか自前で捕まえていないため、
            #   想定外の例外（Notion側の未知エラー等）1社分でrun全体が止まらないようにする。
            #   その社の保存済みデータ・チェックポイントは無事なので、次回そのまま再開できる。
            _log("company_unexpected_error",
                 f"  ✗ 想定外のエラーで中断（他社は続行）: {type(e).__name__}: {e}",
                 project=pjt["project"], error=f"{type(e).__name__}: {e}")
            results.append({"project": pjt["project"], "domain": pjt["domain"], "done": False,
                            "saved_total": 0, "saved_this_run": 0, "pages": 0, "error": str(e)})

    # 最後のページ以降に溜まった分を書く（通常はページ単位で書き終えているので少数）
    await _flush_people()
    if _people_written:
        print(f"\n人物辞書: この実行で {len(_people_written)} 人を登録/更新")

    saved_run = sum(r.get("saved_this_run", 0) for r in results)
    # 未着手の会社が残っている場合も「続きあり」とする
    all_done = len(results) == len(targets) and all(r.get("done") for r in results)
    el = time.monotonic() - t0
    print(f"\n===== このrun: 保存{saved_run}件 / {len(results)}社 / {el:.1f}秒 =====")
    for r in results:
        mark = "✅" if r.get("done") else "→"
        print(f"  {mark} {r['project']}: このrun{r.get('saved_this_run', 0)}件 "
              f"/ 累計{r.get('saved_total', 0)}件"
              + (f" / エラー: {r['error']}" if r.get("error") else ""))
    _log("run_end", "✅ 全完了" if all_done else "→ 続きあり。もう一度 run() を実行してください",
         done=all_done, saved_this_run=saved_run, companies=len(results),
         elapsed=round(el, 1))
    return {"done": all_done, "saved_this_run": saved_run, "results": results}
