"""notion_zoom_過去情報.py — 期間内の全Zoom録画を raw DB に一括保存する（バッチ用・2段方式）

★2026-07-25 改修: 「1回で全部やる」方式をやめ、**Notion 自身を作業キューにする2段方式**にした。
  code_execute には約250秒の実行窓しかなく（実測: 314秒成功 / 385秒失敗）、
  さらにサンドボックスセッションは30分TTLで不意に消える（tmp/ のしおりも消える）。
  そこで「中間状態を tmp ではなく Notion に置く」ことで、何度死んでも続きから完走できる。

  1段目 run_ingest(from, to)  : 録画一覧 → token → VTT DL → **文字起こしだけ保存**
        （LLM を呼ばないので軽い。project="未処理" を目印に置き、VTT はページ本文へ）
  2段目 run_summarize()       : project="未処理" の行を1件ずつ拾って議事録生成 →
        同じ行を更新（content=議事録 / project=本当のPJ名）→ キューから外れる

  どちらも時間予算（TIME_BUDGET_SECONDS）で自分から止まり、「✅完了」まで再実行するだけ。
  再開位置は **Notion の状態そのもの**（parents=uuid の有無／project="未処理" の有無）なので、
  tmp/ が消えても・token が失効しても影響を受けない。

- 保存先は全ソース統合の raw DB（source=meeting。Gmail/Slack と同じDB）
- 重複判定は parents=meeting_uuid ＋ source=meeting の存在チェック（再実行に安全）
- 議事録は content（title）に、生VTTはページ本文ブロックに保存（本文は新規作成時のみ）
- VTT は tmp/vtt/ にもキャッシュするが、あくまで高速化用（消えても Notion 側で完結する）
- エントリポイント:
    run_ingest("2026-06-01", "2026-06-30")  … 文字起こしの取り込み
    run_summarize()                          … 議事録生成（未処理キューを消化）
    run("2026-06-01", "2026-06-30")          … 上2つを時間予算内で続けて実行
"""
import asyncio
import json
import os
import random
import re
import time
import urllib.request
from datetime import datetime, timedelta, timezone
from urllib.parse import quote

from agent_sdk import (
    NOTION_FETCH_ALL_BLOCK_CONTENTS,
    NOTION_QUERY_DATABASE_WITH_FILTER,
    NOTION_UPSERT_ROW_DATABASE,
    ZOOM_GET_MEETING_RECORDINGS,
    ZOOM_GET_PAST_MEETING_PARTICIPANTS,
    ZOOM_LIST_ALL_RECORDINGS,
    llm_call,
)

PARALLEL = 16    # token取得・VTTダウンロードの並列数（実行環境実行スロット上限）
# ※議事録生成は「未処理キューを1件ずつ」に変えたため並列数（旧 BATCH_SIZE）は不要になった

# ===== 時間予算（code_execute の実行窓 約300秒に対する安全圏） =====
# 実測: 250秒予算→実際314秒で成功 / 385秒相当→ReadTimeoutError。
# 予算を超えた後も「処理中の1件」は走り切るため、1件分のコストを見込んで手前で止める。
TIME_BUDGET_SECONDS = 250
INGEST_CHUNK = 8          # 1バッチで token→DL→保存 する会議数
INGEST_COST_EST = 60      # 1バッチ(INGEST_CHUNK件)の想定所要秒。残りがこれ未満なら次に入らない
SUMMARIZE_COST_EST = 100  # 議事録1件の想定所要秒（LLM込み）。残りがこれ未満なら次に入らない

# 文字起こしだけ保存した行の目印（project 列に入れる）。
# ★列を追加せずに「未処理キュー」を作るための仕掛け。議事録生成時に本当のPJ名で上書きされる
MARKER_PROJECT = "未処理"
# 議事録が作れなかった行（文字起こしが実質空）の目印。
# ★これに変えることでキューの先頭で永久に詰まるのを防ぐ（人が見て判断する対象）
MARKER_SKIPPED = "未処理_VTT不足"
# ページ本文に入れるVTTブロックの上限（Notion API の children 上限100に対する安全値）
MAX_VTT_BLOCKS = 90

# raw DB ID（固定。全ソース統合DB「社内情報集約PJ > raw」。Gmail/Slack と同じ保存先）
RAW_DB_ID = "<RAW_DB_ID>"
# プロジェクト判定の候補リスト取得元: ドメイン情報DB（Gmail と同じ。タイトル=pjt_プロジェクト名）
# ※旧プロジェクト一覧DB（36e4a092...）は旧構成のため使わない
DOMAIN_INFO_DB_ID = "<DOMAIN_MAP_DB_ID>"
DOMAIN_PROP = "domain"   # domain_map（旧ドメイン情報DB）のドメイン列名
# 人物辞書DB ID（meta > people。登場人物の自動登録先。空にすると人物登録をスキップ）
PEOPLE_DB_ID = "<PEOPLE_DB_ID>"
# 進捗記録DB ID（meta > sync_state。key ごとに1行だけ持ち、実行のたびに value を上書きする）
# ★tmp/ はセッションが切れると消えるので、走査の進捗はここに置く
SYNC_STATE_DB_ID = "<SYNC_STATE_DB_ID>"
PROP_S_KEY   = "key"     # title: 進捗の種類
PROP_S_VALUE = "value"   # rich_text: 進捗の中身（JSON文字列）
PROP_S_NOTE  = "note"    # rich_text: 人が読むためのメモ
SWEEP_STATE_KEY = "zoom_ingest_sweep"   # 月スイープの進捗を入れる行の key

# raw DB のプロパティ名（Notion側の列名が変わったらここだけ直す）
PROP_CONTENT   = "content"    # title: 議事録全文（生VTTはページ本文へ）
PROP_SUBJECT   = "subject"    # rich_text: 会議タイトル
PROP_TIMESTAMP = "timestamp"  # date: 開始日時（JST・ISO 8601）
PROP_SENDER    = "sender"     # rich_text: 参加者のメールアドレス（カンマ結合・ホスト含む）
PROP_NAME      = "name"       # rich_text: 参加者名（カンマ結合）
PROP_PROJECT   = "project"    # rich_text: プロジェクト名
PROP_SOURCE    = "source"     # select: 情報ソース
PROP_PARENTS   = "parents"    # rich_text: meeting_uuid（重複判定キー）
SOURCE_MEETING = "meeting"    # source select の選択肢名

NOTION_RICH_TEXT_LIMIT = 2000
DOWNLOAD_TIMEOUT = 120

# 前処理後の文字起こしがこれ未満なら議事録を作らない（文字起こし未完了/無発話の防御）
_MIN_TRANSCRIPT_CHARS = 200
# LLM に渡す文字起こしの上限（長時間会議でのコンテキスト/コスト爆発防止）
_MAX_TRANSCRIPT_CHARS = 120_000
JST = timezone(timedelta(hours=9))


# ===== デバッグ用 JSONL ログ =====
# ★ReadTimeoutError やセッション破棄が起きると stdout は丸ごと失われる。
#   このファイルはイベントごとに追記＋即フラッシュするので、落ちた後でも
#   「どこまで何をやったか」を read_log() で追跡できる。
DEBUG_LOG_PATH = "tmp/logs/zoom_過去情報.jsonl"
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

    使い方: read_log(60) で直近60イベント / read_log(20, "meeting_error") で失敗だけ。
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


# ===== 共通ユーティリティ =====

def _page_title(page: dict) -> str:
    """ページの title 型プロパティを名前によらず探して平文で返す。"""
    props = page.get("properties") or {}
    for _, v in props.items():
        if (v or {}).get("type") == "title":
            parts = v.get("title") or []
            return "".join((p.get("plain_text") or "") for p in parts)
    return ""


async def _retry(coro_factory, *, label: str = "api", attempts: int = 3):
    last = None
    for i in range(attempts):
        try:
            return await coro_factory()
        except Exception as e:
            last = e
            # ジッター付き指数バックオフ（同時リトライの衝突＝レート制限の再発を防ぐ）
            wait = 2 ** i + random.uniform(0, 1)
            print(f"  {label} retry {i+1}/{attempts} after {wait:.1f}s: {e}")
            await asyncio.sleep(wait)
    raise RuntimeError(f"{label} が {attempts} 回失敗: {last}")


def _split_text(text: str, limit: int = NOTION_RICH_TEXT_LIMIT) -> list[str]:
    chunks = []
    while len(text) > limit:
        cut = text.rfind("\n", 0, limit)
        if cut < limit // 2:
            cut = limit
        chunks.append(text[:cut])
        text = text[cut:].lstrip("\n")
    if text:
        chunks.append(text)
    return chunks


def _rt(value: str) -> list[dict]:
    """rich_text/title 用の配列を作る。Notion の 2000字上限に合わせてチャンク分割する。"""
    s = value or ""
    if not s:
        return []
    return [{"text": {"content": s[i:i + 2000]}} for i in range(0, len(s), 2000)]


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
# 並列の保存中に都度書くと同一人物のマージが競合するため）
_people_buffer: list[dict] = []


def _valid_person_email(email: str) -> str:
    """人物辞書に登録可能なメールなら小文字正規化して返す。不可なら空文字。"""
    e = (email or "").strip().lower()
    if not e or "@" not in e:
        return ""
    if any(b in e for b in _PEOPLE_EMAIL_BLOCKLIST):
        return ""
    return e


_PEOPLE_QUERY_CHUNK = 50   # OR条件・バッチupsertともにこの件数ずつまとめる
_PEOPLE_WRITE_CONCURRENCY = 2  # 書き込みチャンクの並列数（Notionのレート制限に配慮）


async def _upsert_people(persons: list[dict]) -> None:
    """登場人物を人物辞書DBに upsert する（name / email / aliases / company の4列）。

    ★2026-07-25 バッチ化: 以前は1人ずつ「読み取り→書き込み」を直列でやっていた
      （Gmail版と同じ問題）。読み取りは全員ぶんを1回のOR検索でまとめ、
      書き込みは items=[...] のバッチupsertにする
      （件数が多い場合は _PEOPLE_QUERY_CHUNK 件ずつ）。

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
                resp = await _retry(
                    lambda kw=kwargs: NOTION_QUERY_DATABASE_WITH_FILTER(**kw),
                    label="people_query",
                )
            except Exception as e:
                print(f"[WARN] 人物辞書の一括読み取りに失敗（{len(chunk_emails)}件）: {e}")
                break
            data = resp.get("data", resp) if isinstance(resp, dict) else {}
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

    # ---- バッチ書き込み（チャンク分割＋並列） ----
    chunks = [items[i:i + _PEOPLE_QUERY_CHUNK] for i in range(0, len(items), _PEOPLE_QUERY_CHUNK)]
    sem = asyncio.Semaphore(_PEOPLE_WRITE_CONCURRENCY)

    async def _write_chunk(chunk: list[dict]) -> None:
        async with sem:
            try:
                result = await _retry(
                    lambda c=chunk: NOTION_UPSERT_ROW_DATABASE(
                        database_id=PEOPLE_DB_ID, items=c
                    ),
                    label="people_upsert",
                )
            except Exception as e:
                print(f"[WARN] 人物辞書の一括更新に失敗（{len(chunk)}件）: {e}")
                return
            err = _upsert_failed_detail(result)
            if err:
                print(f"[WARN] 人物辞書の一括更新でNotion側の一部失敗: {err}")

    await asyncio.gather(*(_write_chunk(c) for c in chunks))


def _clean_participant_name(name: str) -> str:
    """Zoom表示名から敬称・括弧内の所属・区切り以降のサフィックスを除いた照合用の名前を返す。

    例: "Ayano Ito さん" → "Ayano Ito" / "田中太郎(営業部)" → "田中太郎" /
        "Taro Tanaka - ABC商事" → "Taro Tanaka"
    """
    n = (name or "").strip()
    n = re.sub(r"[（(\[【].*?[）)\]】]", "", n)
    n = re.split(r"\s+[-–—|｜/／]\s*", n)[0]
    n = re.sub(r"\s*(さん|様|さま|君|くん|ちゃん|先生|氏)$", "", n)
    return n.strip()


def _name_tokens(cleaned: str) -> list[str]:
    """照合用トークン（姓・名など）に分割する。

    空白（全半角）区切りに加え、"AyanoIto" のような camelCase 境界でも分割する。
    1文字トークンは誤ヒットの元なので除外する。
    """
    s = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", cleaned)
    return [t for t in re.split(r"[\s\u3000]+", s) if len(t) >= 2]


async def _resolve_emails_from_people(names: list[str]) -> list[dict]:
    """メールが取れなかった参加者名を人物辞書DBで逆引きする。

    実際のZoom表示名は「Ayano Ito さん」「田中太郎(営業部)」「Ito Ayano」のような
    揺れが多いため、2段階で照合する:
      段階1: 敬称・括弧・サフィックスを除去した名前で contains 照合
      段階2: 0件ならトークン分割（空白/camelCase境界）の AND 照合
             （姓名逆順・スペース無し表記に対応）
    どの段階でも一致が**ちょうど1件**のときだけ採用する
    （複数ヒット＝同名別人の可能性があるため推測しない）。
    戻り値: [{"name": 元の名前, "email": 解決したメール}]
    """
    if not PEOPLE_DB_ID or not names:
        return []

    async def _query_people(filter_obj: dict) -> list[dict]:
        resp = await _retry(
            lambda f=filter_obj: NOTION_QUERY_DATABASE_WITH_FILTER(
                database_id=PEOPLE_DB_ID, filter=f, page_size=2,
            ),
            label="people_reverse",
        )
        data = resp.get("data", resp) if isinstance(resp, dict) else {}
        return data.get("results") or []

    def _contains_group(q: str) -> dict:
        return {"or": [
            {"property": PROP_P_NAME, "title": {"contains": q}},
            {"property": PROP_P_ALIASES, "rich_text": {"contains": q}},
        ]}

    def _email_of(row: dict) -> str:
        props = row.get("properties") or {}
        return "".join(
            t.get("plain_text", "")
            for t in ((props.get(PROP_P_EMAIL) or {}).get("rich_text") or [])
        ).strip()

    # ★2026-07-25 並列化: 以前は名前ごとに直列（1人あたり最大2回のNotion問い合わせ）
    #   だったが、名前同士の解決は互いに独立しているため asyncio.gather で並列化する。
    #   1人の中の「段階1→段階2」の順序（段階1で決着すれば段階2は呼ばない）は維持する。
    sem = asyncio.Semaphore(6)  # Notionのレート制限に配慮しつつ並列化

    async def _resolve_one(name: str) -> dict | None:
        n = (name or "").strip()
        cleaned = _clean_participant_name(n)
        if not cleaned:
            return None
        async with sem:
            try:
                # 段階1: クリーニング済みの名前そのままで照合
                rows = await _query_people(_contains_group(cleaned))
                if len(rows) > 1:
                    print(f"  ⚠ 人物辞書で複数ヒットのためメール解決を見送り: {n}")
                    return None
                if not rows:
                    # 段階2: トークンAND照合（姓名逆順・スペース無し・camelCase対応）
                    tokens = _name_tokens(cleaned)
                    if len(tokens) >= 2:
                        rows = await _query_people({"and": [_contains_group(t) for t in tokens]})
                    if len(rows) > 1:
                        print(f"  ⚠ 人物辞書で複数ヒットのためメール解決を見送り: {n}")
                        return None
                if not rows:
                    # 未登録＝辞書がまだ育っていないだけ。後日この人がメール付きで
                    # 観測されれば辞書に入り、以降の会議では自動解決される
                    print(f"  人物辞書に未登録のためメール無しで記録: {n}")
                    return None
                email = _email_of(rows[0])
                if email:
                    print(f"  人物辞書からメール解決: {n} → {email}")
                    return {"name": n, "email": email}
                return None
            except Exception as e:
                print(f"  ⚠ 人物辞書の逆引き失敗: {n} / {str(e)[:80]}")
                return None

    results = await asyncio.gather(*(_resolve_one(name) for name in names))
    return [r for r in results if r]


def _to_jst(utc_str: str) -> str:
    """JST の ISO 8601 形式（例: 2026-07-09T13:53:01+09:00）で返す。

    raw DB の timestamp（date型）にそのまま入れられる形式。
    Gmail/Slack スクリプトと表記を統一する。
    """
    if not utc_str:
        return ""
    try:
        dt = datetime.fromisoformat(utc_str.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(JST).isoformat(timespec="seconds")
    except Exception as e:
        print(f"⚠ 時刻変換失敗 ({utc_str}): {e}")
        return utc_str


def _normalize(s: str) -> str:
    """大文字小文字・空白・括弧内テキストを正規化する（名前比較用）"""
    s = (s or "").strip()
    s = re.sub(r"[（(\[].*?[）)\]]", "", s)
    s = re.sub(r"\s+", "", s)
    return s.lower()


def _encode_uuid_double(uuid: str) -> str:
    """Zoom UUID は '/' や '==' を含む場合ダブル URL エンコードが必要（録画取得API用）"""
    if uuid.startswith("/") or "//" in uuid or "+" in uuid or "=" in uuid:
        return quote(quote(uuid, safe=""), safe="")
    return uuid


def _encode_uuid_single(uuid: str) -> str:
    """参加者取得APIはシングル URL エンコード"""
    if uuid.startswith("/") or "//" in uuid or "+" in uuid or "=" in uuid:
        return quote(uuid, safe="")
    return uuid


def _month_chunks(from_date: str, to_date: str) -> list[tuple[str, str]]:
    start = datetime.strptime(from_date, "%Y-%m-%d")
    end = datetime.strptime(to_date, "%Y-%m-%d")
    chunks = []
    cur = start
    while cur <= end:
        next_month = (cur.replace(day=1) + timedelta(days=32)).replace(day=1)
        chunk_end = min(next_month - timedelta(days=1), end)
        chunks.append((cur.strftime("%Y-%m-%d"), chunk_end.strftime("%Y-%m-%d")))
        cur = next_month
    return chunks


# ===== Phase1: 録画一覧取得 + token取得 =====

async def _list_all_recordings(from_date: str, to_date: str) -> list[dict]:
    all_meetings: list[dict] = []
    seen_uuids: set = set()
    for chunk_from, chunk_to in _month_chunks(from_date, to_date):
        print(f"  録画取得中: {chunk_from} ～ {chunk_to}")
        token = ""
        while True:
            kwargs: dict = {"user_id": "me", "page_size": 300}
            kwargs.update({"from": chunk_from, "to": chunk_to})
            if token:
                kwargs["next_page_token"] = token
            resp = await ZOOM_LIST_ALL_RECORDINGS(**kwargs)
            data = resp.get("data", resp) if isinstance(resp, dict) else {}
            for m in data.get("meetings", []):
                uid = m.get("uuid", "")
                if uid and uid not in seen_uuids:
                    seen_uuids.add(uid)
                    all_meetings.append(m)
            token = data.get("next_page_token") or ""
            if not token:
                break
    return all_meetings


async def _fetch_record(i: int, meeting: dict, sem: asyncio.Semaphore) -> dict | None:
    """1件の録画からtokenを取得してレコードを返す。スキップ/失敗時はNoneを返す。"""
    async with sem:
        meeting_uuid = meeting.get("uuid", "")
        meeting_id = str(meeting.get("id", ""))
        topic = meeting.get("topic", "（無題）")
        recording_files = meeting.get("recording_files", [])

        transcript_file = next(
            (f for f in recording_files
             if f.get("file_type") == "TRANSCRIPT" or f.get("recording_type") == "audio_transcript"),
            None,
        )
        if not transcript_file:
            print(f"  (一覧{i}) ⚠ TRANSCRIPT なし、スキップ: {topic}")
            return None

        print(f"  (一覧{i}) token取得中: {topic}")
        access_token = ""
        for mid in [_encode_uuid_double(meeting_uuid), meeting_id]:
            if not mid:
                continue
            try:
                resp = await ZOOM_GET_MEETING_RECORDINGS(
                    meetingId=mid, include_fields="download_access_token"
                )
                data = resp.get("data", resp) if isinstance(resp, dict) else {}
                tok = data.get("download_access_token", "")
                if tok:
                    access_token = tok
                    break
            except Exception as e:
                print(f"    ⚠ token取得失敗 ({mid[:12]}...): {e}")

        if not access_token:
            print(f"  (一覧{i}) ⚠ token取得不可、スキップ")
            return None

        return {
            "_sort_key": i,
            "meeting_topic": topic,
            "meeting_start_time": meeting.get("start_time", ""),
            "meeting_uuid": meeting_uuid,
            "meeting_id": meeting_id,
            "host_email": meeting.get("host_email", ""),
            "download_url": transcript_file["download_url"],
            "access_token": access_token,
        }


# ===== Phase2: VTT ダウンロード（実行環境内） =====

def _download_sync(url: str, token: str) -> bytes:
    """Bearer 認証付きで URL の内容を取得する（同期。to_thread 経由で呼ぶ）"""
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    with urllib.request.urlopen(req, timeout=DOWNLOAD_TIMEOUT) as resp:
        return resp.read()


async def _download_one(rec: dict, sem: asyncio.Semaphore, *, attempts: int = 3) -> dict:
    """1件の VTT をダウンロードして tmp/vtt/{index:03d}_{uuid8}.vtt に保存する"""
    idx = rec["index"]
    topic = rec["meeting_topic"]
    uuid8 = re.sub(r'[^A-Za-z0-9]', '', rec["meeting_uuid"])[:8]
    out_path = f"tmp/vtt/{idx:03d}_{uuid8}.vtt"

    if os.path.exists(out_path):
        print(f"  [{idx:03d}] スキップ (既存): {topic}")
        return {"index": idx, "status": "exists", "path": out_path}

    async with sem:
        last = None
        for attempt in range(attempts):
            try:
                content = await asyncio.to_thread(
                    _download_sync, rec["download_url"], rec["access_token"]
                )
                head = content[:200].decode("utf-8", errors="replace")
                if len(content) < 100 or "WEBVTT" not in head:
                    raise RuntimeError(f"VTT 内容が不正: size={len(content)} 先頭={head[:80]!r}")
                with open(out_path, "wb") as f:
                    f.write(content)
                print(f"  [{idx:03d}] ✓ ダウンロード OK ({len(content)} bytes): {topic}")
                return {"index": idx, "status": "ok", "path": out_path, "size": len(content)}
            except Exception as e:
                last = e
                wait = 2 ** attempt
                print(f"  [{idx:03d}] retry {attempt+1}/{attempts} after {wait}s: {e}")
                await asyncio.sleep(wait)
        print(f"  [{idx:03d}] ⚠ ダウンロード失敗: {last}")
        return {"index": idx, "status": "error", "error": str(last)}


# ===== 参加者取得（Zoom API + VTT発言者、名前のみ） =====

async def _fetch_zoom_participants(meeting_uuid: str) -> list[dict]:
    """Zoom API から参加者の {name, email} 一覧を取得する。

    メールアドレスは取れる場合（同一アカウント/サインイン参加者等）のみ入る。
    """
    if not meeting_uuid:
        return []
    participants: list[dict] = []
    token = ""
    while True:
        kwargs: dict = {"meetingId": _encode_uuid_single(meeting_uuid), "page_size": 300}
        if token:
            kwargs["next_page_token"] = token
        try:
            resp = await ZOOM_GET_PAST_MEETING_PARTICIPANTS(**kwargs)
        except Exception as e:
            print(f"  ⚠ 参加者取得失敗: {e}")
            break
        if isinstance(resp, dict) and resp.get("successful") is False:
            print(f"  ⚠ 参加者取得エラー: {resp.get('error')}")
            break
        data = resp.get("data", resp) if isinstance(resp, dict) else {}
        for user in data.get("participants", []):
            name = (user.get("name") or "").strip()
            email = (user.get("user_email") or user.get("email") or "").strip()
            if name or email:
                participants.append({"name": name, "email": email})
        token = data.get("next_page_token") or ""
        if not token:
            break
    return participants


def _extract_speakers_from_vtt(vtt: str) -> list[str]:
    """
    VTT テキストの各行から "発言者名: テキスト" 形式の発言者名を抽出する。
    タイムスタンプ・ヘッダ・シーケンス番号行はスキップする。
    """
    speakers: set[str] = set()
    for line in vtt.splitlines():
        s = line.strip()
        if not s or s == "WEBVTT":
            continue
        if re.match(r'^\d+$', s) or re.match(r'^\d{2}:\d{2}', s):
            continue
        m = re.match(r'^(.{1,40})[:：]\s+\S', s)
        if m:
            name = m.group(1).strip()
            if not re.match(r'^\d{1,2}:\d{2}', name):
                speakers.add(name)
    return list(speakers)


# Zoom API が返す実在しないダミー/テスト用の参加者名（正規化キーで比較して除外。
# 2026-07-14 実走で "aeiou" の混入を確認）
_PARTICIPANT_BLOCKLIST = {"aeiou"}


def _merge_participants(zoom_participants: list[dict],
                        vtt_names: list[str]) -> tuple[list[str], list[str], list[str]]:
    """(参加者名リスト, メールアドレスリスト, メール未解決の名前リスト) を返す。

    - 名前: Zoom API 参加者 + VTT 発言者を正規化キー（大文字小文字・空白無視）で
      重複排除（name プロパティ用）
    - メール: 取得できた参加者分を重複排除して集める（sender プロパティ用）
    - メール未解決: 名前はあるがメールが取れなかった参加者
      （呼び出し元が人物辞書DBで逆引き補完する）
    Zoom API 側の取得は uuid だけで呼べるため、呼び出し元で他の処理と
    並行実行してからここでマージする。
    """
    entries = [((p.get("name") or "").strip(), (p.get("email") or "").strip())
               for p in zoom_participants]
    entries += [(n.strip(), "") for n in vtt_names]

    # メールが取れている名前の正規化キー（未解決判定に使う）
    keys_with_email = {_normalize(n) for n, e in entries if e and n}

    seen_norm: set[str] = set()
    seen_email: set[str] = set()
    names: list[str] = []
    emails: list[str] = []
    for name, email in entries:
        key = _normalize(name or email)
        if not key or key in _PARTICIPANT_BLOCKLIST:
            continue
        if email and email.lower() not in seen_email:
            seen_email.add(email.lower())
            emails.append(email)
        if key not in seen_norm:
            seen_norm.add(key)
            names.append(name or email)
    unresolved = [n for n in names if "@" not in n and _normalize(n) not in keys_with_email]
    return names, emails, unresolved


# ===== Notion 操作 =====

async def _query_all(db_id: str) -> list[dict]:
    rows, cursor = [], None
    while True:
        kwargs: dict = {"database_id": db_id, "page_size": 100}
        if cursor:
            kwargs["start_cursor"] = cursor
        # フィルタなし全件取得も WITH_FILTER で行う（filter は省略可・Gmail側で実績あり）
        r = await _retry(
            lambda k=kwargs: NOTION_QUERY_DATABASE_WITH_FILTER(**k),
            label=f"query({db_id[:8]})",
        )
        d = r.get("data", r) if isinstance(r, dict) else {}
        rows.extend(d.get("results", []))
        if not d.get("has_more"):
            break
        cursor = d.get("next_cursor")
        if not cursor:
            break
    return rows


async def _load_all_projects() -> list[dict]:
    """ドメイン情報DBから [{name, domain}] 一覧を返す。

    Gmail スクリプトと同じ取得元（タイトル=pjt_ prefix付き）にすることで、
    raw DB の project 列の値が全ソースで一致する。domain は参加者メールの
    ドメイン一致判定（静的ルール）に使う。
    """
    rows = await _query_all(DOMAIN_INFO_DB_ID)
    projects = []
    for row in rows:
        name = _page_title(row).strip()
        props = row.get("properties") or {}
        d = props.get(DOMAIN_PROP) or {}
        domain = "".join(x.get("plain_text", "") for x in (d.get("rich_text") or []))
        domain = domain.strip().lstrip("@").lower()
        if name:
            projects.append({"name": name, "domain": domain})
    return projects


def _match_by_topic(topic: str, projects: list[dict]) -> tuple[str, str] | None:
    """
    meeting topic にプロジェクト名が含まれていれば (project_name, strategy) を返す。
    完全一致 → 部分一致の順で最初に見つかったものを採用する。
    """
    norm_topic = _normalize(topic)
    if not norm_topic:
        return None

    for proj in projects:
        if not proj["name"]:
            continue
        if _normalize(proj["name"]) == norm_topic:
            return proj["name"], "topic完全一致"

    for proj in projects:
        if not proj["name"]:
            continue
        norm_proj = _normalize(proj["name"])
        if norm_proj and norm_proj in norm_topic:
            return proj["name"], f"topic部分一致 ({proj['name']})"

    return None


def _match_by_participant_domain(participant_emails: list[str],
                                 projects: list[dict]) -> tuple[str, str] | None:
    """参加者メールのドメインをドメイン情報DBと突き合わせる静的ルール。

    社外参加者のドメインが登録済みならそのプロジェクトに確定する。
    LLM推定より確実なため topic 一致の次・LLM の前に評価する
    （自社ドメインはDBに登録されていないため誤マッチしない）。
    """
    seen: list[str] = []
    for e in participant_emails:
        if "@" in e:
            d = e.split("@", 1)[1].lower()
            if d not in seen:
                seen.append(d)
    for d in seen:
        for proj in projects:
            if proj.get("domain") and proj["domain"] == d:
                return proj["name"], f"参加者ドメイン一致 ({d})"
    return None


async def _match_by_content(
    topic: str, transcript: str, projects: list[dict]
) -> tuple[str, str] | None:
    """
    meeting topic + VTT テキスト（抜粋）を LLM に渡してプロジェクトを推定する。
    信頼度が medium 以上の場合のみ結果を採用する。
    """
    valid_projects = [p for p in projects if p["name"]]
    if not valid_projects:
        return None

    project_list_text = "\n".join(f"- {p['name']}" for p in valid_projects)

    prompt = f"""以下の会議情報から、最も関連するプロジェクトを選んでください。

会議タイトル: {topic}

プロジェクト一覧:
{project_list_text}

【文字起こし（抜粋）】
{transcript[:3000]}

一致するプロジェクト名をそのままの表記で返してください。
判断が難しい場合や一致するものがない場合は project_name を空文字列にしてください。
"""
    schema = {
        "type": "object",
        "properties": {
            "project_name": {
                "type": "string",
                "description": "最も一致するプロジェクト名。なければ空文字列。",
            },
            "confidence": {
                "type": "string",
                "enum": ["high", "medium", "low"],
                "description": "判断の信頼度。不明・複数候補あり・判断根拠薄の場合は low。",
            },
            "reason": {"type": "string", "description": "判断根拠（短く）"},
        },
        "required": ["project_name", "confidence"],
    }

    try:
        result = await llm_call(prompt=prompt, schema=schema)
        data = result["data"]
        matched_name = (data.get("project_name") or "").strip()
        confidence = data.get("confidence", "low")
        reason = data.get("reason", "")
        print(f"  内容推定: '{matched_name}' 信頼度={confidence} 理由={reason}")

        if not matched_name or confidence == "low":
            print("  ⚠ 内容推定を棄却 (信頼度が low またはプロジェクト名なし)")
            return None

        for proj in valid_projects:
            if proj["name"] == matched_name:
                return proj["name"], f"内容推定 (信頼度:{confidence})"

        norm_matched = _normalize(matched_name)
        for proj in valid_projects:
            if _normalize(proj["name"]) == norm_matched:
                return proj["name"], f"内容推定 (信頼度:{confidence})"

    except Exception as e:
        print(f"  ⚠ LLM プロジェクト推定失敗: {e}")

    return None


def _all_participants_internal(participant_emails: list[str], host_email: str) -> bool:
    """参加者全員がホストと同じドメイン（＝自社の人だけ）かどうかを判定する。

    ★2026-07-25 実測で判明: 参加者が全員社内（@example.com）の定例MTGが、
      LLM推定で無関係な顧客プロジェクト（pjt_cre 等）に誤って割り当てられた。
      社外の参加者が1人もいない会議はどの顧客のものでもない、という単純で
      確実な事実があるため、この場合は LLM に推測させず未分類にする。
    ホストのドメインが取れない場合は判定不能として False（＝従来どおりLLMに委ねる）。
    """
    host_domain = (host_email or "").split("@")[-1].strip().lower()
    if not host_domain:
        return False
    domains = {e.split("@", 1)[1].strip().lower() for e in participant_emails if "@" in e}
    return bool(domains) and domains <= {host_domain}


async def _find_project(
    topic: str, transcript: str, projects: list[dict],
    participant_emails: list[str], *, host_email: str = "",
) -> tuple[str, str]:
    """
    プロジェクトを次の優先順で判定し (project_name, strategy) を返す:
      1. topic 一致
      2. 参加者メールのドメイン一致（静的ルール。LLMより確実）
      3. 参加者が全員自社の人だけ → 未分類（LLM推定はしない。社内会議の誤判定防止）
      4. 内容 (LLM) 推定
    どれも一致しなければ未分類で処理を続行する。
    """
    result = _match_by_topic(topic, projects)
    if result:
        return result

    result = _match_by_participant_domain(participant_emails, projects)
    if result:
        return result

    if _all_participants_internal(participant_emails, host_email):
        print(f"  参加者が全員自社ドメインのため '未分類' で保存（LLM推定はスキップ）")
        return "未分類", "社内会議（外部参加者なし）"

    result = await _match_by_content(topic, transcript, projects)
    if result:
        return result

    print(f"  ⚠ プロジェクト未特定: topic='{topic}' → '未分類' で保存")
    return "未分類", "フォールバック"


def _heading_block(level: str, text: str) -> dict:
    return {"object": "block", "type": level,
            level: {"rich_text": [{"text": {"content": text}}]}}


def _build_child_blocks(minutes_text: str, vtt_raw: str) -> list[dict]:
    """ページ本文のブロック（議事録＋生VTT）を Notion 生API形式で構築する。

    content プロパティにも議事録を入れるが、本文にも常に両方を残す
    （VTT の行き先はページ本文＝2026-07-14 ユーザー確定）。
    """
    blocks = [_heading_block("heading_2", "議事録")]
    for chunk in _split_text(minutes_text):
        blocks.append({"object": "block", "type": "paragraph",
                       "paragraph": {"rich_text": [{"text": {"content": chunk}}]}})
    blocks.append({"object": "block", "type": "divider", "divider": {}})
    blocks.append(_heading_block("heading_3", "生の文字起こし (VTT)"))
    for chunk in _split_text(vtt_raw):
        blocks.append({"object": "block", "type": "paragraph",
                       "paragraph": {"rich_text": [{"text": {"content": chunk}}]}})
    return blocks


# ===== VTT 前処理 =====

def _preprocess_vtt(vtt: str) -> str:
    """
    タイムスタンプ・ヘッダ・シーケンス番号を除去し、
    ローリングキャプション（前行が次行の前方一致）の重複を削除する。
    """
    lines = []
    for line in vtt.splitlines():
        s = line.strip()
        if not s or s == "WEBVTT":
            continue
        if re.match(r'^\d+$', s):
            continue
        if re.match(r'^\d{2}:\d{2}:\d{2}', s):
            continue
        lines.append(s)

    deduped = []
    for i, line in enumerate(lines):
        text = re.sub(r'^[^:：]+[:：]\s*', '', line)
        next_line = lines[i + 1] if i + 1 < len(lines) else ""
        next_text = re.sub(r'^[^:：]+[:：]\s*', '', next_line)
        if text and next_text.startswith(text):
            continue
        deduped.append(line)

    result = "\n".join(deduped)
    if len(result) > _MAX_TRANSCRIPT_CHARS:
        # 3時間級の会議などで LLM のコンテキスト/コストが爆発しないよう上限を設ける
        print(f"⚠ 文字起こしが {len(result)} 字と長いため {_MAX_TRANSCRIPT_CHARS} 字に切り詰め")
        result = result[:_MAX_TRANSCRIPT_CHARS] + "\n（以降省略）"
    return result


# ===== LLM 議事録生成 =====

def _build_prompt(topic: str, meeting_at: str, transcript: str) -> str:
    return f"""以下のZoom会議の文字起こしを元に、詳細な議事録を作成してください。

会議タイトル: {topic}
開始日時: {meeting_at}

【出力フォーマット（この順序・見出しで必ず出力）】

## 会議概要（テーマ）
（会議の目的・背景・主な議題を2〜3文で簡潔に）

## 主な議論内容
- 議論されたトピックを箇条書きで

## 決定事項
- 会議で合意・決定された内容を箇条書きで

## 課題・懸念点
- 未解決の問題、懸念、リスクを箇条書きで

## 次のアクション
- 担当者: タスク内容（期限があれば記載）

## 参加者の主要発言
- 発言者名: 印象的・重要な発言を要約

【重要】
- 前置き・挨拶・断り書きは一切書かず、必ず「## 会議概要（テーマ）」の見出しから直接出力すること
- 文字起こしから読み取れない内容は創作せず、不明な箇所は「不明」と明記すること

【文字起こし（前処理済み）】
{transcript}
"""


async def _summarize(topic: str, meeting_at: str, transcript: str) -> str:
    last_err = None
    for attempt in range(3):
        try:
            result = await llm_call(prompt=_build_prompt(topic, meeting_at, transcript))
            text = result["text"] or ""
            # プロンプトで禁止していても前置きが付くことがあるため、
            # 最初の見出しより前を捨てる（2026-07-14 実走で混入を確認）
            head = text.find("## ")
            if head > 0:
                text = text[head:]
            if text and len(text) > 50:
                return text
            raise RuntimeError(f"LLM 応答短すぎ: {len(text)} chars")
        except Exception as e:
            last_err = e
            wait = 2 ** attempt
            print(f"  LLM retry {attempt+1}/3 after {wait}s: {e}")
            await asyncio.sleep(wait)
    raise RuntimeError(f"LLM が 3 回失敗: {last_err}")


# ===== Phase2.5: 文字起こしだけ保存（LLM を呼ばない＝軽い） =====

def _build_vtt_blocks(vtt_raw: str) -> list[dict]:
    """ページ本文に入れる「生の文字起こし」ブロックを作る。

    ★このブロックが Phase3（議事録生成）の入力になる＝Notion が中間状態の置き場。
      Notion の children 上限（100）に対して MAX_VTT_BLOCKS で頭打ちにする。
    """
    blocks = [_heading_block("heading_3", "生の文字起こし (VTT)")]
    chunks = _split_text(vtt_raw)
    if len(chunks) > MAX_VTT_BLOCKS - 1:
        print(f"  ⚠ VTTが長いためブロックを{MAX_VTT_BLOCKS - 1}個で打ち切り（全{len(chunks)}個）")
        chunks = chunks[:MAX_VTT_BLOCKS - 1]
    for chunk in chunks:
        blocks.append({"object": "block", "type": "paragraph",
                       "paragraph": {"rich_text": [{"text": {"content": chunk}}]}})
    return blocks


async def _existing_keys_in_month(month_from: str, month_to: str) -> set[str]:
    """その月に保存済みの会議の parents（=meeting_uuid）を集める。

    ★「取り込み済みかどうか」は Notion に行があるかで判断する＝これが tmp のしおりの代わり。
      raw DB 全体ではなくその月だけを聞くので、DBが何万行に育っても
      1ヶ月あたり1クエリで一定（全件取得だと会議3000件で30クエリ＝1分近くかかる）。
    """
    keys: set[str] = set()
    cursor: str | None = None
    while True:
        kwargs: dict = {
            "database_id": RAW_DB_ID,
            "filter": {"and": [
                {"property": PROP_SOURCE, "select": {"equals": SOURCE_MEETING}},
                {"property": PROP_TIMESTAMP, "date": {"on_or_after": month_from}},
                {"property": PROP_TIMESTAMP, "date": {"on_or_before": month_to}},
            ]},
            "page_size": 100,
        }
        if cursor:
            kwargs["start_cursor"] = cursor
        resp = await _retry(
            lambda kw=kwargs: NOTION_QUERY_DATABASE_WITH_FILTER(**kw), label="existing_keys"
        )
        data = resp.get("data", resp) if isinstance(resp, dict) else {}
        for row in data.get("results", []):
            key = _row_text(row.get("properties") or {}, PROP_PARENTS)
            if key:
                keys.add(key)
        if not data.get("has_more"):
            break
        cursor = data.get("next_cursor")
        if not cursor:
            break
    return keys


# ===== 進捗記録（sync_state DB に1行だけ持ち、実行のたびに上書きする） =====

async def _load_sync_state(key: str) -> dict:
    """sync_state から進捗を読む。行が無い/壊れている場合は空dict（＝最初から）。"""
    if not SYNC_STATE_DB_ID:
        return {}
    try:
        resp = await _retry(
            lambda: NOTION_QUERY_DATABASE_WITH_FILTER(
                database_id=SYNC_STATE_DB_ID,
                filter={"property": PROP_S_KEY, "title": {"equals": key}},
                page_size=1,
            ),
            label="load_sync_state",
        )
        data = resp.get("data", resp) if isinstance(resp, dict) else {}
        rows = data.get("results") or []
        if not rows:
            return {}
        raw = _row_text(rows[0].get("properties") or {}, PROP_S_VALUE)
        return json.loads(raw) if raw else {}
    except Exception as e:
        print(f"⚠ 進捗の読み込みに失敗（最初から走査します）: {e}")
        return {}


async def _save_sync_state(key: str, value: dict, *, note: str = "") -> None:
    """sync_state の同じ行を上書きする（失敗しても本体を止めない）。"""
    if not SYNC_STATE_DB_ID:
        return
    props = {
        PROP_S_KEY:   {"title":     _rt(key)},
        PROP_S_VALUE: {"rich_text": _rt(json.dumps(value, ensure_ascii=False))},
        PROP_S_NOTE:  {"rich_text": _rt(note)},
    }
    try:
        await _retry(
            lambda: NOTION_UPSERT_ROW_DATABASE(
                database_id=SYNC_STATE_DB_ID,
                items=[{
                    "match":  {"property": PROP_S_KEY, "equals": key},
                    "create": {"properties": props},
                    "update": {"properties": props},
                }],
            ),
            label="save_sync_state",
        )
    except Exception as e:
        print(f"⚠ 進捗の保存に失敗（次回は走査をやり直します）: {e}")


def _row_text(props: dict, name: str) -> str:
    """Notion の行プロパティから rich_text を平文で取り出す。"""
    rt = (props.get(name) or {}).get("rich_text") or []
    return "".join(t.get("plain_text", "") for t in rt).strip()


async def _save_transcript_row(rec: dict, vtt_raw: str) -> dict:
    """1会議の文字起こしを raw DB に新規作成する（議事録はまだ作らない）。

    project=MARKER_PROJECT を目印に置き、Phase3 がこれを拾う。
    ★update は parents だけ＝既存行（すでに議事録がある行）を絶対に上書きしない。
    """
    topic = rec["meeting_topic"]
    uid = rec["meeting_uuid"]
    dedup_key = uid or f"zoom:{rec.get('meeting_start_time', '')}:{topic}"
    meeting_at = _to_jst(rec.get("meeting_start_time", ""))

    props = {
        # 議事録が入るまでは会議タイトルを暫定表示（Phase3 で議事録全文に置き換わる）
        PROP_CONTENT: {"title":     _rt(topic)},
        PROP_SUBJECT: {"rich_text": _rt(topic)},
        PROP_SENDER:  {"rich_text": _rt(rec.get("host_email", ""))},
        PROP_PROJECT: {"rich_text": _rt(MARKER_PROJECT)},
        PROP_SOURCE:  {"select":    {"name": SOURCE_MEETING}},
        PROP_PARENTS: {"rich_text": _rt(dedup_key)},
    }
    if meeting_at:
        props[PROP_TIMESTAMP] = {"date": {"start": meeting_at}}

    try:
        result = await _retry(
            lambda: NOTION_UPSERT_ROW_DATABASE(
                database_id=RAW_DB_ID,
                items=[{
                    "match":  {"property": PROP_PARENTS, "equals": dedup_key},
                    "create": {"properties": props, "children": _build_vtt_blocks(vtt_raw)},
                    "update": {"properties": {PROP_PARENTS: {"rich_text": _rt(dedup_key)}}},
                }],
            ),
            label=f"ingest({dedup_key[:12]})",
        )
    except Exception as e:
        return {"status": "error_notion", "topic": topic, "error": str(e)}
    err = _upsert_failed_detail(result)
    if err:
        return {"status": "error_notion", "topic": topic, "error": err}
    return {"status": "ingested", "topic": topic, "key": dedup_key}


# ===== Phase3: 未処理キューを消化（ページ本文のVTT → 議事録） =====

async def _read_vtt_from_page(page_id: str) -> str:
    """ページ本文のブロックから生の文字起こしを復元する（Phase3 の入力）。

    「文字起こし」を含む見出し以降の paragraph を連結する。
    見出しが無い場合は全 paragraph を連結する（旧フローで作られた行への保険）。

    ★このツールに start_cursor は無い（引数は block_id/page_url/page_size/recursive/
      max_depth/max_blocks のみ）。実行環境は未知引数を黙って無視するため、カーソルを
      渡してループすると同じ先頭ページを無限に取り続ける。よって**1回だけ呼ぶ**。
      1段目が書き込むブロックは見出し1＋VTT最大 MAX_VTT_BLOCKS 個＝100未満に収めてあるため、
      page_size=100 の1回で全量が取れる。
    """
    resp = await _retry(
        lambda: NOTION_FETCH_ALL_BLOCK_CONTENTS(block_id=page_id, page_size=100),
        label="read_vtt",
    )
    if isinstance(resp, list):
        blocks, has_more = resp, False
    else:
        data = resp.get("data", resp) if isinstance(resp, dict) else {}
        blocks = data.get("results", []) or []
        has_more = bool(data.get("has_more"))
    if has_more:
        # 想定外（旧フローで作られた行など）。取れた分だけで議事録を作る
        print("    ⚠ ブロックが100個を超えています（取得できた分だけで議事録を作成します）")

    texts: list[str] = []
    started = False
    saw_heading = False
    for b in blocks:
        btype = b.get("type", "")
        rich = (b.get(btype) or {}).get("rich_text", []) if btype else []
        text = "".join(t.get("plain_text", "") for t in rich)
        if btype.startswith("heading"):
            saw_heading = True
            started = "文字起こし" in text
            continue
        if btype == "paragraph" and (started or not saw_heading):
            texts.append(text)
    return "\n".join(texts)


async def _fetch_pending_rows(limit: int = 20) -> list[dict]:
    """project=MARKER_PROJECT の行（=議事録がまだ無い行）を古い順に取得する。

    ★これが「未処理キュー」。Notion が状態を持つので tmp/ のしおりが要らない。
    """
    resp = await _retry(
        lambda: NOTION_QUERY_DATABASE_WITH_FILTER(
            database_id=RAW_DB_ID,
            filter={"and": [
                {"property": PROP_PROJECT, "rich_text": {"equals": MARKER_PROJECT}},
                {"property": PROP_SOURCE, "select": {"equals": SOURCE_MEETING}},
            ]},
            sorts=[{"property": PROP_TIMESTAMP, "direction": "ascending"}],
            page_size=min(max(limit, 1), 100),
        ),
        label="pending_rows",
    )
    data = resp.get("data", resp) if isinstance(resp, dict) else {}
    rows: list[dict] = []
    for page in data.get("results", []):
        props = page.get("properties") or {}
        title_rt = (props.get(PROP_CONTENT) or {}).get("title") or []
        rows.append({
            "page_id":   page.get("id", ""),
            "url":       page.get("url", ""),
            "topic":     _row_text(props, PROP_SUBJECT)
                         or "".join(t.get("plain_text", "") for t in title_rt).strip(),
            "key":       _row_text(props, PROP_PARENTS),
            "host":      _row_text(props, PROP_SENDER),
            "timestamp": ((props.get(PROP_TIMESTAMP) or {}).get("date") or {}).get("start", "") or "",
        })
    return rows


async def _mark_project(key: str, value: str) -> None:
    """既存行の project 列だけを書き換える（未処理キューから外すため）。"""
    try:
        await _retry(
            lambda: NOTION_UPSERT_ROW_DATABASE(
                database_id=RAW_DB_ID,
                items=[{
                    "match":  {"property": PROP_PARENTS, "equals": key},
                    "create": {"properties": {PROP_PARENTS: {"rich_text": _rt(key)}}},
                    "update": {"properties": {PROP_PROJECT: {"rich_text": _rt(value)}}},
                }],
            ),
            label="mark_project",
        )
    except Exception as e:
        print(f"    ⚠ project の書き換えに失敗（次回も未処理として拾われます）: {e}")


async def _summarize_pending_row(row: dict, projects: list[dict]) -> dict:
    """未処理行1件を議事録化して同じ行を更新する（Phase3 の本体）。

    VTT はページ本文から読むので tmp/ に依存しない。成功すると project が
    本当のPJ名に変わるため、この行は未処理キューから自動的に外れる。
    """
    topic = row["topic"]
    key = row["key"]
    meeting_at = row.get("timestamp") or ""
    # parents が "zoom:..." 形式のときは uuid ではないので参加者APIは呼べない
    meeting_uuid = key if key and not key.startswith("zoom:") else ""
    print(f"  ▶ {topic} ({meeting_at})")

    vtt_raw = await _read_vtt_from_page(row["page_id"])
    transcript = _preprocess_vtt(vtt_raw)
    _log("vtt_restored", topic=topic, page=row["page_id"][:8],
         body_chars=len(vtt_raw), transcript_chars=len(transcript))
    if len(transcript) < _MIN_TRANSCRIPT_CHARS:
        # 文字起こしが実質空。毎回キューの先頭で詰まらないよう目印を変えて外す
        _log("skipped_short_vtt", f"    スキップ (文字起こしが短すぎる: {len(transcript)}字)",
             topic=topic, transcript_chars=len(transcript))
        await _mark_project(key, MARKER_SKIPPED)
        return {"status": "skipped_short_vtt", "topic": topic}

    # 参加者取得は uuid だけで呼べるため LLM 議事録生成と並行して開始する
    participants_task = (
        asyncio.create_task(_fetch_zoom_participants(meeting_uuid)) if meeting_uuid else None
    )
    try:
        minutes_text = await _summarize(topic, meeting_at, transcript)
    except Exception as e:
        if participants_task:
            participants_task.cancel()
        # 一時的な失敗は project を変えない＝次回の再実行で再挑戦する
        return {"status": "error_llm", "topic": topic, "error": str(e)}

    zoom_participants = await participants_task if participants_task else []
    participants, participant_emails, unresolved_names = _merge_participants(
        zoom_participants, _extract_speakers_from_vtt(vtt_raw)
    )
    # メールが取れなかった参加者は人物辞書で逆引き補完（一意ヒット時のみ）
    resolved_people = await _resolve_emails_from_people(unresolved_names)
    for r in resolved_people:
        if r["email"].lower() not in {e.lower() for e in participant_emails}:
            participant_emails.append(r["email"])
    host_email = (row.get("host") or "").strip()
    if host_email and host_email.lower() not in {e.lower() for e in participant_emails}:
        participant_emails.insert(0, host_email)

    project_name, strategy = await _find_project(
        topic, transcript, projects, participant_emails, host_email=host_email
    )
    print(f"    参加者{len(participants)}名 / プロジェクト: {project_name} ({strategy})")

    props = {
        PROP_CONTENT: {"title":     _rt(minutes_text)},
        PROP_SUBJECT: {"rich_text": _rt(topic)},
        PROP_SENDER:  {"rich_text": _rt(", ".join(participant_emails))},
        PROP_NAME:    {"rich_text": _rt(", ".join(participants))},
        # ★これで未処理キューから外れる
        PROP_PROJECT: {"rich_text": _rt(project_name)},
        PROP_SOURCE:  {"select":    {"name": SOURCE_MEETING}},
        PROP_PARENTS: {"rich_text": _rt(key)},
    }
    if meeting_at:
        props[PROP_TIMESTAMP] = {"date": {"start": meeting_at}}

    # 本文（VTT）は Phase2.5 の create 時に入っているので children は渡さない
    try:
        result = await _retry(
            lambda: NOTION_UPSERT_ROW_DATABASE(
                database_id=RAW_DB_ID,
                items=[{
                    "match":  {"property": PROP_PARENTS, "equals": key},
                    "create": {"properties": props},
                    "update": {"properties": props},
                }],
            ),
            label=f"summarize({key[:12]})",
        )
    except Exception as e:
        return {"status": "error_notion", "topic": topic, "error": str(e)}
    err = _upsert_failed_detail(result)
    if err:
        return {"status": "error_notion", "topic": topic, "error": err}

    # 人物辞書用に参加者＋ホストを収集（書き込みは run_summarize の最後にまとめて行う）
    _people_buffer.extend(
        list(zoom_participants) + resolved_people + [{"name": "", "email": host_email}]
    )
    print(f"    ✓ 議事録を保存（{len(minutes_text)}字）")
    return {"status": "ok", "topic": topic, "project": project_name, "page_url": row.get("url")}


# ===== エントリポイント =====

async def _ingest_one_month(month_from: str, month_to: str, *, seq_start: int,
                            sem: asyncio.Semaphore, stats: dict) -> int:
    """1ヶ月ぶんを取り込む（一覧→既存除外→token→VTT DL→文字起こし保存）。次の seq を返す。

    ★token は短命なので「取ったらすぐ使う」を月単位で閉じる
      （旧設計のように全期間の token を先に集めると、使う前に失効する）。
    """
    label = month_from[:7]
    meetings = await _list_all_recordings(month_from, month_to)
    if not meetings:
        _log("month_empty", f"  {label}: 録画なし", month=label)
        return seq_start

    done_keys = await _existing_keys_in_month(month_from, month_to)
    pending = [m for m in meetings if (m.get("uuid") or "") not in done_keys]
    if not pending:
        _log("month_all_done", f"  {label}: 録画{len(meetings)}件（すべて取り込み済み）",
             month=label, recordings=len(meetings), already=len(done_keys))
        return seq_start
    _log("month_start", f"  {label}: 録画{len(meetings)}件 / 未取り込み{len(pending)}件",
         month=label, recordings=len(meetings), already=len(done_keys), pending=len(pending))

    seq = seq_start
    for i in range(0, len(pending), INGEST_CHUNK):
        chunk = pending[i:i + INGEST_CHUNK]
        raw = await asyncio.gather(*[_fetch_record(j, m, sem) for j, m in enumerate(chunk)])
        records = [r for r in raw if r is not None]
        for r in records:
            r.pop("_sort_key", None)
            r["index"] = seq
            seq += 1
        stats["skipped_no_vtt"] += len(chunk) - len(records)   # TRANSCRIPT無し/token取得不可
        if not records:
            continue

        dls = await asyncio.gather(*[_download_one(r, sem) for r in records])
        for rec, dl in zip(records, dls):
            if dl.get("status") not in ("ok", "exists"):
                stats["skipped_no_vtt"] += 1
                _log("vtt_download_failed", month=label, topic=rec["meeting_topic"],
                     uuid=rec["meeting_uuid"][:16], error=dl.get("error", ""))
                continue
            try:
                with open(dl["path"], encoding="utf-8", errors="replace") as f:
                    vtt_raw = f.read()
            except OSError as e:
                _log("vtt_read_failed", f"    ⚠ VTT読み込み失敗: {e}",
                     month=label, topic=rec["meeting_topic"], error=str(e))
                stats["skipped_no_vtt"] += 1
                continue
            res = await _save_transcript_row(rec, vtt_raw)
            if res["status"] == "ingested":
                stats["ingested"] += 1
                _log("meeting_ingested", f"    ✓ 取り込み: {res['topic']}",
                     month=label, topic=res["topic"], uuid=rec["meeting_uuid"][:16],
                     vtt_chars=len(vtt_raw), blocks=min(len(_split_text(vtt_raw)), MAX_VTT_BLOCKS - 1) + 1)
            else:
                stats["errors"] += 1
                _log("meeting_ingest_error",
                     f"    ✗ 取り込み失敗: {res['topic']} - {res.get('error', '')}",
                     month=label, topic=res["topic"], error=res.get("error", ""))
    return seq


async def run_ingest(from_date: str, to_date: str, *,
                     budget_seconds: float = TIME_BUDGET_SECONDS) -> dict:
    """1段目: **新しい月から古い月へ順に走査**して、文字起こしだけ raw DB に保存する。

    LLM を呼ばないので軽い。時間予算で自分から止まるので「✅ 全期間の走査完了」まで再実行する。

    ★月単位で処理する理由: 期間を一括で扱うと、一覧取得と既存チェックだけで時間予算を
      使い切って1件も処理できず、再実行しても同じことを繰り返して永遠に進まなくなる
      （2026-07-25 に実際に発生）。月ごとなら1回の実行で必ず前進する。
    ★進捗（次に処理する月）は **sync_state DB** に1行だけ書いて上書きする。
      tmp/ はセッションが切れると消えるが、ここに書けば何度死んでも続きから走査できる。
    """
    global _run_id
    t0 = time.monotonic()
    _run_id = "ing-" + datetime.now(JST).strftime("%m%d-%H%M%S")
    _log("run_ingest_start", f"=== 取り込み（文字起こし）{from_date} ～ {to_date} ===",
         from_date=from_date, to_date=to_date, budget=budget_seconds)
    os.makedirs("tmp/vtt", exist_ok=True)

    months = list(reversed(_month_chunks(from_date, to_date)))   # 新しい月が先頭
    range_key = f"{from_date}~{to_date}"

    # ---- 進捗を Notion から読み、続きの月を決める ----
    state = await _load_sync_state(SWEEP_STATE_KEY)
    start = 0
    if state.get("range") == range_key:
        if state.get("done"):
            print("✅ この期間は走査済み（もう一度やり直すなら sync_state の行を消してください）")
            return {"phase": "ingest", "done": True, "remaining_months": 0,
                    "ingested": 0, "skipped_no_vtt": 0, "errors": 0}
        next_month = state.get("next_month")
        if next_month:
            for i, (mf, _) in enumerate(months):
                if mf[:7] == next_month:
                    start = i
                    break
    print(f"走査対象: {len(months) - start}ヶ月（{months[start][0][:7]} から古い方へ）"
          + ("" if start == 0 else " ※前回の続き"))

    sem = asyncio.Semaphore(PARALLEL)
    stats = {"ingested": 0, "skipped_no_vtt": 0, "errors": 0}
    seq, idx = 0, start
    while idx < len(months):
        if budget_seconds - (time.monotonic() - t0) <= INGEST_COST_EST:
            break
        month_from, month_to = months[idx]
        seq = await _ingest_one_month(month_from, month_to,
                                      seq_start=seq, sem=sem, stats=stats)
        idx += 1
        # ★1ヶ月終えるたびに進捗を上書き保存（次回はここから）
        done = idx >= len(months)
        await _save_sync_state(
            SWEEP_STATE_KEY,
            {"range": range_key, "next_month": None if done else months[idx][0][:7], "done": done},
            note=(f"{range_key} 走査完了" if done
                  else f"次は {months[idx][0][:7]} から（{from_date}〜{to_date}）"),
        )
        _log("month_done", month=month_from[:7],
             next_month=None if done else months[idx][0][:7],
             elapsed=round(time.monotonic() - t0, 1), **stats)

    remaining_months = len(months) - idx
    el = time.monotonic() - t0
    _log("run_ingest_end",
         f"\n取り込み: 新規{stats['ingested']} / VTTなし{stats['skipped_no_vtt']} / "
         f"失敗{stats['errors']} / {idx - start}ヶ月を走査 / {el:.1f}秒",
         months_swept=idx - start, remaining_months=remaining_months,
         elapsed=round(el, 1), **stats)
    print("✅ 全期間の走査完了" if remaining_months == 0
          else f"→ 残り{remaining_months}ヶ月（次は{months[idx][0][:7]}）。"
               f"もう一度 run_ingest を実行してください")
    return {"phase": "ingest", "done": remaining_months == 0,
            "remaining_months": remaining_months, **stats}


async def run_summarize(*, budget_seconds: float = TIME_BUDGET_SECONDS) -> dict:
    """2段目: project="未処理" の行を1件ずつ議事録化する（LLM）。

    ★Notion が唯一の状態なので、セッションが何度死んでも
      「未処理が無くなるまで再実行」で必ず完走する。
    """
    t0 = time.monotonic()
    _people_buffer.clear()
    print("=== 議事録生成（未処理キューの消化） ===")
    projects = await _load_all_projects()
    print(f"プロジェクト候補: {len(projects)} 件")

    results: list[dict] = []
    attempted: set[str] = set()   # 同一run内で同じ行を何度も掴まないため（失敗行での無限ループ防止）
    queue: list[dict] = []
    while budget_seconds - (time.monotonic() - t0) > SUMMARIZE_COST_EST:
        queue = [r for r in queue if r["page_id"] not in attempted]
        if not queue:
            queue = [r for r in await _fetch_pending_rows(limit=20)
                     if r["page_id"] not in attempted]
            if not queue:
                break
        row = queue.pop(0)
        attempted.add(row["page_id"])
        results.append(await _summarize_pending_row(row, projects))

    if _people_buffer:
        print(f"\n人物辞書を更新中...（収集 {len(_people_buffer)} 件）")
        await _upsert_people(_people_buffer)

    ok = sum(1 for r in results if r["status"] == "ok")
    skipped = sum(1 for r in results if r["status"].startswith("skipped"))
    errors = sum(1 for r in results if r["status"].startswith("error"))
    rest = len(await _fetch_pending_rows(limit=100))
    el = time.monotonic() - t0
    print(f"\n議事録: OK={ok} / スキップ={skipped} / エラー={errors} / {el:.1f}秒")
    for r in results:
        if r["status"].startswith("error"):
            print(f"  ✗ {r['status']}: {r['topic']} - {r.get('error', '')}")
    print("✅ 議事録完了（未処理なし）" if rest == 0
          else f"→ 未処理あと{rest}件以上。もう一度 run_summarize を実行してください")
    return {"phase": "summarize", "done": rest == 0, "pending": rest,
            "ok": ok, "skipped": skipped, "errors": errors, "results": results}


async def run(from_date: str, to_date: str) -> dict:
    """互換エントリ: 取り込み → 議事録生成 を1回の時間予算内で続けて実行する。

    どちらも途中で止まりうるので「✅ 全完了」が出るまで繰り返し実行する。
    ※取り込みだけ・議事録だけを回したいときは run_ingest / run_summarize を直接呼ぶ。
    """
    t0 = time.monotonic()
    ing = await run_ingest(from_date, to_date)
    rest_budget = TIME_BUDGET_SECONDS - (time.monotonic() - t0)
    if rest_budget > SUMMARIZE_COST_EST:
        sm = await run_summarize(budget_seconds=rest_budget)
    else:
        print(f"\n（残り{rest_budget:.0f}秒のため議事録生成は次回に回します）")
        sm = {"phase": "summarize", "done": False, "deferred": True}
    done = bool(ing.get("done") and sm.get("done"))
    print("\n✅ 全完了" if done else "\n→ 続きあり。もう一度 run() を実行してください")
    return {"done": done, "ingest": ing, "summarize": sm}
