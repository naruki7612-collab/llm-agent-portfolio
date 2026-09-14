"""agent_platform DB（DataTable）＋ Note を横断検索する検索OS v4。

notion_集約_v3.py の agent_platform 版。**検索の仕様は v3 と同じ**（レシピ・キーワード横断・
人物名寄せ・スレッド復元・深読み・出典必須の合成）。変わったのは「どこに問い合わせるか」と
「サーバーにできない処理をどこでやるか」だけ。

★v3 からの実装上の変更点（agent_platform 側の制約から来るもの・仕様変更ではない）
  1. **contains（部分一致）と $or が無い**。find() が受けるのは
     完全一致 / $gt,$gte,$lt,$lte,$ne / $in / トップレベル複数キー=AND だけ。
     → v3 の `_keyword_group` / `person_conditions`（contains の OR 展開）は
       **同じ意味論のまま Python 側の照合に移した**（_keyword_match / _person_match）。
  2. **sort を付けるとページングできない**（after_id 併用不可）。
     → キーワード無し = sort降順で1発取得 / キーワード有り = after_id で全走査して
       ローカル照合→降順、と取得モードを切り替える（_top_rows / _scan_rows）。
  3. **find({"_id": ...}) は backend が 422 で明示拒否**（行IDでの引き当て不可）。
     → Relation から行を復元することはできない。添付はメール行の links に入っている
       note_path の完全一致で引く（_fetch_attachment_rows）。
  4. **Notion のページ本文が無い**。meeting の議事録全文は content 列にそのまま入っている。
     → 深読みは「file 行の note 実体を読む」に置き換えた。notes/ は自分たちの資産なので、
       v3 の「外部ファイルは検索時に再取得しない」原則には反しない。
  5. DB ID 定数が無い。テーブルは会話にアタッチされた Project 配下に自動スコープされる。
     ＝ **Project が attach された会話でしか動かない**（未 attach は 422 で落ちる）。

構成（v3 と同じ）:
  取得層: fetch_records / search（自然言語）
  解決層: resolve_person（人物辞書3層・複数人は候補提示）/ resolve_company / company_members
  文脈層: expand_thread（parents をソース別に解釈）/ expand_attachments（links → file行）
  深読み層: read_note / fetch_full_text（meeting=content全文・file=note実体）
  合成層: answer_query（プランナー→取得→再計画→深読み→出典必須の構造化回答）

【呼び出し方（code_execute）】質問はトリプルシングルクォートで直接埋め込む
（二重引用符3連続はこのモジュールの docstring を壊した事故実績があるため使わない）:
  import sys
  if "rag" not in sys.path:
      sys.path.insert(0, "rag")
  from agent_platform_集約_v4 import answer_query
  result = await answer_query('''アルファリースの見積もりいくらだった？''')
  print(result)
"""

from __future__ import annotations

import asyncio
import os
import random
import re
import unicodedata
from datetime import datetime, timedelta, timezone

from agent_sdk import llm_call, ToolCallError
from agent_sdk import db as adb

# --- デバッグ出力（実行サンドボックスは stdout のみ返るため print ベース） ---
DEBUG = os.environ.get("ASKHUB_DEBUG", "1") != "0"


def _dbg(msg: str) -> None:
    if DEBUG:
        print(f"[DEBUG] {msg}", flush=True)


def _warn(msg: str) -> None:
    """異常系は DEBUG フラグに関係なく常に出す。"""
    print(f"[WARN] {msg}", flush=True)


# ========== テーブル名（書き込み側 agent_platform_db_init.py と揃えること） ==========

RAW_TABLE        = "raw"
PEOPLE_TABLE     = "people"
DOMAIN_MAP_TABLE = "domain_map"

JST = timezone(timedelta(hours=9))


# ========== raw テーブルのセマンティクス表（v4の心臓部） ==========
# 「同じ列でもソースごとに意味が違う」の解釈は必ずこの表を経由する。

_ALL_SOURCES: list[str] = ["slack", "gmail", "drive", "meeting", "file"]

# ソース別: キーワード横断検索の対象フィールド
# （drive/file の parents はフォルダ名・出どころなので検索対象に含める。
#  他ソースの parents は ts/threadId/uuid でキーワード検索には無意味なので含めない）
_BODY_FIELDS_BY_SOURCE: dict[str, list[str]] = {
    "slack":   ["content", "links"],
    "gmail":   ["content", "subject"],
    "drive":   ["content", "subject", "parents"],
    "meeting": ["content", "subject"],
    # file = Gmail 添付を notes/ に置いた行。drive と同じ扱い
    # （content=要約JSON / subject=ファイル名 / parents=出どころ）
    "file":    ["content", "subject", "parents"],
}

# ソース別: 人物（名前・メール）を探すフィールド（resolve_person の展開先）
_PERSON_FIELDS_BY_SOURCE: dict[str, dict[str, list[str]]] = {
    "slack":   {"email": ["sender"], "name": ["name"]},
    "gmail":   {"email": ["sender", "to", "cc_bcc"], "name": ["name"]},
    "drive":   {"email": ["sender"], "name": ["name"]},
    "meeting": {"email": ["sender"], "name": ["name"]},
    "file":    {"email": ["sender"], "name": ["name"]},
}

# parents 列の意味（検索OSはこの表を通してだけ解釈する）
PARENTS_SEMANTICS: dict[str, str] = {
    "slack":   "thread_parent_ts",  # 親メッセージの時刻(ISO文字列)。空=スレッド親 or 単発
    "gmail":   "thread_id",         # スレッド全行が同じ値を持つグループID
    "meeting": "meeting_uuid",      # 会議の一意キー（復元対象なし）
    "drive":   "origin",            # フォルダ名チェーン / "slack" / "gmail"
    "file":    "origin",            # 添付の出どころ（"gmail" 固定）
}

# gmail 行の dedup_key（"gmail:<messageId>"）から原文URLを復元するパターン
# ★v3 は content 末尾の [mail:ID] を見ていたが、agent_platform 版は専用の dedup_key 列を持つ
_MAIL_DEDUP_RE = re.compile(r"^gmail:([0-9a-fA-F]+)$")

# プロジェクト未分類の行（Zoom会議等）の project 値
UNCLASSIFIED_PROJECT = "未分類"

# ========== LLM モデル設定（v3 と同じ） ==========

SEARCH_PARSER_MODEL = "anthropic/claude-haiku-4-5"
STRUCTURED_FALLBACK_MODEL = "anthropic/claude-sonnet-4-6"

# ========== 取得の上限 ==========

# 1ソースの取得上限（暴走防止。超えたら警告して打ち切り）
MAX_ROWS_PER_SOURCE = 500
# find() 1回の上限（backend 仕様。これ以上は 422）
FIND_LIMIT_MAX = 5000
# after_id ページングの1ページ
FIND_PAGE_SIZE = 1000
# キーワード照合のために全走査するときの上限（超えたら警告して打ち切り）
SCAN_HARD_MAX = 20_000

_MAX_RETRIES = 3
_INITIAL_BACKOFF_SEC = 1.0


# ========== リトライ付き db 呼び出し ==========

async def _db_retry(coro_factory, what: str):
    """db 呼び出しを一時エラーのときだけリトライする。

    ★型付き例外（PermissionDenied / ValidationFailed / Conflict / TableNotFound）は
      リトライしても回復しないのでそのまま投げる。素の DbError だけ再試行する。
    """
    last_err = None
    for attempt in range(_MAX_RETRIES):
        try:
            return await coro_factory()
        except (adb.PermissionDenied, adb.ValidationFailed,
                adb.Conflict, adb.TableNotFound):
            raise
        except adb.DbError as e:
            last_err = e
            if attempt == _MAX_RETRIES - 1:
                break
            wait = _INITIAL_BACKOFF_SEC * (2 ** attempt) + random.uniform(0, 1)
            _warn(f"{what} retry {attempt + 1}/{_MAX_RETRIES} after {wait:.1f}s: {e}")
            await asyncio.sleep(wait)
    raise ToolCallError(f"{what} が{_MAX_RETRIES}回失敗: {last_err}")


# ========== 構造化出力 llm_call（v3 からそのまま移植） ==========

async def _call_structured_llm(
    prompt: str, *, model: str, schema: dict, label: str, required: list[str] | None = None
) -> dict:
    """schema指定のllm_callを呼び、構造化出力の欠落を検知してリトライする。

    同じモデルでのリトライが無意味なケースがある（v2実機で確認）ため、
    2回目の試行は STRUCTURED_FALLBACK_MODEL に切り替える。
    """
    required = required if required is not None else schema.get("required", [])
    models = [model] + [m for m in [STRUCTURED_FALLBACK_MODEL] if m != model]
    last_error: str | None = None
    for attempt, attempt_model in enumerate(models, start=1):
        _dbg(f"{label}: llm_call model={attempt_model} (試行{attempt}/{len(models)})")
        try:
            res = await llm_call(prompt=prompt, model=attempt_model, schema=schema)
        except Exception as e:
            last_error = f"{type(e).__name__}: {e}"
            _warn(f"{label}: llm_call 例外 (試行{attempt}): {last_error}")
            continue
        if not isinstance(res, dict):
            last_error = f"レスポンスがdictでない ({type(res).__name__})"
            continue
        if res.get("error"):
            last_error = str(res["error"])
            _warn(f"{label}: llm_call エラー (試行{attempt}): {last_error}")
            continue
        data = res.get("data") or {}
        missing = [k for k in required if k not in data]
        if missing:
            last_error = f"必須キー不足 {missing}"
            _warn(f"{label}: 構造化出力が不正 (試行{attempt}): {last_error}")
            continue
        return data
    raise ToolCallError(f"{label}: 構造化出力を取得できませんでした: {last_error}")


# ========== 行の整形 ==========

def _normalize_row(row: dict) -> dict:
    """find() の戻り（列値 + "_id"）に v3 互換のキーを足す。

    v3 は Notion 行から id / url / created_time を持っていた。agent_platform の行に
    ページURLは無いので、url は「一次情報に届くURL」だけを入れる（無ければ空文字）。
    """
    out = dict(row or {})
    out["id"] = row.get("_id")
    out["url"] = gmail_source_url(out) or ""
    return out


def gmail_source_url(row: dict) -> str | None:
    """gmail行の dedup_key から Gmail 原文URLを復元する。

    出典として提示すると一次情報（メール原文）に1クリックで届く。
    gmail 以外の行・IDが無い行は None。
    """
    if (row.get("source") or "") != "gmail":
        return None
    m = _MAIL_DEDUP_RE.match((row.get("dedup_key") or "").strip())
    return f"https://mail.google.com/mail/u/0/#all/{m.group(1)}" if m else None


# ========== 照合用の正規化（サーバー側 contains の代替） ==========

def _norm(s: str) -> str:
    """照合用に正規化する（NFKC＋小文字化）。

    ★v3 はサーバー側 contains（Notion の素の部分一致）だった。ローカル照合に
      移したのを機に、全角/半角・大小文字のゆれは吸収する（取りこぼしが減る方向）。
    """
    return unicodedata.normalize("NFKC", s or "").casefold()


def _keyword_match(row: dict, source: str, keywords: list[str]) -> bool:
    """v3 の _keyword_group と同じ意味論（本文系フィールド × キーワードの OR）。"""
    kws = [_norm(k) for k in (keywords or []) if k and k.strip()]
    if not kws:
        return True
    fields = _BODY_FIELDS_BY_SOURCE.get(source, [])
    haystack = " \n".join(_norm(str(row.get(f) or "")) for f in fields)
    return any(kw in haystack for kw in kws)


def _person_match(row: dict, source: str, person: dict | None) -> bool:
    """v3 の person_conditions と同じ意味論（メール列 × email / 名前列 × 表記候補の OR）。"""
    if not person:
        return True
    email = (person.get("email") or "").strip().lower()
    variants = [_norm(v) for v in (person.get("variants") or []) if v and len(v) >= 2]
    fields = _PERSON_FIELDS_BY_SOURCE.get(source, {})
    if email:
        pool = " \n".join(_norm(str(row.get(f) or "")) for f in fields.get("email", []))
        if _norm(email) in pool:
            return True
    if variants:
        pool = " \n".join(_norm(str(row.get(f) or "")) for f in fields.get("name", []))
        if any(v in pool for v in variants):
            return True
    return False


# ========== サーバー側フィルタの組み立て ==========

def _project_variants(name: str) -> list[str]:
    """project 照合の候補値（pjt_ prefix の有無を吸収）。"""
    n = (name or "").strip()
    if not n:
        return []
    return sorted({n, n if n.startswith("pjt_") else f"pjt_{n}"})


def _build_server_filter(source: str, project: str | None,
                         start_date: str | None, end_date: str | None,
                         *, include_unclassified: bool) -> dict:
    """1ソースぶんのサーバー側フィルタ（完全一致 / $in / 日付範囲のみ）。

    ★v3 の _build_source_filter との違い: キーワードと人物条件はここに入らない
      （contains / $or がサーバー側に無いため、取得後に Python で照合する）。
    """
    filter_obj: dict = {"source": source}

    variants = _project_variants(project or "")
    if variants:
        if include_unclassified:
            # キーワード検索時は「未分類」も対象に含める
            # （Zoom の未分類会議を取りこぼさないため。v3 と同じ挙動）
            variants = sorted(set(variants) | {UNCLASSIFIED_PROJECT})
        # 同じ列の複数値は $in で表現できる（$or が無くても等価）
        filter_obj["project"] = variants[0] if len(variants) == 1 else {"$in": variants}

    date_cond: dict = {}
    if start_date:
        date_cond["$gte"] = f"{start_date}T00:00:00+09:00"
    if end_date:
        date_cond["$lte"] = f"{end_date}T23:59:59+09:00"
    if date_cond:
        filter_obj["timestamp"] = date_cond

    return filter_obj


# ========== 取得層 ==========

async def _top_rows(filter_obj: dict, *, label: str, limit: int) -> list[dict]:
    """timestamp 降順の上位 limit 件を1回で取る（キーワード照合が不要なとき）。

    ★sort を付けると backend はページングを受け付けないので、1回で取り切る。
      limit に達したら「新しい順の上位だけ」であることを警告する。
    """
    n = min(limit, FIND_LIMIT_MAX)
    rows = await _db_retry(
        lambda: adb.table(RAW_TABLE).find(filter_obj, sort=[("timestamp", -1)], limit=n),
        f"find({label})",
    )
    if len(rows) >= n:
        try:
            total = await _db_retry(
                lambda: adb.table(RAW_TABLE).count_documents(filter_obj), f"count({label})")
        except Exception:
            total = -1
        _warn(f"_top_rows({label}): {n}件で打ち切り（該当{total if total >= 0 else '不明'}件）"
              f"。新しい順の上位のみです。期間やプロジェクトで絞ってください")
    _dbg(f"_top_rows({label}): {len(rows)}件")
    return [_normalize_row(r) for r in rows]


async def _scan_rows(filter_obj: dict, *, label: str, hard_max: int = SCAN_HARD_MAX) -> list[dict]:
    """条件に合う行を after_id で全走査する（キーワード照合を Python 側でやるため）。

    ★sort は付けない（付けるとページングできない）。挿入順で返るので、
      呼び出し側が照合後に timestamp 降順へ並べ替えること。
    """
    rows: list[dict] = []
    after_id: int | None = None
    while True:
        page = await _db_retry(
            lambda a=after_id: adb.table(RAW_TABLE).find(
                filter_obj, limit=FIND_PAGE_SIZE, after_id=a),
            f"find({label})",
        )
        if not page:
            break
        rows.extend(page)
        after_id = page[-1]["_id"]
        if len(page) < FIND_PAGE_SIZE:
            break
        if len(rows) >= hard_max:
            _warn(f"_scan_rows({label}): 全走査が{hard_max}件に達したため打ち切り"
                  f"（キーワードに合う古い行を取りこぼしている可能性があります）")
            break
    _dbg(f"_scan_rows({label}): {len(rows)}件を走査")
    return [_normalize_row(r) for r in rows]


def _sort_desc(rows: list[dict]) -> list[dict]:
    """timestamp 降順（新しい順）。timestamp 無しは末尾。"""
    return sorted(rows or [], key=lambda r: r.get("timestamp") or "", reverse=True)


async def fetch_records(
    sources: list[str] | None = None,
    project: str | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
    keywords: list[str] | None = None,
    person: dict | None = None,
) -> dict[str, list[dict]]:
    """raw テーブルを横断取得する（v4 取得層の公開API）。

    Args:
        sources: 対象ソース。省略時は全ソース。
        project: プロジェクト名（pjt_ prefix の有無は自動吸収）。
        start_date / end_date: "YYYY-MM-DD"。timestamp でサーバー側絞り込み。
        keywords: 本文系フィールド横断のキーワード（照合は Python 側・対象はソース別）。
        person: resolve_person の戻り。ソース別の人物フィールドに対して照合する。

    Returns:
        {source: [row, ...]}（timestamp 降順）。取得失敗したソースは空リスト＋WARN。
        キーワード指定時は project 絞りに「未分類」も自動で含める（v3 と同じ）。

    ★v3 の extra_conditions_by_source（Notion filter の断片）は person に置き換えた。
      サーバー側に contains / $or が無く、filter 断片を渡す意味が無くなったため。
    """
    targets = [s for s in (sources or _ALL_SOURCES) if s in _ALL_SOURCES]
    needs_local_match = bool(keywords) or bool(person)
    _dbg(
        f"fetch_records: sources={targets} project={project!r} "
        f"date={start_date}~{end_date} keywords={keywords} "
        f"person={(person or {}).get('email') or (person or {}).get('status')}"
    )

    async def _one(source: str) -> list[dict]:
        filter_obj = _build_server_filter(
            source, project, start_date, end_date,
            include_unclassified=bool(keywords))
        if not needs_local_match:
            return await _top_rows(filter_obj, label=source, limit=MAX_ROWS_PER_SOURCE)

        scanned = await _scan_rows(filter_obj, label=source)
        hits = [
            r for r in scanned
            if _keyword_match(r, source, keywords or []) and _person_match(r, source, person)
        ]
        _dbg(f"fetch_records: {source} 走査{len(scanned)}件 → 照合{len(hits)}件")
        if len(hits) > MAX_ROWS_PER_SOURCE:
            _warn(f"fetch_records({source}): 照合{len(hits)}件のうち新しい順"
                  f"{MAX_ROWS_PER_SOURCE}件に絞り込みました")
        return _sort_desc(hits)[:MAX_ROWS_PER_SOURCE]

    results = await asyncio.gather(*(_one(s) for s in targets), return_exceptions=True)
    out: dict[str, list[dict]] = {}
    for source, r in zip(targets, results):
        if isinstance(r, BaseException):
            _warn(f"fetch_records: {source} 取得失敗: {type(r).__name__}: {r}")
            out[source] = []
        else:
            _dbg(f"fetch_records: {source} → {len(r)}件")
            out[source] = r
    return out


# ========== 自然言語検索 ==========

_SEARCH_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "project_name": {"type": ["string", "null"], "description": "プロジェクト名または会社名。無ければ null"},
        "sources": {
            "type": "array",
            "items": {"type": "string", "enum": _ALL_SOURCES},
            "description": (
                "対象ソース。メール→gmail / チャット・Slack→slack / "
                "ドライブの資料→drive / 会議・議事録・MTG→meeting / "
                "メール添付ファイル→file。指定なければ全部"
            ),
        },
        "start_date": {"type": ["string", "null"], "description": "取得開始日 YYYY-MM-DD。無ければ null"},
        "end_date": {"type": ["string", "null"], "description": "取得終了日 YYYY-MM-DD。無ければ null"},
        "keywords": {
            "type": "array",
            "items": {"type": "string"},
            "description": "本文横断検索キーワード（表記ゆれ・省略形バリアント込み、最大5個）。不要なら空配列",
        },
    },
    "required": ["project_name", "sources", "start_date", "end_date", "keywords"],
}


async def search(query: str) -> dict[str, list[dict]]:
    """自然言語クエリを解析して raw テーブルから横断取得する。

    例:
      await search('''アルファリースの6月のメール''')
      await search('''見積もりについて話した会議''')
    """
    _dbg(f"search: query={query!r}")
    today = datetime.now(JST).strftime("%Y-%m-%d")
    prompt = (
        f"今日の日付: {today}\n\n"
        f"以下のクエリを解析して検索パラメータに変換してください。\n"
        f"年が省略された日付は今年として扱ってください。\n"
        f"keywordsには表記ゆれ（カタカナ/英語など）と、「〜エージェント」等の接尾辞を外した\n"
        f"省略形も含めてください（最大5個）。人名が含まれる場合はその人名もkeywordsに入れること。\n\n"
        f"クエリ: {query}"
    )
    params = await _call_structured_llm(
        prompt, model=SEARCH_PARSER_MODEL, schema=_SEARCH_SCHEMA, label="search",
        required=[],  # 欠落キーは下のデフォルトで補完する（キー欠落だけで失敗にしない）
    )
    project = params.get("project_name") or None
    sources = params.get("sources") or None
    start_date = params.get("start_date")
    end_date = params.get("end_date")
    keywords = params.get("keywords") or []
    _dbg(f"search: 解析結果 project={project!r} sources={sources} "
         f"date={start_date}~{end_date} keywords={keywords}")

    return await fetch_records(
        sources=sources, project=project,
        start_date=start_date, end_date=end_date, keywords=keywords,
    )


async def search_from_file(query_path: str = "tmp/query.txt") -> dict[str, list[dict]]:
    """質問をファイル経由で受け取る版（引用符エスケープ事故の回避用）。"""
    with open(query_path, encoding="utf-8") as f:
        return await search(f.read().strip())


# ==========================================================================
# 解決層: 人物・会社をエンティティとして解決する
# ==========================================================================

_VARIANTS_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "variants": {
            "type": "array",
            "items": {"type": "string"},
            "description": "同一人物を指しうる表記の候補（最大6個）",
        },
    },
    "required": ["variants"],
}


def _clean_person_query(query: str) -> str:
    """人物クエリから敬称・括弧を除いた照合用の文字列を返す。"""
    n = (query or "").strip()
    n = re.sub(r"[（(\[【].*?[）)\]】]", "", n)
    n = re.sub(r"\s*(さん|様|さま|君|くん|ちゃん|先生|氏)$", "", n)
    return n.strip()


async def _generate_name_variants(query: str) -> list[str]:
    """クエリの人物名から表記変換候補を生成する（漢字⇄かな⇄ローマ字）。

    人物辞書の aliases は「観測された表記」しか持たないため、未観測の表記
    （ひらがな呼び等）はクエリ側で変換して照合する。
    """
    cleaned = _clean_person_query(query)
    if not cleaned:
        return []
    variants = [cleaned]
    prompt = (
        f"「{cleaned}」という人物名について、同一人物を指しうる別表記を生成してください。\n"
        f"日本語名⇄ローマ字⇄ひらがな/カタカナの相互変換（姓・名それぞれ単体も含む）で、\n"
        f"メールの表示名やSlack名として実際に使われそうな表記のみ、最大6個。"
    )
    try:
        data = await _call_structured_llm(
            prompt, model=SEARCH_PARSER_MODEL, schema=_VARIANTS_SCHEMA, label="name_variants"
        )
        for v in data.get("variants") or []:
            v = (v or "").strip()
            if v and len(v) >= 2 and v not in variants:
                variants.append(v)
    except Exception as e:
        _warn(f"_generate_name_variants: 変換生成失敗（元の表記のみで続行）: {e}")
    return variants[:7]


async def _load_people() -> list[dict]:
    """人物辞書を全件読む。

    ★people テーブルに contains 検索が無いため、v3 のような「候補で OR 検索」は
      できない。人物辞書は小さい（社内＋取引先の登場人物）ので全件取ってから
      ローカル照合する。件数が増えて重くなったら company 等での事前絞りを足すこと。
    """
    rows: list[dict] = []
    after_id: int | None = None
    while True:
        page = await _db_retry(
            lambda a=after_id: adb.table(PEOPLE_TABLE).find(limit=FIND_PAGE_SIZE, after_id=a),
            "find(people)",
        )
        if not page:
            break
        rows.extend(page)
        after_id = page[-1]["_id"]
        if len(page) < FIND_PAGE_SIZE:
            break
    return rows


async def resolve_person(query: str) -> dict:
    """人物を名寄せする（解決層）。v3 と同じ3層＋安全弁:

    ① クエリ正規化＋LLM表記変換 → ② 人物辞書を候補で照合
       - 1人に確定 → email＋aliases一式を検索用バリアントとして返す
       - 複数人 → ★混ぜずに candidates を返す（呼び出し元がユーザーに候補提示）
    ③ 0件なら raw の逆引き（slack/drive/file 行は1行=1人で name⇄sender が正確なペア）
    ④ それでも不明 → 入力文字列そのままの照合にフォールバック

    戻り値:
      {"status": "resolved",   "query", "email", "variants": [...]}
      {"status": "ambiguous",  "query", "candidates": [{"name","email","company"}]}
      {"status": "unresolved", "query", "variants": [元の表記]}
    """
    variants = await _generate_name_variants(query)
    _dbg(f"resolve_person: {query!r} → 変換候補 {variants}")
    if not variants:
        return {"status": "unresolved", "query": query, "variants": [query]}

    # ①② 人物辞書を表記候補で照合（name / aliases / email のいずれかに含まれれば候補）
    people: list[dict] = []
    try:
        people = await _load_people()
    except Exception as e:
        _warn(f"resolve_person: 人物辞書の読み込み失敗: {e}")

    nvariants = [_norm(v) for v in variants]
    by_email: dict[str, dict] = {}
    for r in people:
        email = (r.get("email") or "").strip().lower()
        if not email:
            continue
        pool = " \n".join(_norm(str(r.get(f) or "")) for f in ("name", "aliases", "email"))
        if any(v in pool for v in nvariants):
            by_email.setdefault(email, r)

    if len(by_email) == 1:
        email, r = next(iter(by_email.items()))
        all_variants = list(variants)
        for a in (r.get("aliases") or "").split("\n"):
            a = a.strip()
            if a and a not in all_variants:
                all_variants.append(a)
        name = (r.get("name") or "").strip()
        if name and name not in all_variants:
            all_variants.append(name)
        _dbg(f"resolve_person: 確定 {query!r} → {email}（表記{len(all_variants)}種）")
        return {"status": "resolved", "query": query, "email": email, "variants": all_variants}

    if len(by_email) > 1:
        candidates = [
            {"name": r.get("name") or "", "email": e, "company": r.get("company") or ""}
            for e, r in by_email.items()
        ]
        _warn(f"resolve_person: {query!r} は複数人にヒット（{len(candidates)}人）→ 候補提示")
        return {"status": "ambiguous", "query": query, "candidates": candidates}

    # ③ raw 逆引き（slack/drive/file 行は 1行=1人 で name⇄sender が正確に対応する）
    for source in ("slack", "drive", "file"):
        try:
            rows = await _scan_rows({"source": source}, label=f"person_fallback:{source}",
                                    hard_max=FIND_PAGE_SIZE * 3)
        except Exception as e:
            _warn(f"resolve_person: raw逆引き失敗 ({source}): {e}")
            continue
        for v in nvariants:
            emails = {
                (r.get("sender") or "").strip().lower()
                for r in rows
                if v in _norm(str(r.get("name") or ""))
                and "@" in (r.get("sender") or "") and "," not in (r.get("sender") or "")
            }
            if len(emails) == 1:
                email = next(iter(emails))
                _dbg(f"resolve_person: raw逆引きで確定 {query!r} → {email}")
                return {"status": "resolved", "query": query, "email": email, "variants": variants}

    _dbg(f"resolve_person: 未解決 {query!r} → 文字列一致にフォールバック")
    return {"status": "unresolved", "query": query, "variants": variants}


def _normalize_name(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"[（(\[].*?[）)\]]", "", s)
    s = re.sub(r"\s+", "", s)
    return s.lower()


async def _load_domain_map() -> list[dict]:
    """domain_map（PJ⇄ドメイン・手動管理）を全件読む。[{project, domain}]。"""
    rows = await _db_retry(
        lambda: adb.table(DOMAIN_MAP_TABLE).find(limit=FIND_LIMIT_MAX), "find(domain_map)")
    return [
        {"project": (r.get("project") or "").strip(),
         "domain": (r.get("domain") or "").strip().lstrip("@").lower()}
        for r in rows if (r.get("project") or "").strip()
    ]


async def resolve_company(query: str) -> dict:
    """会社名/プロジェクト名を {project, domain} に解決する（domain_map 参照）。

    戻り値: {"status": "resolved"|"ambiguous"|"unresolved",
             "project", "domain", "candidates": [{project, domain}]}
    """
    q = re.sub(r"\s*(さん|様|さま)$", "", (query or "").strip())
    mapping = await _load_domain_map()
    nq = _normalize_name(re.sub(r"^pjt_", "", q))
    matches = []
    for m in mapping:
        np = _normalize_name(re.sub(r"^pjt_", "", m["project"]))
        if nq and np and (nq in np or np in nq):
            matches.append(m)
        elif "." in q and m["domain"] and m["domain"] == q.lower():
            matches.append(m)  # ドメイン文字列での直接指定
    if len(matches) == 1:
        _dbg(f"resolve_company: {query!r} → {matches[0]}")
        return {"status": "resolved", **matches[0], "candidates": matches}
    if len(matches) > 1:
        _warn(f"resolve_company: {query!r} は複数候補 {matches}")
        return {"status": "ambiguous", "project": None, "domain": None, "candidates": matches}
    _dbg(f"resolve_company: {query!r} は domain_map に見つからず")
    return {"status": "unresolved", "project": None, "domain": None, "candidates": []}


async def company_members(domain: str, *, limit: int = 30) -> list[dict]:
    """人物辞書を company=ドメイン で絞って担当者一覧を返す。"""
    d = (domain or "").strip().lower()
    if not d:
        return []
    return await _db_retry(
        lambda: adb.table(PEOPLE_TABLE).find({"company": d}, limit=min(limit, FIND_LIMIT_MAX)),
        "find(people/company)",
    )


# ==========================================================================
# 文脈層: parents をソース別に解釈してスレッドを復元する
# ==========================================================================

async def expand_thread(row: dict) -> list[dict]:
    """行が属する会話全体を復元する（PARENTS_SEMANTICS に基づくソース別処理）。

    gmail: 同じ threadId の全メール / slack: 親メッセージ＋返信 /
    meeting・drive・file: その行のみ（復元対象なし）。時系列昇順で返す。
    """
    source = row.get("source") or ""
    if source == "gmail":
        tid = (row.get("parents") or "").strip()
        if not tid:
            return [row]
        rows = await _top_rows({"source": "gmail", "parents": tid},
                               label="gmail_thread", limit=100)
        return sorted(rows, key=lambda r: r.get("timestamp") or "") or [row]

    if source == "slack":
        # 自分が返信なら parents=親ts、自分が親/単発なら自分の timestamp が親ts
        parent_ts = (row.get("parents") or "").strip() or (row.get("timestamp") or "").strip()
        if not parent_ts:
            return [row]
        base: dict = {"source": "slack"}
        project = (row.get("project") or "").strip()
        if project:
            # 親tsはチャンネル間で理論上衝突しうるため project でも絞る
            base["project"] = project
        children = await _top_rows({**base, "parents": parent_ts},
                                   label="slack_thread", limit=100)
        # 親メッセージ本体: 日時の等値比較は精度が怖いので±1分で取ってクライアント側一致
        parent_rows: list[dict] = []
        try:
            pdt = datetime.fromisoformat(parent_ts)
            lo = (pdt - timedelta(minutes=1)).isoformat(timespec="seconds")
            hi = (pdt + timedelta(minutes=1)).isoformat(timespec="seconds")
            around = await _top_rows({**base, "timestamp": {"$gte": lo, "$lte": hi}},
                                     label="slack_thread_parent", limit=20)
            parent_rows = [r for r in around if (r.get("timestamp") or "")[:19] == parent_ts[:19]]
        except ValueError:
            _warn(f"expand_thread: 親ts をパースできません: {parent_ts!r}")
        # 種の行が正規化前（find() の生 dict）で渡ることもあるため id は両方見る
        merged = {(r.get("id") or r.get("_id")): r
                  for r in children + parent_rows + [row]}
        return sorted(merged.values(), key=lambda r: r.get("timestamp") or "")

    return [row]


async def expand_attachments(row: dict) -> list[dict]:
    """メール行に紐づく添付（source=file）の行を返す。

    ★Relation（relations_query）ではなく links 列を使う。Relation から返るのは
      相手の record_id だけで、agent_platform の find() は `_id` での引き当てを
      422 で明示的に拒否する（行を復元する手段が無い）ため。
      links にはメール保存時に書いた note_path が改行区切りで入っている。
    """
    paths = [p.strip() for p in (row.get("links") or "").split("\n") if p.strip()]
    out: list[dict] = []
    for path in paths:
        try:
            found = await _db_retry(
                lambda p=path: adb.table(RAW_TABLE).find({"note_path": p}, limit=1),
                "find(attachment)",
            )
        except Exception as e:
            _warn(f"expand_attachments: {path} の取得に失敗: {e}")
            continue
        if found:
            out.append(_normalize_row(found[0]))
    return out


# ==========================================================================
# 深読み層: 必要な行だけ本文まで読む
# ==========================================================================

_HTTP_URL_RE = re.compile(r"https?://[^\s<>|\"')]+")

MAX_AUTO_DEEP_READS = 5      # 1クエリで自動深読みする最大件数
DEEP_READ_TIER1_MAX = 5      # 候補がこの件数以下なら選別せず全部読む
MAX_THREAD_EXPANSION_SEEDS = 10  # スレッド展開する種メッセージ数の上限
MAX_CHUNK_ROWS = 40          # map-reduce 中間要約1チャンクの最大行数

PLANNER_MODEL = "anthropic/claude-sonnet-5"
SYNTHESIS_MODEL = PLANNER_MODEL
LITE_MODEL = "gemini/gemini-3.1-flash-lite"

# 深読みで読む note 本文の上限（1件あたり）
MAX_NOTE_CHARS = 20_000

# sandbox に置かれたプレースホルダ（実体は remote にあり、バイト列は読めない）の目印
_GHOST_MAGIC = "ASKHUB-GHOSTv1"


def _extract_urls(text: str) -> list[str]:
    seen: set[str] = set()
    urls: list[str] = []
    for m in _HTTP_URL_RE.finditer(text or ""):
        u = m.group(0).rstrip("。、.,")
        if u not in seen:
            seen.add(u)
            urls.append(u)
    return urls


def _classify_url(url: str) -> str | None:
    if "drive.google.com" in url or "docs.google.com" in url:
        return "drive"
    if any(d in url for d in ("notion.so", "notion.com", "notion.site")):
        return "notion"
    return None


def _notes_root() -> str:
    """notes ツリーの実パスを返す。見つからなければ例外（黙って別の場所を読まない）。"""
    env = os.environ.get("ASKHUB_NOTES_DIR")
    home = os.environ.get("HOME") or os.path.expanduser("~")
    for candidate in ([env] if env else []) + [os.path.join(home, "workspace", "notes"), "notes"]:
        if candidate and os.path.isdir(candidate):
            return candidate
    raise RuntimeError("notes/ ディレクトリが見つかりません（この会話に Project が attach されていますか）")


def read_note(note_path: str, *, max_chars: int = MAX_NOTE_CHARS) -> dict:
    """notes/<note_path> のテキストを読む（深読みの実体）。

    戻り値: {"note_path", "text", "truncated", "reason"}

    ★読めないケースを黙って空文字にしない:
      - ghost: セッション開始時の復元でプレースホルダだけが置かれた状態
        （バイナリ・大きいファイルはこうなる）。中身は remote にあるが sandbox では読めない。
      - binary: テキストとしてデコードできない（PDF/xlsx 等）。
      どちらも reason に理由を入れて返し、呼び出し側は要約JSONにフォールバックする。
    """
    out = {"note_path": note_path, "text": "", "truncated": False, "reason": ""}
    try:
        abs_path = os.path.join(_notes_root(), note_path)
    except RuntimeError as e:
        out["reason"] = str(e)
        return out
    if not os.path.exists(abs_path):
        out["reason"] = "notes/ に実体がありません（未 sync か削除済み）"
        return out
    try:
        with open(abs_path, "rb") as f:
            head = f.read(len(_GHOST_MAGIC.encode("utf-8")))
            if head.decode("utf-8", "ignore").startswith(_GHOST_MAGIC):
                out["reason"] = "ghost（sandbox にはプレースホルダのみ。実体は remote）"
                return out
            body = head + f.read(max_chars * 4)
    except OSError as e:
        out["reason"] = f"読み込み失敗: {e}"
        return out
    try:
        text = body.decode("utf-8")
    except UnicodeDecodeError:
        out["reason"] = "binary（テキストとして読めない形式）"
        return out
    out["truncated"] = len(text) > max_chars
    out["text"] = text[:max_chars]
    return out


def _collect_deep_read_candidates(rows_by_source: dict, threads: dict) -> list[dict]:
    """深読み候補を集める。

    ★v3 は「meeting行のNotionページ本文」だけが候補だった。agent_platform では
      meeting の議事録全文が content 列にそのまま入っているので取り直す必要が無く、
      代わりに「file 行の note 実体」が唯一の“DBに載っていない本文”になる。
      notes/ は自分たちの資産なので、v3 の「外部ファイルは再取得しない」原則に反しない。
    """
    seen: set[str] = set()
    candidates: list[dict] = []
    for r in rows_by_source.get("file") or []:
        path = (r.get("note_path") or "").strip()
        if not path or path in seen:
            continue
        seen.add(path)
        candidates.append({
            "note_path": path, "kind": "attachment_note",
            "hint": f"{r.get('subject') or ''} {(r.get('content') or '')[:120]}"[:200],
        })
    _dbg(f"_collect_deep_read_candidates: 候補{len(candidates)}件（添付noteの実体）")
    return candidates


_TRIAGE_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "needed_paths": {
            "type": "array", "items": {"type": "string"},
            "description": "本文まで読む必要がある note_path（最大5件、必要性の高い順）",
        },
    },
    "required": ["needed_paths"],
}


async def _triage_deep_read(user_query: str, candidates: list[dict]) -> list[dict]:
    """候補が多い場合に軽量LLMで「本文まで読む必要があるか」を選別する（Tier2）。"""
    lines = "\n".join(
        f"- note_path={c['note_path']} kind={c['kind']} 手がかり={c['hint']}" for c in candidates)
    prompt = (
        f"以下は検索でヒットした候補です。質問に答えるために本文まで読む必要がある\n"
        f"候補だけを最大5件、必要性の高い順に needed_paths に入れてください。\n"
        f"手がかりだけで答えられそうなものは選ばないこと。\n\n質問: {user_query}\n\n候補:\n{lines}"
    )
    try:
        data = await _call_structured_llm(prompt, model=LITE_MODEL, schema=_TRIAGE_SCHEMA, label="triage")
    except Exception as e:
        _warn(f"_triage_deep_read: 選別失敗→先頭{MAX_AUTO_DEEP_READS}件を読む: {e}")
        return candidates[:MAX_AUTO_DEEP_READS]
    needed = set(data.get("needed_paths") or [])
    return [c for c in candidates if c["note_path"] in needed]


async def _deep_read_notes(selected: list[dict]) -> list[dict]:
    """選別済みの候補（添付noteの実体）を読む（1件の失敗で全体を止めない）。"""
    results: list[dict] = []
    for c in selected[:MAX_AUTO_DEEP_READS]:
        try:
            got = await asyncio.to_thread(read_note, c["note_path"])
        except Exception as e:
            _warn(f"_deep_read_notes: 失敗 {c['note_path']}: {e}")
            continue
        if not got.get("text"):
            # 読めない理由は握りつぶさず記録する（合成側は要約JSONで代替する）
            _warn(f"_deep_read_notes: 本文を読めず（要約で代替）: "
                  f"{c['note_path']} / {got.get('reason')}")
            continue
        results.append({"kind": c["kind"], **got})
    _dbg(f"_deep_read_notes: {len(results)}件を深読み完了")
    return results


# ==========================================================================
# プランナー
# ==========================================================================

# ファクト（正確な事実・原文）が要求される分野。共通項は「後から言った/言わないになり得るもの」
_FACT_DOMAINS = """\
1. 顧客の発言・要望 / 2. 要件定義・仕様の合意内容 / 3. 議事録の決定事項（誰が・いつ・何を）
4. 金額・見積もり・契約条件 / 5. 納期・期日のコミットメント / 6. 承認・責任の所在
7. クレーム・トラブル対応の経緯 / 8. 数値実績・KPI / 9. 対外的に送った公式な文面
10. セキュリティ・データ取り扱いの取り決め
"""

_DB_KNOWLEDGE = """\
【raw テーブルの構造（全ソース共通の列・source で種別分け）】
- source=slack:   content=メッセージ本文, name=投稿者, links=共有URLとその要約
- source=gmail:   content=メール本文, subject=件名, sender/to/cc_bcc=アドレス, name=送受信者名
- source=drive:   content=ファイルの要約, subject=ファイル名, parents=フォルダ階層/出どころ
- source=meeting: content=議事録全文, subject=会議タイトル, name=参加者, sender=参加者メール
- source=file:    メールの添付ファイル。content=要約JSON, subject=ファイル名,
                  note_path=notes/ 内の実体の場所（本文が必要ならここを深読みする）
- 共通: project=pjt_プロジェクト名, timestamp=日時, dedup_key=重複判定キー

【project_name の注意】
- pjt_ prefix の有無は自動吸収される。
- 「〜エージェント」「〜ツール」のような社内の仕組み名は project 名と一致しない場合が
  あるので、ユーザーの言葉やその略称を必ず keywords にも入れて本文横断検索できるようにする。
- 何のプロジェクトか不明なら project_name は null にして keywords で横断検索する。
"""

_RECIPES: list[dict] = [
    {"key": "decision_rationale", "label": "決定事項の根拠確認",
     "how": "keywordsで絞り込み→slack/gmailスレッド復元→meetingは全文。経緯を時系列で提示。"},
    {"key": "past_reference_docs", "label": "過去資料の参照",
     "how": "drive/file をファイル名/要約/フォルダ名で検索→上位候補の実体を読む。"},
    {"key": "latest_minutes_location", "label": "最新議事録の場所",
     "how": "source=meetingをprojectで絞り新しい順に返す。"},
    {"key": "catchup_by_period", "label": "期間キャッチアップ",
     "how": "指定期間で全ソース横断取得して要約。"},
    {"key": "catchup_whole_project", "label": "PJ全体の進捗まとめ",
     "how": "期間なしで全ソース取得し月次チャンクでmap-reduce要約。"},
    {"key": "cross_project_status", "label": "人物の横断ステータス",
     "how": "resolve_personで人物解決→その人が登場する直近の行を全PJ横断で取得して要約。"},
    {"key": "client_temperature_check", "label": "顧客の温度感チェック",
     "how": "resolve_companyで会社解決→直近meeting1-2件（全文）+直近gmail数通のみの軽量取得。"},
    {"key": "client_profile", "label": "顧客像の把握",
     "how": "resolve_company→全期間取得。要望・こだわりの抽出に焦点。"},
    {"key": "custom", "label": "上記に当てはまらない自由形式の調査",
     "how": "project/sources/keywords/期間/人物を自由に組み合わせて横断取得する。"},
]

_PLAN_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "recipe_key": {"type": "string", "enum": [r["key"] for r in _RECIPES]},
        "project_name": {"type": ["string", "null"], "description": "プロジェクト名/会社名。不明ならnull"},
        "person_hint": {"type": ["string", "null"],
                        "description": "対象人物の名前/メール。「自分」を指す場合はnull（requester_hintで解決）"},
        "keywords": {"type": "array", "items": {"type": "string"},
                     "description": "表記ゆれバリアント込みのキーワード。不要なら空配列"},
        "sources": {"type": "array", "items": {"type": "string", "enum": _ALL_SOURCES}},
        "start_date": {"type": ["string", "null"]},
        "end_date": {"type": ["string", "null"]},
        "deep_read": {"type": "boolean", "description": "添付ファイル本文の2段階取得を行うか"},
        "fact_critical": {"type": "boolean", "description": "正確な事実・原文が要求される質問か"},
        "notes": {"type": "string", "description": "実行方針の補足。無ければ空文字"},
    },
    "required": ["recipe_key", "project_name", "person_hint", "keywords", "sources",
                 "start_date", "end_date", "deep_read", "fact_critical", "notes"],
}


async def plan_query(user_query: str, *, requester_hint: str | None = None,
                     replan_feedback: str | None = None) -> dict:
    """レシピ選択＋調査計画をLLMに立てさせる（プランナー）。"""
    today = datetime.now(JST).strftime("%Y-%m-%d")
    recipe_lines = "\n".join(f"- {r['key']}: {r['label']} — {r['how']}" for r in _RECIPES)
    feedback = (f"\n\n前回の実行結果に関するフィードバック: {replan_feedback}\n条件を調整してください。"
                if replan_feedback else "")
    prompt = (
        f"今日の日付: {today}\n"
        f"依頼者ヒント（「自分」等が指す人物。未指定ならnull）: {requester_hint or 'null'}\n\n"
        f"{_DB_KNOWLEDGE}\n"
        f"以下のレシピから最も適切なものを1つ選び、実行パラメータを決めてください。\n{recipe_lines}\n\n"
        f"年が省略された日付は今年として扱ってください。\n"
        f"keywordsには表記ゆれ（カタカナ/英語）と省略形を必ず含めてください（最大5個）。\n\n"
        f"質問が以下の分野に該当する場合は fact_critical=true にしてください:\n{_FACT_DOMAINS}"
        f"該当しない場合（雑談・概要把握・温度感などの解釈系）は false。"
        f"{feedback}\n\n"
        f"ユーザーの質問: {user_query}"
    )
    plan = await _call_structured_llm(
        prompt, model=PLANNER_MODEL, schema=_PLAN_SCHEMA, label="plan_query",
        required=["recipe_key"],  # 短い質問では他キーがnullで正常のため受理判定を緩める
    )
    defaults = {"project_name": None, "person_hint": None, "keywords": [],
                "sources": list(_ALL_SOURCES), "start_date": None, "end_date": None,
                "deep_read": False, "fact_critical": False, "notes": ""}
    plan = {**defaults, **plan}
    _dbg(f"plan_query: recipe={plan.get('recipe_key')} plan={plan}")
    return plan


# ==========================================================================
# レシピ実行
# ==========================================================================

def _collect_projects(rows_by_source: dict) -> list[str]:
    projects: set[str] = set()
    for rows in rows_by_source.values():
        if isinstance(rows, list):
            for r in rows:
                p = (r.get("project") or "").strip()
                if p:
                    projects.add(p)
    return sorted(projects)


async def _resolve_plan_person(plan: dict, requester_hint: str | None) -> dict | None:
    """plan の person_hint（or 依頼者）を解決する。指定なしなら None。"""
    person_query = plan.get("person_hint") or ""
    if not person_query and plan.get("recipe_key") == "cross_project_status":
        person_query = requester_hint or ""
    if not person_query:
        return None
    return await resolve_person(person_query)


async def _expand_slack_threads(rows_by_source: dict) -> list[dict]:
    """slackヒットの上位を種にスレッド展開する（会話単位の文脈を作る）。"""
    seeds = (rows_by_source.get("slack") or [])[:MAX_THREAD_EXPANSION_SEEDS]
    threads: list[dict] = []
    seen_keys: set[str] = set()
    for seed in seeds:
        key = (seed.get("parents") or "").strip() or (seed.get("timestamp") or "").strip()
        if not key or key in seen_keys:
            continue
        seen_keys.add(key)
        msgs = await expand_thread(seed)
        if len(msgs) > 1:
            threads.append({"thread_id": key, "messages": msgs})
    _dbg(f"_expand_slack_threads: {len(threads)}スレッド復元")
    return threads


async def _handle_custom(plan: dict, *, user_query: str, person: dict | None) -> dict:
    # ★レシピはソースを絞らない＝常に全ソースを見る（統合rawテーブルの意味。plan の
    #   sources は無視）。絞り込みは project / キーワード / 期間 / 人物で行う。
    p = person if person and person.get("status") != "ambiguous" else None
    rows = await fetch_records(
        sources=None, project=plan.get("project_name"),
        start_date=plan.get("start_date"), end_date=plan.get("end_date"),
        keywords=plan.get("keywords") or [], person=p,
    )
    threads = {"slack": await _expand_slack_threads(rows)} if rows.get("slack") else {}
    candidates = _collect_deep_read_candidates(rows, threads)
    fact = bool(plan.get("fact_critical"))
    if not candidates:
        deep_reads: list[dict] = []
    elif fact or plan.get("deep_read") or len(candidates) <= DEEP_READ_TIER1_MAX:
        deep_reads = await _deep_read_notes(candidates[:MAX_AUTO_DEEP_READS])
    else:
        deep_reads = await _deep_read_notes(await _triage_deep_read(user_query, candidates))
    return {"rows": rows, "threads": threads, "projects": _collect_projects(rows),
            "deep_reads": deep_reads}


async def _handle_latest_minutes(plan: dict) -> dict:
    # 「最新の議事録」質問だが、全ソースを取得して meeting を新しい順に前面に出す
    # （議事録の場所が slack/gmail で共有されているケースも拾える）。
    rows_all = await fetch_records(project=plan.get("project_name"),
                                   keywords=plan.get("keywords") or [])
    rows = {
        "meeting": _sort_desc(rows_all.get("meeting"))[:5],
        "gmail":   _sort_desc(rows_all.get("gmail"))[:3],
        "slack":   _sort_desc(rows_all.get("slack"))[:3],
        "drive":   _sort_desc(rows_all.get("drive"))[:3],
        "file":    _sort_desc(rows_all.get("file"))[:3],
    }
    return {"rows": rows, "projects": _collect_projects(rows), "deep_reads": []}


async def _handle_past_reference_docs(plan: dict, *, user_query: str) -> dict:
    # 資料探しでも全ソースを見る（資料が slack/gmail で言及・共有されている場合を拾う）。
    rows = await fetch_records(
        project=plan.get("project_name"),
        start_date=plan.get("start_date"), end_date=plan.get("end_date"),
        keywords=plan.get("keywords") or [],
    )
    candidates = _collect_deep_read_candidates(rows, {})
    deep_reads = await _deep_read_notes(candidates[:MAX_AUTO_DEEP_READS]) if candidates else []
    return {"rows": rows, "projects": _collect_projects(rows), "deep_reads": deep_reads}


async def _handle_temperature_check(plan: dict) -> dict:
    # 温度感でも全ソースを見る（slack/driveにやり取りが出ることが多い）。
    # 「軽さ」はソースを削ることではなく「各ソース新しい順に少数」で担保する。
    company = await resolve_company(plan.get("project_name") or "")
    project = company.get("project") or plan.get("project_name")
    rows_all = await fetch_records(project=project, keywords=plan.get("keywords") or [])
    rows = {
        "meeting": _sort_desc(rows_all.get("meeting"))[:2],
        "gmail":   _sort_desc(rows_all.get("gmail"))[:10],
        "slack":   _sort_desc(rows_all.get("slack"))[:10],
        "drive":   _sort_desc(rows_all.get("drive"))[:5],
        "file":    _sort_desc(rows_all.get("file"))[:5],
    }
    candidates = _collect_deep_read_candidates(rows, {})
    deep_reads = await _deep_read_notes(candidates[:MAX_AUTO_DEEP_READS]) if candidates else []
    return {"rows": rows, "company": company, "projects": _collect_projects(rows),
            "deep_reads": deep_reads}


async def _handle_cross_project_status(plan: dict, *, requester_hint: str | None,
                                       person: dict | None) -> dict:
    if person is None:
        return {"rows": {}, "deep_reads": [],
                "error_note": "対象人物が特定できません（person_hint も requester_hint も無し）"}
    week_ago = (datetime.now(JST) - timedelta(days=7)).strftime("%Y-%m-%d")
    today = datetime.now(JST).strftime("%Y-%m-%d")
    rows = await fetch_records(
        start_date=plan.get("start_date") or week_ago,
        end_date=plan.get("end_date") or today,
        person=person,
    )
    return {"rows": rows, "person": person, "projects": _collect_projects(rows), "deep_reads": []}


async def _execute_plan(plan: dict, *, requester_hint: str | None, user_query: str) -> dict:
    recipe = plan.get("recipe_key") or "custom"
    _dbg(f"_execute_plan: recipe={recipe}")
    person = await _resolve_plan_person(plan, requester_hint)
    if person and person.get("status") == "ambiguous":
        return {"rows": {}, "deep_reads": [], "ambiguous_person": person}

    if recipe == "latest_minutes_location":
        return await _handle_latest_minutes(plan)
    if recipe == "past_reference_docs":
        return await _handle_past_reference_docs(plan, user_query=user_query)
    if recipe == "client_temperature_check":
        return await _handle_temperature_check(plan)
    if recipe == "cross_project_status":
        return await _handle_cross_project_status(plan, requester_hint=requester_hint, person=person)
    if recipe == "catchup_whole_project":
        bundle = await _handle_custom(plan, user_query=user_query, person=person)
        bundle["needs_map_reduce"] = True
        return bundle
    # decision_rationale / catchup_by_period / client_profile / custom は
    # 横断取得＋スレッド展開＋深読みTierの共通フローで賄える
    return await _handle_custom(plan, user_query=user_query, person=person)


def _total_row_count(rows_obj) -> int:
    if isinstance(rows_obj, list):
        return len(rows_obj)
    if isinstance(rows_obj, dict):
        return sum(_total_row_count(v) for v in rows_obj.values())
    return 0


def _needs_replan(bundle: dict) -> str | None:
    total = _total_row_count(bundle.get("rows") or {})
    _dbg(f"_needs_replan: 取得総数={total}件")
    if total == 0:
        return "該当レコードが0件でした。project名やキーワードを見直してください。"
    if total > 300:
        return f"該当レコードが{total}件と多すぎます。キーワードや期間で絞り込んでください。"
    return None


# ==========================================================================
# 合成層
# ==========================================================================

async def _map_reduce_summarize(rows_by_source: dict) -> list[dict]:
    """月次チャンクで中間要約（軽量LLM）を作る。レコード数が多い全体まとめ専用。"""
    summaries: list[dict] = []
    for source, rows in rows_by_source.items():
        if not isinstance(rows, list) or not rows:
            continue
        by_month: dict[str, list[dict]] = {}
        for r in rows:
            month = (r.get("timestamp") or "")[:7] or "unknown"
            by_month.setdefault(month, []).append(r)
        for month, month_rows in sorted(by_month.items()):
            for i in range(0, len(month_rows), MAX_CHUNK_ROWS):
                chunk = month_rows[i:i + MAX_CHUNK_ROWS]
                text = "\n---\n".join(
                    f"[{_best_ref(r)}] {r.get('timestamp', '')}: "
                    f"{(r.get('content') or r.get('subject') or '')[:500]}"
                    for r in chunk
                )
                try:
                    res = await llm_call(
                        prompt=(f"{source}の{month}分のログです。主要な出来事・決定事項を"
                                f"箇条書きで要約してください（各行に出典を残すこと）。\n\n{text}"),
                        model=LITE_MODEL,
                    )
                except Exception as e:
                    _warn(f"_map_reduce_summarize({source}/{month}): 失敗 {e}")
                    continue
                summaries.append({"source": source, "month": month,
                                  "summary": res.get("text", ""), "row_count": len(chunk),
                                  "url": (_best_url(chunk[0]) if chunk else "") or "",
                                  "ref": (_best_ref(chunk[0]) if chunk else "") or ""})
    _dbg(f"_map_reduce_summarize: 中間要約{len(summaries)}件")
    return summaries


_SYNTHESIS_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "findings": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "statement": {"type": "string"},
                    "sources": {
                        "type": "array", "minItems": 1,
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "source_id": {"type": "string",
                                              "description": "材料の先頭に付いた [S番号] をそのまま入れる。"
                                                             "URLは書かない。番号を創作しない"},
                                "quote": {"type": ["string", "null"],
                                          "description": "ファクト分野なら材料からの逐語引用。該当なしはnull"},
                            },
                            "required": ["source_id", "quote"],
                        },
                    },
                },
                "required": ["statement", "sources"],
            },
        },
        "not_found": {"type": "array", "items": {"type": "string"},
                      "description": "根拠が見つからなかった論点"},
    },
    "required": ["findings", "not_found"],
}


def _best_url(row: dict) -> str:
    """行の代表URL（http で始まる一次情報）を優先順で1つ選ぶ。

    メール原文 > links 内の Drive リンク。無ければ空文字。
    ★LLMにURLを書かせず、この関数が選んだ本物URLを source_map に格納しておき、
      合成後に source_id から差し替える（URL破壊・幻URLの防止）。
    """
    gm = gmail_source_url(row)
    if gm:
        return gm
    for u in _extract_urls(row.get("links") or ""):
        if _classify_url(u) == "drive":
            return u
    return ""


def _best_ref(row: dict) -> str:
    """URLが無い行のための「たどれる場所」の文字列表現。

    agent_platform の行には Notion のようなページURLが無いので、
    file 行は notes のパス、それ以外は raw テーブルの行番号を出典に使う。
    """
    url = _best_url(row)
    if url:
        return url
    path = (row.get("note_path") or "").strip()
    if path:
        return f"notes/{path}"
    rid = row.get("_id") or row.get("id")
    return f"{RAW_TABLE}#{rid}" if rid else ""


def _materials_text(bundle: dict, deep_reads: list[dict],
                    max_rows_per_source: int = 60, *, body_chars: int = 400) -> tuple[str, dict]:
    """合成LLMに渡す材料テキストと出典対応表を返す。

    各材料に短いID [S番号] を振り、source_map[Sn] = {title, date, url, ref} を作る。
    LLM は source_id（Sn）だけを引用し、URL/場所はこちらが復元する。
    """
    parts: list[str] = []
    source_map: dict[str, dict] = {}
    counter = 0

    def _sid(title: str, date: str | None, url: str, ref: str) -> str:
        nonlocal counter
        counter += 1
        sid = f"S{counter}"
        source_map[sid] = {"title": title or "", "date": date or None,
                           "url": url or "", "ref": ref or ""}
        return sid

    projects = bundle.get("projects") or []
    if projects:
        parts.append(f"### ヒットの所属プロジェクト: {', '.join(projects)}")

    rows_by_source = bundle.get("rows") or {}
    for source, rows in rows_by_source.items():
        if not isinstance(rows, list) or not rows:
            continue
        parts.append(f"### {source} ({len(rows)}件中最大{max_rows_per_source}件)")
        for r in rows[:max_rows_per_source]:
            body = (r.get("content") or r.get("subject") or "")[:body_chars]
            who = r.get("name") or r.get("sender") or ""
            title = r.get("subject") or (r.get("content") or "")[:40] or source
            sid = _sid(title, r.get("timestamp"), _best_url(r), _best_ref(r))
            parts.append(f"- [{sid}] [{r.get('timestamp', '')}] {who}: {body}")

    MAX_THREADS = 15
    MAX_MSGS = 30
    for kind, threads in (bundle.get("threads") or {}).items():
        if not isinstance(threads, list) or not threads:
            continue
        parts.append(f"### {kind} スレッド全文 ({len(threads)}スレッド中最大{MAX_THREADS}件)")
        for t in threads[:MAX_THREADS]:
            parts.append(f"--- スレッド {t.get('thread_id')} ---")
            for m in (t.get("messages") or [])[:MAX_MSGS]:
                body = (m.get("content") or "")[:max(body_chars, 300)]
                title = (m.get("content") or "")[:40] or kind
                sid = _sid(title, m.get("timestamp"), _best_url(m), _best_ref(m))
                parts.append(f"- [{sid}] [{m.get('timestamp', '')}] "
                             f"{m.get('name') or m.get('sender', '')}: {body}")

    for dr in deep_reads:
        text = (dr.get("text") or "")[:2000]
        path = dr.get("note_path") or "深読み"
        sid = _sid(f"深読み:{path}", None, "", f"notes/{path}")
        parts.append(f"### [{sid}] 深読み ({dr.get('kind')}): {path}\n{text}")

    result = "\n".join(parts)
    _dbg(f"_materials_text: {len(result)}文字 出典{len(source_map)}件 (body_chars={body_chars})")
    return result, source_map


def _resolve_sources(data: dict, source_map: dict) -> dict:
    """LLMが返した source_id を、こちらが保持する本物の出典に差し替える。

    対応表に無いID（LLMの創作＝幻の出典）は捨てる。有効な出典が1つも無い finding も落とす。
    """
    out_findings: list[dict] = []
    dropped = 0
    for f in data.get("findings") or []:
        resolved: list[dict] = []
        for s in f.get("sources") or []:
            info = source_map.get((s.get("source_id") or "").strip())
            if not info:
                dropped += 1
                continue
            resolved.append({"title": info["title"], "date": info["date"],
                             "url": info["url"], "ref": info["ref"], "quote": s.get("quote")})
        if resolved:
            out_findings.append({"statement": f.get("statement", ""), "sources": resolved})
    if dropped:
        _warn(f"_resolve_sources: 対応表に無い source_id を {dropped} 件破棄")
    return {"findings": out_findings, "not_found": data.get("not_found") or []}


async def synthesize_answer(user_query: str, bundle: dict, deep_reads: list[dict], *,
                            chunk_summaries: list[dict] | None = None,
                            fact_critical: bool = False) -> dict:
    """材料から回答を合成する。出典は source_id 参照→URL/場所はコード側で復元。"""
    if chunk_summaries:
        # 月次チャンクにもIDを振り、代表出典（チャンク先頭行）を対応表に載せる
        source_map: dict[str, dict] = {}
        lines: list[str] = []
        for i, c in enumerate(chunk_summaries, 1):
            sid = f"S{i}"
            source_map[sid] = {"title": f"{c['source']}/{c['month']}",
                               "date": c["month"], "url": c.get("url") or "",
                               "ref": c.get("ref") or ""}
            lines.append(f"[{sid}] {c['summary']} (対象{c['row_count']}件)")
        materials = "\n".join(lines)
    else:
        materials, source_map = _materials_text(
            bundle, deep_reads, body_chars=1200 if fact_critical else 400)

    prompt = (
        f"以下の材料だけを根拠に、ユーザーの質問に答えてください。\n"
        f"各findingのsourcesには、材料の先頭に付いている [S番号] を source_id にそのまま入れてください。\n"
        f"URLは書かない・番号は創作しない（対応表に無い番号は無効になります）。\n"
        f"材料から分からないことは推測せず not_found に列挙してください。\n\n"
        f"以下のファクト分野に該当する記述は、sources.quote に材料の原文を一字一句変えず\n"
        f"逐語引用してください:\n{_FACT_DOMAINS}"
        f"該当しないfindingのquoteはnullでよい。\n\n"
        f"質問: {user_query}\n\n材料:\n{materials}"
    )
    try:
        data = await _call_structured_llm(
            prompt, model=SYNTHESIS_MODEL, schema=_SYNTHESIS_SCHEMA, label="synthesize"
        )
    except Exception as e:
        _warn(f"synthesize_answer: 合成失敗→空の結果にフォールバック: {e}")
        return {"findings": [], "not_found": ["回答の生成に失敗しました"]}
    return _resolve_sources(data, source_map)


# ==========================================================================
# エントリポイント
# ==========================================================================

async def answer_query(user_query: str, *, requester_hint: str | None = None) -> dict:
    """検索OSのエントリポイント: プランナー→取得→（1回だけ再計画）→深読み→出典付き合成。

    requester_hint: 「自分」等が指す依頼者の名前/メール（分かる場合に渡す）。

    戻り値: {"plan", "findings": [{statement, sources}], "not_found": [...]}
            人物が複数候補の場合は {"clarification": "...どちらですか"} を含む。
    """
    _dbg(f"===== answer_query 開始: {user_query!r} =====")
    try:
        plan = await plan_query(user_query, requester_hint=requester_hint)
        bundle = await _execute_plan(plan, requester_hint=requester_hint, user_query=user_query)

        if bundle.get("ambiguous_person"):
            amb = bundle["ambiguous_person"]
            listing = " / ".join(f"{c['name']}（{c['email']}）" for c in amb["candidates"])
            return {"plan": plan, "findings": [], "not_found": [],
                    "clarification": f"「{amb['query']}」に該当する人物が複数います: {listing}。"
                                     f"どの方か指定して聞き直してください。"}

        reason = _needs_replan(bundle)
        if reason:
            _dbg(f"answer_query: 再計画: {reason}")
            plan = await plan_query(user_query, requester_hint=requester_hint, replan_feedback=reason)
            bundle = await _execute_plan(plan, requester_hint=requester_hint, user_query=user_query)
            if bundle.get("ambiguous_person"):
                amb = bundle["ambiguous_person"]
                listing = " / ".join(f"{c['name']}（{c['email']}）" for c in amb["candidates"])
                return {"plan": plan, "findings": [], "not_found": [],
                        "clarification": f"「{amb['query']}」に該当する人物が複数います: {listing}。"}
    except Exception as e:
        _warn(f"answer_query: 取得段階で失敗: {type(e).__name__}: {e}")
        return {"plan": None, "findings": [],
                "not_found": [f"検索処理に失敗しました ({type(e).__name__}: {e})"]}

    chunk_summaries = None
    if bundle.get("needs_map_reduce"):
        chunk_summaries = await _map_reduce_summarize(bundle.get("rows") or {})

    result = await synthesize_answer(
        user_query, bundle, bundle.get("deep_reads") or [],
        chunk_summaries=chunk_summaries, fact_critical=bool(plan.get("fact_critical")),
    )
    _dbg(f"===== answer_query 完了: findings={len(result.get('findings', []))} "
         f"not_found={len(result.get('not_found', []))} =====")
    return {"plan": plan, "findings": result.get("findings", []),
            "not_found": result.get("not_found", [])}


async def answer_query_from_file(query_path: str = "tmp/query.txt", *,
                                 requester_hint: str | None = None) -> dict:
    """質問をファイル経由で受け取る版（引用符エスケープ事故の回避用）。"""
    with open(query_path, encoding="utf-8") as f:
        return await answer_query(f.read().strip(), requester_hint=requester_hint)
