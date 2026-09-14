"""raw DB（4ソース統合）を横断検索するモジュール v3 — 検索OS本体。

v2（notion_集約_v2.py・旧4DB構成）からの根本変更:
  1. 検索対象を統合 raw DB 1本に変更（source select で slack/gmail/drive/meeting を区別）。
  2. timestamp が date 型になったため、日付絞り込みはサーバー側の date フィルタ1発
     （v2 の「starts_with を日数ぶんOR展開」トリックは廃止）。
  3. プロパティ型は既知の固定スキーマ（_PROP_TYPES）なので NOTION_FETCH_DATABASE 不要。
     Phase 1 で使うツールは NOTION_QUERY_DATABASE_WITH_FILTER と llm_call の2つだけ。
  4. 「同じ列でもソースごとに意味が違う」を全部セマンティクス表（_BODY_FIELDS_BY_SOURCE /
     _PERSON_FIELDS_BY_SOURCE / PARENTS_SEMANTICS）に閉じ込め、上位層は表経由でのみ解釈する。

構成（検索OS_v3_設計.md 参照）:
  取得層: fetch_records / search（自然言語）
  解決層: resolve_person（人物辞書3層・複数人は候補提示）/ resolve_company / company_members
  文脈層: expand_thread（parents をソース別に解釈してスレッド復元）
  深読み層: fetch_meeting_minutes（meeting行のページ本文＝議事録+VTT全文。自分のDBのみ）
    ※Drive/Miro/Canva等の外部ファイルは検索時に再取得しない。書く側の要約JSON＋リンク提示で足りる
     （外部アプリごとにリーダーを増やすとキリがないため。要約は保存時に済ませる設計原則③）
  合成層: answer_query（プランナー→取得→再計画→深読み→出典必須の構造化回答）

【呼び出し方（code_execute）】質問やキーワードはトリプルシングルクォートで直接埋め込む
（二重引用符3連続はこのモジュールの docstring を壊した事故実績があるため使わない）:
  import sys
  if "rag" not in sys.path:
      sys.path.insert(0, "rag")
  from notion_集約_v3 import answer_query
  result = await answer_query('''アルファリースの見積もりいくらだった？''')
  print(result)
"""

from __future__ import annotations

import asyncio
import os
import random
import re
from datetime import datetime, timedelta, timezone

from agent_sdk import (
    NOTION_FETCH_ALL_BLOCK_CONTENTS,
    NOTION_QUERY_DATABASE_WITH_FILTER,
    llm_call,
    ToolCallError,
)

# --- デバッグ出力（実行サンドボックスは stdout のみ返るため print ベース） ---
DEBUG = os.environ.get("ASKHUB_DEBUG", "1") != "0"


def _dbg(msg: str) -> None:
    if DEBUG:
        print(f"[DEBUG] {msg}", flush=True)


def _warn(msg: str) -> None:
    """異常系は DEBUG フラグに関係なく常に出す。"""
    print(f"[WARN] {msg}", flush=True)


# ========== DB ID（固定） ==========

# 統合 raw DB「社内情報集約PJ > raw」（全4ソースの保存先）
RAW_DB_ID = "<RAW_DB_ID>"
# 人物辞書「meta > 人」（Phase 2 の resolve_person が使用）
PEOPLE_DB_ID = "<PEOPLE_DB_ID>"
# domain_map（プロジェクト⇄ドメイン対応・手動管理。Phase 2 の resolve_company が使用）
DOMAIN_MAP_DB_ID = "<DOMAIN_MAP_DB_ID>"

JST = timezone(timedelta(hours=9))

# ========== raw DB のセマンティクス表（v3の心臓部） ==========
# 「同じ列でもソースごとに意味が違う」の解釈は必ずこの表を経由する。

# プロパティ型（raw DB は固定スキーマなのでAPI照会不要）
_PROP_TYPES: dict[str, str] = {
    "content": "title",
    "subject": "rich_text",
    "timestamp": "date",
    "sender": "rich_text",
    "name": "rich_text",
    "project": "rich_text",
    "source": "select",
    "links": "rich_text",
    "to": "rich_text",
    "cc_bcc": "rich_text",
    "parents": "rich_text",
}

_ALL_SOURCES: list[str] = ["slack", "gmail", "drive", "meeting"]

# ソース別: キーワード横断検索の対象フィールド
# （driveのparentsはフォルダ名チェーンなので検索対象に含める。他ソースのparentsは
#  ts/threadId/uuidでありキーワード検索には無意味なので含めない）
_BODY_FIELDS_BY_SOURCE: dict[str, list[str]] = {
    "slack": ["content", "links"],
    "gmail": ["content", "subject"],
    "drive": ["content", "subject", "parents"],
    "meeting": ["content", "subject"],
}

# ソース別: 人物（名前・メール）を探すフィールド（Phase 2 の resolve_person 展開先）
_PERSON_FIELDS_BY_SOURCE: dict[str, dict[str, list[str]]] = {
    "slack": {"email": ["sender"], "name": ["name"]},
    "gmail": {"email": ["sender", "to", "cc_bcc"], "name": ["name"]},
    "drive": {"email": ["sender"], "name": ["name"]},
    "meeting": {"email": ["sender"], "name": ["name"]},
}

# parents 列の意味（検索OSはこの表を通してだけ解釈する）
PARENTS_SEMANTICS: dict[str, str] = {
    "slack": "thread_parent_ts",  # 親メッセージの時刻(ISO文字列)。空=スレッド親 or 単発
    "gmail": "thread_id",         # スレッド全行が同じ値を持つグループID
    "meeting": "meeting_uuid",    # 会議の一意キー（復元対象なし。定例の回ごとに異なる）
    "drive": "origin",            # フォルダ名チェーン / "slack" / "gmail" / ""(SlackURL由来)
}

# gmail行の content 末尾に埋まっている重複判定キーから原文URLを復元するパターン
_MAIL_ID_RE = re.compile(r"\[mail:([0-9a-fA-F]+)\]")

# プロジェクト未分類の行（Zoom会議等）の project 値
UNCLASSIFIED_PROJECT = "未分類"

# ========== LLM モデル設定 ==========

# 自然言語クエリの解析用（v1/v2 の search() で実運用実績のあるモデル）
SEARCH_PARSER_MODEL = "anthropic/claude-haiku-4-5"
# 構造化出力の失敗時に切り替えるフォールバック（同一モデルの再試行は無意味な実績があるため）
STRUCTURED_FALLBACK_MODEL = "anthropic/claude-sonnet-4-6"

# 1ソースの取得上限（暴走防止。超えたら警告して打ち切り）
MAX_ROWS_PER_SOURCE = 500

_MAX_RETRIES = 3
_INITIAL_BACKOFF_SEC = 1.0


# ========== リトライ付きツール呼び出し ==========

async def _call_with_retry(tool, **kwargs):
    """一時的エラーに対してジッター付き指数バックオフでリトライする。"""
    last_err = None
    for attempt in range(_MAX_RETRIES):
        try:
            return await tool(**kwargs)
        except Exception as e:
            last_err = e
            if attempt == _MAX_RETRIES - 1:
                break
            wait = _INITIAL_BACKOFF_SEC * (2 ** attempt) + random.uniform(0, 1)
            _warn(f"{getattr(tool, '__name__', 'tool')} retry {attempt + 1}/{_MAX_RETRIES} after {wait:.1f}s: {e}")
            await asyncio.sleep(wait)
    raise ToolCallError(f"ツール呼び出しが{_MAX_RETRIES}回失敗: {last_err}")


# ========== 構造化出力 llm_call（v2から移植） ==========

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

def _extract_row(page: dict) -> dict:
    """Notion行 → フラットな dict（プロパティを平文化）。"""
    props = page.get("properties") or {}
    out: dict = {
        "id": page.get("id"),
        "created_time": page.get("created_time"),
        "url": page.get("url"),
    }
    for name, info in props.items():
        ptype = info.get("type") or _PROP_TYPES.get(name) or ""
        if ptype == "title":
            out[name] = "".join((x.get("plain_text") or "") for x in (info.get("title") or []))
        elif ptype == "rich_text":
            out[name] = "".join((x.get("plain_text") or "") for x in (info.get("rich_text") or []))
        elif ptype == "select":
            out[name] = (info.get("select") or {}).get("name") or ""
        elif ptype == "date":
            out[name] = (info.get("date") or {}).get("start") or ""
        elif ptype == "url":
            out[name] = info.get("url") or ""
        else:
            out[name] = info.get(ptype)
    return out


def gmail_source_url(row: dict) -> str | None:
    """gmail行の content 末尾の [mail:ID] から Gmail 原文URLを復元する。

    出典として提示すると一次情報（メール原文）に1クリックで届く。
    gmail 以外の行・IDが無い行は None。
    """
    if row.get("source") != "gmail":
        return None
    m = _MAIL_ID_RE.search(row.get("content") or "")
    return f"https://mail.google.com/mail/u/0/#all/{m.group(1)}" if m else None


# ========== フィルタ組み立て ==========

def _leaf_or_group(conditions: list[dict]) -> dict | None:
    """0件→None / 1件→そのまま / 2件以上→or。
    Notion の Compound filter は2階層までなので、and の子は常に
    「リーフ」または「リーフのみのor」に保つ（or の中に or を入れない）。
    """
    conditions = [c for c in conditions if c]
    if not conditions:
        return None
    if len(conditions) == 1:
        return conditions[0]
    return {"or": conditions}


def _project_variants(name: str) -> list[str]:
    """project 照合の候補値（pjt_ prefix の有無を吸収）。"""
    n = (name or "").strip()
    if not n:
        return []
    variants = {n, n if n.startswith("pjt_") else f"pjt_{n}"}
    return sorted(variants)


def _project_group(project: str | None, *, include_unclassified: bool) -> dict | None:
    """project 絞り込みの条件グループを作る。

    include_unclassified=True のとき「未分類」も対象に含める
    （キーワード検索時に Zoom の未分類会議を取りこぼさないため）。
    """
    conditions = [
        {"property": "project", "rich_text": {"equals": v}}
        for v in _project_variants(project or "")
    ]
    if conditions and include_unclassified:
        conditions.append({"property": "project", "rich_text": {"equals": UNCLASSIFIED_PROJECT}})
    return _leaf_or_group(conditions)


def _keyword_group(source: str, keywords: list[str]) -> dict | None:
    """キーワードをソース別の本文系フィールドに OR 展開する。"""
    fields = _BODY_FIELDS_BY_SOURCE.get(source, [])
    kws = [k.strip() for k in (keywords or []) if k and k.strip()]
    if not fields or not kws:
        return None
    conditions = [
        {"property": f, _PROP_TYPES[f]: {"contains": kw}}
        for f in fields for kw in kws
    ]
    return _leaf_or_group(conditions)


def _date_leaves(start_date: str | None, end_date: str | None) -> list[dict]:
    """timestamp（date型）のサーバー側絞り込み条件。"""
    leaves = []
    if start_date:
        leaves.append({"property": "timestamp", "date": {"on_or_after": f"{start_date}T00:00:00+09:00"}})
    if end_date:
        leaves.append({"property": "timestamp", "date": {"on_or_before": f"{end_date}T23:59:59+09:00"}})
    return leaves


def _build_source_filter(
    source: str,
    project: str | None,
    start_date: str | None,
    end_date: str | None,
    keywords: list[str] | None,
    extra_conditions: list[dict] | None = None,
) -> dict:
    """1ソースぶんの Notion filter を組み立てる（and の子は2階層制約を守る）。"""
    groups: list[dict] = [{"property": "source", "select": {"equals": source}}]
    pg = _project_group(project, include_unclassified=bool(keywords))
    if pg:
        groups.append(pg)
    groups.extend(_date_leaves(start_date, end_date))
    kg = _keyword_group(source, keywords or [])
    if kg:
        groups.append(kg)
    for c in extra_conditions or []:
        groups.append(c)
    return {"and": groups} if len(groups) > 1 else groups[0]


# ========== 取得層 ==========

async def _query_raw(filter_obj: dict, *, label: str, max_rows: int = MAX_ROWS_PER_SOURCE) -> list[dict]:
    """raw DB をページネーション付きでクエリし、整形済み行を返す（timestamp降順）。"""
    rows: list[dict] = []
    cursor: str | None = None
    page_num = 0
    while True:
        page_num += 1
        kwargs: dict = {
            "database_id": RAW_DB_ID,
            "page_size": 100,
            "filter": filter_obj,
            "sorts": [{"property": "timestamp", "direction": "descending"}],
        }
        if cursor:
            kwargs["start_cursor"] = cursor
        resp = await _call_with_retry(NOTION_QUERY_DATABASE_WITH_FILTER, **kwargs)
        data = resp.get("data", resp) or {}
        page_rows = data.get("results") or []
        rows.extend(_extract_row(r) for r in page_rows)
        _dbg(f"_query_raw({label}) page={page_num}: {len(page_rows)}件 (累計{len(rows)})")
        if len(rows) >= max_rows:
            _warn(f"_query_raw({label}): {max_rows}件の上限に達したため打ち切り（条件を絞ってください）")
            return rows[:max_rows]
        if not data.get("has_more"):
            break
        cursor = data.get("next_cursor")
        if not cursor:
            break
    return rows


async def fetch_records(
    sources: list[str] | None = None,
    project: str | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
    keywords: list[str] | None = None,
    extra_conditions_by_source: dict[str, list[dict]] | None = None,
) -> dict[str, list[dict]]:
    """raw DB を横断取得する（v3 取得層の公開API）。

    Args:
        sources: 対象ソース。省略時は全4ソース。
        project: プロジェクト名（pjt_ prefix の有無は自動吸収）。
        start_date / end_date: "YYYY-MM-DD"。timestamp（date型）でサーバー側絞り込み。
        keywords: 本文系フィールド横断の contains OR 検索（対象フィールドはソース別）。
        extra_conditions_by_source: ソース別の追加リーフ条件（Phase 2 の人物・スレッド検索が使用）。

    Returns:
        {source: [row, ...]} — 取得失敗したソースは空リスト＋WARN。
        キーワード指定時は project 絞りに「未分類」も自動で含める。
    """
    targets = [s for s in (sources or _ALL_SOURCES) if s in _ALL_SOURCES]
    _dbg(
        f"fetch_records: sources={targets} project={project!r} "
        f"date={start_date}~{end_date} keywords={keywords}"
    )

    async def _one(source: str) -> list[dict]:
        filter_obj = _build_source_filter(
            source, project, start_date, end_date, keywords,
            (extra_conditions_by_source or {}).get(source),
        )
        return await _query_raw(filter_obj, label=source)

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
                "ファイル・資料・ドライブ→drive / 会議・議事録・MTG→meeting。指定なければ全部"
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
    """自然言語クエリを解析して raw DB から横断取得する。

    例:
      await search('''アルファリースの6月のメール''')
      await search('''見積もりについて話した会議''')

    ※ 人名での検索は Phase 2 の resolve_person 実装で精度が上がる。
      現状は名前がキーワードとして本文・name列にヒットする範囲で動く。
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
    _dbg(f"search: 解析結果 project={project!r} sources={sources} date={start_date}~{end_date} keywords={keywords}")

    return await fetch_records(
        sources=sources, project=project,
        start_date=start_date, end_date=end_date, keywords=keywords,
    )


async def search_from_file(query_path: str = "tmp/query.txt") -> dict[str, list[dict]]:
    """質問をファイル経由で受け取る版（引用符エスケープ事故の回避用）。

    1. file_create(sandbox_path="tmp/query.txt", content="<質問そのまま>")
    2. code_execute: from notion_集約_v3 import search_from_file; result = await search_from_file()
    """
    with open(query_path, encoding="utf-8") as f:
        return await search(f.read().strip())


# ==========================================================================
# 解決層（Phase 2）: 人物・会社をエンティティとして解決する
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
    （ひらがな呼び等）はクエリ側で変換して照合する（設計md §3）。
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


async def resolve_person(query: str) -> dict:
    """人物を名寄せする（解決層）。3層＋安全弁:

    ① クエリ正規化＋LLM表記変換 → ② 人物辞書を候補で検索
       - 1人に確定 → email＋aliases一式を検索用バリアントとして返す
       - 複数人 → ★混ぜずに candidates を返す（呼び出し元がユーザーに候補提示）
    ③ 0件なら raw の逆引き（slack/drive行は1行=1人で name⇄sender が正確なペア）
    ④ それでも不明 → 入力文字列そのままの contains にフォールバック

    戻り値:
      {"status": "resolved",   "query", "email", "variants": [...]}
      {"status": "ambiguous",  "query", "candidates": [{"name","email","company"}]}
      {"status": "unresolved", "query", "variants": [元の表記]}
    """
    variants = await _generate_name_variants(query)
    _dbg(f"resolve_person: {query!r} → 変換候補 {variants}")
    if not variants:
        return {"status": "unresolved", "query": query, "variants": [query]}

    # ①② 人物辞書を表記候補で検索
    conditions: list[dict] = []
    for v in variants:
        conditions.append({"property": "name", "title": {"contains": v}})
        conditions.append({"property": "aliases", "rich_text": {"contains": v}})
        conditions.append({"property": "email", "rich_text": {"contains": v.lower()}})
    rows: list[dict] = []
    try:
        resp = await _call_with_retry(
            NOTION_QUERY_DATABASE_WITH_FILTER,
            database_id=PEOPLE_DB_ID, filter=_leaf_or_group(conditions), page_size=10,
        )
        data = resp.get("data", resp) or {}
        rows = [_extract_row(r) for r in (data.get("results") or [])]
    except Exception as e:
        _warn(f"resolve_person: 人物辞書の検索失敗: {e}")

    by_email: dict[str, dict] = {}
    for r in rows:
        email = (r.get("email") or "").strip().lower()
        if email:
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

    # ③ raw 逆引き（slack/drive 行は 1行=1人 で name⇄sender が正確に対応する）
    for v in variants:
        try:
            filter_obj = {"and": [
                _leaf_or_group([
                    {"property": "source", "select": {"equals": "slack"}},
                    {"property": "source", "select": {"equals": "drive"}},
                ]),
                {"property": "name", "rich_text": {"contains": v}},
            ]}
            raw_rows = await _query_raw(filter_obj, label=f"person_fallback:{v}", max_rows=10)
        except Exception as e:
            _warn(f"resolve_person: raw逆引き失敗 ({v!r}): {e}")
            continue
        emails = {
            (r.get("sender") or "").strip().lower()
            for r in raw_rows
            if "@" in (r.get("sender") or "") and "," not in (r.get("sender") or "")
        }
        if len(emails) == 1:
            email = next(iter(emails))
            _dbg(f"resolve_person: raw逆引きで確定 {query!r} → {email}")
            return {"status": "resolved", "query": query, "email": email, "variants": variants}

    _dbg(f"resolve_person: 未解決 {query!r} → 文字列一致にフォールバック")
    return {"status": "unresolved", "query": query, "variants": variants}


def person_conditions(person: dict) -> dict[str, list[dict]]:
    """resolve_person の結果を fetch_records の extra_conditions_by_source に変換する。

    ソース別の人物フィールド表（_PERSON_FIELDS_BY_SOURCE）に従い、
    メールはメール系の列へ・表記は name 列へ contains の OR 条件として展開する。
    """
    email = (person.get("email") or "").strip().lower()
    variants = [v for v in (person.get("variants") or []) if v and len(v) >= 2]
    out: dict[str, list[dict]] = {}
    for source, fields in _PERSON_FIELDS_BY_SOURCE.items():
        conds: list[dict] = []
        if email:
            for f in fields["email"]:
                conds.append({"property": f, _PROP_TYPES[f]: {"contains": email}})
        for v in variants:
            for f in fields["name"]:
                conds.append({"property": f, _PROP_TYPES[f]: {"contains": v}})
        group = _leaf_or_group(conds)
        if group:
            out[source] = [group]
    return out


def _page_title_text(page: dict) -> str:
    """title 型プロパティを名前によらず探して平文で返す（domain_map 用）。"""
    for v in (page.get("properties") or {}).values():
        if (v or {}).get("type") == "title":
            return "".join((x.get("plain_text") or "") for x in (v.get("title") or []))
    return ""


async def _load_domain_map() -> list[dict]:
    """domain_map（PJ⇄ドメイン・手動管理）を全件読む。[{project, domain}]。

    小さいDBなので都度全件でよい。列名は domain（旧「ドメイン」にも両対応）。
    """
    rows: list[dict] = []
    cursor = None
    while True:
        kwargs: dict = {"database_id": DOMAIN_MAP_DB_ID, "page_size": 100}
        if cursor:
            kwargs["start_cursor"] = cursor
        resp = await _call_with_retry(NOTION_QUERY_DATABASE_WITH_FILTER, **kwargs)
        data = resp.get("data", resp) or {}
        for page in data.get("results") or []:
            name = _page_title_text(page).strip()
            props = page.get("properties") or {}
            d = props.get("domain") or props.get("ドメイン") or {}
            domain = "".join(x.get("plain_text", "") for x in (d.get("rich_text") or []))
            if name:
                rows.append({"project": name, "domain": domain.strip().lstrip("@").lower()})
        if not data.get("has_more"):
            break
        cursor = data.get("next_cursor")
        if not cursor:
            break
    return rows


def _normalize_name(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"[（(\[].*?[）)\]]", "", s)
    s = re.sub(r"\s+", "", s)
    return s.lower()


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
    if not (domain or "").strip():
        return []
    resp = await _call_with_retry(
        NOTION_QUERY_DATABASE_WITH_FILTER,
        database_id=PEOPLE_DB_ID,
        filter={"property": "company", "rich_text": {"equals": domain.strip().lower()}},
        page_size=min(limit, 100),
    )
    data = resp.get("data", resp) or {}
    return [_extract_row(r) for r in (data.get("results") or [])]


# ==========================================================================
# 文脈層（Phase 2）: parents をソース別に解釈してスレッドを復元する
# ==========================================================================

async def expand_thread(row: dict) -> list[dict]:
    """行が属する会話全体を復元する（PARENTS_SEMANTICS に基づくソース別処理）。

    gmail: 同じ threadId の全メール / slack: 親メッセージ＋返信 /
    meeting・drive: その行のみ（復元対象なし）。時系列昇順で返す。
    """
    source = row.get("source") or ""
    if source == "gmail":
        tid = (row.get("parents") or "").strip()
        if not tid:
            return [row]
        rows = await _query_raw({"and": [
            {"property": "source", "select": {"equals": "gmail"}},
            {"property": "parents", "rich_text": {"equals": tid}},
        ]}, label="gmail_thread", max_rows=100)
        return sorted(rows, key=lambda r: r.get("timestamp") or "") or [row]

    if source == "slack":
        # 自分が返信なら parents=親ts、自分が親/単発なら自分の timestamp が親ts
        parent_ts = (row.get("parents") or "").strip() or (row.get("timestamp") or "").strip()
        if not parent_ts:
            return [row]
        project = (row.get("project") or "").strip()
        base = [{"property": "source", "select": {"equals": "slack"}}]
        if project:
            # 親tsはチャンネル間で理論上衝突しうるため project でも絞る（設計md §2）
            base.append({"property": "project", "rich_text": {"equals": project}})
        children = await _query_raw(
            {"and": base + [{"property": "parents", "rich_text": {"equals": parent_ts}}]},
            label="slack_thread", max_rows=100,
        )
        # 親メッセージ本体: date型の等値比較は精度が怖いので±1分で取ってクライアント側一致
        parent_rows: list[dict] = []
        try:
            pdt = datetime.fromisoformat(parent_ts)
            lo = (pdt - timedelta(minutes=1)).isoformat(timespec="seconds")
            hi = (pdt + timedelta(minutes=1)).isoformat(timespec="seconds")
            around = await _query_raw(
                {"and": base + [
                    {"property": "timestamp", "date": {"on_or_after": lo}},
                    {"property": "timestamp", "date": {"on_or_before": hi}},
                ]},
                label="slack_thread_parent", max_rows=20,
            )
            parent_rows = [r for r in around if (r.get("timestamp") or "")[:19] == parent_ts[:19]]
        except ValueError:
            _warn(f"expand_thread: 親ts をパースできません: {parent_ts!r}")
        merged = {r["id"]: r for r in children + parent_rows + [row]}
        return sorted(merged.values(), key=lambda r: r.get("timestamp") or "")

    return [row]


# ==========================================================================
# 深読み層（Phase 3）: 必要な行だけ本文まで読む
# ==========================================================================

_HTTP_URL_RE = re.compile(r"https?://[^\s<>|\"')]+")
_PAGE_ID_IN_TEXT_RE = re.compile(r"[0-9a-fA-F]{32}")

MAX_AUTO_DEEP_READS = 5      # 1クエリで自動深読みする最大件数
DEEP_READ_TIER1_MAX = 5      # 候補がこの件数以下なら選別せず全部読む
MAX_THREAD_EXPANSION_SEEDS = 10  # スレッド展開する種メッセージ数の上限
MAX_CHUNK_ROWS = 40          # map-reduce 中間要約1チャンクの最大行数

PLANNER_MODEL = "anthropic/claude-sonnet-5"
SYNTHESIS_MODEL = PLANNER_MODEL
LITE_MODEL = "gemini/gemini-3.1-flash-lite"


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


def _extract_page_id(page_id_or_url: str) -> str:
    s = (page_id_or_url or "").strip().replace("-", "")
    m = _PAGE_ID_IN_TEXT_RE.findall(s)
    if not m:
        raise ValueError(f"NotionページIDを抽出できません: {page_id_or_url!r}")
    h = m[-1].lower()
    return f"{h[0:8]}-{h[8:12]}-{h[12:16]}-{h[16:20]}-{h[20:32]}"


async def fetch_meeting_minutes(page_id_or_url: str) -> dict:
    """meeting行のページ本文（議事録＋生VTT全文）を取得する（深読みのStage2）。

    戻り値: {"page_id", "text", "block_count", "truncated"}
    """
    page_id = _extract_page_id(page_id_or_url)
    try:
        res = await _call_with_retry(NOTION_FETCH_ALL_BLOCK_CONTENTS, block_id=page_id)
    except Exception as e:
        _warn(f"fetch_meeting_minutes: 取得失敗 page_id={page_id}: {e}")
        return {"page_id": page_id, "text": "", "block_count": 0, "truncated": False}
    if isinstance(res, list):
        blocks, has_more = res, False
    else:
        data = res.get("data") or res
        blocks = data.get("results", [])
        has_more = bool(data.get("has_more"))
    if not isinstance(blocks, list):
        blocks = []
    lines = []
    for b in blocks:
        btype = b.get("type", "")
        rich = (b.get(btype) or {}).get("rich_text", []) if btype else []
        t = "".join(rt.get("plain_text", "") for rt in rich)
        if t:
            lines.append(t)
    text = "\n".join(lines)
    _dbg(f"fetch_meeting_minutes: blocks={len(blocks)} truncated={has_more} len={len(text)}")
    return {"page_id": page_id, "text": text, "block_count": len(blocks), "truncated": has_more}


def _collect_deep_read_candidates(rows_by_source: dict, threads: dict) -> list[dict]:
    """深読み候補を集める＝meeting行のページ本文（自分のDB）のみ。

    Drive/Miro/Canva 等の外部ファイルは検索時に再取得しない（要約JSON＋リンク提示で足りる。
    外部アプリごとにリーダーを増やすとキリがないため）。合成層は _materials_text で
    drive行の要約と Drive リンクを出典として渡す。
    """
    seen: set[str] = set()
    candidates: list[dict] = []

    def _add(url: str | None, kind: str, hint: str) -> None:
        if url and url not in seen:
            seen.add(url)
            candidates.append({"url": url, "kind": kind, "hint": (hint or "")[:200]})

    for r in rows_by_source.get("meeting") or []:
        _add(r.get("url"), "meeting_page", f"{r.get('subject') or ''} {(r.get('content') or '')[:120]}")
    _dbg(f"_collect_deep_read_candidates: 候補{len(candidates)}件（meetingページ本文のみ）")
    return candidates


_TRIAGE_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "needed_urls": {
            "type": "array", "items": {"type": "string"},
            "description": "本文まで読む必要があるURL（最大5件、必要性の高い順）",
        },
    },
    "required": ["needed_urls"],
}


async def _triage_deep_read(user_query: str, candidates: list[dict]) -> list[dict]:
    """候補が多い場合に軽量LLMで「本文まで読む必要があるか」を選別する（Tier2）。"""
    lines = "\n".join(f"- url={c['url']} kind={c['kind']} 手がかり={c['hint']}" for c in candidates)
    prompt = (
        f"以下は検索でヒットした候補です。質問に答えるために本文まで読む必要がある\n"
        f"候補だけを最大5件、必要性の高い順に needed_urls に入れてください。\n"
        f"手がかりだけで答えられそうなものは選ばないこと。\n\n質問: {user_query}\n\n候補:\n{lines}"
    )
    try:
        data = await _call_structured_llm(prompt, model=LITE_MODEL, schema=_TRIAGE_SCHEMA, label="triage")
    except Exception as e:
        _warn(f"_triage_deep_read: 選別失敗→先頭{MAX_AUTO_DEEP_READS}件を読む: {e}")
        return candidates[:MAX_AUTO_DEEP_READS]
    needed = set(data.get("needed_urls") or [])
    return [c for c in candidates if c["url"] in needed]


async def _deep_read_urls(selected: list[dict]) -> list[dict]:
    """選別済みの候補（meeting行のページ本文のみ）を深読みする（1件の失敗で全体を止めない）。"""
    results: list[dict] = []
    for c in selected[:MAX_AUTO_DEEP_READS]:
        try:
            minutes = await fetch_meeting_minutes(c["url"])
            results.append({"kind": c["kind"], "url": c["url"], **minutes})
        except Exception as e:
            _warn(f"_deep_read_urls: 失敗 {c['url']}: {e}")
    _dbg(f"_deep_read_urls: {len(results)}件を深読み完了")
    return results


# ==========================================================================
# プランナー（Phase 3）
# ==========================================================================

# ファクト（正確な事実・原文）が要求される分野。共通項は「後から言った/言わないになり得るもの」
_FACT_DOMAINS = """\
1. 顧客の発言・要望 / 2. 要件定義・仕様の合意内容 / 3. 議事録の決定事項（誰が・いつ・何を）
4. 金額・見積もり・契約条件 / 5. 納期・期日のコミットメント / 6. 承認・責任の所在
7. クレーム・トラブル対応の経緯 / 8. 数値実績・KPI / 9. 対外的に送った公式な文面
10. セキュリティ・データ取り扱いの取り決め
"""

_DB_KNOWLEDGE = """\
【raw DB の構造（全ソース共通11プロパティ・source で種別分け）】
- source=slack:   content=メッセージ本文, name=投稿者, links=共有URLとその要約
- source=gmail:   content=メール本文, subject=件名, sender/to/cc_bcc=アドレス, name=送受信者名
- source=drive:   content=ファイルの要約, subject=ファイル名, parents=フォルダ階層/出どころ
- source=meeting: content=議事録全文, subject=会議タイトル, name=参加者, sender=参加者メール
- 共通: project=pjt_プロジェクト名, timestamp=日時

【project_name の注意】
- pjt_ prefix の有無は自動吸収される。
- 「〜エージェント」「〜ツール」のような社内の仕組み名は project 名と一致しない場合が
  あるので、ユーザーの言葉やその略称を必ず keywords にも入れて本文横断検索できるようにする。
- 何のプロジェクトか不明なら project_name は null にして keywords で横断検索する。
"""

_RECIPES: list[dict] = [
    {"key": "decision_rationale", "label": "決定事項の根拠確認",
     "how": "keywordsで絞り込み→slack/gmailスレッド復元→meetingは全文深読み。経緯を時系列で提示。"},
    {"key": "past_reference_docs", "label": "過去資料の参照",
     "how": "source=driveをファイル名/要約/フォルダ名で検索→上位候補の本文をDLして読む。"},
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
        "deep_read": {"type": "boolean", "description": "議事録全文/ファイル本文の2段階取得を行うか"},
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
# レシピ実行（Phase 3）
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
    extra = person_conditions(person) if person and person.get("status") != "ambiguous" else None
    # ★レシピはソースを絞らない＝常に全4ソースを見る（統合rawDBの意味。plan の sources は無視）。
    #   絞り込みは project / キーワード / 期間 / 人物で行う。取りこぼしを設計レベルで防ぐ。
    rows = await fetch_records(
        sources=None, project=plan.get("project_name"),
        start_date=plan.get("start_date"), end_date=plan.get("end_date"),
        keywords=plan.get("keywords") or [],
        extra_conditions_by_source=extra,
    )
    threads = {"slack": await _expand_slack_threads(rows)} if rows.get("slack") else {}
    candidates = _collect_deep_read_candidates(rows, threads)
    fact = bool(plan.get("fact_critical"))
    if not candidates:
        deep_reads: list[dict] = []
    elif fact or plan.get("deep_read") or len(candidates) <= DEEP_READ_TIER1_MAX:
        deep_reads = await _deep_read_urls(candidates[:MAX_AUTO_DEEP_READS])
    else:
        deep_reads = await _deep_read_urls(await _triage_deep_read(user_query, candidates))
    return {"rows": rows, "threads": threads, "projects": _collect_projects(rows), "deep_reads": deep_reads}


def _sort_desc(rows: list[dict]) -> list[dict]:
    """timestamp 降順（新しい順）。timestamp 無しは末尾。"""
    return sorted(rows or [], key=lambda r: r.get("timestamp") or "", reverse=True)


async def _handle_latest_minutes(plan: dict) -> dict:
    # 「最新の議事録」質問だが、全ソースを取得して meeting を新しい順に前面に出す
    # （議事録の場所が slack/gmail で共有されているケースも拾える）。
    rows_all = await fetch_records(project=plan.get("project_name"),
                                   keywords=plan.get("keywords") or [])
    rows = {
        "meeting": _sort_desc(rows_all.get("meeting"))[:5],
        "gmail": _sort_desc(rows_all.get("gmail"))[:3],
        "slack": _sort_desc(rows_all.get("slack"))[:3],
        "drive": _sort_desc(rows_all.get("drive"))[:3],
    }
    return {"rows": rows, "projects": _collect_projects(rows), "deep_reads": []}


async def _handle_past_reference_docs(plan: dict, *, user_query: str) -> dict:
    # 資料探しでも全ソースを見る（資料が slack/gmail で言及・共有されている場合を拾う）。
    # drive行は要約JSON＋リンクで提示、meeting行のみ本文深読み（外部ファイルは再取得しない）。
    rows = await fetch_records(
        project=plan.get("project_name"),
        start_date=plan.get("start_date"), end_date=plan.get("end_date"),
        keywords=plan.get("keywords") or [],
    )
    candidates = _collect_deep_read_candidates(rows, {})
    deep_reads = await _deep_read_urls(candidates[:MAX_AUTO_DEEP_READS]) if candidates else []
    return {"rows": rows, "projects": _collect_projects(rows), "deep_reads": deep_reads}


async def _handle_temperature_check(plan: dict) -> dict:
    # 温度感でも全ソースを見る（slack/driveにやり取りが出ることが多い）。
    # 「軽さ」はソースを削ることではなく「各ソース新しい順に少数」で担保する。
    company = await resolve_company(plan.get("project_name") or "")
    project = company.get("project") or plan.get("project_name")
    rows_all = await fetch_records(project=project, keywords=plan.get("keywords") or [])
    rows = {
        "meeting": _sort_desc(rows_all.get("meeting"))[:2],
        "gmail": _sort_desc(rows_all.get("gmail"))[:10],
        "slack": _sort_desc(rows_all.get("slack"))[:10],
        "drive": _sort_desc(rows_all.get("drive"))[:5],
    }
    deep_reads = await _deep_read_urls(
        [{"url": r.get("url"), "kind": "meeting_page", "hint": r.get("subject") or ""}
         for r in rows["meeting"] if r.get("url")]
    )
    return {"rows": rows, "company": company, "projects": _collect_projects(rows), "deep_reads": deep_reads}


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
        extra_conditions_by_source=person_conditions(person),
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
# 合成層（Phase 3）
# ==========================================================================

async def _map_reduce_summarize(rows_by_source: dict) -> list[dict]:
    """月次チャンクで中間要約（軽量LLM）を作る。レコード数が多い全体まとめ専用。"""
    summaries: list[dict] = []
    for source, rows in rows_by_source.items():
        if not isinstance(rows, list) or not rows:
            continue
        by_month: dict[str, list[dict]] = {}
        for r in rows:
            month = (r.get("timestamp") or r.get("created_time") or "")[:7] or "unknown"
            by_month.setdefault(month, []).append(r)
        for month, month_rows in sorted(by_month.items()):
            for i in range(0, len(month_rows), MAX_CHUNK_ROWS):
                chunk = month_rows[i:i + MAX_CHUNK_ROWS]
                text = "\n---\n".join(
                    f"[{r.get('url', '')}] {r.get('timestamp', '')}: "
                    f"{(r.get('content') or r.get('subject') or '')[:500]}"
                    for r in chunk
                )
                try:
                    res = await llm_call(
                        prompt=(f"{source}の{month}分のログです。主要な出来事・決定事項を"
                                f"箇条書きで要約してください（各行に出典URLを残すこと）。\n\n{text}"),
                        model=LITE_MODEL,
                    )
                except Exception as e:
                    _warn(f"_map_reduce_summarize({source}/{month}): 失敗 {e}")
                    continue
                summaries.append({"source": source, "month": month,
                                  "summary": res.get("text", ""), "row_count": len(chunk),
                                  "url": (chunk[0].get("url") if chunk else "") or ""})
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
    """行の代表URLを優先順で1つ選ぶ: メール原文 > Driveファイル直リンク > Notion行URL。

    ★LLMにURLを書かせず、この関数が選んだ本物URLを source_map に格納しておき、
      合成後に source_id から差し替える（URL破壊・幻URLの防止）。
    """
    gm = gmail_source_url(row)
    if gm:
        return gm
    if (row.get("source") or "") == "drive":
        for u in _extract_urls(row.get("links") or ""):
            if _classify_url(u) == "drive":
                return u
    return row.get("url") or ""


def _materials_text(bundle: dict, deep_reads: list[dict],
                    max_rows_per_source: int = 60, *, body_chars: int = 400) -> tuple[str, dict]:
    """合成LLMに渡す材料テキストと出典対応表を返す。

    各材料に短いID [S番号] を振り、source_map[Sn] = {title, date, url} を作る。
    LLM は source_id（Sn）だけを引用し、URLはこちらが復元する。
    戻り値: (材料テキスト, source_map)
    """
    parts: list[str] = []
    source_map: dict[str, dict] = {}
    counter = 0

    def _sid(title: str, date: str | None, url: str) -> str:
        nonlocal counter
        counter += 1
        sid = f"S{counter}"
        source_map[sid] = {"title": title or "", "date": date or None, "url": url or ""}
        return sid

    projects = bundle.get("projects") or []
    if projects:
        parts.append(f"### ヒットの所属プロジェクト: {', '.join(projects)}")

    rows_by_source = bundle.get("rows") or {}
    for source, rows in rows_by_source.items():
        if not isinstance(rows, list):
            continue
        parts.append(f"### {source} ({len(rows)}件中最大{max_rows_per_source}件)")
        for r in rows[:max_rows_per_source]:
            body = (r.get("content") or r.get("subject") or "")[:body_chars]
            who = r.get("name") or r.get("sender") or ""
            title = r.get("subject") or (r.get("content") or "")[:40] or source
            sid = _sid(title, r.get("timestamp"), _best_url(r))
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
                sid = _sid(title, m.get("timestamp"), _best_url(m))
                parts.append(f"- [{sid}] [{m.get('timestamp', '')}] "
                             f"{m.get('name') or m.get('sender', '')}: {body}")

    for dr in deep_reads:
        text = (dr.get("text") or "")[:2000]
        title = dr.get("name") or dr.get("page_id") or "深読み"
        sid = _sid(f"深読み:{title}", None, dr.get("url") or "")
        parts.append(f"### [{sid}] 深読み ({dr.get('kind')}): {title}\n{text}")

    result = "\n".join(parts)
    _dbg(f"_materials_text: {len(result)}文字 出典{len(source_map)}件 (body_chars={body_chars})")
    return result, source_map


def _resolve_sources(data: dict, source_map: dict) -> dict:
    """LLMが返した source_id を、こちらが保持する本物URL（title/date/url）に差し替える。

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
                             "url": info["url"], "quote": s.get("quote")})
        if resolved:
            out_findings.append({"statement": f.get("statement", ""), "sources": resolved})
    if dropped:
        _warn(f"_resolve_sources: 対応表に無い source_id を {dropped} 件破棄")
    return {"findings": out_findings, "not_found": data.get("not_found") or []}


async def synthesize_answer(user_query: str, bundle: dict, deep_reads: list[dict], *,
                            chunk_summaries: list[dict] | None = None,
                            fact_critical: bool = False) -> dict:
    """材料から回答を合成する。出典は source_id 参照→URLはコード側で復元・ファクトは逐語引用。"""
    if chunk_summaries:
        # 月次チャンクにもIDを振り、代表URL（チャンク先頭行）を対応表に載せる
        source_map: dict[str, dict] = {}
        lines: list[str] = []
        for i, c in enumerate(chunk_summaries, 1):
            sid = f"S{i}"
            source_map[sid] = {"title": f"{c['source']}/{c['month']}",
                               "date": c["month"], "url": c.get("url") or ""}
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
