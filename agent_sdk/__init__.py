"""agent_sdk — 社内エージェント基盤 SDK のスタブ。

このリポジトリのスクリプトは、業務で使っている社内 LLM エージェント基盤の
サンドボックス上で動く。サンドボックスは実行時にこのモジュールを注入するため、
本来スクリプト側に実装は無い。

基盤そのものは社内資産なので、このリポジトリには**公開できる形の空実装**だけを置く。
目的は2つ:

  1. 各スクリプトが「どの機能を、どういう引数で、どう組み合わせて使っているか」を
     読み手が追えるようにすること
  2. import を通し、型・シグネチャの意図を残すこと

呼び出すと NotImplementedError になる。動かすには実際の基盤か、同等の
アダプタ（OpenAI/Anthropic SDK + 任意のベクタ検索 + 任意のDB）を差し込む必要がある。
"""

from __future__ import annotations

from typing import Any, Iterable, Sequence

__all__ = [
    "ToolCallError",
    "llm_call",
    "rag_download", "rag_file_list", "rag_search",
    "tavily_search", "tavily_extract", "tavily_map",
    "file_create", "file_output",
    "tool_bridge",
    "list_composio_tools", "list_composio_tool_specs",
    "db",
]


class ToolCallError(Exception):
    """基盤側のツール呼び出しが失敗したときに投げられる例外。

    スクリプトはこれを捕まえてリトライ／フォールバックの分岐に使っている。
    """


def _stub(name: str) -> "NotImplementedError":
    return NotImplementedError(
        f"agent_sdk.{name} はスタブです。実行にはエージェント基盤、"
        f"または同等のアダプタ実装が必要です。"
    )


# ── LLM ───────────────────────────────────────────────────────────────
async def llm_call(
    prompt: str,
    *,
    model: str | None = None,
    schema: dict | None = None,
    files: Sequence[str] | None = None,
    **kwargs: Any,
) -> dict:
    """LLM を1回呼ぶ。非同期。

    model  : "provider/名前" 形式（例 "gemini/gemini-3.7-flash"）。
    schema : JSON Schema を渡すと構造化出力になる。
    files  : 添付するファイルパス。PDF/docx/xlsx 等に対応（pptx は非対応）。
    戻り値 : {"text": ...} もしくは schema 指定時は構造化された dict。
             ※ モデルによって戻り値のキー構造が変わるので、呼び出し側は
                .get("text") で防御している。
    """
    raise _stub("llm_call")


# ── RAG（会話にぶら下がるファイルストア） ───────────────────────────
async def rag_download(files: list[dict]) -> list[dict]:
    """RAG に登録済みのファイルをサンドボックスのローカルに落とす。

    files: [{"index_name": ..., "file_name": ...}, ...] の dict 配列。
    スクリプト本体を RAG に置いて実行時に取り寄せる、という使い方をしている。
    """
    raise _stub("rag_download")


async def rag_file_list(index_name: str | None = None) -> list[dict]:
    """RAG インデックス内のファイル一覧。file_name が不明なときの代替経路。"""
    raise _stub("rag_file_list")


async def rag_search(query: str, **kwargs: Any) -> list[dict]:
    """RAG のセマンティック検索。"""
    raise _stub("rag_search")


# ── Web 検索 ─────────────────────────────────────────────────────────
def tavily_search(query: str, **kwargs: Any) -> dict:
    """Web 検索。債務者区分判定エージェントが外部情報の収集に使う。"""
    raise _stub("tavily_search")


def tavily_extract(urls: Iterable[str], **kwargs: Any) -> dict:
    """URL の本文抽出。検索結果のスニペットでは足りないときの全文取得に使う。"""
    raise _stub("tavily_extract")


def tavily_map(url: str, **kwargs: Any) -> dict:
    """サイト構造のクロール。"""
    raise _stub("tavily_map")


# ── ファイル ─────────────────────────────────────────────────────────
def file_create(path: str, content: bytes | str, **kwargs: Any) -> dict:
    """生成物を会話に添付する形で書き出す。"""
    raise _stub("file_create")


def file_output(path: str, **kwargs: Any) -> dict:
    """ローカルに作ったファイルを成果物として返す。"""
    raise _stub("file_output")


# ── 外部SaaS（Composio 経由） ────────────────────────────────────────
def tool_bridge(tool_slug: str, params: dict, **kwargs: Any) -> dict:
    """Composio 連携ツールを叩く汎用ブリッジ（Gmail / Drive / Zoom / Notion 等）。"""
    raise _stub("tool_bridge")


def list_composio_tools(**kwargs: Any) -> list[str]:
    """利用可能な連携ツールのスラッグ一覧。"""
    raise _stub("list_composio_tools")


def list_composio_tool_specs(slugs: Sequence[str], **kwargs: Any) -> list[dict]:
    """指定スラッグの入力スキーマ。引数名が実行環境ごとに違うので実行前に確認する用途。"""
    raise _stub("list_composio_tool_specs")


# ── DataTable ────────────────────────────────────────────────────────
class _Table:
    """1テーブルへのハンドル。find / insert / update を持つ。"""

    def __init__(self, name: str) -> None:
        self.name = name

    def find(self, filter: dict | None = None, **kwargs: Any) -> list[dict]:
        """完全一致 / $gt,$gte,$lt,$lte,$ne / $in / トップレベル複数キー=AND のみ。
        contains（部分一致）と $or は無い。sort を付けるとページングできない。
        """
        raise _stub("db.table().find")

    def insert(self, rows: list[dict], **kwargs: Any) -> dict:
        raise _stub("db.table().insert")

    def update(self, filter: dict, values: dict, **kwargs: Any) -> dict:
        raise _stub("db.table().update")

    def delete(self, filter: dict, **kwargs: Any) -> dict:
        raise _stub("db.table().delete")


class _DB:
    def table(self, name: str) -> _Table:
        """会話にアタッチされた Project 配下のテーブルを引く。未アタッチだと 422。"""
        return _Table(name)

    def search(self, query: str, **kwargs: Any) -> list[dict]:
        """DataTable 横断のセマンティック検索。"""
        raise _stub("db.search")


db = _DB()
