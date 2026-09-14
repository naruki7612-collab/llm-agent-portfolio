# -*- coding: utf-8 -*-
"""
ガンマ管理株式会社 リフォーム請求書 抽出＋Excel突合スクリプト（実行環境想定）

## RAG登録
このファイルはRAGインデックス「請求書突合エージェント」に格納する。
エージェント側は実行のたびに
`await agent_sdk.rag_download(files=[{"index_name": "請求書突合エージェント", "file_name": "請求書抽出突合スクリプト.py"}])`
で取得してから読み込むこと（コード自体は会話・システムプロンプトには埋め込まない）。
`files`はdictの配列で渡す点に注意（`index=`/`file_name=`を別々の引数にする書き方は誤り。
実測で確認済み・詳細はメモリ[[rag-script-discovery-by-search]]）。
詳細は同ディレクトリの `システムプロンプト_請求書突合エージェント.md` を参照。

## 全体の流れ
  フェーズ1: 請求書PDF → 構造化データ抽出（Vision, Gemini 3.7 Flash, schema指定）
  フェーズ2: 抽出データ ↔ ２つのExcelそれぞれと突合（テキスト判定, Sonnet 5, schema指定）

## 設計方針（フェーズ1: 抽出）
1. まず「リフォーム一覧（本社管理部）.xlsx」と「修繕売上一覧(明細).xls」の
   ２つのExcelに共通するカラムを抽出する（COMMON_SCHEMA）。
2. 次に、それぞれのExcelにしか無い残りのカラムについて、できる範囲で追加抽出する
   （RIFORM_ONLY_SCHEMA / SHUZEN_ONLY_SCHEMA）。
3. 構造的に請求書からは絶対に取得できないと事前検証済みの列（請求金額(売価)・
   請求先名・掛け率・家主/契約者情報・社内進捗管理項目など）はLLMに問い合わせず、
   最初から固定でNoneを出力する（ハルシネーション防止・コスト削減）。
   これは三菱電機システムサービス／サンプル設備／長谷川畳店の3社サンプルを
   4分割画像まで精査した実地検証に基づく。

## 設計方針（フェーズ2: 突合）
- ２つのExcelは同じ運用ルール上「両方に登録しなければならない」ものだが、運用が
  浅く二重管理になっており、**共通の主キーが無い**（工事Noのような社内番号は
  基幹システム側にしか無く、店舗側スプシとは紐づいていない）。そのため必ず
  別々に候補検索→LLM判定を行う（片方にマッチしてももう片方は分からない）。
- 文字列の完全一致では絶対に判定しない。業者名・物件名の表記ゆれが激しく
  （例:「ミドリガス」⇔「ミドリガス株式会社」、「ソレイユ」⇔「SOLEIL」、
  「SOS佐藤」⇔「2SOS60」）、かつ請求書の金額は税抜・Excelの金額列は税込で
  記録されているため単純な数値一致にもならない（実データ検証: ×1.10で一致）。
  そのためコード側では「候補を絞り込む」ことだけを行い、「どれが正解か」
  「どこが違うか」の最終判断は必ずLLMに自然文で判定させる。
- 候補絞り込みは「業者名の包含マッチ＋日付ウィンドウ」のみ。業者名の類似度
  （difflib等）はしきい値による足切りには使わない。実データ検証で「ミドリガス」
  対「ヒカリガス」のような別業者でも類似度が高く出る一方、「サンプル設備」対
  「SOS佐藤」のような同一業者でも低く出ることを確認しており、閾値では安全に
  切り分けられないため。業者名で1件もヒットしない場合は日付の近さのみで
  フォールバックし、類似度は同着時の参考順位付けにのみ使う。

## 前提（重要・実地検証済み。agent_platform-agent-amazon-bedrock-agentcore-main のソースで裏取り済み）
- 対象PDFはテキスト層がほぼ無く、本文（物件名・金額等）は画像として埋め込まれている。
  → 通常のテキスト抽出・正規表現では中身を読めないため、Vision対応LLMでの
    読み取りが必須。
- ツールは `agent_sdk` パッケージから **import して使う**
  （`from agent_sdk import llm_call, rag_download` → `await llm_call(...)` /
  `await rag_download(...)`）。会話やプロンプトに関数がそのまま生えているわけではない。
  `import agent_sdk as at` として `at.llm_call(...)` と呼ぶ書き方も文法上は可能だが、
  `import agent_sdk as llm_call` のように**関数名をそのままモジュールの別名にする書き方は
  誤り**（`agent_sdk`モジュール全体がその名前になるだけで、関数を直接指さない）。
  また `llm_call` / `rag_download` は **非同期関数**（`async def`）なので
  必ず `await` すること。code_executeのサンドボックスはトップレベルawaitに対応している。
- `llm_call` の添付ファイル引数は `file_paths=[...]`（ソースで確認済み。
  `agent_sdk/llm.py`）。
- モデルIDは `"<provider>/<モデル名>"` の形式で指定する
  （`src/domain/model_pricing.py` の `PTC_SUPPORTED_MODELS` で確認済み）:
    - Gemini 3.7 Flash: `"gemini/gemini-3.7-flash"`
    - Sonnet 5:          `"anthropic/claude-sonnet-5"`
  PDFをVisionで読ませる場合は Gemini系モデルを明示指定すること
  （既知の制約。Claude系＋file_pathsの組み合わせは失敗する実例あり）。
- `llm_call` は `schema` にJSON Schema dictを渡すと構造化出力モードになり、
  戻り値が `{"data": {...}, "usage": {...}, "model": "..."}` になる
  （schema省略時は `{"text": "...", ...}`。エラー時は `{"error": "...", "model": "..."}`）。
  本スクリプトは抽出・突合どちらも `schema` を指定し、`response["data"]` を
  そのまま使う。**ただし `response["error"]` の可能性を必ず先にチェックしてから
  `data`/`text` を取り出すこと**（決め打ちで `response["text"]` のようにアクセスすると、
  エラー応答時にKeyErrorで落ちる。詳細は[[llm-call-claude-response-keyerror]]）。
- 実行環境ツールは未知の引数を黙って無視する既知の癖があるため、本番投入前に
  `inspect.signature(agent_sdk.llm_call)` 等で実際の引数名を確認しておくと安全。
- code_executeの実行窓は約300秒。件数が増えて超えるようならバッチ分割の
  仕組みを別途足すこと（現状はデモ規模の数枚〜十数枚を想定し、シンプルに
  全件を1回の実行内で並列処理する構成にしている）。
- 並列実行は `asyncio.gather` ＋ `asyncio.Semaphore` で行う（ThreadPoolExecutorは
  非同期関数の並列化には使わない）。実行環境の並列実行スロット上限は16（実測）なので
  それ以下に収める。

## このファイルの使い方
1. `call_llm_extract()`（抽出用・Gemini 3.7 Flash）と
   `call_llm_match()`（突合用・Sonnet 5）内の `from agent_sdk import llm_call` 以降を、
   実際の実行環境環境で動作確認しながら必要なら微調整する。
2. 実行前に、請求書PDF一式と2つのExcel（リフォーム一覧・修繕売上一覧）を
   **チャットのルームに添付**しておく（`UPLOADS_DIR` = "uploads/" に自動配置される。
   Excelはファイル名の完全一致を要求せず、`RIFORM_XLSX_KEYWORD`("リフォーム")/
   `SHUZEN_XLS_KEYWORD`("修繕")を含むファイルを自動で探す）。
3. 実行環境のcode_execute上で `await main()`（またはスクリプトとして直接実行するなら
   `asyncio.run(main())`）を実行する。
4. 最終結果は `MATCH_RESULTS_JSON_PATH` / `MATCH_RESULTS_CSV_PATH` に出力される
   （突合対象＝請求書の明細行1件ごとに、リフォーム一覧側・修繕売上一覧側
   それぞれの判定結果を持つ）。
   途中経過として抽出結果だけを見たい場合は `EXTRACTION_RESULTS_JSON_PATH` /
   `EXTRACTION_RESULTS_CSV_PATH` も参照できる。
"""

from __future__ import annotations

import asyncio
import csv
import json
import re
import sys
import traceback
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Optional

import openpyxl
import xlrd

# ==============================================================
# 0. パス設定
# ==============================================================

BASE_DIR = Path(__file__).resolve().parent

# ルーム（チャット）に添付したファイルは、code_executeのサンドボックス上では
# 固定パス "workspace/uploads/<ファイル名>" に自動配置される
# （SandboxLayout.UPLOADS = "workspace/uploads"。ソースで確認済み）。
# スクリプト自体はRAG経由（"rag/<ファイル名>"）で取得するが、
# 頻繁に更新される業務Excelと毎回変わる請求書PDFはRAGには入れず、
# 実行のたびにチャットに添付してもらう運用にする。
UPLOADS_DIR = BASE_DIR / "uploads"
INVOICE_DIR = UPLOADS_DIR                                        # 添付した請求書PDFが並ぶ場所

# Excel2つはファイル名の完全一致を前提にしない（デモ軽量版のように別名で
# 添付されることがあるため。実際に実行環境実行で "FileNotFoundError" になった実例あり）。
# 「リフォーム」「修繕」というキーワードを含むファイルをuploads/内から探す。
RIFORM_XLSX_KEYWORD = "リフォーム"
SHUZEN_XLS_KEYWORD = "修繕"


def find_uploaded_file(keyword: str, extensions: tuple[str, ...]) -> Path:
    """uploads/ 内から、ファイル名にkeywordを含み拡張子がextensionsに合う
    ファイルを探す。完全一致のファイル名を前提にしないための柔軟な検索。"""
    candidates = [
        p for p in UPLOADS_DIR.glob("*")
        if keyword in p.name and p.suffix.lower() in extensions
    ]
    if not candidates:
        raise FileNotFoundError(
            f"uploads/ 内に「{keyword}」を含む{extensions}ファイルが見つかりません。"
            f"ルームに添付されているか確認してください。"
        )
    if len(candidates) > 1:
        print(
            f"[警告] 「{keyword}」を含むファイルが複数見つかりました: "
            f"{[p.name for p in candidates]}。先頭の {candidates[0].name} を使います。",
            file=sys.stderr,
        )
    return candidates[0]

EXTRACTION_RESULTS_JSON_PATH = BASE_DIR / "抽出結果_raw.json"
EXTRACTION_RESULTS_CSV_PATH = BASE_DIR / "抽出結果.csv"
MATCH_RESULTS_JSON_PATH = BASE_DIR / "突合結果_raw.json"
MATCH_RESULTS_CSV_PATH = BASE_DIR / "突合結果.csv"

# モデルID（provider/モデル名の形式。PTC_SUPPORTED_MODELS で確認済み）
EXTRACT_MODEL = "gemini/gemini-3.7-flash"   # PDFのVision読み取りにはGemini系が必須
MATCH_MODEL = "anthropic/claude-sonnet-5"   # テキスト比較のみなのでSonnet 5

# 並列実行数。実行環境の並列実行スロット上限は16（実測）なのでそれ以下に収めること。
MAX_WORKERS = 4

# 候補を絞り込む日付ウィンドウ（前後何日まで候補に含めるか）。
# 受付日/完了日/請求書発行日のどれで登録されているか案件ごとにバラつきがあるため広め。
DATE_WINDOW_DAYS = 60

# 業者名フォールバック時の候補上限（日付が近い順に、この件数までLLMに渡す）。
VENDOR_FALLBACK_LIMIT = 30

# ローカルでパイプラインだけを動作確認したい場合に使う（agent_sdkが無い環境向け）
MOCK_MODE = False


# ==============================================================
# フェーズ1: 抽出対象カラム定義
# ==============================================================

# ---- ①２つのExcelに共通するカラム ----
# 値は (LLMへの説明, 抽出可否の実地検証結果) のタプル。
#   "○" = ほぼ確実に取得可能 / "△" = 業者フォーマット次第 / "✕(稀にあり)" = 基本無い
COMMON_SCHEMA: dict[str, tuple[str, str]] = {
    "vendor_name": (
        "請求書の発行元（業者）の正式名称、または屋号。"
        "債権譲渡がある場合は実際に工事をした側（譲渡人）の名前を優先する。",
        "○",
    ),
    "vendor_amount": (
        "業者への支払額（原価）。請求書の「合計」「ご請求金額」「請求額」に相当する数値。"
        "消費税込みの最終請求金額を数値のみ（カンマ・円記号なし）で。",
        "○",
    ),
    "payment_method": (
        "支払方法。銀行振込の場合は「振込（銀行名 支店名）」の形式で。",
        "○",
    ),
    "completion_date": (
        "工事・作業の完了日、または実施日とみなせる日付（YYYY-MM-DD）。"
        "請求書によっては「受付日」「完了日」の区別が無く単に日付だけの場合があるので、"
        "その場合はその日付をそのまま入れる。",
        "△",
    ),
    "work_type": (
        "工事の種別・カテゴリ（例：原状回復工事、給湯器修理、エアコン工事、畳張替 等）。"
        "明記が無い場合は機種名・作業内容から妥当な種別を推定してよいが、"
        "推定した場合は末尾に「(推定)」と付けること。",
        "△",
    ),
    "billing_content": (
        "請求内容・作業内容の明細。複数の作業がある場合は「;」区切りで列挙する。",
        "△",
    ),
    "property_name": (
        "物件名（マンション名・アパート名等）。分かる場合のみ。"
        "1枚の請求書に複数物件が含まれる場合は line_items 側に記載し、"
        "ここには代表値または空欄でよい。",
        "△",
    ),
    "room_number": (
        "号室。分かる場合のみ。複数物件の場合は line_items 側を参照。",
        "△",
    ),
    "property_no": (
        "物件を特定する社内番号らしき表記（例:「No.90673」「№01418」等）。"
        "無ければ空文字列。",
        "✕(稀にあり)",
    ),
    "reception_date": (
        "工事の受付日・依頼受付日（完了日とは別に明記がある場合のみ、YYYY-MM-DD）。"
        "無ければ空文字列。",
        "✕(稀にあり)",
    ),
    "staff_name": (
        "実際の工事担当者名、または客先の担当者名として明記されている人名。"
        "請求書の回付印・承認印にあるだけの社内担当者名は含めない。無ければ空文字列。",
        "✕(稀にあり)",
    ),
    "expected_payment_date": (
        "支払期限・お支払期日（YYYY-MM-DD）。入金予定日の代用値として使う。無ければ空文字列。",
        "△",
    ),
}

# ---- ②リフォーム一覧（本社管理部）.xlsx にのみ存在するカラムのうち、抽出を試みるもの ----
RIFORM_ONLY_SCHEMA: dict[str, tuple[str, str]] = {
    "invoice_issue_date": (
        "請求書発行日（「ご請求日」「発行日」「作成日」等の日付、YYYY-MM-DD）。",
        "○",
    ),
}

# ---- ③修繕売上一覧(明細).xls にのみ存在するカラムのうち、抽出を試みるもの ----
SHUZEN_ONLY_SCHEMA: dict[str, tuple[str, str]] = {
    "location": (
        "物件の所在地（町名レベルでよい。例:「四日市市茂福」）。分かる場合のみ、無ければ空文字列。",
        "△",
    ),
}

# ---- 1請求書に複数の工事・物件が束ねられている場合の明細行スキーマ ----
LINE_ITEM_SCHEMA: dict[str, str] = {
    "property_name": "その明細の物件名",
    "room_number": "その明細の号室",
    "property_no": "その明細の物件№（あれば。無ければ空文字列）",
    "work_date": "その明細の作業日・完了日（YYYY-MM-DD）",
    "work_description": "その明細の作業内容",
    "line_amount": "その明細の金額（数値のみ）",
}

# 数値として扱うフィールド（JSON Schemaの型付けで使う）
_NUMERIC_HEADER_FIELDS = {"vendor_amount"}
_NUMERIC_LINE_ITEM_FIELDS = {"line_amount"}

# ---- 構造的に請求書からは絶対に取得できないと確認済みの列 ----
# LLMには問い合わせず、常にこの値（None）を出力する。
OUTPUT_FIXED_NA: dict[str, Optional[str]] = {
    # --- ２つのExcel共通 ---
    "billing_amount_to_customer": None,  # 請求金額(売価) / 売上合計
    "billee_name": None,                 # 請求先名（家主・入居者・契約者名）
    "markup_rate": None,                 # 掛け率
    "branch": None,                      # 支店・取扱店舗（回付印のみで一意特定不可）
    "payment_date": None,                # 入金日(経理入力)
    # --- リフォーム一覧のみ ---
    "estimate_date": None,               # 見積書作成日
    "order_date": None,                  # 業者発注日
    "kintone_input": None,               # Kintone入力
    "isp_confirm": None,                 # i-SP確認
    "daily_report_date": None,           # 日報計上日
    "billee_type": None,                 # 請求先(区分)
    "dw_estimate_check": None,           # DW内見積書確認（経理部確認）
    # --- 修繕売上一覧のみ ---
    "construction_no": None,             # 工事No（社内番号。業者側の番号とは別物）
    "building_furigana": None,           # フリガナ
    "landlord_1": None, "landlord_2": None, "landlord_3": None, "lessor": None,
    "tenant_name": None, "occupant_name": None,
    "contract_status": None, "contract_date": None, "original_contract_start": None,
    "contract_start": None, "contract_end": None, "cancel_date": None,
    "next_move_in_date": None, "witness_date": None, "witness_time": None,
    "inquiry_registered": None, "inquiry_no": None, "inquiry_staff": None,
    "contact_person": None, "category": None,
    "category_detail_1": None, "category_detail_2": None,
    "next_action_date": None, "status": None,
    "profit": None, "profit_rate": None,
    "tenant_billing_amount": None, "landlord_billing_amount": None, "total_sales": None,
    "rent_item_tenant": None, "billing_month_tenant": None, "deposit_method_tenant": None,
    "deposit_date_tenant": None, "billing_amount_tenant": None, "deposit_amount_tenant": None,
    "rent_item_landlord": None, "billing_month_landlord": None, "deposit_method_landlord": None,
    "deposit_date_landlord": None, "billing_amount_landlord": None, "deposit_amount_landlord": None,
    "vendor_disbursement_date": None,    # 出金日（業者） — 支払期限はあっても実行日は無い
}


# ==============================================================
# フェーズ1: 抽出プロンプト・JSON Schema構築
# ==============================================================

def _json_prop(desc: str, numeric: bool = False) -> dict[str, Any]:
    # 注意: ここは標準JSON SchemaではなくGeminiのSchema形式（実測で確認済み）。
    # "type": ["string", "null"] のような配列型はバリデーションエラーになる
    # （実行環境実行ログで実際に確認: 20 validation errors for Schema）。
    # 型は単一の文字列で指定し、null許容は書かない（未取得時はモデルが空文字/省略で
    # 返してくる想定。normalize_extraction_result側で欠損はどのみち吸収する）。
    return {"type": "number" if numeric else "string", "description": desc}


def build_extraction_prompt() -> str:
    """LLMに渡す抽出指示文を組み立てる（出力形式そのものはschemaが強制するので、
    ここでは抽出ルール・注意事項の説明に専念する）。"""
    all_fields = {**COMMON_SCHEMA, **RIFORM_ONLY_SCHEMA, **SHUZEN_ONLY_SCHEMA}
    field_lines = "\n".join(
        f'  - "{key}": {desc}' for key, (desc, _grade) in all_fields.items()
    )
    line_item_lines = "\n".join(
        f'    - "{key}": {desc}' for key, desc in LINE_ITEM_SCHEMA.items()
    )

    return f"""あなたは不動産管理会社の経理担当者向けに、業者からの請求書PDFの内容を
構造化データとして抽出するアシスタントです。

添付された請求書PDFを読み取り、指定されたスキーマに従って情報を抽出してください。

## 各キーの意味（header）

{field_lines}

## 明細行（line_items）について
1枚の請求書に複数の工事・複数の物件がまとめて記載されている場合（月締め請求書等）は、
line_items に各明細を配列で入れてください。1件しか工事が無い場合も、その1件を
line_items に1要素として入れてください。

{line_item_lines}

## 注意事項
- 分からない文字列の項目は必ず空文字列 "" にすること。憶測で埋めない
  （ただし work_type のみ、機種名や作業内容から妥当に推定できる場合は
  「(推定)」を付けて埋めてよい）。
- 分からない金額の項目（vendor_amount / line_amount）は 0 にすること。
- 金額は数値で出力すること（カンマ・円記号は除く）。
- 日付は "YYYY-MM-DD" 形式に正規化すること（西暦・和暦どちらの記載でも変換する）。
- 請求書に押されている「支店回付印」（本社・桑名・四日市...などの承認スタンプ）は
  社内の承認記録であり、staff_name や branch には使わないこと。
"""


def build_extraction_schema() -> dict[str, Any]:
    """llm_call の schema= に渡すJSON Schema（構造化出力）を組み立てる。"""
    all_fields = {**COMMON_SCHEMA, **RIFORM_ONLY_SCHEMA, **SHUZEN_ONLY_SCHEMA}
    header_properties = {
        key: _json_prop(desc, numeric=key in _NUMERIC_HEADER_FIELDS)
        for key, (desc, _grade) in all_fields.items()
    }
    line_item_properties = {
        key: _json_prop(desc, numeric=key in _NUMERIC_LINE_ITEM_FIELDS)
        for key, desc in LINE_ITEM_SCHEMA.items()
    }
    return {
        "type": "object",
        "properties": {
            "header": {"type": "object", "properties": header_properties},
            "line_items": {
                "type": "array",
                "items": {"type": "object", "properties": line_item_properties},
            },
        },
        "required": ["header", "line_items"],
    }


EXTRACTION_JSON_SCHEMA = build_extraction_schema()


# ==============================================================
# 共通: llm_call レスポンスの安全な取り出し
# ==============================================================

def unwrap_llm_response(response: Any) -> dict[str, Any]:
    """
    llm_call の戻り値から本体を取り出す。
    構造化出力(schema指定)時は {"data": {...}}、テキストのみなら {"text": "..."}、
    失敗時は {"error": "..."} が返る。"error"を最優先でチェックし、"text"決め打ち
    アクセスによるKeyErrorを避ける（[[llm-call-claude-response-keyerror]] 参照）。
    """
    if not isinstance(response, dict):
        raise ValueError(f"想定外のレスポンス型です: {type(response)!r}")
    if response.get("error"):
        raise RuntimeError(f"llm_callがエラーを返しました: {response['error']}")
    if "data" in response:
        return response["data"]
    if "text" in response:
        return parse_llm_json(response["text"])
    raise ValueError(f"想定外のレスポンス構造です。実際のキー: {list(response.keys())}")


def parse_llm_json(raw_text: str) -> dict[str, Any]:
    """LLM応答（テキストモード）からJSON部分を安全に取り出す（コードフェンス混入に対応）。
    schemaモードでは基本的に使わないが、テキストモードへのフォールバック用に残す。"""
    text = raw_text.strip()
    if text.startswith("```"):
        text = text.strip("`")
        if text.lower().startswith("json"):
            text = text[4:]
        text = text.strip()
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1:
        raise ValueError(f"JSONオブジェクトが見つかりません: {raw_text[:200]!r}")
    return json.loads(text[start : end + 1])


# ==============================================================
# フェーズ1: 抽出LLM呼び出し
# ==============================================================

async def call_llm_extract(pdf_path: Path, prompt: str) -> dict[str, Any]:
    """agent_sdk.llm_call を呼び出し、請求書PDFをVisionで読み取らせる。"""
    from agent_sdk import llm_call  # ローカルMOCK_MODE検証時にimport不要にするための遅延import

    response = await llm_call(
        prompt=prompt,
        model=EXTRACT_MODEL,          # PDF読解にはGemini系が必須
        file_paths=[str(pdf_path)],
        schema=EXTRACTION_JSON_SCHEMA,
    )
    return unwrap_llm_response(response)


async def call_llm_extract_mock(pdf_path: Path, prompt: str) -> dict[str, Any]:
    """MOCK_MODE=True のときに使うダミー応答（パイプライン検証専用）。"""
    return {
        "header": {
            "vendor_name": "（モック）テスト業者",
            "vendor_amount": 12345,
            "payment_method": "振込（テスト銀行 テスト支店）",
            "completion_date": "2026-07-01",
            "work_type": "テスト工事",
            "billing_content": "テスト作業",
            "property_name": None,
            "room_number": None,
            "property_no": None,
            "reception_date": None,
            "staff_name": None,
            "expected_payment_date": None,
            "invoice_issue_date": "2026-07-31",
            "location": None,
        },
        "line_items": [
            {
                "property_name": None,
                "room_number": None,
                "property_no": None,
                "work_date": "2026-07-01",
                "work_description": "テスト作業",
                "line_amount": 12345,
            }
        ],
    }


# ==============================================================
# フェーズ1: レスポンス正規化
# ==============================================================

def _unknown_to_none(value: Any) -> Any:
    """Geminiスキーマはnullを表現できないため、モデルは未取得値を空文字列/0で
    返してくる（プロンプト指示）。以降のロジックでは他の列と同じくNoneで
    「未取得」を統一表現するため、ここで変換する。"""
    if value == "" or value == 0:
        return None
    return value


def normalize_extraction_result(pdf_path: Path, parsed: dict[str, Any]) -> dict[str, Any]:
    """LLM出力を固定NA列とマージし、Excel突合用の1レコードに正規化する。"""
    header = parsed.get("header", {}) or {}
    line_items = parsed.get("line_items", []) or []
    for item in line_items:
        for k, v in list(item.items()):
            item[k] = _unknown_to_none(v)

    record: dict[str, Any] = {"source_file": pdf_path.name}

    all_llm_fields = {**COMMON_SCHEMA, **RIFORM_ONLY_SCHEMA, **SHUZEN_ONLY_SCHEMA}
    for key in all_llm_fields:
        record[key] = _unknown_to_none(header.get(key))

    record.update(OUTPUT_FIXED_NA)  # 取得不可能な列は常にNoneで固定
    record["line_items_json"] = json.dumps(line_items, ensure_ascii=False)
    record["line_item_count"] = len(line_items)
    return record


# ==============================================================
# フェーズ1: 抽出バッチ処理（asyncio.gather + Semaphoreで並列実行）
# ==============================================================

async def extract_one_invoice(pdf_path: Path) -> dict[str, Any]:
    prompt = build_extraction_prompt()
    parsed = await (
        call_llm_extract_mock(pdf_path, prompt) if MOCK_MODE else call_llm_extract(pdf_path, prompt)
    )
    return normalize_extraction_result(pdf_path, parsed)


async def extract_one_invoice_safe(pdf_path: Path, semaphore: asyncio.Semaphore) -> dict[str, Any]:
    async with semaphore:
        try:
            record = await extract_one_invoice(pdf_path)
            print(f"[抽出] OK: {pdf_path.name} -> 業者={record.get('vendor_name')} "
                  f"金額={record.get('vendor_amount')}")
            return record
        except Exception as e:  # noqa: BLE001
            print(f"[抽出] NG: {pdf_path.name} -> {e}", file=sys.stderr)
            traceback.print_exc()
            return {"source_file": pdf_path.name, "error": str(e)}


async def run_extraction_phase() -> list[dict[str, Any]]:
    if not INVOICE_DIR.exists():
        print(f"請求書フォルダが見つかりません: {INVOICE_DIR}", file=sys.stderr)
        sys.exit(1)

    pdf_files = sorted(INVOICE_DIR.glob("*.pdf"))
    if not pdf_files:
        print(f"PDFが見つかりません: {INVOICE_DIR}", file=sys.stderr)
        sys.exit(1)

    print(f"[抽出] 対象PDF: {len(pdf_files)}件（並列数: {MAX_WORKERS}）")

    semaphore = asyncio.Semaphore(MAX_WORKERS)
    records = await asyncio.gather(*(extract_one_invoice_safe(p, semaphore) for p in pdf_files))
    records = sorted(records, key=lambda r: r.get("source_file", ""))

    with open(EXTRACTION_RESULTS_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    if records:
        with open(EXTRACTION_RESULTS_CSV_PATH, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=list(records[0].keys()))
            writer.writeheader()
            writer.writerows(records)

    print(f"[抽出] 保存しました: {EXTRACTION_RESULTS_JSON_PATH.name} / {EXTRACTION_RESULTS_CSV_PATH.name}")
    print("\n[抽出] ===== 生JSON結果 =====")
    print(json.dumps(records, ensure_ascii=False, indent=2, default=str))
    return records


# ==============================================================
# フェーズ2: 文字列正規化・ゆるいマッチ
# ==============================================================

_VENDOR_NOISE_PATTERN = re.compile(
    r"(株式会社|有限会社|合同会社|\(株\)|（株）|\(有\)|（有）|㈱|㈲|"
    r"支店|営業所|様|御中|・|\s|　)"
)


def normalize_vendor(name: Optional[str]) -> str:
    """業者名の表記ゆれを吸収するための軽い正規化（大文字小文字・法人格・空白除去）。"""
    if not name:
        return ""
    s = str(name)
    s = re.sub(r"^\d+", "", s)  # 修繕売上一覧の「2サンプル設備」のような先頭コード除去
    s = _VENDOR_NOISE_PATTERN.sub("", s)
    return s.strip().lower()


def normalize_property(name: Optional[str]) -> str:
    """物件名の表記ゆれ（中黒・スペース・全角半角）を軽く吸収する。"""
    if not name:
        return ""
    s = str(name)
    s = s.replace("・", "").replace(" ", "").replace("　", "")
    return s.strip().lower()


def fuzzy_ratio(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


def parse_date_loose(value: Any) -> Optional[datetime]:
    """Excel由来のdatetime、または 'YYYY-MM-DD' 文字列を datetime に正規化する。"""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(str(value)[:10], fmt)
        except ValueError:
            continue
    return None


def within_window(d1: Optional[datetime], d2: Optional[datetime], days: int) -> bool:
    if d1 is None or d2 is None:
        return True
    return abs((d1 - d2).days) <= days


# ==============================================================
# フェーズ2: Excel読み込み（候補プールの構築）
# ==============================================================

def load_riform_rows() -> list[dict[str, Any]]:
    """リフォーム一覧（本社管理部）.xlsx を全月シート横断でフラットな行リストにする。"""
    path = find_uploaded_file(RIFORM_XLSX_KEYWORD, (".xlsx",))
    wb = openpyxl.load_workbook(path, data_only=True)
    rows: list[dict[str, Any]] = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        header_row_idx = None
        headers: list[Any] = []
        for r in range(1, 6):
            vals = [c.value for c in ws[r]]
            if "業者名" in vals:
                header_row_idx = r
                headers = vals
                break
        if header_row_idx is None:
            continue  # 「相殺不可物件」等、対象外シート

        for excel_row_idx, row in enumerate(
            ws.iter_rows(min_row=header_row_idx + 1, values_only=True),
            start=header_row_idx + 1,
        ):
            d = dict(zip(headers, row))
            if not d.get("業者名"):
                continue
            rows.append(
                {
                    "row_id": f"リフォーム一覧:{sheet_name}:{excel_row_idx}",
                    "sheet_name": sheet_name,
                    "excel_row": excel_row_idx,
                    "vendor_name": d.get("業者名"),
                    "property_name": d.get("物件名"),
                    "room_number": d.get("号室"),
                    "property_no": d.get("物件№"),
                    "work_type": d.get("工事種別"),
                    "billing_content": d.get("請求内容"),
                    "vendor_amount": d.get("業者金額"),
                    "billing_amount_to_customer": d.get("請求金額"),
                    "completion_date": d.get("完了日"),
                    "reception_date": d.get("受付\n日") or d.get("受付日"),
                    "invoice_issue_date": d.get("請求書\n発行日"),
                    "payment_method": d.get("請求方法"),
                    "_vendor_norm": normalize_vendor(d.get("業者名")),
                    "_property_norm": normalize_property(d.get("物件名")),
                }
            )
    return rows


def load_shuzen_rows() -> list[dict[str, Any]]:
    """修繕売上一覧(明細).xls をフラットな行リストにする。"""
    path = find_uploaded_file(SHUZEN_XLS_KEYWORD, (".xls", ".xlsx"))
    wb = xlrd.open_workbook(str(path))
    sh = wb.sheet_by_index(0)
    headers = sh.row_values(4)

    def col(name: str) -> int:
        return headers.index(name)

    profit_col_idx = None
    for name in headers:
        if isinstance(name, str) and name.startswith("利益") and "率" not in name:
            profit_col_idx = headers.index(name)
            break

    rows: list[dict[str, Any]] = []
    for r in range(6, sh.nrows):
        row = sh.row_values(r)
        vendor = row[col("取引業者")] if col("取引業者") < len(row) else None
        if not vendor:
            continue

        billing_amount_to_customer = row[col("売上合計\n(契約者請求＋家主請求)")]
        # 「売上合計」はお客様への請求額（売価）であり、業者への支払額（原価）とは別物。
        # このシートには原価そのものの列が信頼できる形では入っていない
        # （「業者支払分(原価）」列は多くの行で「修繕費」等の文字列プレースホルダに
        # なっていて数値が入っていないことを実データで確認済み）。
        # 原価は「売上合計－利益」で逆算できる（実データで複数件検証済み：
        # 例 売上合計97,240－利益22,440＝74,800＝業者金額と一致）ため、ここで計算する。
        profit = row[profit_col_idx] if profit_col_idx is not None and profit_col_idx < len(row) else None
        vendor_amount = None
        if isinstance(billing_amount_to_customer, (int, float)) and isinstance(profit, (int, float)):
            vendor_amount = billing_amount_to_customer - profit

        rows.append(
            {
                "row_id": f"修繕売上一覧:row{r}",
                "excel_row": r,
                "construction_no": row[col("工事No")],
                "vendor_name": vendor,
                "property_name": row[col("建物名称")],
                "room_number": row[col("場所")],
                "work_type": row[col("工事内容")],
                "billing_content": row[col("工事名称")],
                "vendor_amount": vendor_amount,  # 原価相当（=売上合計-利益）。請求書の金額と比較する対象
                "billing_amount_to_customer": billing_amount_to_customer,  # 売価。金額比較には使わないこと
                "completion_date": row[col("工事完了日")],
                "location": None,
                "_vendor_norm": normalize_vendor(vendor),
                "_property_norm": normalize_property(row[col("建物名称")]),
            }
        )
    return rows


# ==============================================================
# フェーズ2: 候補絞り込み（コード側。判定はしない、絞るだけ）
# ==============================================================

def get_candidates(
    invoice_item: dict[str, Any],
    pool: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """業者名の包含マッチ＋日付ウィンドウで候補行を絞り込む。

    注意: 業者名の類似度（difflib）を絞り込みの「合否判定」には使わない。
    実データで検証した結果、「ミドリガス」対「ヒカリガス」のように業種名が
    共通するだけの**別業者**でも類似度が高く出てしまい、逆に「サンプル設備」対
    「SOS佐藤」のように**同じ業者**でも類似度が低く出ることがあり、閾値では
    安全に切り分けられないことを確認した（両者を区別する判断は、まさに
    ユーザー側の要望通りLLMに委ねるべき領域）。
    そのためコード側は
      1) 包含マッチ（法人格・支店名などの表記ゆれ吸収後に部分一致するか）
      2) 日付ウィンドウ
    という機械的・安全な絞り込みだけを行う。包含マッチで1件もヒットしない
    場合（社内コード化されていて正規化しても一致しないケース）は、日付が
    近い順に候補を作ってLLMに渡す（並べ替えはヒントであり、除外はしていない）。
    """
    vendor_norm = normalize_vendor(invoice_item.get("vendor_name"))
    inv_date = parse_date_loose(invoice_item.get("completion_date"))

    def date_ok_lenient(row: dict[str, Any]) -> bool:
        # 日付が片方でも無ければ「除外しない」（強い業者マッチには日付一致を必須にしない）
        row_date = parse_date_loose(row.get("completion_date"))
        if inv_date is None or row_date is None:
            return True
        return within_window(inv_date, row_date, DATE_WINDOW_DAYS)

    def date_ok_strict(row: dict[str, Any]) -> bool:
        # 業者名で絞れない弱いフォールバックでは、日付が無い行まで含めると
        # 候補が際限なく広がるため、日付が両方揃っていて近いことを必須にする
        row_date = parse_date_loose(row.get("completion_date"))
        if inv_date is None or row_date is None:
            return False
        return within_window(inv_date, row_date, DATE_WINDOW_DAYS)

    vendor_matched = [
        row
        for row in pool
        if vendor_norm
        and row["_vendor_norm"]
        and (vendor_norm in row["_vendor_norm"] or row["_vendor_norm"] in vendor_norm)
        and date_ok_lenient(row)
    ]
    if vendor_matched:
        return vendor_matched

    # フォールバック: 業者名では絞れなかったので日付の近さだけで候補を作る。
    # 並べ替えは「日付が近い順」を優先する（業者名の類似度は当てにならないことを
    # 確認済みなので、件数を絞るための主軸には使わない。同着の場合のみ類似度で補助）。
    fallback = [row for row in pool if date_ok_strict(row)]

    def sort_key(row: dict[str, Any]) -> tuple[int, float]:
        row_date = parse_date_loose(row.get("completion_date"))
        day_diff = abs((inv_date - row_date).days) if (inv_date and row_date) else 9999
        return (day_diff, -fuzzy_ratio(vendor_norm, row["_vendor_norm"]))

    fallback.sort(key=sort_key)
    return fallback[:VENDOR_FALLBACK_LIMIT]


# ==============================================================
# フェーズ2: LLM判定プロンプト・JSON Schema構築
# ==============================================================

MATCH_JSON_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "matched_row_id": {
            "type": "string",
            "description": "候補のrow_id。該当する行が無ければ空文字列。",
        },
        "confidence": {
            "type": "string",
            "enum": ["high", "medium", "low", "no_match"],
        },
        "reasoning": {
            "type": "string",
            "description": "一致した理由、または一致しない/候補が複数ある理由を具体的に。",
        },
        "mismatched_fields": {
            "type": "array",
            "items": {"type": "string"},
            "description": "一致しなかった項目名と差分の説明。無ければ空配列。",
        },
        "alternative_candidates": {
            "type": "array",
            "items": {"type": "string"},
            "description": "他に可能性がある候補のrow_id。無ければ空配列。",
        },
    },
    "required": ["matched_row_id", "confidence", "reasoning", "mismatched_fields", "alternative_candidates"],
}


def build_match_prompt(
    invoice_item: dict[str, Any],
    candidates: list[dict[str, Any]],
    excel_label: str,
) -> str:
    invoice_json = json.dumps(invoice_item, ensure_ascii=False, indent=2, default=str)
    candidates_json = json.dumps(candidates, ensure_ascii=False, indent=2, default=str)

    return f"""あなたは不動産管理会社の経理担当者です。業者から届いた請求書1件分の情報と、
社内Excel「{excel_label}」の候補行一覧を比較し、どの行が同じ工事案件を指しているかを
判定してください（出力形式は指定されたスキーマに従ってください）。

## 判定のルール
- 表記のゆれは同じものとみなしてよい。例:
  - 業者名:「ミドリガス」と「ミドリガス株式会社」は同じ。「SOS佐藤」と「2SOS60」の
    ように見た目が全く違っても、日付・物件・金額が一致していれば同一業者の
    社内コード表記の可能性が高いと判断してよい。
  - 物件名:「ソレイユ」と「SOLEIL」、「アムールノース」と「アムール・ノース」は同じ。
- 金額比較には、候補行の "vendor_amount"（業者への支払額＝原価相当）を使うこと。
  "billing_amount_to_customer" が候補行にあっても、それは家主・入居者への
  請求額（売価。原価に社内の掛け率を掛けた金額で、業者への支払額とは別物）
  なので、請求書の金額とは比較しないこと。
- 金額は請求書側が税抜、Excel側（vendor_amount）が税込で記録されていることがある。
  請求書の金額に消費税10%を加算した額とExcel側のvendor_amountが一致する場合は、
  金額が一致しているとみなしてよい。
- 号室・物件名が完全一致していなくても、業者名・日付・金額の組み合わせから
  同一案件と強く推測できる場合は一致と判定してよい。ただし金額が
  税込換算しても明確に異なる場合は、他がどれだけ似ていても「不一致」と
  判定し、具体的にいくら違うかを指摘すること。
- 候補の中に該当する行が無い場合は、無理に1件を選ばず matched_row_id を
  空文字列にし、no_match と判定すること。
  **ただしno_matchの場合でも、alternative_candidatesには「完全には一致しないが
  人が見て確認する価値がある」候補を1〜3件は必ず挙げること**（同じ業者名の別日、
  日付だけ近い、金額だけ近い、物件名が似ている、等の理由でよい）。候補行一覧に
  1件も近いものが無いと判断した場合のみ空配列にしてよい。
- 候補が複数あり甲乙つけがたい場合も、無理に1件を選ばず low とし、
  該当しそうな行を alternative_candidates に列挙すること。

## 請求書側の情報
{invoice_json}

## 候補行一覧（{excel_label}）
{candidates_json}
"""


# ==============================================================
# フェーズ2: 突合LLM呼び出し
# ==============================================================

async def call_llm_match(prompt: str) -> dict[str, Any]:
    """agent_sdk.llm_call を呼び出し、突合判定をさせる（テキストのみ、Sonnet 5）。"""
    from agent_sdk import llm_call  # ローカルMOCK_MODE検証時にimport不要にするための遅延import

    response = await llm_call(
        prompt=prompt,
        model=MATCH_MODEL,
        schema=MATCH_JSON_SCHEMA,
    )
    return unwrap_llm_response(response)


async def call_llm_match_mock(prompt: str) -> dict[str, Any]:
    """MOCK_MODE=True のときに使うダミー応答（パイプライン検証専用）。"""
    return {
        "matched_row_id": None,
        "confidence": "no_match",
        "reasoning": "（モック応答）実際の判定は実行環境環境で行われます。",
        "mismatched_fields": [],
        "alternative_candidates": [],
    }


# ==============================================================
# フェーズ2: 突合対象（明細単位）の組み立て
# ==============================================================

def build_match_targets(invoice_record: dict[str, Any]) -> list[dict[str, Any]]:
    """
    抽出結果1件（請求書1枚）から、突合対象（＝Excel1行と比較する単位）を作る。

    line_itemsは「物件名・号室・作業日が同じもの」を同一訪問とみなして1つの
    突合対象にまとめる（グループ化）。理由: サンプル設備の請求書のように、同じ部屋への
    1回の訪問で行った複数の小作業（ソケット交換・ライト交換・電球交換...）を
    別々の明細行として書いてくる業者がいるが、Excel側はこれらをまとめて1行で
    登録している可能性が高い。明細を分割したままLLMに個別で判定させると
    llm_callの回数が無駄に増えるだけでなく、Excel側の「まとめられた1行」との
    突合精度もかえって下がる（バラバラの短い作業内容よりまとめた内容の方が
    Excel側の記載と一致しやすいため）。
    line_itemsが無ければヘッダーそのものを1件とする。
    """
    if invoice_record.get("error"):
        return []

    line_items = json.loads(invoice_record.get("line_items_json") or "[]")
    targets: list[dict[str, Any]] = []

    if not line_items:
        targets.append(
            {
                "source_file": invoice_record.get("source_file"),
                "line_index": None,
                "vendor_name": invoice_record.get("vendor_name"),
                "property_name": invoice_record.get("property_name"),
                "room_number": invoice_record.get("room_number"),
                "property_no": invoice_record.get("property_no"),
                "completion_date": invoice_record.get("completion_date"),
                "work_type": invoice_record.get("work_type"),
                "billing_content": invoice_record.get("billing_content"),
                "vendor_amount": invoice_record.get("vendor_amount"),
            }
        )
        return targets

    # (物件名, 号室, 作業日) が同じ明細を1グループ＝1突合対象にまとめる
    groups: dict[tuple[Any, Any, Any], dict[str, Any]] = {}
    group_order: list[tuple[Any, Any, Any]] = []
    for item in line_items:
        prop = item.get("property_name") or invoice_record.get("property_name")
        room = item.get("room_number") or invoice_record.get("room_number")
        date = item.get("work_date") or invoice_record.get("completion_date")
        key = (prop, room, date)
        if key not in groups:
            groups[key] = {
                "property_no": item.get("property_no") or invoice_record.get("property_no"),
                "descriptions": [],
                "amount_total": None,
            }
            group_order.append(key)
        g = groups[key]
        if not g["property_no"]:
            g["property_no"] = item.get("property_no")
        desc = item.get("work_description")
        if desc:
            g["descriptions"].append(desc)
        amt = item.get("line_amount")
        if amt is not None:
            g["amount_total"] = (g["amount_total"] or 0) + amt

    for line_index, (prop, room, date) in enumerate(group_order):
        g = groups[(prop, room, date)]
        targets.append(
            {
                "source_file": invoice_record.get("source_file"),
                "line_index": line_index,
                "vendor_name": invoice_record.get("vendor_name"),
                "property_name": prop,
                "room_number": room,
                "property_no": g["property_no"],
                "completion_date": date,
                "work_type": invoice_record.get("work_type"),
                "billing_content": "; ".join(g["descriptions"]) or invoice_record.get("billing_content"),
                "vendor_amount": g["amount_total"] if g["amount_total"] is not None else invoice_record.get("vendor_amount"),
            }
        )
    return targets


# ==============================================================
# フェーズ2: 突合本体（asyncio.gather + Semaphoreで並列実行）
# ==============================================================

async def match_one_excel(
    invoice_item: dict[str, Any],
    pool: list[dict[str, Any]],
    excel_label: str,
) -> dict[str, Any]:
    candidates = get_candidates(invoice_item, pool)
    if not candidates:
        return {
            "matched_row_id": None,
            "confidence": "no_match",
            "reasoning": "候補行が0件（同時期・同業者らしき行がExcelに存在しない）",
            "mismatched_fields": [],
            "alternative_candidates": [],
            "candidate_count": 0,
        }

    prompt = build_match_prompt(invoice_item, candidates, excel_label)
    result = await (call_llm_match_mock(prompt) if MOCK_MODE else call_llm_match(prompt))
    result["candidate_count"] = len(candidates)
    if not result.get("matched_row_id"):  # Geminiスキーマの都合上、無ければ空文字列で返る
        result["matched_row_id"] = None
    return result


async def process_match_target(
    invoice_item: dict[str, Any],
    riform_pool: list[dict[str, Any]],
    shuzen_pool: list[dict[str, Any]],
    semaphore: asyncio.Semaphore,
) -> dict[str, Any]:
    record = dict(invoice_item)
    async with semaphore:
        # リフォーム一覧・修繕売上一覧は独立した突合なので同時に走らせてよい
        riform_task = match_one_excel(invoice_item, riform_pool, "リフォーム一覧（本社管理部）")
        shuzen_task = match_one_excel(invoice_item, shuzen_pool, "修繕売上一覧(明細)")
        riform_result, shuzen_result = await asyncio.gather(
            riform_task, shuzen_task, return_exceptions=True
        )
    if isinstance(riform_result, Exception):
        riform_result = {"confidence": "error", "reasoning": str(riform_result)}
    if isinstance(shuzen_result, Exception):
        shuzen_result = {"confidence": "error", "reasoning": str(shuzen_result)}

    record["match_riform"] = riform_result
    record["match_shuzen"] = shuzen_result
    label = f"{record['source_file']}#{record['line_index']}"
    print(
        f"[突合] {label} -> リフォーム一覧:{riform_result.get('confidence')} / "
        f"修繕売上一覧:{shuzen_result.get('confidence')}"
    )
    return record


def flatten_for_csv(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    flat = []
    for r in records:
        row = {k: v for k, v in r.items() if k not in ("match_riform", "match_shuzen")}
        for prefix, sub in (("riform", r.get("match_riform", {})), ("shuzen", r.get("match_shuzen", {}))):
            for k, v in (sub or {}).items():
                row[f"{prefix}_{k}"] = json.dumps(v, ensure_ascii=False) if isinstance(v, (list, dict)) else v
        flat.append(row)
    return flat


async def run_matching_phase(invoice_records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    print("[突合] Excelを読み込み中...")
    riform_pool = load_riform_rows()
    shuzen_pool = load_shuzen_rows()
    print(f"[突合] リフォーム一覧: {len(riform_pool)}行 / 修繕売上一覧: {len(shuzen_pool)}行")

    match_targets: list[dict[str, Any]] = []
    for invoice_record in invoice_records:
        match_targets.extend(build_match_targets(invoice_record))
    print(f"[突合] 突合対象（明細単位）: {len(match_targets)}件（並列数: {MAX_WORKERS}）")

    semaphore = asyncio.Semaphore(MAX_WORKERS)
    results = await asyncio.gather(
        *(process_match_target(item, riform_pool, shuzen_pool, semaphore) for item in match_targets)
    )
    results = sorted(results, key=lambda r: (r.get("source_file", ""), r.get("line_index") or 0))

    with open(MATCH_RESULTS_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2, default=str)

    flat = flatten_for_csv(results)
    if flat:
        with open(MATCH_RESULTS_CSV_PATH, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=list(flat[0].keys()))
            writer.writeheader()
            writer.writerows(flat)

    print(f"[突合] 保存しました: {MATCH_RESULTS_JSON_PATH.name} / {MATCH_RESULTS_CSV_PATH.name}")

    # ファイルを開かなくても実行ログだけで中身が分かるように、生JSONと
    # 自然文サマリーの両方を標準出力に直接書く（ユーザー指摘: ファイルを
    # 別途開いて貼ってもらうのは非現実的）。
    print("\n[突合] ===== 生JSON結果 =====")
    print(json.dumps(results, ensure_ascii=False, indent=2, default=str))

    print_natural_language_summary(results)
    return results


def print_natural_language_summary(results: list[dict[str, Any]]) -> None:
    """突合結果を人間が読める自然文で標準出力に書く。
    エージェントがチャット上でCSVへのリンクだけを貼って済ませてしまう事故を
    防ぐため、ここで作った文章をそのまま報告に使ってもらう想定。"""
    print("\n[突合] ===== 自然文サマリー =====")
    if not results:
        print("突合対象がありませんでした。")
        return

    counts = {"match": 0, "mismatch_or_unclear": 0, "no_match": 0}
    for r in results:
        vendor = r.get("vendor_name") or "(業者名不明)"
        prop = r.get("property_name") or "(物件名不明)"
        room = r.get("room_number")
        date = r.get("completion_date") or "(日付不明)"
        amount = r.get("vendor_amount")
        label = f"{r.get('source_file')}" + (f"（明細{r.get('line_index')}）" if r.get("line_index") is not None else "")

        print(f"\n■ {label}: {vendor} / {prop}" + (f" {room}" if room else "") + f" / {date} / {amount}円")
        for excel_key, excel_label in (("match_riform", "リフォーム一覧"), ("match_shuzen", "修繕売上一覧")):
            m = r.get(excel_key) or {}
            confidence = m.get("confidence", "unknown")
            reasoning = m.get("reasoning", "")
            matched_id = m.get("matched_row_id")
            mismatched = m.get("mismatched_fields") or []
            cand_count = m.get("candidate_count")

            if confidence == "high" and matched_id:
                counts["match"] += 1
                print(f"  {excel_label}: 一致（{matched_id}）— {reasoning}")
            elif confidence == "no_match" or not matched_id:
                counts["no_match"] += 1
                reason = reasoning or f"候補{cand_count}件の中に該当なし"
                print(f"  {excel_label}: 該当なし — {reason}")
            else:
                counts["mismatch_or_unclear"] += 1
                extra = f"（差分: {'; '.join(mismatched)}）" if mismatched else ""
                print(f"  {excel_label}: {confidence} — {reasoning}{extra}")

    total = sum(counts.values())
    print(
        f"\n[突合] サマリー合計: 一致{counts['match']}件 / "
        f"不一致・要確認{counts['mismatch_or_unclear']}件 / "
        f"該当なし{counts['no_match']}件（全{total}件）"
    )


# ==============================================================
# メイン（フェーズ1 → フェーズ2 を通しで実行）
# ==============================================================

async def main() -> None:
    invoice_records = await run_extraction_phase()
    await run_matching_phase(invoice_records)


if __name__ == "__main__":
    asyncio.run(main())
