"""出力:
  output/取引成立台帳_{貸主名}.xlsx
  output/full_log.json
  output/review_{貸主名}.html
"""
import subprocess
try:
    import pypdfium2
except ImportError:
    subprocess.run(["pip", "install", "pypdfium2", "--quiet"], check=True)

try:
    from PIL import Image
except ImportError:
    subprocess.run(["pip", "install", "Pillow", "--quiet"], check=True)
    from PIL import Image

import asyncio
import base64
import csv
import io
import json
import re
import os
import html as _html
from datetime import datetime
from pathlib import Path
import pypdfium2 as pdfium
import openpyxl
from openpyxl.styles import PatternFill
from agent_sdk import llm_call, file_create, file_output, ToolCallError

# =====================================================================
# 設定
# =====================================================================

MAX_CONCURRENT = 16
PAGE_TIMEOUT = 60
DPI = 150
EXTRACT_MODEL = "gemini/gemini-3.5-flash"
INTEGRATE_MODEL = "openai/gpt-5.4-mini"

# =====================================================================
# 定数
# =====================================================================

ITEMS: list[str] = [
    "成立年月日", "入居引渡年月日",
    "貸主氏名", "貸主住所", "貸主電話番号",
    "代理人氏名", "代理人住所", "代理人電話番号",
    "借主氏名", "借主住所", "借主電話番号",
    "入居者氏名", "入居者電話番号",
    "所在地", "物件の名称", "構造",
    "管理委託先の名称", "管理委託先の住所", "管理委託先の電話番号",
    "専有面積",
    "契約開始年", "契約開始月", "契約開始日",
    "契約終了年", "契約終了月", "契約終了日",
    "所在階", "部屋番号",
    "賃料", "敷金", "礼金", "償却",
    "駐車場賃料", "駐車場敷金",
    "火災保険料", "鍵交換費用", "初回保証料", "管理費等",
    "更新料", "解約予告期日", "間取り", "設備情報",
    "仲介手数料受領日", "買主氏名",
    "仲介手数料税抜", "仲介手数料消費税", "仲介手数料合計",
]

_CIT = {
    "type": "string",
    "description": (
        'JSON配列の文字列。形式: [{"key":"項目名","quote":"抽出値の前後10文字を含むPDF原文スニペット"}]。'
        '空欄にした項目は含めない。'
        '例: {"key":"賃料","quote":"月額賃料 155,000円/月 管理"}'
    ),
}
_UNC = {
    "type": "string",
    "description": 'JSON配列の文字列。形式: [{"key":"項目名","reason":"理由"}]。自信があれば "[]"',
}
_COORD = {
    "type": "string",
    "description": (
        'JSON配列の文字列。形式: [{"key":"項目名","box_2d":[ymin,xmin,ymax,xmax]}]。'
        'box_2d は画像上の正規化座標（0〜1000）。ymin/xmin=左上, ymax/xmax=右下。'
        '値を返した全項目について含めること。空欄項目は含めない。'
    ),
}

EXTRACT_SCHEMA: dict = {
    "type": "object",
    "properties": {
        **{item: {"type": "string"} for item in ITEMS},
        "_citations_json": _CIT,
        "_uncertainties_json": _UNC,
        "_coordinates_json": _COORD,
    },
}

INTEGRATE_SCHEMA: dict = {
    "type": "object",
    "properties": {
        **{item: {"type": "string"} for item in ITEMS},
        **{f"{item}_source": {"type": "string"} for item in ITEMS},
        "_uncertainties_json": _UNC,
    },
}

EXTRACT_PROMPT = """\
あなたは賃貸契約書類から情報を抜き出す機械です。
画像は契約書類の1ページです。以下のルールで47項目を抽出してください。
このページに記載がない項目は必ず空欄（""）にすること。創作しないこと。
値には書類に書かれている内容そのものだけを書く。記載の有無・省略についての説明や理由を値に書くことは絶対に禁止。

## 人物エンティティの一貫性ルール（最重要）
貸主・代理人・借主はそれぞれ独立した契約当事者として扱う。

- 貸主（甲・賃貸人）: 「貸主」「甲」「賃貸人」はすべて同じ当事者を指す。
  「賃貸人の表示」欄に記載された法人名・氏名を貸主として記載すること。
  登記簿（甲区欄）の所有者は貸主と異なる場合があるため、混同しないこと。
  信託銀行が登記上の所有者であっても、「賃貸人の表示」に別の法人が記載されていればそちらを優先する。
  管理委託先・管理会社・仲介業者・代理人は貸主ではない。
- 代理人: 貸主欄に「代理人」と明記されている者のみ記載する。記載がなければ必ず空欄（""）にすること。
- 借主（乙・賃借人）: 「乙」「賃借人」として契約に署名する当事者。

【重要】氏名・住所・電話番号は必ず同一人物/法人の情報をセットで記載すること。

## 抽出ルール
- 成立年月日: 契約が成立した日付。「YYYY年M月D日」形式のみ有効
- 入居引渡年月日: 鍵引渡し・入居開始日。「YYYY年M月D日」形式。記載なければ契約期間の開始日
- 貸主氏名: 物件所有者の氏名または法人名。管理会社名と混同しないこと
- 貸主電話番号: 自分が抽出した貸主氏名と同じ法人名・氏名に直接紐づく電話番号のみ記載すること。貸主欄の近くに記載がなければ空欄にする。管理委託先・仲介業者・代理人の電話番号を流用しないこと
- 代理人氏名: 貸主の代理人。法人なら「法人名 役職 氏名」形式。記載なければ空欄
- 借主氏名: 賃貸借契約を締結する借主。外国人名はアルファベット表記のまま
- 専有面積: 数値のみ（単位「㎡」「m²」は含めない。例: 55.64）
- 契約開始年/月/日、契約終了年/月/日: 西暦数値のみ（例: 2026、4、1）
- 所在階: 数値のみ（「階」を含めない。例: 14）
- 部屋番号: 数値またはアルファベット+数値のみ（「号室」「号」を含めない。例: 1403、F）
- 賃料・各金額（賃料/敷金/礼金/償却/管理費等/駐車場賃料/駐車場敷金/火災保険料/鍵交換費用/初回保証料/仲介手数料3項目）:
  整数値のみ。「円」「¥」「,」「/月」「ヶ月」等の単位・記号を一切含めないこと（例: 120000）
- 火災保険料: 入居者が加入する火災保険（家財保険）の保険料のみ。
  「火災保険」「家財保険」と明記された金額を記載する。
  社宅費用・管理費・保証料・社宅保険料など他の費用と混同しないこと。記載がなければ空欄
- 構造: 略称のまま（RC、SRC、木造等）。「鉄筋コンクリート造」→ RC
- 更新料: 算出方法の文字列（例: 新賃料の1ヶ月分）
- 解約予告期日: 月数の数値のみ（「ヶ月」「ヵ月」を含めない。例: 1）
- 仲介手数料受領日: 仲介手数料を借主から受領した日付。「YYYY年M月D日」形式
- 買主氏名: 仲介手数料を実際に支払った人の氏名。通常は借主本人と同一

## 必須出力
- _citations_json: 空でない値を返した全項目について、抽出値の前後10文字を含む原文スニペットをJSON配列で返すこと
- _uncertainties_json: 自信がない項目があればJSON配列で報告すること
- _coordinates_json: 空でない値を返した全項目について、その値が書かれている画像上の位置をバウンディングボックスで返すこと。box_2d は [ymin, xmin, ymax, xmax] の4要素配列。各値は 0〜1000 の正規化座標（ymin/xmin=左上, ymax/xmax=右下）。空欄項目は含めない。省略禁止
  出力例（架空の例）: 賃料「月額 98,000円」と借主氏名「山田 太郎」を抽出した場合
  → [{"key":"賃料","box_2d":[412,180,430,265]},{"key":"借主氏名","box_2d":[120,300,142,410]}]\
"""

PRIORITY_FIRST: str = "契約書"
PRIORITY_EQUAL: list[str] = ["マイソク", "申込書", "請求書"]

DOC_LABELS: dict[str, str] = {
    "契約書":  "契約書",
    "マイソク": "マイソク",
    "申込書":  "申込書",
    "請求書":  "請求書",
}

MAPPING: dict[str, list[str]] = {
    "成立年月日":           ["H15"],
    "入居引渡年月日":       ["AA15"],
    "貸主氏名":             ["H17"],
    "貸主住所":             ["H19"],
    "貸主電話番号":         ["AM17"],
    "代理人氏名":           ["H21"],
    "代理人住所":           ["H23"],
    "代理人電話番号":       ["AM21"],
    "借主氏名":             ["H25"],
    "借主住所":             ["H27"],
    "借主電話番号":         ["AM25"],
    "入居者氏名":           ["H29"],
    "入居者電話番号":       ["AM29"],
    "所在地":               ["H32"],
    "物件の名称":           ["H34"],
    "構造":                 ["H36"],
    "管理委託先の名称":     ["L43"],
    "管理委託先の住所":     ["L45"],
    "管理委託先の電話番号": ["AM43"],
    "専有面積":             ["AI36"],
    "契約開始年":           ["H56"],
    "契約開始月":           ["N56"],
    "契約開始日":           ["R56"],
    "契約終了年":           ["X56"],
    "契約終了月":           ["AD56"],
    "契約終了日":           ["AH56"],
    "所在階":               ["AK34"],
    "部屋番号":             ["AQ34"],
    "賃料":                 ["K48"],
    "敷金":                 ["X48"],
    "礼金":                 ["H52"],
    "償却":                 ["X50"],
    "駐車場賃料":           ["AA52"],
    "駐車場敷金":           ["X54"],
    "火災保険料":           ["AM48"],
    "鍵交換費用":           ["AM50"],
    "初回保証料":           ["AM52"],
    "管理費等":             ["K50"],
    "更新料":               ["M58"],
    "解約予告期日":         ["AB58"],
    "間取り":               ["AD38"],
    "設備情報":             ["H40"],
    "仲介手数料受領日":     ["L65"],
    "買主氏名":             ["X65"],
    "仲介手数料税抜":       ["L67", "AD69"],
    "仲介手数料消費税":     ["AE67", "AP69"],
    "仲介手数料合計":       ["AP67", "L69"],
}

SHEET_NAME = "733-3.取引成立台帳（居住用建物賃貸借）"
STATE_PATH = "tmp/state.json"
YELLOW_FILL = PatternFill("solid", fgColor="FFFF00")

NO_YELLOW_ITEMS: set[str] = {
    "代理人氏名", "代理人住所", "代理人電話番号",
}

CURRENCY_ITEMS: set[str] = {
    "賃料", "敷金", "礼金", "償却", "管理費等",
    "駐車場賃料", "駐車場敷金", "火災保険料", "鍵交換費用", "初回保証料",
    "仲介手数料税抜", "仲介手数料消費税", "仲介手数料合計",
}

DATE_ITEMS: set[str] = {"成立年月日", "入居引渡年月日", "仲介手数料受領日"}

SECTIONS: list[tuple[str, list[str]]] = [
    ("日付", ["成立年月日", "入居引渡年月日", "契約開始年", "契約開始月", "契約開始日", "契約終了年", "契約終了月", "契約終了日"]),
    ("貸主・代理人", ["貸主氏名", "貸主住所", "貸主電話番号", "代理人氏名", "代理人住所", "代理人電話番号"]),
    ("借主・入居者", ["借主氏名", "借主住所", "借主電話番号", "入居者氏名", "入居者電話番号"]),
    ("物件", ["所在地", "物件の名称", "構造", "専有面積", "所在階", "部屋番号", "間取り", "設備情報"]),
    ("管理委託先", ["管理委託先の名称", "管理委託先の住所", "管理委託先の電話番号"]),
    ("金額", ["賃料", "管理費等", "敷金", "礼金", "償却", "駐車場賃料", "駐車場敷金", "火災保険料", "鍵交換費用", "初回保証料", "更新料", "解約予告期日"]),
    ("仲介手数料", ["仲介手数料受領日", "買主氏名", "仲介手数料税抜", "仲介手数料消費税", "仲介手数料合計"]),
]

# =====================================================================
# ログ（printの代わり）
# =====================================================================

_log_entries: list[dict] = []

def _log(event: str, **kwargs) -> None:
    _log_entries.append({"t": datetime.now().isoformat(), "event": event, **kwargs})

# =====================================================================
# ヘルパー
# =====================================================================

def parse_int(val: str) -> int | None:
    if not val:
        return None
    cleaned = re.sub(r"[¥円,\s\\]", "", val)
    m = re.match(r"[\d.]+", cleaned)
    if not m:
        _log("parse_int_no_match", raw=val, cleaned=cleaned)
        return None
    num_str = m.group()
    try:
        return int(num_str)
    except ValueError:
        try:
            result = int(float(num_str))
            _log("parse_int_float_fallback", raw=val, result=result)
            return result
        except ValueError:
            _log("parse_int_failed", raw=val, num_str=num_str)
            return None


def is_empty(val: str) -> bool:
    if not val:
        return True
    s = val.strip()
    return s in ("", "不明", "未記入", "未定", "N/A", "-", "−", "ー") or s.lower() in ("null", "none")


def is_valid_date(val: str) -> bool:
    return bool(re.match(r"^\d{4}年\d{1,2}月\d{1,2}日$", val.strip())) if val else False


def is_valid_amount(val: str) -> bool:
    if not val:
        return False
    return bool(re.match(r"^\d", re.sub(r"[,¥円\s]", "", val)))


def format_cell(item: str, val: str) -> int | float | str | None:
    if item in CURRENCY_ITEMS:
        result = parse_int(val)
        _log("format_cell", item=item, raw=val, result=result, type="currency")
        return result
    if item == "専有面積":
        m = re.search(r"[\d.]+", val)
        if m:
            try:
                result = float(m.group())
                _log("format_cell", item=item, raw=val, result=result, type="area")
                return result
            except ValueError:
                _log("format_cell_area_fallback", item=item, raw=val, matched=m.group())
                return m.group()
        return val
    if item == "所在階":
        m = re.search(r"\d+", val)
        if m:
            result = int(m.group())
            _log("format_cell", item=item, raw=val, result=result, type="floor")
            return result
        return val
    if item == "部屋番号":
        m = re.search(r"[A-Za-z]*\d+", val.strip())
        if m:
            _log("format_cell", item=item, raw=val, result=m.group(), type="room")
            return m.group()
        return val.strip() if val.strip() else val
    if item in {"契約開始年", "契約開始月", "契約開始日",
                "契約終了年", "契約終了月", "契約終了日"}:
        m = re.search(r"\d+", val)
        if m:
            try:
                result = int(m.group())
                _log("format_cell", item=item, raw=val, result=result, type="date_part")
                return result
            except ValueError:
                _log("format_cell_date_fallback", item=item, raw=val, matched=m.group())
                return m.group()
        return val
    return val


def parse_json_field(raw: str) -> list[dict]:
    try:
        items = json.loads(raw) if isinstance(raw, str) else (raw or [])
    except Exception:
        items = []
    return [x for x in items if isinstance(x, dict) and "key" in x]

# =====================================================================
# fitz: PDF → PNG リスト
# =====================================================================

def render_pages(pdf_path: str, doc_type: str, page_range: range) -> list[str]:
    doc = pdfium.PdfDocument(pdf_path)
    max_dim = int(DPI / 72 * 842)
    os.makedirs("tmp/pages", exist_ok=True)
    paths: list[str] = []
    for i in page_range:
        if i >= len(doc):
            _log("render_pages_skip", doc_type=doc_type, page=i+1, reason="page_index_out_of_range")
            break
        page = doc[i]
        w, h = page.get_size()
        scale = min(max_dim / w, max_dim / h, 1.0)
        bmp = page.render(scale=scale)
        img_path = f"tmp/pages/{doc_type}_p{i + 1}.png"
        bmp.to_pil().save(img_path)
        paths.append(img_path)
        _log("render_page_done", doc_type=doc_type, page=i+1, size=f"{w:.0f}x{h:.0f}", scale=f"{scale:.3f}")
    _log("render_pages_complete", doc_type=doc_type, total_rendered=len(paths))
    return paths


def count_pages(pdf_path: str) -> int:
    doc = pdfium.PdfDocument(pdf_path)
    return len(doc)

# =====================================================================
# llm_call: 1ページ画像 → 抽出（引用・不確信付き）
# =====================================================================

RECOVER_COORDS_SCHEMA: dict = {
    "type": "object",
    "properties": {"_coordinates_json": _COORD},
}


async def recover_coords(img_path: str, targets: list[dict]) -> list[dict]:
    """座標が返ってこなかった項目だけ、同じページ画像から位置を再取得する。

    抽出済みの値・引用は捨てずに保持し、位置だけを聞く（全ページ再抽出より安く確実）。
    targets: [{key, value, quote}]
    """
    lines = [f'- {t["key"]}: 「{t.get("quote") or t.get("value", "")}」' for t in targets]
    prompt = (
        "画像内で、以下の各テキストが書かれている位置を特定してください。\n"
        + "\n".join(lines)
        + "\n\nbox_2d は [ymin, xmin, ymax, xmax]、0〜1000の正規化座標（ymin/xmin=左上, ymax/xmax=右下）。"
        "そのテキスト行だけを囲む最小の矩形を返すこと。画像内に見つからない項目は含めない。\n"
        "\n"
        "## 出力例（架空の例）\n"
        "依頼:\n"
        "- 賃料: 「賃料 月額 98,000円」\n"
        "- 借主氏名: 「賃借人 山田 太郎」\n"
        "_coordinates_json:\n"
        '[{"key":"賃料","box_2d":[412,180,430,265]},{"key":"借主氏名","box_2d":[120,300,142,410]}]'
    )
    result = await llm_call(prompt, image_path=img_path, model=EXTRACT_MODEL, schema=RECOVER_COORDS_SCHEMA)
    coords = parse_json_field(result.get("data", {}).get("_coordinates_json", "[]"))
    usage = result.get("usage", {})
    _log("recover_coords_done", img_path=img_path,
         keys=[t["key"] for t in targets], returned=[c.get("key") for c in coords],
         input_tokens=usage.get("input_tokens", 0), output_tokens=usage.get("output_tokens", 0),
         model=result.get("model", ""))
    return coords


async def call_page(img_path: str, _retry: int = 0) -> tuple[dict, list[dict], list[dict], list[dict], int]:
    _log("call_page_start", img_path=img_path, retry=_retry)
    try:
        result = await asyncio.wait_for(
            llm_call(
                EXTRACT_PROMPT,
                image_path=img_path,
                model=EXTRACT_MODEL,
                schema=EXTRACT_SCHEMA,
            ),
            timeout=PAGE_TIMEOUT,
        )
    except asyncio.TimeoutError:
        _log("call_page_timeout", img_path=img_path, timeout=PAGE_TIMEOUT, retry=_retry)
        if _retry < 1:
            return await call_page(img_path, _retry=_retry + 1)
        return {}, [], [], [], 0
    raw = result.get("data", {})
    citations = parse_json_field(raw.pop("_citations_json", "[]"))
    uncertainties = parse_json_field(raw.pop("_uncertainties_json", "[]"))
    coordinates = parse_json_field(raw.pop("_coordinates_json", "[]"))
    usage = result.get("usage", {})
    filled_keys = {k for k, v in raw.items() if k in ITEMS and v and not is_empty(str(v))}
    coord_keys = {c.get("key") for c in coordinates if c.get("key")}
    missing_coords = filled_keys - coord_keys
    _log("call_page_done", img_path=img_path, filled=len(filled_keys), citations=len(citations),
         uncertainties=len(uncertainties), coordinates=len(coordinates),
         input_tokens=usage.get("input_tokens", 0), output_tokens=usage.get("output_tokens", 0),
         model=result.get("model", ""),
         missing_coords=list(missing_coords) if missing_coords else None)
    if missing_coords:
        quote_by_key = {c.get("key"): c.get("quote", "") for c in citations}
        targets = [{"key": k, "value": str(raw.get(k, "")), "quote": quote_by_key.get(k, "")}
                   for k in sorted(missing_coords)]
        try:
            recovered = await asyncio.wait_for(recover_coords(img_path, targets), timeout=PAGE_TIMEOUT)
        except (asyncio.TimeoutError, ToolCallError) as e:
            _log("recover_coords_failed", img_path=img_path, error=str(e))
            recovered = []
        recovered = [c for c in recovered
                     if c.get("key") in missing_coords
                     and isinstance(c.get("box_2d"), list) and len(c["box_2d"]) >= 4]
        if recovered:
            coordinates = coordinates + recovered
    return raw, citations, uncertainties, coordinates, usage.get("input_tokens", 0)


def merge_page(
    page_records: list[dict[str, str]],
    page_data: dict[str, str],
    page_citations: list[dict],
    page_uncertainties: list[dict],
    doc_type: str,
    page_num: int,
    all_citations: list[dict],
    all_uncertainties: list[dict],
) -> None:
    page_records.append(page_data)
    for c in page_citations:
        all_citations.append({**c, "source": doc_type, "page": page_num})
    for u in page_uncertainties:
        all_uncertainties.append({**u, "source": doc_type, "page": page_num})


def resolve_majority(page_records: list[dict[str, str]]) -> tuple[dict[str, str], dict[str, int]]:
    """書類内の全ページで最頻出の値を採用。採用ページ(0-based)も返す。"""
    freq: dict[str, dict[str, int]] = {item: {} for item in ITEMS}
    first_page: dict[str, dict[str, int]] = {item: {} for item in ITEMS}
    for page_idx, record in enumerate(page_records):
        for item in ITEMS:
            val = (record.get(item) or "").strip()
            if val and not is_empty(val):
                freq[item][val] = freq[item].get(val, 0) + 1
                if val not in first_page[item]:
                    first_page[item][val] = page_idx
    result: dict[str, str] = {}
    adopted_pages: dict[str, int] = {}
    conflicts = []
    for item in ITEMS:
        counts = freq[item]
        if counts:
            winner = max(counts, key=counts.get)
            result[item] = winner
            adopted_pages[item] = first_page[item][winner]
            if len(counts) > 1:
                conflicts.append({"item": item, "winner": winner, "candidates": counts})
        else:
            result[item] = ""
    if conflicts:
        _log("resolve_majority_conflicts", count=len(conflicts), details=conflicts)
    return result, adopted_pages

# =====================================================================
# STEP1: PDF → 全体抽出（16並列バッチ）
# =====================================================================

def render_all_pdfs(pdf_map: dict[str, str]) -> dict[str, list[str]]:
    """全PDFを同期的にレンダリングしてページ画像パスを返す。"""
    result: dict[str, list[str]] = {}
    for doc_type, pdf_path in pdf_map.items():
        total = count_pages(pdf_path)
        _log("render_start", doc_type=doc_type, total_pages=total)
        result[doc_type] = render_pages(pdf_path, doc_type, range(total))
    return result


async def extract_all_pages(
    source_pages: dict[str, list[str]],
    sem: asyncio.Semaphore,
) -> tuple[dict[str, tuple[dict[str, str], list[dict], list[dict]]], dict[str, dict[int, list[dict]]], dict[str, dict[str, int]]]:
    """全書類の全ページを一括で並列抽出。座標データと採用ページマップも返す。"""

    async def sem_call(doc_type: str, page_idx: int, img_path: str) -> tuple[str, int, dict, list, list, list]:
        async with sem:
            try:
                data, cit, unc, coords, _ = await call_page(img_path)
                return doc_type, page_idx, data, cit, unc, coords
            except ToolCallError as e:
                _log("page_error", doc_type=doc_type, page=page_idx + 1, error=str(e))
                return doc_type, page_idx, {}, [], [], []

    all_tasks = [
        sem_call(dt, i, img)
        for dt, images in source_pages.items()
        for i, img in enumerate(images)
    ]
    _log("extract_start", total_tasks=len(all_tasks))
    page_results = await asyncio.gather(*all_tasks)

    per_doc_records: dict[str, list[dict[str, str]]] = {}
    per_doc_citations: dict[str, list[dict]] = {}
    per_doc_uncertainties: dict[str, list[dict]] = {}
    page_coords: dict[str, dict[int, list[dict]]] = {}

    for doc_type, page_idx, data, cit, unc, coords in page_results:
        records = per_doc_records.setdefault(doc_type, [])
        assert len(records) == page_idx, (
            f"page_records index mismatch: expected {page_idx}, got {len(records)} for {doc_type}"
        )
        cits = per_doc_citations.setdefault(doc_type, [])
        uncs = per_doc_uncertainties.setdefault(doc_type, [])
        merge_page(records, data, cit, unc, doc_type, page_idx + 1, cits, uncs)
        for c in coords:
            c.setdefault("value", data.get(c.get("key", ""), ""))
        page_coords.setdefault(doc_type, {})[page_idx] = coords

    extractions: dict[str, tuple[dict[str, str], list[dict], list[dict]]] = {}
    adopted_page_map: dict[str, dict[str, int]] = {}
    for doc_type in source_pages:
        records = per_doc_records.get(doc_type, [])
        accumulated, adopted_pages = resolve_majority(records)
        adopted_page_map[doc_type] = adopted_pages
        cits = per_doc_citations.get(doc_type, [])
        uncs = per_doc_uncertainties.get(doc_type, [])
        filled = sum(1 for v in accumulated.values() if not is_empty(v))
        _log("extract_done", doc_type=doc_type, filled=filled, total=len(ITEMS))
        extractions[doc_type] = (accumulated, cits, uncs)

    return extractions, page_coords, adopted_page_map

# =====================================================================
# STEP2: LLM 統合
# =====================================================================

async def integrate_with_llm(
    extractions: dict[str, dict[str, str]],
    all_citations: list[dict],
) -> tuple[dict[str, str], dict[str, str], list[dict]]:
    """LLMに統合を委任。Returns: (merged, sources, integration_uncertainties)"""
    _log("integrate_start", doc_types=list(extractions.keys()), citation_count=len(all_citations))

    # 引用マップ: {doc_type: {item: [{page, quote}]}}
    cit_by_doc: dict[str, dict[str, list[dict]]] = {}
    for c in all_citations:
        src = c.get("source", "")
        key = c.get("key", "")
        if src and key:
            cit_by_doc.setdefault(src, {}).setdefault(key, []).append(c)

    all_doc_types = [PRIORITY_FIRST] + PRIORITY_EQUAL
    lines: list[str] = []
    for doc_type in all_doc_types:
        ext = extractions.get(doc_type, {})
        if not ext:
            continue
        lines.append(f"## {DOC_LABELS.get(doc_type, doc_type)}")
        for item in ITEMS:
            val = ext.get(item, "")
            if is_empty(val):
                continue
            lines.append(f"  {item}: {val}")
            for c in cit_by_doc.get(doc_type, {}).get(item, []):
                lines.append(f'    引用(p{c.get("page","?")}): 「{c.get("quote","")}」')
        lines.append("")

    extraction_text = "\n".join(lines)

    integrate_prompt = f"""\
あなたは賃貸契約の専門家です。4種類の書類から抽出された情報を統合して、{len(ITEMS)}項目の最終値を決定してください。

{extraction_text}
## 統合ルール
- 優先順位: 契約書 > マイソク = 申込書 = 請求書
- 契約書に記載がなければ他書類から補完する
- 日付は「YYYY年M月D日」形式のみ有効
- 金額は整数のみ（円・¥・カンマなし）
- 代理人は明示的な記載がある場合のみ。なければフィールドごと省略
- 各項目について採用ソース（契約書/マイソク/申込書/請求書）を {{item名}}_source フィールドに必ず記録すること
- 【絶対厳守】いずれかの書類に値が存在する項目は絶対に空欄にしないこと。不確信・矛盾があっても値を必ず記入し、空欄にしてよいのは「全書類に記載が一切ない場合」のみ
- 確信が持てない場合や複数書類で矛盾がある場合は、値を記入した上で _uncertainties_json に理由を記録すること。空欄にすることは禁止
- 全書類に記載がない項目は、そのフィールド自体を出力しないこと。「null」「なし」「不明」「-」等の文字列を値として書くことは絶対に禁止
"""

    result = await llm_call(integrate_prompt, model=INTEGRATE_MODEL, schema=INTEGRATE_SCHEMA)
    usage = result.get("usage", {})
    raw = result.get("data", {})

    merged = {item: raw.get(item, "") for item in ITEMS}
    sources = {
        item: raw[f"{item}_source"]
        for item in ITEMS
        if raw.get(f"{item}_source")
    }
    uncertainties = parse_json_field(raw.get("_uncertainties_json", "[]"))
    for u in uncertainties:
        u["source"] = "統合AI"

    _log(
        "integrate_done",
        filled=sum(1 for v in merged.values() if not is_empty(v)),
        uncertain_count=len(uncertainties),
        input_tokens=usage.get("input_tokens", 0),
        output_tokens=usage.get("output_tokens", 0),
        model=result.get("model", ""),
    )
    return merged, sources, uncertainties


async def save_csv(merged: dict[str, str]) -> None:
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(["項目名", "値"])
    for k, v in merged.items():
        writer.writerow([k, v])
    await file_create("tmp/info.csv", "﻿" + buf.getvalue())

# =====================================================================
# STEP3: Excel転記
# =====================================================================

def _fix_named_styles(wb: openpyxl.Workbook) -> None:
    """openpyxl が保存時に cellStyleXfs を圧縮するため、高い xfId を持つ
    named style の参照先が消えて Excel が破損修復を要求する。保存前に除去する。"""
    safe = [ns for ns in wb._named_styles
            if not hasattr(ns, 'name') or ns.name in ('Normal',)]
    removed = len(wb._named_styles) - len(safe)
    if removed:
        wb._named_styles.clear()
        wb._named_styles.extend(safe)
        _log("fix_named_styles", removed=removed)


def fill_excel(merged: dict[str, str], excel_path: str) -> str:
    _log("fill_excel_start", template=excel_path, item_count=len(merged))
    wb = openpyxl.load_workbook(excel_path)
    ws = wb[SHEET_NAME]

    yellow_count = 0
    written_count = 0
    for item, cells in MAPPING.items():
        raw = merged.get(item, "")
        if is_empty(raw):
            if item not in NO_YELLOW_ITEMS:
                for cell in cells:
                    ws[cell].fill = YELLOW_FILL
                yellow_count += 1
                _log("fill_excel_yellow", item=item, cells=cells)
        else:
            value = format_cell(item, raw)
            if value is not None and value != "":
                for cell in cells:
                    ws[cell] = value
                written_count += 1
                _log("fill_excel_write", item=item, cells=cells, value=value, value_type=type(value).__name__)
            else:
                for cell in cells:
                    ws[cell] = raw
                _log("fill_excel_raw_fallback", item=item, cells=cells, raw=raw)

    start_year  = merged.get("契約開始年", "")
    start_month = merged.get("契約開始月", "")
    start_day   = merged.get("契約開始日", "")

    if start_year and not is_empty(start_year):
        try:
            ws["AK2"] = int(start_year) - 2018
        except ValueError:
            ws["AK2"].fill = YELLOW_FILL
    else:
        ws["AK2"].fill = YELLOW_FILL

    if start_month and not is_empty(start_month):
        try:
            ws["AP2"] = int(re.search(r"\d+", start_month).group())
        except (ValueError, AttributeError):
            ws["AP2"] = start_month
    else:
        ws["AP2"].fill = YELLOW_FILL

    if start_day and not is_empty(start_day):
        try:
            ws["AT2"] = int(re.search(r"\d+", start_day).group())
        except (ValueError, AttributeError):
            ws["AT2"] = start_day
    else:
        ws["AT2"].fill = YELLOW_FILL

    _log("fill_excel_header_done", yellow=yellow_count, written=written_count)

    _fix_named_styles(wb)
    lender = merged.get("貸主氏名", "") or "貸主名不明"
    out_path = f"output/取引成立台帳_{lender}.xlsx"
    wb.save(out_path)
    _log("fill_excel_saved", path=out_path)
    return out_path

# =====================================================================
# レビューHTML生成
# =====================================================================

def _encode_pages_b64(source_pages: dict[str, list[str]]) -> dict[str, list[str]]:
    result: dict[str, list[str]] = {}
    for doc_type, paths in source_pages.items():
        uris: list[str] = []
        for p in paths:
            with open(p, "rb") as f:
                uris.append("data:image/png;base64," + base64.b64encode(f.read()).decode())
        result[doc_type] = uris
    return result


def _pick_best_page(
    item: str,
    val: str,
    src: str,
    cits: list[dict],
    page_coords: dict[str, dict[int, list[dict]]] | None,
    adopted_page_map: dict[str, dict[str, int]] | None,
) -> tuple[int, str]:
    """項目クリック時の遷移先ページを決定する。1-based page を返す。

    優先順位:
      1. adopted_page（resolve_majorityが決めたページ）に座標がある → そのページ
      2. 同ソース内で座標がある最小ページ
      3. adopted_page（座標なしでも値は正しい）
      4. citation から最小ページ
      5. 見つからない → 0
    """
    src_coords = (page_coords or {}).get(src, {})
    adopted_pidx = (adopted_page_map or {}).get(src, {}).get(item)

    def _coord_pages(match_val: bool) -> list[int]:
        return sorted(
            pidx for pidx, cs in src_coords.items()
            if any(
                c.get("key") == item and (not match_val or c.get("value") == val)
                for c in cs
            )
        )

    if adopted_pidx is not None:
        adopted_coords = [
            c for c in src_coords.get(adopted_pidx, [])
            if c.get("key") == item
        ]
        if any(c.get("value") == val for c in adopted_coords):
            return adopted_pidx + 1, src

    exact_pages = _coord_pages(match_val=True)
    if exact_pages:
        return exact_pages[0] + 1, src

    any_pages = _coord_pages(match_val=False)
    if any_pages:
        return any_pages[0] + 1, src

    if adopted_pidx is not None:
        return adopted_pidx + 1, src

    src_cits = [c for c in cits if c.get("source") == src]
    best_cit = (src_cits or cits or [None])[0]
    if best_cit:
        return best_cit.get("page", 1), best_cit.get("source", src)

    return 0, src


def build_review_html(
    merged: dict[str, str],
    extractions: dict[str, dict[str, str]],
    sources: dict[str, str],
    all_citations: list[dict],
    all_uncertainties: list[dict],
    source_pages: dict[str, list[str]] | None = None,
    page_coords: dict[str, dict[int, list[dict]]] | None = None,
    adopted_page_map: dict[str, dict[str, int]] | None = None,
) -> str:
    _log("build_html_start", merged_count=len(merged), citation_count=len(all_citations),
         uncertainty_count=len(all_uncertainties),
         has_pages=source_pages is not None, has_coords=page_coords is not None)
    e = _html.escape

    page_data_uris: dict[str, list[str]] = {}
    if source_pages:
        page_data_uris = _encode_pages_b64(source_pages)

    cmap: dict[str, list[dict]] = {}
    for c in all_citations:
        cmap.setdefault(c["key"], []).append(c)

    umap: dict[str, list[dict]] = {}
    for u in all_uncertainties:
        umap.setdefault(u["key"], []).append(u)

    all_doc_types = [PRIORITY_FIRST] + PRIORITY_EQUAL

    empty_keys: set[str] = set()
    has_raw_value: set[str] = set()
    for k, v in merged.items():
        if not is_empty(v):
            continue
        any_source_has_value = any(
            not is_empty(extractions.get(src, {}).get(k, ""))
            for src in all_doc_types
        )
        if any_source_has_value:
            has_raw_value.add(k)
        else:
            empty_keys.add(k)

    unc_keys: set[str] = set(has_raw_value)
    integrate_uncertain = {u["key"] for u in all_uncertainties if u.get("source") == "統合AI"}
    unc_keys |= integrate_uncertain

    total = len(merged)
    empty_cnt = len(empty_keys)
    unc_cnt = len(unc_keys - empty_keys)
    ok_cnt = total - empty_cnt - unc_cnt
    prop_name = merged.get("物件の名称", "") or merged.get("所在地", "物件名不明")
    room = merged.get("部屋番号", "")

    items_json: list[dict] = []
    no = 0
    for sec_name, keys in SECTIONS:
        for item in keys:
            if item not in merged:
                continue
            no += 1
            val = merged.get(item, "")
            src = sources.get(item, "")
            src_label = DOC_LABELS.get(src, src)
            cits = cmap.get(item, [])
            best_page, best_src = _pick_best_page(
                item, val, src, cits, page_coords, adopted_page_map,
            )

            if item in empty_keys:
                status = "empty"
            elif item in unc_keys:
                status = "warn"
            else:
                status = "ok"

            other_vals = []
            for dt in all_doc_types:
                if dt == src:
                    continue
                v2 = extractions.get(dt, {}).get(item, "")
                if v2 and not is_empty(v2):
                    other_vals.append({"src": DOC_LABELS.get(dt, dt), "val": v2})

            cite_list = [
                {"src": DOC_LABELS.get(c.get("source", ""), ""), "page": c.get("page", 0), "quote": c.get("quote", "")}
                for c in cits
            ]
            unc_list = [
                {"src": DOC_LABELS.get(u.get("source", ""), ""), "page": u.get("page", 0), "reason": u.get("reason", "")}
                for u in umap.get(item, [])
            ]

            items_json.append({
                "no": no, "section": sec_name, "field": item,
                "value": val, "status": status,
                "src": src, "srcLabel": src_label,
                "page": best_page, "docType": best_src,
                "others": other_vals, "citations": cite_list, "uncertainties": unc_list,
                "type": "num" if item in CURRENCY_ITEMS else "text",
            })

    doc_pages_json: dict[str, int] = {}
    for dt in all_doc_types:
        if dt in page_data_uris:
            doc_pages_json[dt] = len(page_data_uris[dt])
        elif dt in extractions:
            doc_pages_json[dt] = 0

    doc_uris_js = json.dumps(
        {dt: uris for dt, uris in page_data_uris.items()},
        ensure_ascii=False,
    ) if page_data_uris else "{}"

    coords_js = json.dumps(
        {dt: {str(p): [{"key": c.get("key",""), "box_2d": c.get("box_2d",[])} for c in cs]
              for p, cs in pages.items()}
         for dt, pages in (page_coords or {}).items()},
        ensure_ascii=False,
    )

    return f"""<!DOCTYPE html>
<html lang="ja"><head><meta charset="utf-8">
<title>台帳レビュー — {e(prop_name)}</title>
<style>
:root {{
  --ground: #F3F4F7; --surface: #FFFFFF; --text: #1A1D2B; --text-2: #5A6178;
  --accent: #2D4FCA; --accent-light: #EDF0FB;
  --ok: #0F7B5F; --ok-bg: #ECFAF4;
  --warn: #B45309; --warn-bg: #FFF8EB;
  --empty: #DC2626; --empty-bg: #FEF2F2;
  --border: #DDE0E9;
  --mono: ui-monospace, 'SF Mono', 'Cascadia Code', 'Consolas', monospace;
  --sans: system-ui, -apple-system, 'Segoe UI', sans-serif;
  --viewer-bg: #2A2D3A;
}}
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ background: var(--ground); color: var(--text); font-family: var(--sans); font-size: 14px; line-height: 1.5; height: 100vh; display: flex; flex-direction: column; overflow: hidden; }}

.topbar {{
  background: var(--text); color: #fff; padding: 12px 24px;
  display: flex; align-items: center; gap: 20px; flex-shrink: 0; z-index: 10;
}}
.topbar-label {{ font-size: 11px; opacity: 0.5; text-transform: uppercase; letter-spacing: 0.06em; font-weight: 600; }}
.topbar-bukken {{ font-size: 17px; font-weight: 700; }}
.topbar-meta {{ margin-left: auto; display: flex; gap: 12px; font-size: 12px; }}
.stat {{ padding: 3px 10px; border-radius: 6px; font-weight: 600; font-family: var(--mono); }}
.stat-ok {{ background: var(--ok-bg); color: var(--ok); }}
.stat-warn {{ background: var(--warn-bg); color: var(--warn); }}
.stat-empty {{ background: var(--empty-bg); color: var(--empty); }}

.main-split {{ display: flex; flex: 1; min-height: 0; }}

.left-pane {{
  width: 480px; min-width: 380px; flex-shrink: 0;
  background: var(--surface); border-right: 1px solid var(--border);
  display: flex; flex-direction: column; overflow: hidden;
}}
.pane-header {{
  padding: 10px 16px; border-bottom: 1px solid var(--border);
  display: flex; align-items: center; gap: 10px; flex-shrink: 0;
  background: var(--surface); height: 44px;
}}
.pane-header h2 {{ font-size: 13px; font-weight: 700; }}
.sort-btns {{ display: flex; gap: 0; margin-left: auto; }}
.sort-btn {{
  padding: 4px 10px; font-size: 11px; font-weight: 600;
  border: 1px solid var(--border); background: var(--ground);
  color: var(--text-2); cursor: pointer; font-family: var(--sans);
}}
.sort-btn:first-child {{ border-radius: 5px 0 0 5px; }}
.sort-btn:last-child {{ border-radius: 0 5px 5px 0; border-left: none; }}
.sort-btn.active {{ background: var(--accent); color: #fff; border-color: var(--accent); }}
.filter-btns {{ display: flex; gap: 0; margin-left: 8px; }}
.filter-btn {{
  padding: 4px 10px; font-size: 11px; font-weight: 600;
  border: 1px solid var(--border); background: var(--ground);
  color: var(--text-2); cursor: pointer; font-family: var(--sans);
}}
.filter-btn:first-child {{ border-radius: 5px 0 0 5px; }}
.filter-btn:last-child {{ border-radius: 0 5px 5px 0; border-left: none; }}
.filter-btn.active {{ background: var(--accent); color: #fff; border-color: var(--accent); }}

.extract-list {{ flex: 1; overflow-y: auto; }}
.section-divider {{
  padding: 6px 16px; background: var(--ground); font-size: 11px;
  font-weight: 700; color: var(--text-2); text-transform: uppercase;
  letter-spacing: 0.04em; border-bottom: 1px solid var(--border);
  position: sticky; top: 0; z-index: 2;
}}
.extract-row {{
  display: grid; grid-template-columns: 36px 1fr auto auto;
  gap: 0; padding: 10px 16px; border-bottom: 1px solid #F0F1F5;
  cursor: pointer; transition: background 0.1s; align-items: center;
}}
.extract-row:hover {{ background: #F8F9FC; }}
.extract-row.selected {{ background: var(--accent-light); border-left: 3px solid var(--accent); padding-left: 13px; }}
.extract-no {{
  font-family: var(--mono); font-size: 11px; font-weight: 700; color: var(--text-2);
  width: 28px; height: 28px; border-radius: 6px; background: var(--ground);
  display: flex; align-items: center; justify-content: center;
}}
.extract-row.selected .extract-no {{ background: var(--accent); color: #fff; }}
.extract-info {{ padding: 0 12px; }}
.extract-field {{ font-size: 13px; font-weight: 700; color: var(--text); }}
.extract-val {{ font-weight: 400; font-size: 13px; margin-top: 2px; color: var(--text-2); }}
.extract-val.is-num {{ font-family: var(--mono); font-variant-numeric: tabular-nums; }}
.extract-val.is-empty {{ opacity: 0.4; font-style: italic; }}
.extract-src {{
  font-size: 10px; font-weight: 600; padding: 2px 6px; border-radius: 4px;
  background: #6c757d; color: #fff; font-family: var(--mono); margin-right: 8px;
  white-space: nowrap;
}}
.extract-status {{
  font-size: 10px; font-weight: 700; padding: 2px 8px; border-radius: 4px;
  white-space: nowrap;
}}
.extract-status.s-ok {{ background: var(--ok-bg); color: var(--ok); }}
.extract-status.s-warn {{ background: var(--warn-bg); color: var(--warn); }}
.extract-status.s-empty {{ background: var(--empty-bg); color: var(--empty); }}

.right-pane {{
  flex: 1; background: var(--viewer-bg);
  display: flex; flex-direction: column; min-width: 0;
}}
.viewer-toolbar {{
  padding: 8px 16px; display: flex; align-items: center; gap: 12px;
  background: rgba(0,0,0,0.3); flex-shrink: 0; height: 44px;
}}
.doc-tabs {{ display: flex; gap: 4px; }}
.doc-tab {{
  padding: 4px 12px; border-radius: 5px; font-size: 11px; font-weight: 600;
  color: rgba(255,255,255,0.6); background: rgba(255,255,255,0.08);
  cursor: pointer; border: none; font-family: var(--sans);
}}
.doc-tab:hover {{ background: rgba(255,255,255,0.15); }}
.doc-tab.active {{ background: var(--accent); color: #fff; }}
.page-nav {{ display: flex; align-items: center; gap: 8px; margin-left: auto; }}
.page-btn {{
  width: 32px; height: 32px; border-radius: 6px;
  border: 1px solid rgba(255,255,255,0.15); background: rgba(255,255,255,0.08);
  color: #fff; font-size: 16px; cursor: pointer;
  display: flex; align-items: center; justify-content: center; font-family: var(--sans);
}}
.page-btn:hover {{ background: rgba(255,255,255,0.15); }}
.page-btn:disabled {{ opacity: 0.3; cursor: default; }}
.page-indicator {{
  color: #fff; font-family: var(--mono); font-size: 13px; font-weight: 600;
  min-width: 60px; text-align: center;
}}
.page-indicator span {{ opacity: 0.5; font-weight: 400; }}
.viewer-thumbs {{ display: flex; gap: 4px; margin-left: 12px; }}
.thumb {{
  width: 28px; height: 36px; border-radius: 3px; border: 2px solid transparent;
  cursor: pointer; font-family: var(--mono); font-size: 9px; font-weight: 700;
  color: rgba(255,255,255,0.6); display: flex; align-items: center; justify-content: center;
  background: rgba(255,255,255,0.08); flex-shrink: 0;
}}
.thumb:hover {{ background: rgba(255,255,255,0.15); }}
.thumb.active {{ border-color: var(--accent); background: rgba(45,79,202,0.3); color: #fff; }}

.viewer-body {{
  flex: 1; display: flex; align-items: flex-start; justify-content: center;
  padding: 0 16px 16px; overflow: auto;
}}
.img-wrapper {{
  position: relative; display: none; transform-origin: top left;
}}
.img-wrapper.visible {{ display: inline-block; }}
.img-wrapper img {{
  display: block; max-width: 100%; height: auto;
  border-radius: 4px; box-shadow: 0 8px 32px rgba(0,0,0,0.4);
}}
.marker-overlay {{
  position: absolute; border: none;
  background: rgba(215, 55, 63, 0.18); pointer-events: none; border-radius: 4px;
  box-shadow: 0 0 0 2px rgba(215, 55, 63, 0.10);
  transition: opacity 0.3s ease;
}}
.zoom-btns {{ display: flex; gap: 4px; margin-left: 12px; }}
.zoom-btn {{
  width: 32px; height: 32px; border-radius: 6px;
  border: 1px solid rgba(255,255,255,0.15); background: rgba(255,255,255,0.08);
  color: #fff; font-size: 16px; cursor: pointer;
  display: flex; align-items: center; justify-content: center; font-family: var(--sans);
}}
.zoom-btn:hover {{ background: rgba(255,255,255,0.15); }}
.zoom-level {{
  color: rgba(255,255,255,0.6); font-family: var(--mono); font-size: 11px;
  min-width: 36px; text-align: center; display: flex; align-items: center;
  justify-content: center;
}}
.no-image {{
  color: rgba(255,255,255,0.3); font-size: 16px; font-weight: 600;
  margin-top: 40px; text-align: center;
}}

.detail-panel {{
  position: fixed; bottom: 0; left: 0; right: 0; z-index: 20;
  background: var(--surface); border-top: 2px solid var(--border);
  max-height: 40vh; overflow-y: auto; display: none;
  box-shadow: 0 -4px 16px rgba(0,0,0,0.1);
}}
.detail-panel.open {{ display: block; }}
.detail-header {{
  padding: 10px 20px; display: flex; align-items: center; gap: 12px;
  border-bottom: 1px solid var(--border); position: sticky; top: 0;
  background: var(--surface); z-index: 1;
}}
.detail-header h3 {{ font-size: 14px; font-weight: 700; }}
.cand-btn {{
  margin-left: auto; border: 1px solid var(--accent); background: var(--accent-light);
  color: var(--accent); font-size: 12px; font-weight: 700; padding: 5px 12px;
  border-radius: 6px; cursor: pointer; font-family: var(--sans); white-space: nowrap;
}}
.cand-btn:hover {{ background: var(--accent); color: #fff; }}
.detail-close {{
  border: none; background: none; font-size: 18px;
  cursor: pointer; color: var(--text-2); padding: 4px 8px;
}}
.detail-body {{ padding: 12px 20px; font-size: 13px; }}
.detail-row {{ display: flex; gap: 12px; padding: 4px 0; border-bottom: 1px solid #F0F1F5; }}
.detail-label {{ width: 100px; color: var(--text-2); font-weight: 600; flex-shrink: 0; }}
.detail-val {{ flex: 1; }}
.cite-line {{ color: var(--text-2); font-size: 12px; font-style: italic; margin-top: 2px; }}
.unc-line {{ color: var(--warn); font-size: 12px; margin-top: 2px; }}

@media (max-width: 900px) {{
  .main-split {{ flex-direction: column; }}
  .left-pane {{ width: 100%; min-width: 0; max-height: 50vh; }}
  .right-pane {{ min-height: 300px; }}
}}
</style>
</head><body>

<div class="topbar">
  <div>
    <div class="topbar-label">台帳レビュー</div>
    <div class="topbar-bukken">{e(prop_name)}{(' ' + e(room)) if room else ''}</div>
  </div>
  <div class="topbar-meta">
    <span class="stat stat-ok">{ok_cnt} OK</span>
    <span class="stat stat-warn">{unc_cnt} 要確認</span>
    <span class="stat stat-empty">{empty_cnt} 空欄</span>
  </div>
</div>

<div class="main-split">
  <div class="left-pane">
    <div class="pane-header">
      <h2>抽出結果 ({total})</h2>
      <div class="sort-btns">
        <button class="sort-btn active" onclick="sortBy('section')">項目順</button>
        <button class="sort-btn" onclick="sortBy('page')">ページ順</button>
      </div>
      <div class="filter-btns">
        <button class="filter-btn active" onclick="filterBy('all')">全て</button>
        <button class="filter-btn" onclick="filterBy('ok')">OK</button>
        <button class="filter-btn" onclick="filterBy('warn')">確認</button>
        <button class="filter-btn" onclick="filterBy('empty')">空欄</button>
      </div>
    </div>
    <div class="extract-list" id="extractList"></div>
  </div>

  <div class="right-pane">
    <div class="viewer-toolbar">
      <div class="doc-tabs" id="docTabs"></div>
      <div class="page-nav">
        <button class="page-btn" onclick="goPage(currentPage-1)" id="prevBtn">&#8249;</button>
        <div class="page-indicator" id="pageIndicator"></div>
        <button class="page-btn" onclick="goPage(currentPage+1)" id="nextBtn">&#8250;</button>
      </div>
      <div class="viewer-thumbs" id="thumbs"></div>
      <div class="zoom-btns">
        <button class="zoom-btn" onclick="zoom(-1)">&#8722;</button>
        <div class="zoom-level" id="zoomLevel">100%</div>
        <button class="zoom-btn" onclick="zoom(1)">&#43;</button>
        <button class="zoom-btn" onclick="zoomReset()">&#8634;</button>
      </div>
    </div>
    <div class="viewer-body" id="viewerBody"></div>
  </div>
</div>

<div class="detail-panel" id="detailPanel">
  <div class="detail-header">
    <h3 id="detailTitle"></h3>
    <button class="cand-btn" id="candBtn" onclick="nextCandidate()" style="display:none"></button>
    <button class="detail-close" onclick="closeDetail()">&#10005;</button>
  </div>
  <div class="detail-body" id="detailBody"></div>
</div>

<script>
const ITEMS = {json.dumps(items_json, ensure_ascii=False)};
const DOC_PAGES = {json.dumps(doc_pages_json, ensure_ascii=False)};
const DOC_URIS = {doc_uris_js};
const DOC_ORDER = {json.dumps([dt for dt in all_doc_types if dt in doc_pages_json], ensure_ascii=False)};
const COORDS = {coords_js};
const STATUS_LABELS = {{"ok":"OK","warn":"要確認","empty":"空欄"}};

let currentDoc = DOC_ORDER[0] || "";
let currentPage = 1;
let currentFilter = "all";
let currentSort = "section";
let selectedNo = null;
let zoomPct = 100;

function esc(s) {{ const d = document.createElement('div'); d.textContent = s; return d.innerHTML; }}

function renderDocTabs() {{
  const c = document.getElementById('docTabs');
  c.innerHTML = DOC_ORDER.map(dt =>
    '<button class="doc-tab' + (dt === currentDoc ? ' active' : '') +
    '" onclick="switchDoc(\\''+dt+'\\')">' + esc(dt) + '</button>'
  ).join('');
}}

function renderThumbs() {{
  const c = document.getElementById('thumbs');
  const total = DOC_PAGES[currentDoc] || 0;
  c.innerHTML = '';
  for (let i = 1; i <= total; i++) {{
    const t = document.createElement('div');
    t.className = 'thumb' + (i === currentPage ? ' active' : '');
    t.textContent = i;
    t.onclick = () => goPage(i);
    c.appendChild(t);
  }}
}}

function renderViewer() {{
  const body = document.getElementById('viewerBody');
  body.innerHTML = '';
  const uris = DOC_URIS[currentDoc] || [];
  if (!uris.length) {{
    body.innerHTML = '<div class="no-image">PDF画像なし</div>';
    return;
  }}
  uris.forEach((uri, i) => {{
    const wrapper = document.createElement('div');
    wrapper.className = 'img-wrapper' + ((i + 1 === currentPage) ? ' visible' : '');
    wrapper.id = 'wrapper-' + (i + 1);
    const img = document.createElement('img');
    img.src = uri;
    wrapper.appendChild(img);
    body.appendChild(wrapper);
  }});
}}

function goPage(n) {{
  const total = DOC_PAGES[currentDoc] || 0;
  if (n < 1 || n > total) return;
  currentPage = n;
  clearMarkers();
  document.querySelectorAll('.img-wrapper').forEach(el => el.classList.remove('visible'));
  const target = document.getElementById('wrapper-' + n);
  if (target) target.classList.add('visible');
  document.getElementById('pageIndicator').innerHTML = n + ' <span>/ ' + total + '</span>';
  document.getElementById('prevBtn').disabled = n <= 1;
  document.getElementById('nextBtn').disabled = n >= total;
  renderThumbs();
  if (selectedNo) showMarkersForItem(selectedNo);
}}

function switchDoc(dt) {{
  currentDoc = dt;
  currentPage = 1;
  renderDocTabs();
  renderViewer();
  goPage(1);
}}

function renderList() {{
  const list = document.getElementById('extractList');
  let sorted = [...ITEMS];
  if (currentSort === 'page') {{
    const docIdx = {{}};
    DOC_ORDER.forEach((d, i) => docIdx[d] = i);
    sorted.sort((a, b) => {{
      const da = docIdx[a.docType] !== undefined ? docIdx[a.docType] : 99;
      const db = docIdx[b.docType] !== undefined ? docIdx[b.docType] : 99;
      if (da !== db) return da - db;
      if (a.page !== b.page) return a.page - b.page;
      return a.no - b.no;
    }});
  }}
  let html = '';
  let lastGroup = '';
  sorted.forEach(item => {{
    if (currentFilter !== 'all' && item.status !== currentFilter) return;
    const group = currentSort === 'page'
      ? (item.docType || '不明') + ' P' + item.page
      : item.section;
    if (group !== lastGroup) {{
      html += '<div class="section-divider">' + esc(group) + '</div>';
      lastGroup = group;
    }}
    const sel = selectedNo === item.no ? ' selected' : '';
    const valClass = 'extract-val' + (item.type === 'num' ? ' is-num' : '') + (!item.value ? ' is-empty' : '');
    const statusClass = 's-' + item.status;
    html += '<div class="extract-row' + sel + '" data-no="' + item.no + '" onclick="selectItem(' + item.no + ')">' +
      '<div class="extract-no">' + item.no + '</div>' +
      '<div class="extract-info">' +
        '<div class="extract-field">' + esc(item.field) + '</div>' +
        '<div class="' + valClass + '">' + esc(item.value || '(なし)') + '</div>' +
      '</div>' +
      '<span class="extract-src">' + esc(item.srcLabel) + '</span>' +
      '<span class="extract-status ' + statusClass + '">' + STATUS_LABELS[item.status] + '</span>' +
    '</div>';
  }});
  list.innerHTML = html;
}}

function selectItem(no) {{
  selectedNo = no;
  const item = ITEMS.find(i => i.no === no);
  if (!item) return;
  candList = buildCandidates(item);
  candIdx = 0;
  renderList();
  showDetail(item);
  if (candList.length) {{
    jumpToCandidate(0);
  }} else {{
    if (item.docType && DOC_ORDER.includes(item.docType) && item.docType !== currentDoc) {{
      switchDoc(item.docType);
    }}
    if (item.page > 0) goPage(item.page);
    updateCandBtn();
  }}
}}

// ---- 候補巡回: 同じ項目の座標が複数ページ/書類にあるとき、ボタンで順に飛ぶ ----
let candList = [];
let candIdx = 0;

function buildCandidates(item) {{
  const list = [];
  DOC_ORDER.forEach(doc => {{
    const pages = COORDS[doc] || {{}};
    Object.keys(pages).sort((a, b) => Number(a) - Number(b)).forEach(p => {{
      (pages[p] || []).forEach(c => {{
        if (c.key === item.field && Array.isArray(c.box_2d) && c.box_2d.length >= 4) {{
          list.push({{doc: doc, page: Number(p) + 1, box: c.box_2d}});
        }}
      }});
    }});
  }});
  // 採用ページの候補を先頭に（安定ソート）
  return list.map((c, i) => [c, i]).sort((a, b) => {{
    const aa = (a[0].doc === item.docType && a[0].page === item.page) ? 0 : 1;
    const bb = (b[0].doc === item.docType && b[0].page === item.page) ? 0 : 1;
    return aa - bb || a[1] - b[1];
  }}).map(x => x[0]);
}}

function jumpToCandidate(i) {{
  const c = candList[i];
  if (!c) return;
  candIdx = i;
  if (c.doc !== currentDoc) switchDoc(c.doc);
  goPage(c.page);
  clearMarkers();
  const wrapper = document.getElementById('wrapper-' + c.page);
  if (wrapper) drawBox(wrapper, c.box);
  updateCandBtn();
}}

function nextCandidate() {{
  if (!candList.length) return;
  jumpToCandidate((candIdx + 1) % candList.length);
}}

function updateCandBtn() {{
  const btn = document.getElementById('candBtn');
  if (!btn) return;
  if (candList.length < 2) {{ btn.style.display = 'none'; return; }}
  const c = candList[candIdx];
  btn.style.display = '';
  btn.textContent = '\\uD83D\\uDCCD ' + c.doc + ' p' + c.page + '（' + (candIdx + 1) + '/' + candList.length + '）　次の候補 ▸';
}}

function drawBox(wrapper, box) {{
  const b = (box || []).map(v => Math.max(0, Math.min(1000, v)));
  if (b.length < 4) return;
  const PAD = 12;
  let y0 = Math.min(b[0], b[2]), x0 = Math.min(b[1], b[3]);
  let y1 = Math.max(b[0], b[2]), x1 = Math.max(b[1], b[3]);
  y0 = Math.max(0, y0 - PAD); x0 = Math.max(0, x0 - PAD);
  y1 = Math.min(1000, y1 + PAD); x1 = Math.min(1000, x1 + PAD);
  const overlay = document.createElement('div');
  overlay.className = 'marker-overlay';
  overlay.style.left = (x0 / 10) + '%';
  overlay.style.top = (y0 / 10) + '%';
  overlay.style.width = ((x1 - x0) / 10) + '%';
  overlay.style.height = ((y1 - y0) / 10) + '%';
  wrapper.appendChild(overlay);
}}

function showMarkersForItem(no) {{
  clearMarkers();
  const item = ITEMS.find(i => i.no === no);
  if (!item) return;
  const pageIdx = currentPage - 1;
  const pageCoords = ((COORDS[currentDoc] || {{}})[String(pageIdx)]) || [];
  const matching = pageCoords.filter(c => c.key === item.field);
  const wrapper = document.getElementById('wrapper-' + currentPage);
  if (!wrapper || !matching.length) return;
  matching.forEach(coord => drawBox(wrapper, coord.box_2d));
}}

function clearMarkers() {{
  document.querySelectorAll('.marker-overlay').forEach(el => el.remove());
}}

function showDetail(item) {{
  document.getElementById('detailTitle').textContent = item.field;
  let html = '<div class="detail-row"><div class="detail-label">採用値</div><div class="detail-val"><strong>' + esc(item.value || '(なし)') + '</strong></div></div>';
  html += '<div class="detail-row"><div class="detail-label">採用ソース</div><div class="detail-val">' + esc(item.srcLabel) + '</div></div>';
  if (item.others.length) {{
    html += '<div class="detail-row"><div class="detail-label">他ソース</div><div class="detail-val">' +
      item.others.map(o => esc(o.src) + ': ' + esc(o.val)).join('<br>') + '</div></div>';
  }}
  if (item.citations.length) {{
    html += '<div class="detail-row"><div class="detail-label">引用</div><div class="detail-val">' +
      item.citations.map(c => '<div class="cite-line">(' + esc(c.src) + ' p' + c.page + ') 「' + esc(c.quote) + '」</div>').join('') + '</div></div>';
  }}
  if (item.uncertainties.length) {{
    html += '<div class="detail-row"><div class="detail-label">不確信</div><div class="detail-val">' +
      item.uncertainties.map(u => '<div class="unc-line">(' + esc(u.src) + ') ' + esc(u.reason) + '</div>').join('') + '</div></div>';
  }}
  document.getElementById('detailBody').innerHTML = html;
  document.getElementById('detailPanel').classList.add('open');
}}

function closeDetail() {{
  document.getElementById('detailPanel').classList.remove('open');
}}

function applyZoom() {{
  const scale = zoomPct / 100;
  document.querySelectorAll('.img-wrapper').forEach(w => {{
    w.style.transform = 'scale(' + scale + ')';
  }});
  document.getElementById('zoomLevel').textContent = zoomPct + '%';
}}

function zoom(dir) {{
  zoomPct = Math.max(50, Math.min(300, zoomPct + dir * 25));
  applyZoom();
}}

function zoomReset() {{
  zoomPct = 100;
  applyZoom();
}}

function sortBy(mode) {{
  currentSort = mode;
  document.querySelectorAll('.sort-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.sort-btn').forEach(b => {{
    if (b.getAttribute('onclick') === "sortBy('" + mode + "')") b.classList.add('active');
  }});
  renderList();
}}

function filterBy(mode) {{
  currentFilter = mode;
  document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.filter-btn').forEach(b => {{
    if (b.getAttribute('onclick') === "filterBy('" + mode + "')") b.classList.add('active');
  }});
  renderList();
}}

renderDocTabs();
renderViewer();
goPage(1);
renderList();
</script>
</body></html>"""

# =====================================================================
# 状態保存・復元（handle_edit 用）
# =====================================================================

def _save_state(
    merged: dict[str, str],
    sources: dict[str, str],
    extractions: dict[str, dict[str, str]],
    all_citations: list[dict],
    all_uncertainties: list[dict],
    page_coords: dict,
    source_pages: dict[str, list[str]],
    excel_path: str,
    pdf_map: dict[str, str],
    adopted_page_map: dict[str, dict[str, int]] | None = None,
) -> None:
    os.makedirs("tmp", exist_ok=True)
    state = {
        "merged": merged,
        "sources": sources,
        "extractions": extractions,
        "all_citations": all_citations,
        "all_uncertainties": all_uncertainties,
        "page_coords": {
            dt: {str(p): cs for p, cs in pages.items()}
            for dt, pages in page_coords.items()
        },
        "source_pages": source_pages,
        "excel_path": excel_path,
        "pdf_map": pdf_map,
        "adopted_page_map": adopted_page_map,
    }
    with open(STATE_PATH, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, default=str)
    _log("state_saved", path=STATE_PATH)


def _load_state() -> dict | None:
    if not os.path.exists(STATE_PATH):
        _log("state_not_found", path=STATE_PATH)
        return None
    with open(STATE_PATH, encoding="utf-8") as f:
        state = json.load(f)
    page_coords_raw = state.get("page_coords", {})
    state["page_coords"] = {
        dt: {int(p): cs for p, cs in pages.items()}
        for dt, pages in page_coords_raw.items()
    }
    _log("state_loaded", path=STATE_PATH, keys=list(state.get("merged", {}).keys())[:5])
    return state


async def handle_edit(edits: dict[str, str]) -> tuple[dict, dict[str, str], str]:
    """前回の抽出結果を復元し、edits で上書き → Excel/HTML を再生成。
    戻り値: (diff, new_merged, output_path)
    """
    _log("handle_edit_start", edits=edits)
    state = _load_state()
    if state is None:
        raise RuntimeError("前回の実行結果がありません。先に main() を実行してください")

    old_merged = state["merged"]
    sources = state["sources"]
    extractions = state["extractions"]
    all_citations = state["all_citations"]
    all_uncertainties = state["all_uncertainties"]
    page_coords = state["page_coords"]
    source_pages = state["source_pages"]
    excel_path = state["excel_path"]
    pdf_map = state["pdf_map"]
    adopted_page_map = state.get("adopted_page_map")

    new_merged = {**old_merged}
    for key, value in edits.items():
        if key not in ITEMS:
            _log("handle_edit_unknown_key", key=key)
            continue
        if value is None:
            _log("handle_edit_clear", key=key, old=old_merged.get(key))
            new_merged[key] = ""
        else:
            _log("handle_edit_apply", key=key, old=old_merged.get(key), new=value)
            new_merged[key] = str(value)

    os.makedirs("output", exist_ok=True)
    out_path = fill_excel(new_merged, excel_path)
    _log("handle_edit_excel_done", path=out_path)
    await file_output(out_path)

    html_content = build_review_html(
        new_merged, extractions, sources, all_citations, all_uncertainties,
        source_pages=source_pages,
        page_coords=page_coords,
        adopted_page_map=adopted_page_map,
    )
    lender = new_merged.get("貸主氏名", "") or "貸主名不明"
    html_path = f"output/review_{lender}.html"
    Path(html_path).write_text(html_content, encoding="utf-8")
    await file_output(html_path)
    _log("handle_edit_html_done", path=html_path)

    _save_state(
        new_merged, sources, extractions, all_citations, all_uncertainties,
        page_coords, source_pages, excel_path, pdf_map, adopted_page_map,
    )

    diff = {}
    for key in set(old_merged) | set(new_merged):
        if old_merged.get(key) != new_merged.get(key):
            diff[key] = {"before": old_merged.get(key), "after": new_merged.get(key)}
    _log("handle_edit_done", diff_count=len(diff), diff_keys=list(diff.keys()))
    return diff, new_merged, out_path


# =====================================================================
# メイン
# =====================================================================

async def main(pdf_sources_arg: dict[str, str] | None = None) -> dict | None:
    _pdf_sources: dict[str, str] = pdf_sources_arg or globals().get("pdf_sources", {})
    if not _pdf_sources:
        _log("error", message="pdf_sources が定義されていません")
        return

    _log("main_pdf_sources_received", count=len(_pdf_sources), keys=list(_pdf_sources.keys()))

    pdf_map: dict[str, str] = {}
    for doc_type, path in _pdf_sources.items():
        if doc_type in DOC_LABELS:
            pdf_map[doc_type] = path
            _log("file_mapped", doc_type=doc_type, path=path)
        else:
            _log("unknown_doc_type", doc_type=doc_type, path=path)

    if not pdf_map:
        _log("error", message="有効な書類が pdf_sources に含まれていません", received=list(_pdf_sources.keys()))
        return

    # ------------------------------------------------------------------
    _log("step1_start", docs=list(pdf_map.keys()))
    os.makedirs("tmp/pages", exist_ok=True)

    source_pages = await asyncio.to_thread(render_all_pdfs, pdf_map)
    total_images = sum(len(v) for v in source_pages.values())
    _log("step1_render_done", total_images=total_images, docs={k: len(v) for k, v in source_pages.items()})

    sem = asyncio.Semaphore(MAX_CONCURRENT)
    raw_results, page_coords, adopted_page_map = await extract_all_pages(source_pages, sem)
    _log("step1_extract_done", doc_types=list(raw_results.keys()),
         coord_pages={dt: list(p.keys()) for dt, p in page_coords.items()})

    extractions: dict[str, dict[str, str]] = {}
    all_citations: list[dict] = []
    all_uncertainties: list[dict] = []
    for dt, (acc, cits, uncs) in raw_results.items():
        extractions[dt] = acc
        all_citations.extend(cits)
        all_uncertainties.extend(uncs)

    # ------------------------------------------------------------------
    _log("step2_start")
    merged, sources, integration_uncertainties = await integrate_with_llm(extractions, all_citations)
    all_uncertainties.extend(integration_uncertainties)
    _log("step2_integrate_done", sources_count=len(sources),
         integration_uncertainties=len(integration_uncertainties))

    # 統合AIが空にした項目をソースから補完
    all_doc_types_ordered = [PRIORITY_FIRST] + PRIORITY_EQUAL
    for item in ITEMS:
        if is_empty(merged.get(item, "")):
            for doc_type in all_doc_types_ordered:
                source_val = extractions.get(doc_type, {}).get(item, "")
                if not is_empty(source_val):
                    merged[item] = source_val
                    sources[item] = doc_type
                    _log("fallback_applied", item=item, source=doc_type, value=source_val)
                    break

    await save_csv(merged)

    empty_items = [k for k, v in merged.items() if is_empty(v)]
    _log("step2_done", empty_count=len(empty_items), empty_items=empty_items)

    # ------------------------------------------------------------------
    _log("step3_start")
    excel_path = ""
    for search_dir in ("rag", "uploads"):
        try:
            candidates = [f for f in os.listdir(search_dir) if "daicho" in f.lower() and f.endswith(".xlsx")]
        except FileNotFoundError:
            continue
        if candidates:
            excel_path = f"{search_dir}/{candidates[0]}"
            break
    if not excel_path:
        _log("error", message="daicho_format.xlsx が rag/ にも uploads/ にも見つかりません")
        return
    _log("step3_template_found", excel_path=excel_path)

    os.makedirs("output", exist_ok=True)
    out_path = fill_excel(merged, excel_path)
    await file_output(out_path)
    _log("step3_done", output=out_path)

    # ------------------------------------------------------------------
    _log("step4_start")

    full_log = {
        "timestamp": datetime.now().isoformat(),
        "extract_model": EXTRACT_MODEL,
        "integrate_model": INTEGRATE_MODEL,
        "sources_found": list(extractions.keys()),
        "extraction_raw": extractions,
        "merged": merged,
        "adopted_sources": sources,
        "all_citations": all_citations,
        "all_uncertainties": all_uncertainties,
        "page_coordinates": page_coords,
        "empty_items": empty_items,
        "process_log": _log_entries,
    }
    log_path = "output/full_log.json"
    Path(log_path).write_text(
        json.dumps(full_log, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    await file_output(log_path)

    html_content = build_review_html(
        merged, extractions, sources, all_citations, all_uncertainties,
        source_pages=source_pages,
        page_coords=page_coords,
        adopted_page_map=adopted_page_map,
    )
    lender = merged.get("貸主氏名", "") or "貸主名不明"
    html_path = f"output/review_{lender}.html"
    Path(html_path).write_text(html_content, encoding="utf-8")
    await file_output(html_path)
    _log("step4_done", html=html_path, log=log_path)

    # 状態保存（handle_edit 用）
    _save_state(
        merged, sources, extractions, all_citations, all_uncertainties,
        page_coords, source_pages, excel_path, pdf_map, adopted_page_map,
    )

    confirm_items = [
        {"key": u["key"], "value": merged.get(u["key"], ""), "reason": u.get("reason", "")}
        for u in all_uncertainties if u.get("key")
    ]
    return {
        "lender": lender,
        "ok": len(ITEMS) - len(empty_items) - len(confirm_items),
        "confirm": len(confirm_items),
        "empty": len(empty_items),
        "empty_items": empty_items,
        "confirm_items": confirm_items,
    }


if __name__ == "__main__":
    asyncio.run(main())
