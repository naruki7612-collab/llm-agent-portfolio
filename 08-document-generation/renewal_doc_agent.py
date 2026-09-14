"""
更新書類作成エージェント v2
main.py の完全代替。

変更点:
  - extract / compute / write を明確に分離
  - print → 構造化ログ（JSON）。ファイル出力はオプション
  - 抽出項目にページ出典を付与
  - HITL レビュー HTML（PDFページ画像埋め込み）を自動生成

使い方（実行環境上）:
  from main_ver2 import main
  info, vals, output_path = await main(CHANGES)
"""

import asyncio
import base64
import json
import os
import re
import shutil
import math
import calendar
import datetime
import time

import fitz
import openpyxl
from openpyxl.utils import coordinate_to_tuple
from copy import copy
from agent_sdk import llm_call


# ──────────────────────────────────────────
# VALID_KEYS（main / handle_edit 共有）
# ──────────────────────────────────────────

VALID_KEYS = {
    "new_rent", "rent_up_amount", "rent_up_pct", "rent_rate",
    "new_mgmt", "mgmt_up_amount", "mgmt_up_pct", "mgmt_rate",
    "new_parking",
    "new_deposit", "deposit_up_amount", "deposit_up_pct", "deposit_rate",
    "other_fees", "kanri_itaku", "kanri_itaku_pct",
    "koshinryo_months", "fire_insurance_fee",
    "new_start", "new_end",
    "kashunushi_name", "kashunushi_address", "kariunushi_name",
    "kashunushi_is_corporate", "kariunushi_is_corporate",
    "bukken_name", "goshitsu",
    "shiharai_kijitsu",
    "fire_insurance_type",
    "kurashiido",
    "kashunushi_fees",
    "keigo",
}


# ──────────────────────────────────────────
# ログ
# ──────────────────────────────────────────

_log_entries: list[dict] = []


def _log(level: str, event: str, **data):
    entry = {"level": level, "event": event, **data}
    _log_entries.append(entry)


def _save_log(path: str):
    def _default(o):
        if isinstance(o, (datetime.date, datetime.datetime)):
            return o.isoformat()
        return str(o)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(_log_entries, f, ensure_ascii=False, indent=2, default=_default)


# ──────────────────────────────────────────
# 抽出（Extract）
# ──────────────────────────────────────────

CONTRACT_SCHEMA: dict = {
    "type": "object",
    "properties": {
        "bukken_name":        {"type": "string", "description": "物件名"},
        "bukken_address":     {"type": "string", "description": "建物所在地（住居表示または登記簿所在地）"},
        "goshitsu":           {"type": "string", "nullable": True, "description": "号室（戸建て等で号室がない場合は null）"},
        "kashunushi_name":    {"type": "string", "description": "貸主（甲）の名前（頭書5「貸主及び管理業者」欄または末尾署名欄）"},
        "kashunushi_is_corporate": {"type": "boolean", "nullable": True, "description": "貸主が法人（会社・団体／株式会社・合同会社等の法人格あり、または代表取締役の記載あり）なら true、個人なら false。このページに貸主名が無い場合のみ null。貸主名を抽出したら必ず true/false のどちらかを返すこと（不明禁止）"},
        "kashunushi_address": {"type": "string", "description": "貸主（甲）の住所（頭書5「貸主及び管理業者」欄または末尾署名欄の住所）"},
        "kariunushi_name":    {"type": "string", "description": "借主名"},
        "kariunushi_is_corporate": {"type": "boolean", "nullable": True, "description": "借主が法人（会社・団体／株式会社・合同会社等の法人格あり、または代表取締役の記載あり）なら true、個人なら false。このページに借主名が無い場合のみ null。借主名を抽出したら必ず true/false のどちらかを返すこと（不明禁止）"},
        "current_start":      {"type": "string", "description": "現行契約の開始日（契約書の表記のまま。例: 令和3年10月9日）"},
        "current_end":        {"type": "string", "description": "現行契約の終了日（契約書の表記のまま。例: 令和5年10月8日）"},
        "current_rent":       {"type": "integer", "nullable": True, "description": "現行賃料（数値）"},
        "current_mgmt":       {"type": "integer", "nullable": True, "description": "現行管理費・共益費（数値）"},
        "current_parking":    {"type": "integer", "nullable": True, "description": "現行駐車場代（数値、記載なければ null）"},
        "current_deposit":    {"type": "integer", "nullable": True, "description": "現行敷金（数値）"},
        "koshinryo_months":   {"type": "number", "nullable": True, "description": "更新料の月数（例: 1ヶ月分なら 1、0.5ヶ月分なら 0.5。記載なければ null）"},
        "fire_insurance_fee": {"type": "integer", "nullable": True, "description": "火災保険料（数値、記載なければ null）"},
        "furigana":           {"type": "string", "nullable": True, "description": "借主のフリガナ（記載なければ null）"},
        "birthday":           {"type": "string", "nullable": True, "description": "借主の生年月日（記載のまま。例: 平成2年3月5日。記載なければ null）"},
        "address":            {"type": "string", "nullable": True, "description": "借主の住所（記載なければ null）"},
        "phone":              {"type": "string", "nullable": True, "description": "借主の携帯番号（記載なければ null）"},
        "workplace":          {"type": "string", "nullable": True, "description": "借主の勤務先名称（記載なければ null）"},
        "emergency_contact":  {"type": "string", "nullable": True, "description": "緊急連絡先（記載なければ null）"},
        "signing_reiwa_year":  {"type": "integer", "nullable": True, "description": "書類の署名日・効力発生日の令和年（電子署名日＞契約開始日の順で優先。記載なければ null）"},
        "signing_reiwa_month": {"type": "integer", "nullable": True, "description": "署名日・効力発生日の月"},
        "signing_reiwa_day":   {"type": "integer", "nullable": True, "description": "署名日・効力発生日の日"},
    },
    "required": [
        "bukken_name", "bukken_address", "goshitsu",
        "kashunushi_name", "kashunushi_is_corporate",
        "kashunushi_address", "kariunushi_name", "kariunushi_is_corporate",
        "current_start", "current_end",
        "current_rent", "current_mgmt", "current_parking", "current_deposit",
        "koshinryo_months", "fire_insurance_fee",
        "furigana", "birthday", "address", "phone", "workplace", "emergency_contact",
        "signing_reiwa_year", "signing_reiwa_month", "signing_reiwa_day",
    ],
}

ITEM_DEFS = [
    ("bukken_name",        "物件名",     "text"),
    ("bukken_address",     "所在地",     "text"),
    ("goshitsu",           "号室",       "text"),
    ("kashunushi_name",    "貸主（甲）", "text"),
    ("kashunushi_address", "貸主住所",   "text"),
    ("kariunushi_name",    "借主（乙）", "text"),
    ("current_start",      "契約開始",   "text"),
    ("current_end",        "契約終了",   "text"),
    ("current_rent",       "現行賃料",   "num"),
    ("current_mgmt",       "現行管理費", "num"),
    ("current_parking",    "現行駐車場", "num"),
    ("current_deposit",    "現行敷金",   "num"),
    ("koshinryo_months",   "更新料月数", "text"),
    ("fire_insurance_fee", "火災保険料", "num"),
    ("furigana",           "フリガナ",   "text"),
    ("birthday",           "生年月日",   "text"),
    ("address",            "住所",       "text"),
    ("phone",              "携帯番号",   "text"),
    ("workplace",          "勤務先",     "text"),
    ("emergency_contact",  "緊急連絡先", "text"),
    ("signing_reiwa_year", "署名年",     "text"),
    ("signing_reiwa_month","署名月",     "text"),
    ("signing_reiwa_day",  "署名日",     "text"),
]

_PAGE_EXTRACT_PROMPT = """この画像は賃貸借契約書の{page_num}ページ目です。
このページに記載がある契約情報だけを抽出してください。
このページに記載がない項目は必ず null にしてください（他ページの内容を推測して埋めないこと）。

- 数値は数字のみ（円・カンマなし）、日付は記載のまま（例: 令和3年10月9日）
- 署名日・効力発生日は電子署名日＞契約開始日の順で優先し、令和の年・月・日を数値で返す
- 貸主名・借主名を抽出したら、その法人/個人を表す *_is_corporate（true=法人 / false=個人）も必ずセットで返すこと。
  「不明」は禁止。法人か個人かを必ず true/false で判定する。
  ※ そのページに当該の名前自体が無い場合のみ、名前も *_is_corporate も null にする
- _coordinates_json: 空でない値を返した全項目について、その値が書かれている画像上の位置を
  バウンディングボックスで返すこと。box_2d は [ymin, xmin, ymax, xmax] の4要素配列。
  各値は 0〜1000 の正規化座標（ymin/xmin=左上, ymax/xmax=右下）。null の項目は含めない。省略禁止。
  → [{{"key":"current_rent","box_2d":[412,180,430,265]}},{{"key":"kariunushi_name","box_2d":[120,300,142,410]}}]
"""

# 座標は「JSON配列の文字列」で受け取る（ネスト配列を安定生成させるため。台帳と同方式）
_COORDS_PROP = {
    "type": "string",
    "nullable": True,
    "description": (
        'JSON配列の文字列。形式: [{"key":"項目名","box_2d":[ymin,xmin,ymax,xmax]}]。'
        'box_2d は画像上の正規化座標（0〜1000）。ymin/xmin=左上, ymax/xmax=右下。'
        '値がnullの項目は含めない。'
    ),
}

# 1ページ単位の抽出用スキーマ（全項目 nullable: そのページに記載がなければ null）
_PAGE_EXTRACT_SCHEMA = {
    "type": "object",
    "properties": {
        **{k: {**v, "nullable": True} for k, v in CONTRACT_SCHEMA["properties"].items()},
        "_coordinates_json": _COORDS_PROP,
    },
    "required": list(CONTRACT_SCHEMA["properties"].keys()) + ["_coordinates_json"],
}


def _parse_coords_json(raw) -> list[dict]:
    """_coordinates_json（JSON文字列）を [{key, box_2d}] に正規化する。壊れていても落とさない。"""
    if not raw:
        return []
    if isinstance(raw, list):
        items = raw
    else:
        try:
            items = json.loads(raw)
        except Exception:
            return []
        if not isinstance(items, list):
            return []
    out = []
    for c in items:
        if not isinstance(c, dict):
            continue
        key, box = c.get("key"), c.get("box_2d")
        if key and isinstance(box, list) and len(box) >= 4:
            try:
                out.append({"key": key, "box_2d": [int(v) for v in box[:4]]})
            except (TypeError, ValueError):
                continue
    return out


def _pdf_to_images(pdf_path: str) -> list[str]:
    work_dir = "tmp/pdf_pages"
    os.makedirs(work_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    paths: list[str] = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
        out_path = f"{work_dir}/page_{i + 1:03d}.png"
        pix.save(out_path)
        paths.append(out_path)
    doc.close()
    _log("info", "pdf_to_images", pdf=pdf_path, pages=len(paths))
    return paths


async def _extract_page(page_num: int, image_path: str, max_retries: int = 3) -> tuple[int, dict | None]:
    """1ページの画像から構造化抽出する。(page_num, data|None) を返す。"""
    for attempt in range(1, max_retries + 1):
        resp = await llm_call(
            prompt=_PAGE_EXTRACT_PROMPT.format(page_num=page_num),
            image_path=image_path,
            schema=_PAGE_EXTRACT_SCHEMA,
            model="gemini/gemini-3.5-flash",
        )
        if isinstance(resp, dict) and "error" in resp:
            _log("warn", "page_extract_retry", page=page_num, attempt=attempt, error=resp["error"])
            if attempt < max_retries:
                await asyncio.sleep(2 * attempt)
                continue
            _log("error", "page_extract_failed", page=page_num)
            return page_num, None
        data = resp.get("data") if isinstance(resp, dict) else None
        found = [k for k in (data or {}) if data.get(k) is not None]
        _log("info", "page_extract_result", page=page_num, found_keys=found)
        return page_num, data
    return page_num, None


async def _fill_furigana(info: dict):
    """フリガナが契約書から取れなかった場合、借主名から LLM で補完する。"""
    if info.get("furigana") or not info.get("kariunushi_name"):
        return
    name = info["kariunushi_name"]
    prompt = (
        "次の氏名のフリガナ（全角カタカナ）だけを返してください。\n"
        "姓と名の間は全角スペース1つ。フリガナ以外は一切出力しないこと。\n\n"
        f"氏名: {name}"
    )
    try:
        resp = await llm_call(prompt, model="gemini/gemini-3.5-flash")
        furigana = resp.get("text", "").strip() if isinstance(resp, dict) else ""
        if furigana and furigana != "None":
            info["furigana"] = furigana
            _log("info", "furigana_filled", name=name, furigana=furigana)
            return
        _log("warn", "furigana_fill_failed", name=name, resp=str(resp)[:200])
    except Exception as e:
        _log("error", "furigana_fill_error", name=name, error=str(e))


async def extract(pdf_path: str, *, fill_furigana: bool = True) -> dict:
    """各ページ画像から構造化抽出し、ページ出典つきの info を返す。
    各値の出典ページ = その値を抽出した画像のページ番号（機械的に確定）。
    """
    page_paths = _pdf_to_images(pdf_path)
    _log("info", "extract_start", pages=len(page_paths))

    sem = asyncio.Semaphore(16)

    async def _with_limit(page_num, path):
        async with sem:
            return await _extract_page(page_num, path)

    tasks = [_with_limit(i + 1, p) for i, p in enumerate(page_paths)]
    results = await asyncio.gather(*tasks)

    # ページ順にマージ: 各キーは最初に非null値を返したページを採用し、そのページを出典とする
    # 名前と法人フラグは「同一ページから原子的に」確定させる（別ページの値が混入しないように）
    NAME_FLAG_PAIRS = {
        "kariunushi_name": "kariunushi_is_corporate",
        "kashunushi_name": "kashunushi_is_corporate",
    }
    info: dict = {}
    page_sources: dict[str, int | None] = {key: None for key, _, _ in ITEM_DEFS}
    # 座標は抽出と同時に取得済み（_coordinates_json）。ページ→候補リストで保持する。
    page_coords: dict[int, list[dict]] = {}
    for page_num, data in sorted(results, key=lambda r: r[0]):
        if not data:
            continue
        coords = _parse_coords_json(data.get("_coordinates_json"))
        if coords:
            page_coords[page_num] = coords
        for key, _, _ in ITEM_DEFS:
            if info.get(key) is None and data.get(key) is not None:
                info[key] = data[key]
                page_sources[key] = page_num
                # 名前を採用したら、その法人フラグも同じpage_numのデータから採用する。
                # （名前のページに flag が無ければ None。別ページの flag は拾わない＝セット保証）
                if key in NAME_FLAG_PAIRS:
                    info[NAME_FLAG_PAIRS[key]] = data.get(NAME_FLAG_PAIRS[key])

    for key, _, _ in ITEM_DEFS:
        info.setdefault(key, None)
    for flag_key in NAME_FLAG_PAIRS.values():
        info.setdefault(flag_key, None)

    if all(info.get(k) is None for k, _, _ in ITEM_DEFS):
        raise ValueError("PDF から内容を読み取れませんでした")

    # フリガナが契約書になければ借主名から補完（出典ページなし＝生成値）
    # 補足資料（extra_pdf_paths）とマージする場合は、実物のフリガナを優先するため
    # マージ完了後に main() 側で1回だけ実行する（fill_furigana=False で抑止）
    if fill_furigana:
        await _fill_furigana(info)

    info["_page_sources"] = page_sources
    info["_page_coords"] = page_coords
    # 採用ページの座標を _boxes[key] に展開（既存のハイライト表示との後方互換）
    boxes: dict[str, list[float]] = {}
    for key, _, _ in ITEM_DEFS:
        pg = page_sources.get(key)
        if not pg:
            continue
        for c in page_coords.get(pg, []):
            if c["key"] == key:
                ymin, xmin, ymax, xmax = c["box_2d"]
                boxes[key] = [xmin / 10, ymin / 10, xmax / 10, ymax / 10]
                break
    info["_boxes"] = boxes
    _log("info", "extract_page_sources",
         assigned={k: v for k, v in page_sources.items() if v is not None},
         unassigned=[k for k, v in page_sources.items() if v is None])
    _log("info", "extract_coords",
         pages=sorted(page_coords), total=sum(len(v) for v in page_coords.values()),
         boxes=len(boxes))
    _log("info", "extract_done", info=info)
    return info


def _merge_extra_info(info: dict, extra_info: dict, source: str) -> list[str]:
    """補足資料（申込書・請求書等）の抽出結果から、契約書で取れなかった項目だけ補完する。
    - 既に値がある項目は上書きしない（契約書優先）
    - 名前(kariunushi/kashunushi)を採用する場合は、法人フラグも同じ資料から原子的に採用
    - 補完した項目の出典ページは None のまま（レビューHTMLのページは契約書のみのため）
    採用したキーのリストを返す。
    """
    NAME_FLAG_PAIRS = {
        "kariunushi_name": "kariunushi_is_corporate",
        "kashunushi_name": "kashunushi_is_corporate",
    }
    adopted: list[str] = []
    for key, _, _ in ITEM_DEFS:
        if info.get(key) is None and extra_info.get(key) is not None:
            info[key] = extra_info[key]
            adopted.append(key)
            # 補完元の資料名を記録（レビューHTMLで出典バッジとして表示する）
            info.setdefault("_extra_sources", {})[key] = os.path.basename(source)
            if key in NAME_FLAG_PAIRS:
                info[NAME_FLAG_PAIRS[key]] = extra_info.get(NAME_FLAG_PAIRS[key])
    _log("info", "merge_extra_info", source=source, adopted=adopted)
    return adopted


# ──────────────────────────────────────────
# 座標取得（Gemini box_2d）
# ──────────────────────────────────────────

_BOX_SCHEMA = {
    "type": "object",
    "properties": {
        "results": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "key": {"type": "string"},
                    "box_2d": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "minItems": 4, "maxItems": 4,
                    },
                },
                "required": ["key", "box_2d"],
            },
        }
    },
    "required": ["results"],
}


async def _locate_boxes(info: dict, pdf_path: str):
    """抽出済みの値について、Gemini box_2d で原本上の座標を取得し info["_boxes"] に格納する。"""
    page_sources = info.get("_page_sources", {})
    by_page: dict[int, list[tuple[str, str, str]]] = {}
    for key, label, _ in ITEM_DEFS:
        pg = page_sources.get(key)
        val = info.get(key)
        if pg is None or val is None:
            continue
        by_page.setdefault(pg, []).append((key, label, str(val)))

    if not by_page:
        _log("info", "locate_boxes_skip", reason="no page sources")
        info["_boxes"] = {}
        return

    _log("info", "locate_boxes_start",
         pages=list(by_page.keys()),
         items_per_page={pg: len(items) for pg, items in by_page.items()})
    page_paths = _pdf_to_images(pdf_path)

    async def _locate_page(pg, items_on_page):
        if pg < 1 or pg > len(page_paths):
            return []
        img_path = page_paths[pg - 1]
        if not img_path:
            return []
        item_lines = "\n".join(f"- key={k}, label={lb}, value={v}" for k, lb, v in items_on_page)
        prompt = (
            f"画像は賃貸借契約書の{pg}ページ目です。\n"
            f"以下の各項目の値が画像上のどこに書かれているか、バウンディングボックスで返してください。\n"
            f"全ての項目について必ず box_2d を返すこと。省略禁止。\n\n"
            f"{item_lines}\n\n"
            f"box_2d は [ymin, xmin, ymax, xmax] の4要素配列。各値は 0〜1000 の正規化座標。"
        )
        try:
            resp = await llm_call(prompt, image_path=img_path, schema=_BOX_SCHEMA,
                                  model="gemini/gemini-3.5-flash")
            results = resp.get("data", {}).get("results", [])
            _log("info", "locate_boxes_page", page=pg,
                 requested=len(items_on_page), returned=len(results),
                 keys=[r.get("key") for r in results])
            return results
        except Exception as e:
            _log("error", "locate_boxes_page_error", page=pg, error=str(e))
            return []

    tasks = [_locate_page(pg, items) for pg, items in by_page.items()]
    results = await asyncio.gather(*tasks)

    boxes: dict[str, list[float]] = {}
    for result_list in results:
        for r in result_list:
            key = r.get("key")
            box = r.get("box_2d")
            if key and box and len(box) == 4:
                ymin, xmin, ymax, xmax = box
                boxes[key] = [xmin / 10, ymin / 10, xmax / 10, ymax / 10]

    info["_boxes"] = boxes
    _log("info", "locate_boxes_done", found=len(boxes), total=sum(len(v) for v in by_page.values()))


# ──────────────────────────────────────────
# 計算（Compute）
# ──────────────────────────────────────────

def _reiwa_to_date(s: str | None) -> datetime.date | None:
    if not s:
        return None
    m = re.search(r"令和\s*(\d+)\s*年\s*(\d+)\s*月\s*(\d+)\s*日", s)
    if m:
        return datetime.date(int(m.group(1)) + 2018, int(m.group(2)), int(m.group(3)))
    m = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", s)
    if m:
        return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def _date_to_reiwa(d: datetime.date | None) -> str | None:
    if not d:
        return None
    return f"令和{d.year - 2018}年{d.month}月{d.day}日"


def _add_months(d: datetime.date, months: int) -> datetime.date:
    total = d.month - 1 + months
    year = d.year + total // 12
    month = total % 12 + 1
    day = min(d.day, calendar.monthrange(year, month)[1])
    return datetime.date(year, month, day)


def _derive_new_period(start, end):
    if not start or not end:
        return None, None
    new_start = end + datetime.timedelta(days=1)
    months = (new_start.year - start.year) * 12 + (new_start.month - start.month)
    if new_start.day < start.day:
        months -= 1
    new_end = _add_months(new_start, months) - datetime.timedelta(days=1)
    return new_start, new_end


def calc_new_value(current, changes: dict, key: str) -> int:
    alias_map = {
        "rent_multiplier": ("rent", "rate"),
        "rent_times":      ("rent", "rate"),
        "mgmt_multiplier": ("mgmt", "rate"),
        "parking_fee":     ("parking", "new"),
    }
    for alias, (akey, atype) in alias_map.items():
        if alias in changes and akey == key:
            canon = f"new_{akey}" if atype == "new" else f"{akey}_rate"
            changes = {**changes, canon: changes[alias]}

    if changes.get(f"new_{key}") is not None:
        return int(changes[f"new_{key}"])
    elif changes.get(f"{key}_up_amount") is not None:
        return int(current or 0) + int(changes[f"{key}_up_amount"])
    elif changes.get(f"{key}_up_pct") is not None:
        return int((current or 0) * (1 + changes[f"{key}_up_pct"] / 100))
    elif changes.get(f"{key}_rate") is not None:
        return int((current or 0) * changes[f"{key}_rate"])
    else:
        return int(current or 0)


def _compute_tsumimashi(info: dict, vals: dict) -> dict:
    result: dict = {"daily_amount": 0, "daily_desc": None, "deposit_diff": 0}
    current_rent_val = int(info.get("current_rent") or 0)
    rent_diff = vals["new_rent"] - current_rent_val
    start_date = info.get("new_start_date")

    if rent_diff > 0 and start_date:
        days_in_month = calendar.monthrange(start_date.year, start_date.month)[1]
        remaining_days = days_in_month - start_date.day + 1
        result["daily_amount"] = math.ceil(rent_diff * remaining_days / days_in_month)
        end_of_month = datetime.date(start_date.year, start_date.month, days_in_month)
        result["daily_desc"] = (
            f"{start_date.month}月{start_date.day}日"
            f"〜{end_of_month.month}月{end_of_month.day}日　{remaining_days}日分"
        )
        _log("info", "tsumimashi_daily", rent_diff=rent_diff, remaining_days=remaining_days,
             days_in_month=days_in_month, daily_amount=result["daily_amount"])

    current_dep_val = int(info.get("current_deposit") or 0)
    if rent_diff > 0 and current_rent_val > 0 and current_dep_val > 0:
        result["deposit_diff"] = math.ceil(rent_diff * current_dep_val / current_rent_val)
        _log("info", "tsumimashi_deposit", rent_diff=rent_diff, current_deposit=current_dep_val,
             deposit_diff=result["deposit_diff"])
    return result


def compute(info: dict, changes: dict) -> dict:
    """info + changes → vals を計算し、info に期間情報を追加する。"""
    _log("info", "compute_start", changes=changes)
    info["current_start_date"] = _reiwa_to_date(info.get("current_start"))
    info["current_end_date"] = _reiwa_to_date(info.get("current_end"))
    info["new_start_date"], info["new_end_date"] = _derive_new_period(
        info["current_start_date"], info["current_end_date"]
    )
    if changes.get("new_start"):
        info["new_start_date"] = _reiwa_to_date(changes["new_start"])
    if changes.get("new_end"):
        info["new_end_date"] = _reiwa_to_date(changes["new_end"])
    _log("info", "compute_period",
         current_start=info.get("current_start_date"),
         current_end=info.get("current_end_date"),
         new_start=info.get("new_start_date"),
         new_end=info.get("new_end_date"))

    for key in ("current_start", "current_end", "new_start", "new_end"):
        d = info.get(f"{key}_date")
        info[f"{key}_ts"] = (
            int(datetime.datetime(d.year, d.month, d.day, tzinfo=datetime.timezone.utc).timestamp())
            if d else None
        )

    new_rent = calc_new_value(info.get("current_rent"), changes, "rent")
    new_mgmt = calc_new_value(info.get("current_mgmt"), changes, "mgmt")
    _log("info", "compute_rent_mgmt",
         current_rent=info.get("current_rent"), new_rent=new_rent,
         current_mgmt=info.get("current_mgmt"), new_mgmt=new_mgmt)

    deposit_explicitly_set = any(
        changes.get(k) is not None
        for k in ["new_deposit", "deposit_up_amount", "deposit_up_pct", "deposit_rate"]
    )
    if deposit_explicitly_set:
        new_deposit = calc_new_value(info.get("current_deposit"), changes, "deposit")
        _log("info", "compute_deposit", mode="explicit", current=info.get("current_deposit"), new=new_deposit)
    else:
        current_rent_val = int(info.get("current_rent") or 0)
        current_dep_val = int(info.get("current_deposit") or 0)
        rent_diff = new_rent - current_rent_val
        if rent_diff > 0 and current_rent_val > 0 and current_dep_val > 0:
            new_deposit = current_dep_val + math.ceil(rent_diff * current_dep_val / current_rent_val)
            _log("info", "compute_deposit", mode="auto_increase", current=current_dep_val,
                 new=new_deposit, rent_diff=rent_diff)
        else:
            new_deposit = current_dep_val
            _log("info", "compute_deposit", mode="unchanged", current=current_dep_val, new=new_deposit)

    raw_other_fees: list[dict] = changes.get("other_fees", [])
    kanri_from_other = 0
    other_fees: list[dict] = []
    for f in raw_other_fees:
        if "管理委託" in f.get("name", ""):
            kanri_from_other += int(f.get("amount", 0))
        else:
            other_fees.append(f)
    other_fee_total = sum(int(f.get("amount", 0)) for f in other_fees)
    other_fee_label = "・".join(f.get("name", "") for f in other_fees) if other_fees else ""
    if raw_other_fees:
        _log("info", "compute_other_fees", raw=raw_other_fees, kanri_from_other=kanri_from_other,
             other_fees=other_fees, total=other_fee_total)

    if changes.get("koshinryo_months") is not None:
        koshinryo_months = float(changes["koshinryo_months"])
        _log("info", "compute_koshinryo_months", source="changes", value=koshinryo_months)
    else:
        koshinryo_months = info.get("koshinryo_months")
        if not koshinryo_months or koshinryo_months <= 0:
            _log("info", "compute_koshinryo_months", source="default", extracted=info.get("koshinryo_months"))
            koshinryo_months = 1.0
        else:
            _log("info", "compute_koshinryo_months", source="extracted", value=koshinryo_months)

    kashunushi_koshinryo = int(new_rent * 0.5 * 1.1)
    koshin_tesuryo = int(new_rent * 0.25 * 1.1)
    _log("info", "compute_fees", koshinryo=int(new_rent * koshinryo_months),
         koshin_tesuryo=koshin_tesuryo, kashunushi_koshinryo=kashunushi_koshinryo)

    if changes.get("new_parking") is not None:
        new_parking = int(changes["new_parking"])
        _log("info", "compute_parking", source="changes", value=new_parking)
    elif info.get("current_parking") is not None:
        new_parking = int(info["current_parking"])
        _log("info", "compute_parking", source="extracted", value=new_parking)
    else:
        new_parking = None
        _log("info", "compute_parking", source="none")

    monthly_total = new_rent + new_mgmt
    if new_parking is not None:
        monthly_total += new_parking
    _log("info", "compute_monthly_total", total=monthly_total,
         rent=new_rent, mgmt=new_mgmt, parking=new_parking)

    # 支払期日: 明示指定があればそれを優先。なければ「新契約開始日の5日前」を初期値にする
    if changes.get("shiharai_kijitsu") is not None:
        shiharai_kijitsu = changes["shiharai_kijitsu"]
        _log("info", "compute_shiharai_kijitsu", source="changes", value=shiharai_kijitsu)
    else:
        new_start = info.get("new_start_date")
        if new_start:
            shiharai_kijitsu = _date_to_reiwa(new_start - datetime.timedelta(days=5))
            _log("info", "compute_shiharai_kijitsu", source="default_5days_before",
                 new_start=new_start, value=shiharai_kijitsu)
        else:
            shiharai_kijitsu = None
            _log("info", "compute_shiharai_kijitsu", source="none", reason="new_start unknown")

    fire_insurance_type = changes.get("fire_insurance_type")

    if "kurashiido" in changes:
        kurashiido = changes["kurashiido"]
        if kurashiido is not None:
            kurashiido = int(kurashiido)
        _log("info", "compute_kurashiido", source="changes", value=kurashiido)
    else:
        kurashiido = 16500
        _log("info", "compute_kurashiido", source="default", value=kurashiido)

    keigo_val = changes.get("keigo")

    raw_kashunushi_fees = changes.get("kashunushi_fees", [])
    if len(raw_kashunushi_fees) > 3:
        raise ValueError("kashunushi_fees は最大3件までです")
    kashunushi_fees = []
    for fee in raw_kashunushi_fees:
        name = fee.get("name", "")
        amount = int(fee.get("amount", 0))
        if not name:
            raise ValueError("kashunushi_fees の name は必須です")
        if amount < 0:
            raise ValueError("kashunushi_fees の amount は0以上である必要があります")
        kashunushi_fees.append({"name": name, "amount": amount})

    if changes.get("kanri_itaku") is not None:
        kanri_itaku = int(changes["kanri_itaku"])
        _log("info", "compute_kanri_itaku", source="direct", value=kanri_itaku)
    elif kanri_from_other:
        kanri_itaku = kanri_from_other
        _log("info", "compute_kanri_itaku", source="other_fees", value=kanri_itaku)
    elif changes.get("kanri_itaku_pct") is not None:
        kanri_itaku = int((new_rent + new_mgmt) * changes["kanri_itaku_pct"] / 100)
        _log("info", "compute_kanri_itaku", source="pct", pct=changes["kanri_itaku_pct"], value=kanri_itaku)
    else:
        kanri_itaku = 0

    if changes.get("fire_insurance_fee") is not None:
        fire_insurance = int(changes["fire_insurance_fee"])
        _log("info", "compute_fire_insurance", source="changes", value=fire_insurance)
    else:
        fire_insurance = int(info.get("fire_insurance_fee") or 0)
        _log("info", "compute_fire_insurance", source="extracted", value=fire_insurance)

    vals = {
        "new_rent": new_rent, "new_mgmt": new_mgmt, "new_parking": new_parking,
        "new_deposit": new_deposit, "other_fee_total": other_fee_total,
        "other_fee_label": other_fee_label,
        "kashunushi_koshinryo": kashunushi_koshinryo, "koshinryo_months": koshinryo_months,
        "koshin_tesuryo": koshin_tesuryo, "monthly_total": monthly_total,
        "kanri_itaku": kanri_itaku, "fire_insurance": fire_insurance,
        "shiharai_kijitsu": shiharai_kijitsu,
        "fire_insurance_type": fire_insurance_type,
        "kurashiido": kurashiido,
        "keigo": keigo_val,
        "kashunushi_fees": kashunushi_fees,
    }
    _log("info", "compute_done", vals=vals)
    return vals


# ──────────────────────────────────────────
# Excel 書込み（Write）
# ──────────────────────────────────────────

def write(info: dict, vals: dict, template_path: str, output_path: str) -> list[dict]:
    """Excel に書き込み、書込みログを返す。"""
    _log("info", "write_start", template=template_path, output=output_path)
    shutil.copy(template_path, output_path)
    wb = openpyxl.load_workbook(output_path)
    _log("info", "write_sheets", sheets=wb.sheetnames)
    tsumimashi = _compute_tsumimashi(info, vals)

    current_period_str = None
    if info.get("current_start_date") and info.get("current_end_date"):
        current_period_str = f"{_date_to_reiwa(info['current_start_date'])}〜{_date_to_reiwa(info['current_end_date'])}"
    new_period_str = None
    if info.get("new_start_date") and info.get("new_end_date"):
        new_period_str = f"{_date_to_reiwa(info['new_start_date'])}〜{_date_to_reiwa(info['new_end_date'])}"

    write_log: list[dict] = []

    def w(ws, coord, value):
        cell = ws[coord]
        merged = False
        if cell.__class__.__name__ == "MergedCell":
            merged = True
            row, col = coordinate_to_tuple(coord)
            for rng in ws.merged_cells.ranges:
                if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
                    cell = ws.cell(rng.min_row, rng.min_col)
                    break
        existing_font = copy(cell.font)
        cell.value = value
        cell.font = existing_font
        actual = ws[cell.coordinate].value
        ok = actual == value
        write_log.append({
            "sheet": ws.title, "coord": coord, "value": value,
            "ok": ok, "merged": merged,
        })

    def get_sheet(name):
        if name in wb.sheetnames:
            return wb[name]
        # テンプレートのシート名に前後の空白ゆれがあっても拾う
        target = name.strip()
        for sn in wb.sheetnames:
            if sn.strip() == target:
                return wb[sn]
        return None

    # ── 覚書（委任あり・委任なし共通データ書込み） ──
    def _write_oboegaki_common(ws):
        w(ws, "G5", info.get("kashunushi_name"))
        w(ws, "W5", info.get("kariunushi_name"))
        w(ws, "B21", current_period_str)
        w(ws, "X21", new_period_str)
        w(ws, "B25", info.get("current_rent"))
        w(ws, "X25", vals["new_rent"])
        w(ws, "B29", info.get("current_mgmt"))
        w(ws, "X29", vals["new_mgmt"])
        current_parking = info.get("current_parking")
        new_parking_val = vals.get("new_parking")
        if current_parking is not None:
            w(ws, "B33", current_parking)
            w(ws, "X33", new_parking_val if new_parking_val is not None else current_parking)
        w(ws, "B37", info.get("current_deposit"))
        w(ws, "X37", vals["new_deposit"])
        w(ws, "O14", info.get("bukken_address"))
        w(ws, "O15", info.get("bukken_name"))
        goshitsu_af = info.get("goshitsu")
        w(ws, "AF15", re.sub(r"号室?$", "", goshitsu_af) if goshitsu_af else goshitsu_af)
        ry, rm, rd = info.get("signing_reiwa_year"), info.get("signing_reiwa_month"), info.get("signing_reiwa_day")
        if ry is not None and rm is not None and rd is not None:
            w(ws, "AK5", f"{int(ry) + 2018}年{rm}月{rd}日")
        else:
            w(ws, "AK5", None)

    ws1 = get_sheet("覚書")
    if ws1:
        _log("info", "write_sheet", sheet="覚書")
        _write_oboegaki_common(ws1)
        w(ws1, "H50", info.get("kashunushi_address"))
        w(ws1, "H51", info.get("kashunushi_name"))

    ws1b = get_sheet("覚書(委任なし)")
    if ws1b:
        _log("info", "write_sheet", sheet="覚書(委任なし)")
        _write_oboegaki_common(ws1b)
        w(ws1b, "H50", None)
        w(ws1b, "H51", None)
        for cell_ref in ("V49", "V50", "V51", "V52", "V53", "AJ53"):
            w(ws1b, cell_ref, None)

    # 借主の法人フラグ（抽出値）で申込書を出し分ける。法人なら法人、それ以外は個人。
    is_corporate = info.get("kariunushi_is_corporate") is True
    _log("info", "write_application_route", kariunushi_is_corporate=info.get("kariunushi_is_corporate"),
         target="法人" if is_corporate else "個人")

    # ── 更新申込書（個人） ──
    ws2 = get_sheet("更新申込書(個人)")
    if ws2:
        _log("info", "write_sheet", sheet="更新申込書(個人)", filled=not is_corporate)
        w(ws2, "B4", info.get("bukken_name"))
        w(ws2, "H4", info.get("goshitsu"))
        if not is_corporate:
            w(ws2, "C7", info.get("kariunushi_name"))
            w(ws2, "C6", info.get("furigana"))
            w(ws2, "C8", info.get("birthday"))
            w(ws2, "H7", info.get("phone"))

    # ── 更新申込書（法人） ──
    ws3 = get_sheet("更新申込書 (法人)")
    if ws3:
        _log("info", "write_sheet", sheet="更新申込書(法人)", filled=is_corporate)
        w(ws3, "B4", info.get("bukken_name"))
        w(ws3, "H4", info.get("goshitsu"))
        if is_corporate:
            w(ws3, "C7", info.get("kariunushi_name"))
            w(ws3, "C6", info.get("furigana"))

    # ── 請求書 ──
    ws4 = get_sheet("請求書")
    if ws4:
        _log("info", "write_sheet", sheet="請求書")
        w(ws4, "A4", info.get("kariunushi_name"))
        # D4 敬称: 請求書の宛先は借主。手動keigo指定があれば優先、なければ借主の法人フラグで判定
        keigo_val = vals.get("keigo")
        if keigo_val:
            keigo_seikyu = keigo_val
            _log("info", "write_keigo_seikyu", source="manual", value=keigo_seikyu)
        else:
            keigo_seikyu = "御中" if info.get("kariunushi_is_corporate") else "様"
            _log("info", "write_keigo_seikyu", source="is_corporate",
                 is_corporate=info.get("kariunushi_is_corporate"), value=keigo_seikyu)
        w(ws4, "D4", keigo_seikyu)
        w(ws4, "C18", info.get("bukken_name"))
        w(ws4, "H18", info.get("goshitsu"))

        koshinryo_months = vals["koshinryo_months"]
        w(ws4, "H21", int(vals["new_rent"] * koshinryo_months))
        months_str = f"{koshinryo_months:g}" if koshinryo_months != int(koshinryo_months) else str(int(koshinryo_months))
        w(ws4, "D21", f"新賃料の{months_str}ヶ月分")

        w(ws4, "H23", vals["koshin_tesuryo"])

        shiharai_kijitsu = vals.get("shiharai_kijitsu")
        w(ws4, "B13", shiharai_kijitsu)

        fire_type = vals.get("fire_insurance_type")
        fire_fee = vals.get("fire_insurance") or 0
        _log("info", "write_fire_insurance", type=fire_type, fee=fire_fee)
        if fire_type == "A":
            w(ws4, "D25", f"保証会社より賃料と併せて引落し（{fire_fee:,}円）")
        elif fire_type == "B":
            w(ws4, "D25", "火災保険につきましては、保険会社からの案内に従い、別途お手続きをお願いいたします。")
        elif fire_type == "C":
            w(ws4, "D25", f"火災保険料（{fire_fee:,}円）※弊社にて手続きいたします")
        elif fire_type == "D":
            w(ws4, "D25", None)
            w(ws4, "A25", None)

        kurashiido = vals.get("kurashiido")
        _log("info", "write_kurashiido", value=kurashiido, action="write" if kurashiido and kurashiido > 0 else "clear")
        if kurashiido and kurashiido > 0:
            w(ws4, "H27", kurashiido)
        else:
            w(ws4, "D27", None)
            w(ws4, "H27", None)

        inv_row = 29
        daily = tsumimashi["daily_amount"]
        dep = tsumimashi["deposit_diff"]
        _log("info", "write_tsumimashi_invoice", daily=daily, deposit_diff=dep)
        if daily > 0:
            w(ws4, f"A{inv_row}", "積み増し分　日割り賃料")
            w(ws4, f"D{inv_row}", tsumimashi["daily_desc"])
            w(ws4, f"H{inv_row}", daily)
            inv_row += 2
        if dep > 0:
            w(ws4, f"A{inv_row}", "積み増し分　敷金")
            w(ws4, f"D{inv_row}", "賃料増額に伴う増額分")
            w(ws4, f"H{inv_row}", dep)

    # ── 家主精算書 ──
    ws5 = get_sheet("家主精算書")
    if ws5:
        _log("info", "write_sheet", sheet="家主精算書")
        kashunushi_name = info.get("kashunushi_name") or ""
        w(ws5, "A9", kashunushi_name)
        keigo_val = vals.get("keigo")
        if keigo_val:
            keigo = keigo_val
            _log("info", "write_keigo_yanushi", source="manual", value=keigo)
        else:
            keigo = "御中" if info.get("kashunushi_is_corporate") else "様"
            _log("info", "write_keigo_yanushi", source="is_corporate",
                 is_corporate=info.get("kashunushi_is_corporate"), value=keigo)
        w(ws5, "D9", keigo)
        w(ws5, "B13", info.get("bukken_name"))
        w(ws5, "D15", info.get("goshitsu"))
        w(ws5, "B18", current_period_str)
        w(ws5, "B21", vals["new_rent"])
        new_parking_val = vals.get("new_parking")
        w(ws5, "C21", vals["new_mgmt"] + (new_parking_val or 0))
        w(ws5, "E21", vals["monthly_total"])
        w(ws5, "G21", vals["kanri_itaku"])

        koshinryo_months = vals["koshinryo_months"]
        koshinryo_shakuchu = int(vals["new_rent"] * koshinryo_months)
        w(ws5, "G26", koshinryo_shakuchu)
        months_str = f"{koshinryo_months:g}" if koshinryo_months != int(koshinryo_months) else str(int(koshinryo_months))
        w(ws5, "D26", f"新賃料の{months_str}ヶ月分")
        shakuchu_total = koshinryo_shakuchu

        next_row = 27
        if tsumimashi["daily_amount"] > 0:
            w(ws5, f"A{next_row}", "積み増し分　日割り賃料")
            w(ws5, f"D{next_row}", tsumimashi["daily_desc"])
            w(ws5, f"G{next_row}", tsumimashi["daily_amount"])
            shakuchu_total += tsumimashi["daily_amount"]
            next_row += 1
        if tsumimashi["deposit_diff"] > 0:
            w(ws5, f"A{next_row}", "積み増し分　敷金")
            w(ws5, f"D{next_row}", "賃料増額に伴う増額分")
            w(ws5, f"G{next_row}", tsumimashi["deposit_diff"])
            shakuchu_total += tsumimashi["deposit_diff"]

        w(ws5, "G31", shakuchu_total)
        w(ws5, "G35", vals["kashunushi_koshinryo"])
        kashunushi_total = vals["kashunushi_koshinryo"]

        kashunushi_fees = vals.get("kashunushi_fees") or []
        for i, fee in enumerate(kashunushi_fees[:3]):
            row = 36 + i
            w(ws5, f"D{row}", fee["name"])
            w(ws5, f"G{row}", fee["amount"])
            kashunushi_total += fee["amount"]

        w(ws5, "G39", kashunushi_total)
        w(ws5, "G41", shakuchu_total - kashunushi_total)
        _log("info", "write_yanushi_totals",
             shakuchu_total=shakuchu_total, kashunushi_total=kashunushi_total,
             remittance=shakuchu_total - kashunushi_total)

    # ── 更新案内（フレックス） ──
    ws6 = get_sheet("更新案内 (フレックス）")
    if ws6:
        _log("info", "write_sheet", sheet="更新案内(フレックス)")
        goshitsu = info.get("goshitsu")
        w(ws6, "A4", info.get("kariunushi_name"))
        if goshitsu:
            w(ws6, "B10", f"{info.get('bukken_name', '')} {goshitsu}　更新の件")
        else:
            w(ws6, "B10", f"{info.get('bukken_name', '')}　更新の件")
        w(ws6, "A13", info.get("bukken_name"))
        w(ws6, "D13", goshitsu if goshitsu else None)
        w(ws6, "D3", None)
        keigo_val = vals.get("keigo")
        if keigo_val:
            keigo_annai = keigo_val
            _log("info", "write_keigo_annai", source="manual", value=keigo_annai)
        else:
            keigo_annai = "御中" if info.get("kariunushi_is_corporate") else "様"
            _log("info", "write_keigo_annai", source="is_corporate",
                 is_corporate=info.get("kariunushi_is_corporate"), value=keigo_annai)
        w(ws6, "D4", keigo_annai)

    # H/I 列の文字列を数値に変換
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.__class__.__name__ == "MergedCell":
                    continue
                if cell.column in (8, 9) and isinstance(cell.value, str) and not cell.value.startswith("="):
                    m = re.fullmatch(r"[\d,]+", cell.value.strip())
                    if m:
                        old_val = cell.value
                        cell.value = int(m.group(0).replace(",", ""))
                        _log("warn", "auto_numericize", sheet=ws.title, coord=cell.coordinate,
                             old=old_val, new=cell.value)

    # 請求書 H35 合計
    if "請求書" in wb.sheetnames:
        ws_inv = wb["請求書"]
        invoice_total = 0
        cell_breakdown = []
        for row in ws_inv.iter_rows(min_row=21, max_row=34, min_col=8, max_col=9):
            for cell in row:
                if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                    invoice_total += int(cell.value)
                    cell_breakdown.append({"coord": cell.coordinate, "value": int(cell.value)})
        fire_type = vals.get("fire_insurance_type")
        if fire_type in ("A", "C"):
            fire_add = int(vals.get("fire_insurance") or 0)
            invoice_total += fire_add
            cell_breakdown.append({"coord": "fire_insurance", "value": fire_add})
        ws_inv["H35"] = invoice_total
        _log("info", "write_invoice_total", total=invoice_total, breakdown=cell_breakdown)

    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.save(output_path)

    _log("info", "write_done", output=output_path, cells=len(write_log),
         ok=sum(1 for e in write_log if e["ok"]),
         ng=sum(1 for e in write_log if not e["ok"]))
    return write_log


# ──────────────────────────────────────────
# HITL レビュー HTML 生成
# ──────────────────────────────────────────

def _format_display_value(key: str, raw) -> str:
    if raw is None:
        return "(なし)"   # 台帳のレビューHTMLと同じ空欄表記
    if key == "koshinryo_months":
        return f"{raw} ヶ月"
    if key == "signing_reiwa_year":
        return f"令和{raw}年"
    if key == "signing_reiwa_month":
        return f"{raw}月"
    if key == "signing_reiwa_day":
        return f"{raw}日"
    if isinstance(raw, (int, float)):
        return f"{int(raw):,}"
    return str(raw)


def _build_logic_rows(info: dict, vals: dict, changes: dict) -> str:
    """ボトムパネル「変更計算」のHTMLを組み立てる。"""
    tsumimashi = _compute_tsumimashi(info, vals)
    rows: list[str] = []

    def row(label, val_str, formula=""):
        f = f' <span class="logic-formula">{formula}</span>' if formula else ""
        rows.append(f'<div class="logic-row"><span class="logic-label">{label}</span>'
                    f'<span><span class="logic-val">{val_str}</span>{f}</span></div>')

    cr = int(info.get("current_rent") or 0)
    cm = int(info.get("current_mgmt") or 0)
    cp = int(info.get("current_parking") or 0)
    cd = int(info.get("current_deposit") or 0)
    nr, nm = vals["new_rent"], vals["new_mgmt"]
    np_ = vals.get("new_parking")
    nd = vals["new_deposit"]

    def diff_str(old, new):
        d = new - old
        return f"+{d:,}" if d > 0 else (f"{d:,}" if d < 0 else "±0")

    row("賃料", f"{cr:,} → {nr:,}", diff_str(cr, nr))
    row("管理費", f"{cm:,} → {nm:,}", diff_str(cm, nm))
    if np_ is not None:
        row("駐車場", f"{cp:,} → {np_:,}", diff_str(cp, np_))
    deposit_auto = not any(changes.get(k) is not None
                           for k in ["new_deposit", "deposit_up_amount", "deposit_up_pct", "deposit_rate"])
    row("敷金", f"{cd:,} → {nd:,}", "自動" if deposit_auto else diff_str(cd, nd))

    km = vals["koshinryo_months"]
    koshinryo = int(nr * km)
    row("更新料", f"{koshinryo:,}",
        f"{nr // 1000}k×{km}")
    row("更新手数料(税込)", f"{vals['koshin_tesuryo']:,}",
        f"{nr // 1000}k×0.25×1.1")

    if tsumimashi["daily_amount"] > 0:
        start = info.get("new_start_date")
        if start:
            import calendar as _cal
            dim = _cal.monthrange(start.year, start.month)[1]
            rem = dim - start.day + 1
            short_desc = f"{start.month}/{start.day}-{dim} {rem}日"
        else:
            short_desc = tsumimashi.get("daily_desc", "")
        row("積増 日割", f"{tsumimashi['daily_amount']:,}", short_desc)
    if tsumimashi["deposit_diff"] > 0:
        row("積増 敷金", f"{tsumimashi['deposit_diff']:,}", "差額比例")

    if vals["kanri_itaku"]:
        formula = "pct指定" if changes.get("kanri_itaku_pct") is not None else ""
        row("管理委託費", f"{vals['kanri_itaku']:,}", formula)

    row("貸主精算金(税込)", f"{vals['kashunushi_koshinryo']:,}",
        f"{nr // 1000}k×0.5×1.1")

    # 請求合計
    fire_type = vals.get("fire_insurance_type")
    fire_for_total = int(vals.get("fire_insurance") or 0) if fire_type in ("A", "C") else 0
    kurashiido_for_total = int(vals.get("kurashiido") or 0) if vals.get("kurashiido") and vals["kurashiido"] > 0 else 0
    invoice_total = (koshinryo
                     + vals["koshin_tesuryo"]
                     + tsumimashi["daily_amount"]
                     + tsumimashi["deposit_diff"]
                     + fire_for_total
                     + kurashiido_for_total)
    rows.append(
        f'<div class="logic-row logic-row-total"><span class="logic-label">'
        f'<strong>請求合計</strong></span><span class="logic-val">'
        f'<strong>{invoice_total:,}</strong></span></div>'
    )

    return "\n      ".join(rows)


def _build_write_summary(write_log: list[dict]) -> str:
    """ボトムパネル「Excel書込み」のHTMLを組み立てる。"""
    from collections import Counter
    sheet_counts = Counter(e["sheet"] for e in write_log)
    total_cells = len(write_log)
    rows: list[str] = []
    for sheet, cnt in sheet_counts.items():
        rows.append(f'<div class="logic-row"><span class="logic-label">{sheet}</span>'
                    f'<span class="logic-val">{cnt} cells</span></div>')

    warnings: list[str] = []
    merged = [e for e in write_log if e.get("merged")]
    if merged:
        coords = ", ".join(e["coord"] for e in merged[:3])
        warnings.append(f"結合セル書込み({coords})")
    ng = [e for e in write_log if not e.get("ok")]
    if ng:
        coords = ", ".join(e["coord"] for e in ng[:3])
        warnings.append(f"書込み検証NG({coords})")

    warn_html = ""
    if warnings:
        warn_html = ('<div class="logic-warn">'
                     f'<strong>要確認:</strong> {"/ ".join(warnings)}</div>')

    return (f'<div class="logic-title">Excel 書込み {total_cells}セル</div>\n'
            + "\n".join(rows)
            + warn_html)


def generate_review_html(
    info: dict,
    pdf_path: str,
    vals: dict | None = None,
    changes: dict | None = None,
    write_log: list[dict] | None = None,
    processing_time: float | None = None,
    output_path: str = "output/review.html",
) -> str:
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    def _render_pages(path: str) -> list[dict]:
        d = fitz.open(path)
        out = []
        for page in d:
            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
            b64 = base64.b64encode(pix.tobytes("png")).decode("ascii")
            out.append({"src": f"data:image/png;base64,{b64}",
                        "w": pix.width, "h": pix.height})
        d.close()
        return out

    # 書類（契約書＋補足資料）。ラベル＝アップロードされたファイル名そのまま。
    # info["_docs"] が無い場合（旧state・単体呼び出し）は契約書のみで構成する。
    doc_defs = info.get("_docs") or [{
        "label": os.path.splitext(os.path.basename(pdf_path))[0],
        "path": pdf_path,
        "coords": info.get("_page_coords", {}),
        "primary": True,
    }]
    docs_js_data: list[dict] = []
    for d in doc_defs:
        p = d.get("path")
        if not p or not os.path.isfile(p):
            _log("warn", "review_doc_missing", label=d.get("label"), path=p)
            continue
        docs_js_data.append({
            "label": d.get("label") or os.path.splitext(os.path.basename(p))[0],
            "pages": _render_pages(p),
            "coords": {str(k): v for k, v in (d.get("coords") or {}).items()},
            "primary": bool(d.get("primary")),
        })
    if not docs_js_data:  # 全滅時は契約書だけでも出す
        docs_js_data = [{"label": os.path.splitext(os.path.basename(pdf_path))[0],
                         "pages": _render_pages(pdf_path), "coords": {}, "primary": True}]

    page_data = docs_js_data[0]["pages"]
    total_pages = len(page_data)

    page_sources = info.get("_page_sources", {})
    boxes = info.get("_boxes", {})

    # 借主名・貸主名には法人/個人区分を併記（人間がレビューで検算できるように）
    NAME_FLAG_LABEL = {
        "kariunushi_name": "kariunushi_is_corporate",
        "kashunushi_name": "kashunushi_is_corporate",
    }

    def _corp_suffix(key):
        flag_key = NAME_FLAG_LABEL.get(key)
        if flag_key is None:
            return ""
        flag = info.get(flag_key)
        if flag is True:
            return "（法人）"
        if flag is False:
            return "（個人）"
        return "（区分不明）"

    # 補完元（補足資料）バッジ: ファイル名（拡張子なし）をそのまま表示。加工しない。
    # 長い名前の見た目は CSS（ellipsis）と title 属性で対応する。
    extra_sources = info.get("_extra_sources", {}) or {}

    status_map = _item_status(info, vals or {})
    section_of = {k: sec for sec, keys in SECTIONS for k in keys}
    # 「項目順」表示はセクション順に並べる（ITEM_DEFS の並びのままだと見出しが重複するため）
    section_idx = {sec: i for i, (sec, _) in enumerate(SECTIONS)}

    items = []
    for i, (key, label, typ) in enumerate(ITEM_DEFS, 1):
        raw = info.get(key)
        pg = page_sources.get(key)
        box = boxes.get(key)
        value = _format_display_value(key, raw)
        if raw is not None:
            value += _corp_suffix(key)
        src = (os.path.splitext(os.path.basename(extra_sources[key]))[0]
               if key in extra_sources else None)
        # 値はあるがページ出典も補完元もない ＝ 自動生成（フリガナ等）
        gen = raw is not None and src is None and pg is None
        # この項目がどの書類由来か（0=契約書）。補完元があればその書類を指す。
        doc_idx = 0
        if src is not None:
            for di, dd in enumerate(docs_js_data):
                if dd["label"] == src:
                    doc_idx = di
                    break
        st, st_reason = status_map.get(key, ("ok", ""))
        items.append({
            "no": i, "field": label, "key": key,
            "value": value,
            "type": typ, "isNull": raw is None or raw == 0,
            "page": pg, "box": box, "src": src, "gen": gen,
            "docIdx": doc_idx,
            "status": st, "reason": st_reason,
            "section": section_of.get(key, "その他"),
            "secIdx": section_idx.get(section_of.get(key, ""), 99),
        })

    bukken = info.get("bukken_name", "")
    goshitsu = info.get("goshitsu", "")
    title = f"{bukken} {goshitsu}".strip()
    n_ok =sum(1 for it in items if it["status"] == "ok")
    n_warn = sum(1 for it in items if it["status"] == "warn")
    n_empty = sum(1 for it in items if it["status"] == "empty")

    kariunushi = info.get("kariunushi_name", "")
    kashunushi = info.get("kashunushi_name", "")
    time_str = f"{processing_time:.1f}s" if processing_time else ""

    # ── ボトムパネル ──
    logic_html = ""
    write_html = ""
    if vals and changes is not None:
        logic_html = _build_logic_rows(info, vals, changes)
    if write_log:
        write_html = _build_write_summary(write_log)

    # ── ページデータJSON（srcはURIが巨大なのでJSで参照） ──
    pages_js = json.dumps(page_data)
    docs_js = json.dumps(docs_js_data, ensure_ascii=False)

    html = f"""<head><meta charset="utf-8"><title>更新書類 HITL Review</title>
<style>
:root {{
  --ground: #F3F4F7;
  --surface: #FFFFFF;
  --text: #1A1D2B;
  --text-2: #5A6178;
  --accent: #2D4FCA;
  --accent-light: #EDF0FB;
  --success: #0F7B5F;
  --success-bg: #ECFAF4;
  --warn: #B45309;
  --warn-bg: #FFF8EB;
  --ok: #0F7B5F;
  --ok-bg: #ECFAF4;
  --empty: #DC2626;
  --empty-bg: #FEF2F2;
  --border: #DDE0E9;
  --mono: ui-monospace, 'SF Mono', 'Cascadia Code', 'Consolas', monospace;
  --sans: system-ui, -apple-system, 'Segoe UI', sans-serif;
  --viewer-bg: #2A2D3A;
}}
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ background: var(--ground); color: var(--text); font-family: var(--sans);
  font-size: 14px; line-height: 1.5; height: 100vh; display: flex; flex-direction: column; overflow: hidden; }}

.topbar {{
  background: var(--text); color: #fff; padding: 12px 24px;
  display: flex; align-items: center; gap: 20px; flex-shrink: 0; z-index: 10;
}}
.topbar-label {{ font-size: 11px; opacity: 0.5; text-transform: uppercase; letter-spacing: 0.06em; font-weight: 600; }}
.topbar-bukken {{ font-size: 17px; font-weight: 700; letter-spacing: -0.01em; }}
.topbar-meta {{ margin-left: auto; display: flex; gap: 20px; font-size: 12px; }}
.topbar-meta span {{ opacity: 0.6; }}
.topbar-meta strong {{ opacity: 1; font-weight: 600; }}

.main-split {{ display: flex; flex: 1; min-height: 0; }}

.splitter {{
  width: 6px; flex-shrink: 0; cursor: col-resize;
  background: var(--border); position: relative; z-index: 5; transition: background 0.12s;
}}
.splitter::before {{
  content: ''; position: absolute; top: 50%; left: 50%;
  transform: translate(-50%,-50%); width: 2px; height: 28px;
  border-radius: 2px; background: var(--text-2); opacity: 0.4;
}}
.splitter:hover, .splitter.dragging {{ background: var(--accent); }}
.splitter:hover::before, .splitter.dragging::before {{ background: #fff; opacity: 0.9; }}
body.col-resizing {{ cursor: col-resize; user-select: none; }}

.left-pane {{
  width: 480px; min-width: 380px; flex-shrink: 0; background: var(--surface);
  border-right: 1px solid var(--border); display: flex; flex-direction: column; overflow: hidden;
}}
.pane-header {{
  height: 53px; padding: 0 16px; border-bottom: 1px solid var(--border);
  display: flex; align-items: center; gap: 10px; flex-shrink: 0; background: var(--surface);
}}
.pane-header h2 {{ font-size: 13px; font-weight: 700; }}
.next-warn-btn {{ padding: 4px 10px; border: 1px solid rgba(201,151,59,0.5);
  border-radius: 6px; background: rgba(201,151,59,0.12); color: #C9973B; font-size: 11px;
  font-weight: 700; cursor: pointer; white-space: nowrap; }}
.next-warn-btn:hover {{ background: rgba(201,151,59,0.25); }}
.sort-btns {{ margin-left: auto; display: flex; gap: 0; }}
.sort-btn {{
  padding: 4px 12px; font-size: 11px; font-weight: 600; border: 1px solid var(--border);
  background: var(--ground); color: var(--text-2); cursor: pointer; font-family: var(--sans);
}}
.sort-btn:first-child {{ border-radius: 5px 0 0 5px; }}
.sort-btn:last-child {{ border-radius: 0 5px 5px 0; border-left: none; }}
.sort-btn.active {{ background: var(--accent); color: #fff; border-color: var(--accent); }}
.filter-btns {{ display: flex; gap: 0; margin-left: 8px; }}
.filter-btn {{
  padding: 4px 10px; font-size: 11px; font-weight: 600; border: 1px solid var(--border);
  background: var(--ground); color: var(--text-2); cursor: pointer; font-family: var(--sans);
}}
.filter-btn:first-child {{ border-radius: 5px 0 0 5px; }}
.filter-btn:last-child {{ border-radius: 0 5px 5px 0; border-left: none; }}
.filter-btn.active {{ background: var(--accent); color: #fff; border-color: var(--accent); }}
/* 「次の要確認へ」だけの細い帯（ヘッダーが詰まるので1段下げる） */
.pane-subheader {{
  padding: 6px 16px; border-bottom: 1px solid var(--border); background: var(--surface);
  display: flex; align-items: center; flex-shrink: 0; min-height: 0;
}}
.pane-subheader:empty {{ display: none; }}

.extract-list {{ flex: 1; overflow-y: auto; }}
.section-divider {{
  padding: 6px 16px; background: var(--ground); font-size: 11px; font-weight: 700;
  color: var(--text-2); letter-spacing: 0.04em; border-bottom: 1px solid var(--border);
  position: sticky; top: 0; z-index: 2;
}}
.extract-row {{
  display: grid; grid-template-columns: 36px 1fr auto auto; gap: 0;
  padding: 10px 16px; border-bottom: 1px solid #F0F1F5;
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
.extract-val {{ font-weight: 600; font-size: 13px; margin-top: 1px; }}
.extract-val.is-num {{ font-family: var(--mono); font-variant-numeric: tabular-nums; }}
.extract-val.is-null {{ color: var(--text-2); opacity: 0.4; font-style: italic; }}
.extract-key {{ font-family: var(--mono); font-size: 10px; opacity: 0.35; font-weight: 400; }}
.extract-page {{
  display: inline-flex; align-items: center; justify-content: center; min-width: 32px;
  height: 24px; padding: 0 8px; border-radius: 5px; font-family: var(--mono); font-size: 11px;
  font-weight: 700; background: var(--accent-light); color: var(--accent); flex-shrink: 0;
}}
.extract-page.src {{ background: rgba(76,125,216,0.18); color: #7FA6E8; font-size: 10px; font-family: inherit;
  max-width: 110px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; display: inline-block;
  line-height: 24px; text-align: center; }}
.extract-page.gen {{ background: rgba(201,151,59,0.15); color: #C9973B; font-size: 10px; font-family: inherit; }}
.extract-status {{
  font-size: 10px; font-weight: 700; padding: 2px 8px; border-radius: 4px;
  white-space: nowrap; margin-left: 8px;
}}
.extract-status.s-ok {{ background: var(--ok-bg); color: var(--ok); }}
.extract-status.s-warn {{ background: var(--warn-bg); color: var(--warn); }}
.extract-status.s-empty {{ background: var(--empty-bg); color: var(--empty); }}

.right-pane {{ flex: 1; background: var(--viewer-bg); display: flex; flex-direction: column; min-width: 0; }}
.viewer-toolbar {{
  height: 53px; padding: 0 16px; display: flex; align-items: center; gap: 14px;
  background: rgba(0,0,0,0.3); flex-shrink: 0;
}}
.viewer-title {{ font-size: 13px; font-weight: 700; color: #fff; letter-spacing: 0.01em; white-space: nowrap; }}
.viewer-title span {{ opacity: 0.45; font-weight: 500; margin-left: 6px; font-size: 11px; }}
/* 書類タブ（ラベル＝アップロードされたファイル名そのまま） */
/* ラベルが長くても全書類のタブが並ぶよう、1タブずつ省略表示する（フル名は title に出す）*/
.doc-tabs {{ display: flex; gap: 6px; align-items: center; overflow-x: auto; flex: 0 0 auto; }}
.doc-tab {{ padding: 5px 12px; border: 0; border-radius: 6px; background: rgba(255,255,255,0.10);
  color: rgba(255,255,255,0.65); font-size: 12px; font-weight: 700; cursor: pointer; white-space: nowrap;
  max-width: 180px; overflow: hidden; text-overflow: ellipsis; flex-shrink: 0; }}
.doc-tab:hover {{ background: rgba(255,255,255,0.18); color: #fff; }}
.doc-tab.active {{ background: var(--accent); color: #fff; }}
/* 次の候補（同じ項目の別の記載箇所へ巡回） */
.cand-btn {{ padding: 5px 12px; border: 0; border-radius: 6px; background: var(--accent);
  color: #fff; font-size: 12px; font-weight: 700; cursor: pointer; white-space: nowrap; }}
.cand-btn:hover {{ filter: brightness(1.1); }}
.zoom-controls {{
  margin-left: auto; display: flex; align-items: center; gap: 4px;
  background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.12);
  border-radius: 8px; padding: 3px;
}}
.zoom-btn {{
  width: 26px; height: 26px; border-radius: 5px; border: none;
  background: transparent; color: #fff; font-size: 15px; cursor: pointer;
  display: flex; align-items: center; justify-content: center; font-family: var(--sans);
}}
.zoom-btn:hover {{ background: rgba(255,255,255,0.14); }}
.zoom-level {{
  color: #fff; font-family: var(--mono); font-size: 12px; font-weight: 600;
  min-width: 46px; text-align: center; cursor: pointer; user-select: none;
}}
.zoom-level:hover {{ text-decoration: underline; }}
.zoom-fit {{
  font-family: var(--sans); font-size: 11px; font-weight: 600;
  padding: 0 8px; height: 26px; border-radius: 5px; border: none;
  background: transparent; color: rgba(255,255,255,0.75); cursor: pointer;
}}
.zoom-fit:hover {{ background: rgba(255,255,255,0.14); color: #fff; }}
.viewer-toolbar .page-nav {{ display: flex; align-items: center; gap: 8px; }}
.page-btn {{
  width: 32px; height: 32px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.15);
  background: rgba(255,255,255,0.08); color: #fff; font-size: 16px; cursor: pointer;
  display: flex; align-items: center; justify-content: center; font-family: var(--sans);
}}
.page-btn:hover {{ background: rgba(255,255,255,0.15); }}
.page-btn:disabled {{ opacity: 0.3; cursor: default; }}
.page-indicator {{ color: #fff; font-family: var(--mono); font-size: 13px; font-weight: 600; min-width: 60px; text-align: center; }}
.page-indicator span {{ opacity: 0.5; font-weight: 400; }}
/* ページ数が多いときはサムネイル側を縮める（書類タブが切れないように）*/
.viewer-thumbs {{ display: flex; gap: 4px; flex: 0 1 auto; min-width: 0; overflow-x: auto; }}
.thumb {{
  width: 28px; height: 36px; border-radius: 3px; border: 2px solid transparent; cursor: pointer;
  font-family: var(--mono); font-size: 9px; font-weight: 700; color: rgba(255,255,255,0.6);
  display: flex; align-items: center; justify-content: center; background: rgba(255,255,255,0.08);
  transition: all 0.15s; flex-shrink: 0;
}}
.thumb:hover {{ background: rgba(255,255,255,0.15); }}
.thumb.active {{ border-color: #D7373F; background: rgba(215,55,63,0.3); color: #fff; }}
.thumb.has-data {{ background: rgba(215,55,63,0.15); color: rgba(255,255,255,0.8); }}

.viewer-body {{ flex: 1; overflow: auto; position: relative; }}
.page-stage {{
  min-width: 100%; min-height: 100%; box-sizing: border-box; padding: 28px;
  display: flex; align-items: center; justify-content: center;
}}
.page-card {{
  background: #fff; border-radius: 3px; box-shadow: 0 8px 32px rgba(0,0,0,0.45);
  position: relative; flex-shrink: 0; padding: 0; line-height: 0;
}}
.page-card img {{ display: block; width: 100%; height: 100%; border-radius: 3px; }}
.page-label {{
  position: absolute; bottom: 8px; right: 10px;
  font-family: var(--mono); font-size: 11px; line-height: 1.4; color: #fff;
  background: rgba(0,0,0,0.55); padding: 2px 8px; border-radius: 4px;
}}
.page-highlight {{
  position: absolute; border: none; background: rgba(215,55,63,.16);
  border-radius: 2px; pointer-events: none;
  animation: hl-fade 0.35s ease-out;
}}
@keyframes hl-fade {{
  0% {{ opacity: 0; }}
  100% {{ opacity: 1; }}
}}

.bottom-panel {{
  border-top: 1px solid var(--border); background: var(--surface);
  flex-shrink: 0; max-height: 50vh; overflow-y: auto; display: none;
}}
.bottom-panel.open {{ display: block; }}
.bottom-toggle {{
  position: fixed; bottom: 12px; right: 12px; z-index: 20;
  padding: 8px 16px; border-radius: 8px; border: 1px solid var(--border);
  background: var(--surface); color: var(--text); font-size: 12px; font-weight: 600;
  cursor: pointer; box-shadow: 0 2px 8px rgba(0,0,0,0.12); font-family: var(--sans);
  display: flex; align-items: center; gap: 6px;
}}
.bottom-toggle:hover {{ background: var(--ground); }}

.badge {{ font-size: 10px; padding: 2px 7px; border-radius: 8px; font-weight: 600; font-family: var(--mono); }}
.badge-ok {{ background: var(--success-bg); color: var(--success); }}

.logic-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 0; border-bottom: 1px solid var(--border); }}
.logic-card {{ padding: 16px 20px; border-right: 1px solid var(--border); }}
.logic-card:last-child {{ border-right: none; }}
.logic-title {{ font-size: 12px; font-weight: 700; margin-bottom: 10px; color: var(--text-2); text-transform: uppercase; letter-spacing: 0.04em; }}
.logic-row {{ display: flex; justify-content: space-between; padding: 4px 0; font-size: 12px; border-bottom: 1px solid #F5F6F9; }}
.logic-row:last-child {{ border-bottom: none; }}
.logic-row-total {{ border-top: 2px solid var(--border); padding-top: 8px; }}
.logic-label {{ color: var(--text-2); }}
.logic-val {{ font-family: var(--mono); font-weight: 600; font-variant-numeric: tabular-nums; }}
.logic-formula {{ font-size: 10px; color: var(--text-2); font-family: var(--mono); }}
.logic-warn {{
  margin-top: 12px; padding: 8px 10px; background: var(--warn-bg);
  border-radius: 5px; font-size: 11px; color: var(--warn);
}}

@media (max-width: 900px) {{
  .main-split {{ flex-direction: column; }}
  .left-pane {{ width: 100%; min-width: 0; max-height: 50vh; }}
  .right-pane {{ min-height: 300px; }}
  .logic-grid {{ grid-template-columns: 1fr; }}
}}
</style></head><body>

<div class="topbar">
  <div>
    <div class="topbar-label">HITL Review</div>
    <div class="topbar-bukken">{title}</div>
  </div>
  <div class="topbar-meta">
    <div><span>借主 </span><strong>{kariunushi}</strong></div>
    <div><span>貸主 </span><strong>{kashunushi}</strong></div>
    {"<div><span>処理 </span><strong>" + time_str + "</strong></div>" if time_str else ""}
    <div>
      <span class="extract-status s-ok">OK {n_ok}</span>
      <span class="extract-status s-warn">要確認 {n_warn}</span>
      <span class="extract-status s-empty">空欄 {n_empty}</span>
    </div>
  </div>
</div>

<div class="main-split">
  <div class="left-pane">
    <div class="pane-header">
      <h2>抽出結果 ({len(items)})</h2>
      <div class="sort-btns">
        <button class="sort-btn active" onclick="sortBy('no')">項目順</button>
        <button class="sort-btn" onclick="sortBy('page')">ページ順</button>
      </div>
      <div class="filter-btns">
        <button class="filter-btn active" onclick="filterBy('all')">全て</button>
        <button class="filter-btn" onclick="filterBy('ok')">OK</button>
        <button class="filter-btn" onclick="filterBy('warn')">確認</button>
        <button class="filter-btn" onclick="filterBy('empty')">空欄</button>
      </div>
    </div>
    <div class="pane-subheader">
      <button class="next-warn-btn" id="nextWarnBtn" onclick="nextWarnItem()" style="display:none"></button>
    </div>
    <div class="extract-list" id="extractList"></div>
  </div>

  <div class="splitter" id="splitter"></div>
  <div class="right-pane">
    <div class="viewer-toolbar">
      <div class="doc-tabs" id="docTabs"></div>
      <button class="cand-btn" id="candBtn" onclick="nextCandidate()" style="display:none"></button>
      <div class="page-nav">
        <button class="page-btn" onclick="goPage(currentPage-1)" id="prevBtn">&#8249;</button>
        <div class="page-indicator" id="pageIndicator">1 <span>/ {total_pages}</span></div>
        <button class="page-btn" onclick="goPage(currentPage+1)" id="nextBtn">&#8250;</button>
      </div>
      <div class="zoom-controls">
        <button class="zoom-btn" onclick="setZoom(zoom-0.25)" title="縮小">&#8722;</button>
        <div class="zoom-level" id="zoomLevel" onclick="resetZoom()" title="クリックで適合">100%</div>
        <button class="zoom-btn" onclick="setZoom(zoom+0.25)" title="拡大">+</button>
        <button class="zoom-fit" onclick="resetZoom()">適合</button>
      </div>
      <div class="viewer-thumbs" id="thumbs"></div>
    </div>
    <div class="viewer-body" id="viewerBody">
      <div class="page-stage" id="pageStage">
        <div class="page-card" id="pageCard">
          <img id="pageImg" alt="">
          <div class="page-label" id="pageLabel"></div>
        </div>
      </div>
    </div>
  </div>
</div>

{"<button class='bottom-toggle' onclick='toggleBottom()' id='bottomToggle'><span id='toggleIcon'>&#9650;</span> ロジック・書込み詳細</button>" if logic_html or write_html else ""}
{"<div class='bottom-panel' id='bottomPanel'><div class='logic-grid'>" if logic_html or write_html else ""}
{"<div class='logic-card'><div class='logic-title'>変更計算</div>" + logic_html + "</div>" if logic_html else ""}
{"<div class='logic-card'>" + write_html + "</div>" if write_html else ""}
{"</div></div>" if logic_html or write_html else ""}

<script>
const ITEMS={json.dumps(items, ensure_ascii=False)};
const DOCS={docs_js};
var currentDoc=0;
var PAGES=DOCS[0].pages;
var TOTAL=PAGES.length;
var DATA_PAGES=new Set(ITEMS.filter(function(i){{return i.page&&i.docIdx===0}}).map(function(i){{return i.page}}));

function renderDocTabs(){{
  var c=document.getElementById('docTabs');
  if(!c)return;
  c.innerHTML=DOCS.map(function(d,i){{
    var t=document.createElement('div');t.textContent=d.label;  // ファイル名そのまま（エスケープ）
    return '<button class="doc-tab'+(i===currentDoc?' active':'')+'" title="'+t.innerHTML+'" onclick="switchDoc('+i+')">'+t.innerHTML+'</button>';
  }}).join('');
}}

function switchDoc(i){{
  if(i<0||i>=DOCS.length)return;
  currentDoc=i;
  PAGES=DOCS[i].pages;TOTAL=PAGES.length;
  DATA_PAGES=new Set(ITEMS.filter(function(it){{return it.page&&it.docIdx===i}}).map(function(it){{return it.page}}));
  currentPage=1;zoom=1;updateZoomLabel();
  clearHighlights();renderDocTabs();showPage();renderThumbs();
  document.getElementById('pageIndicator').innerHTML='1 <span>/ '+TOTAL+'</span>';
  document.getElementById('prevBtn').disabled=true;
  document.getElementById('nextBtn').disabled=TOTAL<=1;
}}
let currentPage=1,currentSort='no',currentFilter='all',selectedNo=null,zoom=1;
const STATUS_LABELS={{"ok":"OK","warn":"要確認","empty":"空欄"}};

function layoutPage(){{
  var body=document.getElementById('viewerBody');
  var card=document.getElementById('pageCard');
  var p=PAGES[currentPage-1];
  if(!p)return;
  var pad=28*2;
  var availW=Math.max(40,body.clientWidth-pad);
  var availH=Math.max(40,body.clientHeight-pad);
  var fit=Math.min(availW/p.w,availH/p.h);
  var scale=fit*zoom;
  card.style.width=Math.round(p.w*scale)+'px';
  card.style.height=Math.round(p.h*scale)+'px';
}}

function esc(s){{var d=document.createElement('div');d.textContent=s==null?'':s;return d.innerHTML}}

function renderList(){{
  var list=document.getElementById('extractList');
  var sorted=[].concat(ITEMS);
  if(currentSort==='page')sorted.sort(function(a,b){{return(a.page||99)-(b.page||99)||a.no-b.no}});
  else sorted.sort(function(a,b){{return(a.secIdx-b.secIdx)||(a.no-b.no)}});
  var html='',lastGroup='';
  sorted.forEach(function(item){{
    if(currentFilter!=='all'&&item.status!==currentFilter)return;
    var group=currentSort==='page'
      ? (item.page?('P'+item.page):(item.src?item.src:'ページ出典なし'))
      : (item.section||'その他');
    if(group!==lastGroup){{
      html+='<div class="section-divider">'+esc(group)+'</div>';
      lastGroup=group;
    }}
    var sel=selectedNo===item.no?' selected':'';
    var vc='extract-val'+(item.type==='num'?' is-num':'')+(item.isNull?' is-null':'');
    var val=(item.value===''||item.value==null)?'(なし)':item.value;
    // 値が取れなかった項目に出典バッジは出さない（「?」は紛らわしい）
    var pg=item.page?'P'+item.page:(item.src?item.src:(item.gen?'生成':''));
    var pcls='extract-page'+(item.page?'':(item.src?' src':(item.gen?' gen':'')));
    var ptitle=item.src?' title="'+esc(item.src)+'"':'';
    var stitle=item.reason?' title="'+esc(item.reason)+'"':'';
    html+='<div class="extract-row'+sel+'" data-no="'+item.no+'" data-page="'+item.page+'" onclick="selectItem('+item.no+')">'+
      '<div class="extract-no">'+item.no+'</div>'+
      '<div class="extract-info"><div class="extract-field">'+esc(item.field)+'</div>'+
      '<div class="'+vc+'">'+esc(val)+' <span class="extract-key">'+esc(item.key)+'</span></div></div>'+
      (pg?'<div class="'+pcls+'"'+ptitle+'>'+esc(pg)+'</div>':'<span></span>')+
      '<span class="extract-status s-'+item.status+'"'+stitle+'>'+STATUS_LABELS[item.status]+'</span></div>';
  }});
  list.innerHTML=html;
}}

function filterBy(mode){{
  currentFilter=mode;
  document.querySelectorAll('.filter-btn').forEach(function(b){{b.classList.remove('active')}});
  document.querySelectorAll('.filter-btn').forEach(function(b){{
    if(b.getAttribute('onclick')==="filterBy('"+mode+"')")b.classList.add('active');
  }});
  renderList();
}}

function showPage(){{
  var p=PAGES[currentPage-1];
  var img=document.getElementById('pageImg');
  img.onload=function(){{layoutPage()}};
  img.src=p.src;
  document.getElementById('pageLabel').textContent='Page '+currentPage+' / '+TOTAL;
  var body=document.getElementById('viewerBody');
  layoutPage();
  body.scrollTop=0;body.scrollLeft=0;
}}

function renderThumbs(){{
  var c=document.getElementById('thumbs');c.innerHTML='';
  for(var i=1;i<=TOTAL;i++){{
    var t=document.createElement('div');
    t.className='thumb'+(i===currentPage?' active':'')+(DATA_PAGES.has(i)?' has-data':'');
    t.textContent=i;t.onclick=(function(n){{return function(){{goPage(n)}}}})(i);
    c.appendChild(t);
  }}
}}

function goPage(n){{
  if(n<1||n>TOTAL)return;
  currentPage=n;zoom=1;updateZoomLabel();
  clearHighlights();
  showPage();
  document.getElementById('pageIndicator').innerHTML=n+' <span>/ '+TOTAL+'</span>';
  document.getElementById('prevBtn').disabled=n<=1;
  document.getElementById('nextBtn').disabled=n>=TOTAL;
  renderThumbs();
}}

function setZoom(z){{
  zoom=Math.min(6,Math.max(0.25,Math.round(z*100)/100));
  updateZoomLabel();layoutPage();
}}
function resetZoom(){{setZoom(1)}}
function updateZoomLabel(){{
  document.getElementById('zoomLevel').textContent=Math.round(zoom*100)+'%';
}}

function selectItem(no){{
  selectedNo=(selectedNo===no)?null:no;
  clearHighlights();
  if(selectedNo){{
    var item=ITEMS.find(function(i){{return i.no===selectedNo}});
    if(item){{
      candList=buildCandidates(item);candIdx=0;
      if(candList.length){{
        jumpToCandidate(0);
      }} else {{
        if(typeof item.docIdx==='number'&&item.docIdx!==currentDoc)switchDoc(item.docIdx);
        if(item.page)goPage(item.page);
        if(item.box)showHighlight(item);
        updateCandBtn();
      }}
    }}
  }} else {{
    candList=[];candIdx=0;updateCandBtn();
  }}
  renderList();
  updateNextWarnBtn();
}}

// ---- 候補巡回: 同じ項目の値が原本の複数箇所にあるとき、順に飛ぶ（台帳と同設計）----
var candList=[];
var candIdx=0;

function buildCandidates(item){{
  var list=[];
  DOCS.forEach(function(d,di){{
    var pages=d.coords||{{}};
    Object.keys(pages).sort(function(a,b){{return Number(a)-Number(b)}}).forEach(function(p){{
      (pages[p]||[]).forEach(function(c){{
        if(c.key===item.key&&Array.isArray(c.box_2d)&&c.box_2d.length>=4){{
          list.push({{docIdx:di,label:d.label,page:Number(p),box:c.box_2d}});
        }}
      }});
    }});
  }});
  // 採用箇所（その項目の出典書類・出典ページ）を先頭に（安定ソート）
  return list.map(function(c,i){{return [c,i]}}).sort(function(a,b){{
    var aa=(a[0].docIdx===item.docIdx&&a[0].page===item.page)?0:1;
    var bb=(b[0].docIdx===item.docIdx&&b[0].page===item.page)?0:1;
    return aa-bb||a[1]-b[1];
  }}).map(function(x){{return x[0]}});
}}

function jumpToCandidate(i){{
  var c=candList[i];
  if(!c)return;
  candIdx=i;
  if(c.docIdx!==currentDoc)switchDoc(c.docIdx);
  goPage(c.page);
  clearHighlights();
  // box_2d = [ymin,xmin,ymax,xmax] (0-1000) → page-highlight の % に変換
  showHighlight({{box:[c.box[1]/10,c.box[0]/10,c.box[3]/10,c.box[2]/10]}});
  updateCandBtn();
}}

function nextCandidate(){{
  if(!candList.length)return;
  jumpToCandidate((candIdx+1)%candList.length);
}}

function updateCandBtn(){{
  var btn=document.getElementById('candBtn');
  if(!btn)return;
  if(candList.length<2){{btn.style.display='none';return}}
  var c=candList[candIdx];
  var t=document.createElement('div');t.textContent=c.label;
  btn.style.display='';
  btn.innerHTML='\\uD83D\\uDCCD '+t.innerHTML+' p'+c.page+'（'+(candIdx+1)+'/'+candList.length+'）　次の候補 \\u25B8';
}}

// ---- 次の要確認へ: status が ok でない項目（要確認・空欄）を番号順に巡回（台帳と同設計）----
function pendingItems(){{
  return ITEMS.filter(function(i){{return i.status!=='ok'}}).map(function(i){{return i.no}}).sort(function(a,b){{return a-b}});
}}
function nextWarnItem(){{
  var pending=pendingItems();
  if(!pending.length)return;
  var next=null;
  for(var k=0;k<pending.length;k++){{if(pending[k]>(selectedNo||0)){{next=pending[k];break}}}}
  if(next===null)next=pending[0];  // 末尾まで来たら先頭へ戻る
  if(selectedNo===next)selectedNo=null;  // selectItemのトグルで解除されないように
  selectItem(next);
  var row=document.querySelector('.extract-row[data-no="'+next+'"]');
  if(row)row.scrollIntoView({{block:'center'}});
}}
function updateNextWarnBtn(){{
  var btn=document.getElementById('nextWarnBtn');
  if(!btn)return;
  var pending=pendingItems();
  var bar=btn.parentElement;
  if(!pending.length){{btn.style.display='none';if(bar)bar.style.display='none';return}}
  btn.style.display='';if(bar)bar.style.display='';
  btn.textContent='次の要確認へ ▸（残 '+pending.length+'）';
}}
function clearHighlights(){{
  document.querySelectorAll('.page-highlight').forEach(function(e){{e.remove()}});
}}
function showHighlight(item){{
  if(!item.box)return;
  var card=document.getElementById('pageCard');
  var el=document.createElement('div');
  el.className='page-highlight';
  var PAD=1.2;
  el.style.left=(item.box[0]-PAD)+'%';el.style.top=(item.box[1]-PAD)+'%';
  el.style.width=(item.box[2]-item.box[0]+PAD*2)+'%';el.style.height=(item.box[3]-item.box[1]+PAD*2)+'%';
  card.appendChild(el);
}}

function sortBy(mode){{
  currentSort=mode;
  document.querySelectorAll('.sort-btn').forEach(function(b){{b.classList.remove('active')}});
  document.querySelectorAll('.sort-btn').forEach(function(b){{
    if(b.getAttribute('onclick')==="sortBy('"+mode+"')")b.classList.add('active');
  }});renderList();
}}

function toggleBottom(){{
  var panel=document.getElementById('bottomPanel');
  var icon=document.getElementById('toggleIcon');
  if(!panel)return;
  panel.classList.toggle('open');
  icon.innerHTML=panel.classList.contains('open')?'\\u25BC':'\\u25B2';
}}

renderDocTabs();renderList();renderThumbs();showPage();updateZoomLabel();updateNextWarnBtn();
requestAnimationFrame(layoutPage);
window.addEventListener('load',layoutPage);
document.getElementById('prevBtn').disabled=true;

(function(){{
  var splitter=document.getElementById('splitter');
  var left=document.querySelector('.left-pane');
  var dragging=false;
  function onMove(e){{
    if(!dragging)return;
    var x=(e.touches?e.touches[0].clientX:e.clientX);
    var min=280,max=window.innerWidth-360;
    var w=Math.max(min,Math.min(max,x));
    left.style.width=w+'px';layoutPage();
  }}
  function stop(){{
    dragging=false;splitter.classList.remove('dragging');
    document.body.classList.remove('col-resizing');
  }}
  splitter.addEventListener('mousedown',function(e){{
    dragging=true;e.preventDefault();
    splitter.classList.add('dragging');document.body.classList.add('col-resizing');
  }});
  splitter.addEventListener('touchstart',function(){{dragging=true;splitter.classList.add('dragging')}},{{passive:true}});
  window.addEventListener('mousemove',onMove);
  window.addEventListener('touchmove',onMove,{{passive:true}});
  window.addEventListener('mouseup',stop);
  window.addEventListener('touchend',stop);
  splitter.addEventListener('dblclick',function(){{left.style.width='480px';layoutPage()}});
}})();

document.getElementById('viewerBody').addEventListener('wheel',function(e){{
  if(e.ctrlKey||e.metaKey){{e.preventDefault();setZoom(zoom+(e.deltaY<0?0.15:-0.15))}}
}},{{passive:false}});
var _rt;
window.addEventListener('resize',function(){{clearTimeout(_rt);_rt=setTimeout(layoutPage,80)}});
</script></body>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    _log("info", "review_html_generated", path=output_path, pages=total_pages)
    return output_path


# ──────────────────────────────────────────
# 状態保存・復元（ターン2以降の編集用）
# ──────────────────────────────────────────

STATE_PATH = "tmp/state.json"
NAME_KEYS = {"kashunushi_name", "kashunushi_address", "kariunushi_name", "bukken_name", "goshitsu"}
# 抽出済みの法人フラグ。handle_edit では changes ではなく info を直接書き換える
INFO_FLAG_KEYS = {"kashunushi_is_corporate", "kariunushi_is_corporate"}


def _save_state(info, changes, vals, pdf_path, template_path):
    os.makedirs("tmp", exist_ok=True)
    with open(STATE_PATH, "w", encoding="utf-8") as f:
        json.dump(
            {"info": info, "changes": changes, "vals": vals,
             "pdf_path": pdf_path, "template_path": template_path},
            f, ensure_ascii=False, default=str,
        )
    _log("info", "state_saved", path=STATE_PATH)


def _load_state():
    if not os.path.exists(STATE_PATH):
        _log("warn", "state_not_found", path=STATE_PATH)
        return None
    with open(STATE_PATH, encoding="utf-8") as f:
        state = json.load(f)
    _log("info", "state_loaded", path=STATE_PATH, keys=list(state.get("changes", {}).keys()))
    return state


# エントリーポイント
# ──────────────────────────────────────────

def _find_file(suffix: str) -> str | None:
    for f in sorted(os.listdir(".")):
        if f.lower().endswith(suffix) and os.path.isfile(f):
            _log("info", "find_file", suffix=suffix, found=f, location="cwd")
            return f
    for search_dir in ("uploads", "rag"):
        if os.path.isdir(search_dir):
            for f in sorted(os.listdir(search_dir)):
                if f.lower().endswith(suffix):
                    path = f"{search_dir}/{f}"
                    _log("info", "find_file", suffix=suffix, found=path, location=search_dir)
                    return path
    _log("warn", "find_file_not_found", suffix=suffix)
    return None


def _find_all_files(suffix: str) -> list[str]:
    """suffix に一致するファイルを cwd → uploads → rag の順で全て返す。"""
    found: list[str] = []
    for f in sorted(os.listdir(".")):
        if f.lower().endswith(suffix) and os.path.isfile(f):
            found.append(f)
    for search_dir in ("uploads", "rag"):
        if os.path.isdir(search_dir):
            for f in sorted(os.listdir(search_dir)):
                if f.lower().endswith(suffix):
                    found.append(f"{search_dir}/{f}")
    _log("info", "find_all_files", suffix=suffix, found=found)
    return found


# レビューHTML左ペインのセクション見出し（台帳の SECTIONS 相当）
SECTIONS: list[tuple[str, list[str]]] = [
    ("物件", ["bukken_name", "bukken_address", "goshitsu"]),
    ("貸主・借主", ["kashunushi_name", "kashunushi_address", "kariunushi_name"]),
    ("契約期間", ["current_start", "current_end",
                  "signing_reiwa_year", "signing_reiwa_month", "signing_reiwa_day"]),
    ("金額", ["current_rent", "current_mgmt", "current_parking", "current_deposit",
              "koshinryo_months", "fire_insurance_fee"]),
    ("借主情報", ["furigana", "birthday", "address", "phone", "workplace", "emergency_contact"]),
]


def _item_status(info: dict, vals: dict | None = None) -> dict[str, tuple[str, str]]:
    """項目キー → (status, 理由) を返す。status は ok / warn / empty。
    レビューHTMLのバッジ・フィルタと、報告の「ご確認ください」で同じ判定を使う（二重管理しない）。
    - empty: 全書類に記載がなく値が取れなかった
    - warn : 値はあるが人の確認が要る（法人区分が判定不能 など）
    - ok   : 契約書・補足資料から取得済み（フリガナの自動生成もOK扱い）
    """
    vals = vals or {}
    st: dict[str, tuple[str, str]] = {}
    for key, _, _ in ITEM_DEFS:
        if info.get(key) is None:
            st[key] = ("empty", "全書類に記載が見つかりませんでした")
        else:
            st[key] = ("ok", "")
    # 契約書に記載がなく補足資料から補完した値は、契約内容と一致するか人が見る
    for key, src in (info.get("_extra_sources") or {}).items():
        if st.get(key, ("", ""))[0] == "ok":
            st[key] = ("warn", f"契約書に記載がなく「{os.path.basename(src)}」から補完しました")
    # 法人区分が判定できない場合、その名前の項目を要確認にする
    if info.get("kariunushi_name") is not None and info.get("kariunushi_is_corporate") is None:
        st["kariunushi_name"] = ("warn", "法人／個人が判定できませんでした（個人・様として出力）")
    if info.get("kashunushi_name") is not None and info.get("kashunushi_is_corporate") is None:
        st["kashunushi_name"] = ("warn", "法人／個人が判定できませんでした（個人・様として出力）")
    # 火災保険：金額はあるが取り扱いパターン未指定なら要確認
    if info.get("fire_insurance_fee") is not None and vals.get("fire_insurance_type") is None:
        st["fire_insurance_fee"] = ("warn", "取り扱い未指定（A:保証会社 / B:保険会社直接 / C:管理会社代行 / D:加入なし）")
    return st


def _confirm_notes(info: dict, vals: dict) -> list[tuple[str, str]]:
    """報告の「ご確認ください」に出す不足・要確認項目を (項目, 内容) で返す。
    成果物（Excel/review.html）に全値があるので、ここでは確認が要る点だけに絞る。"""
    rows: list[tuple[str, str]] = []
    sk = vals.get("shiharai_kijitsu")
    if sk:
        rows.append(("支払期日", f"「{sk}」に自動設定（新契約開始日の5日前）。変更があれば指示してください"))
    else:
        rows.append(("支払期日", "未設定（新契約開始日が不明）。指定してください"))
    if vals.get("kanri_itaku") == 0:
        rows.append(("管理委託費", "未指定（0円）。必要なら指示してください"))
    if vals.get("fire_insurance_type") is None:
        rows.append(("火災保険", "未指定（A:保証会社引落し / B:保険会社直接 / C:管理会社代行 / D:加入なし）"))
    if info.get("kariunushi_is_corporate") is None:
        rows.append(("借主の法人区分", "判定できませんでした（個人・様として出力）。誤りなら指示してください"))
    if info.get("kashunushi_is_corporate") is None:
        rows.append(("貸主の法人区分", "判定できませんでした（個人・様として出力）。誤りなら指示してください"))
    rows.append(("全般", "内容に変更がある可能性がありますのでご確認ください"))
    return rows


def _confirm_table(rows: list[tuple[str, str]]) -> str:
    """(項目, 内容) のリストを Markdown 表にして返す。"""
    lines = ["【ご確認ください】", "| 項目 | 内容 |", "|------|------|"]
    lines += [f"| {label} | {detail} |" for label, detail in rows]
    return "\n".join(lines)


def _calc_breakdown(info: dict, vals: dict, changes: dict) -> list[tuple[str, str, str]]:
    """(項目, 金額, 計算根拠) のリストを返す。どの式でその値になったかを示す。"""
    tsumimashi = _compute_tsumimashi(info, vals)
    cr = int(info.get("current_rent") or 0)
    cm = int(info.get("current_mgmt") or 0)
    cp = int(info.get("current_parking") or 0)
    cd = int(info.get("current_deposit") or 0)
    nr, nm, nd = vals["new_rent"], vals["new_mgmt"], vals["new_deposit"]
    np_ = vals.get("new_parking")
    km = vals["koshinryo_months"]
    km_str = f"{km:g}"
    rows: list[tuple[str, str, str]] = []
    rows.append(("賃料", f"{nr:,}", f"現行 {cr:,} → 新 {nr:,}"))
    rows.append(("管理費", f"{nm:,}", f"現行 {cm:,} → 新 {nm:,}"))
    if np_ is not None:
        rows.append(("駐車場", f"{np_:,}", f"現行 {cp:,} → 新 {np_:,}"))
    deposit_auto = not any(changes.get(k) is not None
                           for k in ["new_deposit", "deposit_up_amount", "deposit_up_pct", "deposit_rate"])
    if not deposit_auto:
        rows.append(("敷金", f"{nd:,}", "明示指定"))
    elif nd != cd:
        rows.append(("敷金", f"{nd:,}", f"自動増額: 現敷金 {cd:,} ＋ 賃料差額×(現敷金÷現賃料)"))
    else:
        rows.append(("敷金", f"{nd:,}", "据え置き（賃料UPなし）"))
    rows.append(("更新料", f"{int(nr * km):,}", f"新賃料 {nr:,} × {km_str}ヶ月"))
    rows.append(("更新事務手数料", f"{vals['koshin_tesuryo']:,}", f"新賃料 {nr:,} × 0.25 × 1.1（税込・借主負担）"))
    rows.append(("貸主更新業務手数料", f"{vals['kashunushi_koshinryo']:,}", f"新賃料 {nr:,} × 0.5 × 1.1（税込）"))
    if tsumimashi["daily_amount"] > 0:
        rows.append(("積み増し 日割り賃料", f"{tsumimashi['daily_amount']:,}",
                     f"賃料差額×残日数÷月日数（{tsumimashi['daily_desc']}）"))
    if tsumimashi["deposit_diff"] > 0:
        rows.append(("積み増し 敷金差額", f"{tsumimashi['deposit_diff']:,}", "賃料差額×(現敷金÷現賃料)"))
    if vals.get("kanri_itaku"):
        src = "（新賃料＋管理費）×％" if changes.get("kanri_itaku_pct") is not None else "指定額"
        rows.append(("管理委託費", f"{vals['kanri_itaku']:,}", src))
    rows.append(("月額合計", f"{vals['monthly_total']:,}",
                 "賃料＋管理費" + ("＋駐車場" if np_ is not None else "")))
    # 請求合計
    koshinryo = int(nr * km)
    fire_type = vals.get("fire_insurance_type")
    fire = int(vals.get("fire_insurance") or 0) if fire_type in ("A", "C") else 0
    kurashiido = int(vals.get("kurashiido")) if (vals.get("kurashiido") and vals["kurashiido"] > 0) else 0
    invoice_total = (koshinryo + vals["koshin_tesuryo"] + tsumimashi["daily_amount"]
                     + tsumimashi["deposit_diff"] + fire + kurashiido)
    parts = [f"更新料{koshinryo:,}", f"事務手数料{vals['koshin_tesuryo']:,}"]
    if tsumimashi["daily_amount"] > 0:
        parts.append(f"日割り{tsumimashi['daily_amount']:,}")
    if tsumimashi["deposit_diff"] > 0:
        parts.append(f"敷金差額{tsumimashi['deposit_diff']:,}")
    if fire > 0:
        parts.append(f"火災{fire:,}")
    if kurashiido > 0:
        parts.append(f"くらしーど{kurashiido:,}")
    rows.append(("請求合計", f"{invoice_total:,}", "＋".join(parts)))
    return rows


def _calc_table(rows: list[tuple[str, str, str]]) -> str:
    """(項目, 金額, 計算根拠) のリストを Markdown 表にして返す。"""
    lines = ["【計算根拠】", "| 項目 | 金額 | 計算根拠 |", "|------|------|----------|"]
    lines += [f"| {label} | {amount} | {basis} |" for label, amount, basis in rows]
    return "\n".join(lines)


async def main(changes: dict | None = None, *, pdf_path: str | None = None,
               extra_pdf_paths: list[str] | None = None):
    """戻り値: (info, vals, output_path)

    pdf_path: 契約書PDF（正本。抽出・出典ページ・レビューHTMLの基準）
    extra_pdf_paths: 同一物件の補足資料PDF（更新申込書・請求書・マイソク等）。
        契約書から取れなかった項目だけをここから補完する（契約書の値は上書きしない）。
    """
    t0 = time.time()
    _log_entries.clear()
    os.makedirs("tmp", exist_ok=True)
    os.makedirs("output", exist_ok=True)

    if changes is None:
        changes = {}

    unknown = set(changes) - VALID_KEYS
    if unknown:
        raise ValueError(
            f"未対応キー: {unknown}\n"
            f"使えるキー: {sorted(VALID_KEYS)}\n"
            f"※「3万円」「15,000円増やす」などは数値に変換してから渡すこと"
        )

    if pdf_path is None:
        # pdf_path 未指定は「PDFが1つだけ」の場合のみ許容する。
        # 複数PDFがアップロードされている場合は、どれを処理するか曖昧なので
        # 呼び出し側（AI）に pdf_path の明示指定を要求する。
        pdf_candidates = _find_all_files(".pdf")
        if not pdf_candidates:
            raise FileNotFoundError("PDF ファイルが見つかりません")
        if len(pdf_candidates) > 1:
            listing = "\n".join(f"  - {p}" for p in pdf_candidates)
            raise ValueError(
                f"PDF が複数見つかりました。pdf_path を明示指定してください。\n"
                f"候補:\n{listing}\n"
                f"例: await main(CHANGES, pdf_path=\"{pdf_candidates[0]}\")\n"
                f"※ 複数物件を処理する場合は、1物件ずつ pdf_path を指定して実行すること"
            )
        pdf_path = pdf_candidates[0]
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"指定された PDF が存在しません: {pdf_path}")
    extra_pdf_paths = extra_pdf_paths or []
    for p in extra_pdf_paths:
        if not os.path.isfile(p):
            raise FileNotFoundError(f"指定された補足資料 PDF が存在しません: {p}")
    template_path = _find_file(".xlsx")
    if template_path is None:
        raise FileNotFoundError("xlsx ファイルが見つかりません")

    _log("info", "start", pdf=pdf_path, extra_pdfs=extra_pdf_paths,
         template=template_path, changes=changes)
    output_path = "output/更新書類_記入済み.xlsx"

    # Phase 1: 抽出（契約書 → 補足資料の順。フリガナ生成はマージ後に1回だけ）
    # 座標は抽出スキーマ（_coordinates_json）に同梱されるため、追加のLLM呼び出しは発生しない。
    _log("info", "phase_start", phase="extract")
    info = await extract(pdf_path, fill_furigana=False)

    # 書類ごとの情報（レビューHTMLの書類タブ・候補巡回用）。ラベル＝アップロードされたファイル名そのまま。
    docs: list[dict] = [{
        "label": os.path.splitext(os.path.basename(pdf_path))[0],
        "path": pdf_path,
        "coords": info.get("_page_coords", {}),
        "primary": True,
    }]

    # Phase 1.2: 補足資料から不足項目を補完（契約書の値は上書きしない）
    for extra_path in extra_pdf_paths:
        try:
            extra_info = await extract(extra_path, fill_furigana=False)
        except Exception as e:
            # 補足資料が読めなくても本処理は止めない（マイソク等、抽出0件もあり得る）
            _log("warn", "extract_extra_failed", pdf=extra_path, error=str(e))
            continue
        _merge_extra_info(info, extra_info, source=extra_path)
        docs.append({
            "label": os.path.splitext(os.path.basename(extra_path))[0],
            "path": extra_path,
            "coords": extra_info.get("_page_coords", {}),
            "primary": False,
        })

    # フリガナ: 実物（契約書・補足資料）に無かった場合のみ借主名から生成
    await _fill_furigana(info)
    info["_docs"] = docs
    _log("info", "phase_end", phase="extract",
         docs=[d["label"] for d in docs],
         coords_per_doc={d["label"]: sum(len(v) for v in d["coords"].values()) for d in docs})

    # Phase 1.5: 座標リカバリ（抽出で座標が1件も取れなかった場合のみ追加取得）
    if not info.get("_boxes"):
        _log("info", "phase_start", phase="locate_boxes_recover")
        await _locate_boxes(info, pdf_path)
        if info.get("_boxes"):
            docs[0]["coords"] = info.get("_page_coords", {}) or docs[0]["coords"]
        _log("info", "phase_end", phase="locate_boxes_recover")

    # Phase 2: 計算
    _log("info", "phase_start", phase="compute")
    vals = compute(info, changes)
    _log("info", "phase_end", phase="compute")

    # Phase 3: Excel 書込み
    _log("info", "phase_start", phase="write")
    write_log = write(info, vals, template_path, output_path)
    _log("info", "phase_end", phase="write")

    # 実行環境 print: 対象見出し＋計算根拠＋不足・要確認（全値の再掲はしない＝成果物にある）
    print(f"【対象】{(info.get('bukken_name') or '')} {(info.get('goshitsu') or '')}".rstrip())
    print()
    print(_calc_table(_calc_breakdown(info, vals, changes)))
    print()
    print(_confirm_table(_confirm_notes(info, vals)))

    # Phase 4: HITL レビュー HTML
    _log("info", "phase_start", phase="review_html")
    elapsed = time.time() - t0
    review_path = generate_review_html(
        info, pdf_path,
        vals=vals, changes=changes, write_log=write_log,
        processing_time=elapsed,
    )

    _save_state(info, changes, vals, pdf_path, template_path)
    _log("info", "main_complete", elapsed=f"{elapsed:.1f}s", output=output_path)
    _save_log("tmp/log.json")
    _save_log("output/log.json")  # output/ 配下は毎回自動でファイル提示される（JSONログ必須化）

    return info, vals, output_path


async def handle_edit(edits: dict):
    """ターン2以降: 状態復元 → edits マージ → 再計算 → Excel再書込み"""
    _log_entries.clear()
    _log("info", "handle_edit_start", edits=edits)
    os.makedirs("output", exist_ok=True)
    state = _load_state()
    if state is None:
        raise RuntimeError("前回の実行結果がありません。先に main(CHANGES) を実行してください")

    info = state["info"]
    old_changes = state["changes"]
    old_vals = state["vals"]
    pdf_path = state["pdf_path"]
    template_path = state["template_path"]

    # 旧バージョンの state.json には法人フラグが無い。欠落していたら None 補完して警告。
    # （フラグ無し → 黙って個人・様に化けるのを防ぐ。今回 edits で上書きされる分は除く）
    legacy_flags = [
        k for k in INFO_FLAG_KEYS
        if k not in info and k not in edits
    ]
    if legacy_flags:
        for k in legacy_flags:
            info[k] = None
        _log("warn", "legacy_state_missing_flags", keys=legacy_flags)

    unknown = set(edits) - VALID_KEYS
    if unknown:
        raise ValueError(
            f"未対応キー: {unknown}\n"
            f"使えるキー: {sorted(VALID_KEYS)}\n"
            f"※「3万円」「15,000円増やす」などは数値に変換してから渡すこと"
        )

    merged = {**old_changes}
    name_keys_applied = []
    nullified_keys = []
    updated_keys = []
    flag_keys_applied = []
    for key, value in edits.items():
        if key in NAME_KEYS:
            info[key] = value
            name_keys_applied.append(key)
        elif key in INFO_FLAG_KEYS:
            # 法人フラグは抽出値なので info を直接書き換える（True/False/None いずれも可）
            info[key] = value
            flag_keys_applied.append(key)
        elif value is None:
            merged.pop(key, None)
            nullified_keys.append(key)
        else:
            merged[key] = value
            updated_keys.append(key)
    _log("info", "handle_edit_merge",
         name_keys=name_keys_applied, flag_keys=flag_keys_applied,
         nullified=nullified_keys, updated=updated_keys,
         merged_keys=list(merged.keys()))

    new_vals = compute(info, merged)

    output_path = "output/更新書類_記入済み.xlsx"
    write_log = write(info, new_vals, template_path, output_path)
    generate_review_html(info, pdf_path, vals=new_vals, changes=merged, write_log=write_log)

    _save_state(info, merged, new_vals, pdf_path, template_path)
    _save_log("tmp/log.json")
    _save_log("output/log.json")  # output/ 配下は毎回自動でファイル提示される（JSONログ必須化）

    # 編集時は「計算根拠」＋「ご確認ください」を表形式で出す
    notes = _confirm_notes(info, new_vals)
    if legacy_flags:
        notes = [("法人区分(復元)",
                  "前回データに法人/個人区分がありません（個人・様扱い）。指定するか main() からやり直してください")] + notes
    print(f"【対象】{(info.get('bukken_name') or '')} {(info.get('goshitsu') or '')}".rstrip())
    print()
    print(_calc_table(_calc_breakdown(info, new_vals, merged)))
    print()
    print(_confirm_table(notes))

    diff = {}
    for key in set(old_vals) | set(new_vals):
        if old_vals.get(key) != new_vals.get(key):
            diff[key] = {"before": old_vals.get(key), "after": new_vals.get(key)}

    _log("info", "handle_edit_complete", diff_keys=list(diff.keys()), output=output_path)
    return diff, new_vals, output_path
