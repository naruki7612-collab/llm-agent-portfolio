"""
ptc_media_agent.py  —  メディア選定エージェント 実行環境版 (一本化)

使い方（メインAIがcode_executeで呼ぶ）:
    user_request = "金融業界の経営層向けにタイアップで300万以内"
    await main(user_request)

ターン2以降は ptc_state グローバル変数を参照して続きを実行:
    await handle_turn2(user_request_2)
"""

# ============================================================
# 依存ライブラリ
# ============================================================
import asyncio
import difflib
import unicodedata
import io
import json
import os
import re
import sys
from pathlib import Path

import pandas as pd
import pdfplumber
from openpyxl import load_workbook

from agent_sdk import llm_call, tool_bridge
import datetime

import time

# ============================================================
# デバッグロガー（tmp/debug.jsonl に追記）— 診断モード: 截断ゼロ・全段記録
# ============================================================
_DEBUG_PATH = "tmp/debug.jsonl"
_START_TIME = time.time()
_DIAG = True            # 診断モード。Falseにすると重いログを抑制
_SEQ = 0                # イベント通し番号（順序の確定用）

# ★1改修: 媒体資料PDFは画像化(image_path)せず file_paths でPDF書類のまま読む。
# PDF読解はGemini明示時のみ動作するため読取りモデルを固定する（Claude等はfile_pathsで戻り値形が違い落ちる）。
GEMINI_PDF_MODEL = "gemini/gemini-3.5-flash"
# 最後の砦(pdfplumberテキスト抽出)でllm_callに渡すテキスト上限。超過分は末尾を切る（切ったらログに残す）。
_PDF_TEXT_FALLBACK_MAXCHARS = 30000
# ★最重大改修: マスタ(DB)値とPDF値の乖離がこのpt数を超えたら「矛盾」とみなし、
# 一次ソースであるPDF値を採用して警告表示＋DB修正ログを残す（例: Forbes DB57.7% vs PDF34.7% = 23pt乖離）。
_DB_PDF_DIVERGENCE_PT = 15
# llm_call添付のinline上限対策: 上限は環境差がある（旧実装10MB=実ログで確認 / 現行実測64MB）。
# 確実に超えるサイズだけ事前スキップし、実際の上限は下のエラー文字列検知で環境に自動適応する。
_LLM_ATTACH_HARD_LIMIT_BYTES = 64 * 1024 * 1024


class PdfTooLargeError(Exception):
    """llm_call添付がinline上限超過。決定論的エラーであり同じPDFの再送リトライは無意味。"""
    pass

def _reset_debug_path(turn: int):
    """ターン別のdebugファイルパスに切り替え。新ファイルを開始。"""
    global _DEBUG_PATH, _START_TIME, _SEQ
    suffix = f"_turn{turn}" if turn > 1 else ""
    _DEBUG_PATH = f"tmp/debug{suffix}.jsonl"
    if turn > 1:
        _START_TIME = time.time()  # turn2以降のみ再計測
    _SEQ = 0
    os.makedirs("tmp", exist_ok=True)
    with open(_DEBUG_PATH, "w", encoding="utf-8") as f:
        pass

def _count_pdf_pages(path: str) -> int:
    """PDFのページ数を返す。失敗時は-1。"""
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            return len(pdf.pages)
    except Exception:
        try:
            # fallback: バイナリからページ数を推定
            with open(path, "rb") as f:
                content = f.read()
            return content.count(b"/Type /Page") - content.count(b"/Type /Pages")
        except Exception:
            return -1

def _describe(v, _depth=0):
    """任意の値を「型・長さ・中身」付きで完全記述する（截断なし）。"""
    try:
        if v is None:
            return {"type": "NoneType", "value": None}
        if isinstance(v, bool):
            return {"type": "bool", "value": v}
        if isinstance(v, (int, float)):
            return {"type": type(v).__name__, "value": v}
        if isinstance(v, str):
            return {"type": "str", "len": len(v), "value": v}        # 截断しない
        if isinstance(v, dict):
            if _depth > 20:
                return {"type": "dict", "keys": list(v.keys()), "note": "max_depth_20"}
            return {"type": "dict", "keys": list(v.keys()),
                    "items": {str(k): _describe(val, _depth + 1) for k, val in v.items()}}
        if isinstance(v, (list, tuple)):
            if _depth > 20:
                return {"type": type(v).__name__, "len": len(v), "note": "max_depth_20"}
            return {"type": type(v).__name__, "len": len(v),
                    "items": [_describe(x, _depth + 1) for x in v]}
        return {"type": type(v).__name__, "repr": repr(v)}
    except Exception as e:
        return {"type": "DESCRIBE_ERROR", "error": str(e)}

def _dlog(event: str, payload: dict = None):
    """全イベントをJSONLに記録。截断なし。seq/elapsed/呼び出し元行番号を自動付与。"""
    global _SEQ
    _SEQ += 1
    payload = payload or {}
    rec = {
        "seq":     _SEQ,
        "ts":      datetime.datetime.now().isoformat(),
        "elapsed": round(time.time() - _START_TIME, 4),
        "event":   event,
    }
    # 呼び出し元の行番号（どこで吐いたか）
    try:
        import inspect
        fr = inspect.stack()[1]
        rec["src"] = f"{fr.function}:{fr.lineno}"
    except Exception:
        pass
    rec.update(payload)
    os.makedirs("tmp", exist_ok=True)
    with open(_DEBUG_PATH, "a", encoding="utf-8") as f:
        f.write(json.dumps(rec, ensure_ascii=False, default=str) + "\n")

def _dlog_full(event: str, context: str, value, **extra):
    """任意オブジェクトを完全記述（型・全キー・全値・全文）して記録する。"""
    _dlog(event, {"context": context, "full": _describe(value), **extra})


async def _llm_call_safe(prompt: str, schema: dict = None, context: str = "",
                         pdf_path: str = "") -> dict:
    """
    llm_callのラッパー。診断モードでは入力・生レスポンス・fallback・rescueの全段を截断なしで記録。
    pdf_path を渡すと、画像化せず file_paths でPDF書類のまま Gemini に読ませる（★1改修）。
    """
    t0 = time.time()
    cleaned = None
    if schema:
        def _clean_schema(s):
            if isinstance(s, dict):
                s.pop("additionalProperties", None)
                for v in s.values(): _clean_schema(v)
            elif isinstance(s, list):
                for v in s: _clean_schema(v)
        # 元schemaを破壊しないようコピーしてから掃除
        import copy
        cleaned = copy.deepcopy(schema)
        _clean_schema(cleaned)

    _dlog("llm_call_start", {"context": context})
    # 入力を全文記録（プロンプト全文・schema全体）
    _dlog("llm_call_input_full", {
        "context": context,
        "has_schema": schema is not None,
        "has_pdf": bool(pdf_path),
        "pdf_path": pdf_path,
        "prompt_len": len(prompt),
        "prompt": prompt,                       # 截断なし
        "schema_in": _describe(schema),
        "schema_cleaned": _describe(cleaned),
    })
    try:
        # --- 1回目: schema付き ---
        kw = {"prompt": prompt}
        if cleaned:
            kw["schema"] = cleaned
        if pdf_path:
            # ★1改修: 画像化せずPDF書類のまま渡す。読取りはGemini固定（Claude等は落ちる）。
            kw["file_paths"] = [pdf_path]
            kw["model"] = GEMINI_PDF_MODEL
        res = await llm_call(**kw)
        _dlog_full("llm_call_raw_response", context, res, attempt=1)

        # --- 添付サイズ超過は決定論的エラー: 同じPDFで再試行しても必ず失敗するため即raise ---
        if pdf_path and isinstance(res, dict) and res.get("error"):
            _err_s = str(res.get("error", ""))
            if ("inline limit" in _err_s or "over the inline" in _err_s
                    or ("attachments total" in _err_s and "bytes" in _err_s)):
                _dlog("llm_call_attachment_too_large", {"context": context, "error_full": _err_s})
                raise PdfTooLargeError(_err_s)

        # --- エラーならschemaなしでfallback ---
        if isinstance(res, dict) and res.get("error") and schema:
            _dlog("llm_call_fallback_trigger", {"context": context,
                  "error_full": str(res.get("error", ""))})
            fallback_prompt = (prompt +
                "\n\n出力は有効なJSONオブジェクトだけ。説明・markdown・前置きは一切禁止。"
                "rawなJSONのみを返す。")
            _dlog("llm_call_fallback_prompt", {"context": context,
                  "prompt_len": len(fallback_prompt), "prompt": fallback_prompt})
            fb_kw = {"prompt": fallback_prompt}
            if pdf_path:
                fb_kw["file_paths"] = [pdf_path]
                fb_kw["model"] = GEMINI_PDF_MODEL
            res = await llm_call(**fb_kw)
            _dlog_full("llm_call_raw_response", context, res, attempt=2)
            if isinstance(res, dict) and res.get("error"):
                _dlog("llm_call_fallback_failed", {"context": context,
                      "error_full": str(res.get("error", ""))})

        duration = round(time.time() - t0, 3)
        _dlog("llm_call_finish", {"context": context, "duration": duration,
              "keys": list(res.keys()) if isinstance(res, dict) else None,
              "res_type": type(res).__name__})

        if schema:
            data = res.get("data") if isinstance(res, dict) else None
            # 詳細: resの全トップレベルキーとdataの構造を記録
            _dlog("llm_call_data_extraction", {
                "context": context,
                "res_keys": list(res.keys()) if isinstance(res, dict) else None,
                "res_has_data": "data" in res if isinstance(res, dict) else False,
                "res_has_text": "text" in res if isinstance(res, dict) else False,
                "res_has_error": "error" in res if isinstance(res, dict) else False,
                "data_type": type(data).__name__,
                "data_keys": list(data.keys())[:20] if isinstance(data, dict) else None,
                "data_exec_ratio": data.get("exec_ratio") if isinstance(data, dict) else "N/A",
            })
            _dlog_full("llm_call_data_field", context, data, present=data is not None)

            if data is None and isinstance(res, dict) and "text" in res:
                text = (res.get("text") or "")
                _dlog("llm_call_text_field", {"context": context,
                      "text_len": len(text), "text": text})        # 截断なし
                # 複数戦略でJSON救出を試み、全部記録する
                data = _rescue_json(text, context)

            if data is None:
                _dlog("llm_call_data_missing", {"context": context,
                      "res": _describe(res), "error": str(res.get("error", "?")) if isinstance(res, dict) else "?"})
                raise KeyError(f"LLM data missing. context={context}")

            _dlog_full("llm_call_return", context, data)
            return data
        else:
            out = {"text": res.get("text", "") if isinstance(res, dict) else str(res)}
            _dlog_full("llm_call_return", context, out)
            return out
    except Exception as e:
        import traceback
        _dlog("llm_call_exception", {"context": context, "error": str(e),
              "traceback": traceback.format_exc()})
        raise


def _rescue_json(text: str, context: str = ""):
    """textから複数戦略でJSONを救出。各戦略の試行と結果を全部ログに残す。"""
    text = (text or "").strip()
    strategies = []

    # 戦略1: そのままloads
    strategies.append(("raw_loads", text))
    # 戦略2: ```json ... ``` フェンス除去
    m = re.search(r'```(?:json)?\s*(.*?)```', text, re.DOTALL)
    if m:
        strategies.append(("strip_fence", m.group(1).strip()))
    # 戦略3: 最初の { から最後の } まで（貪欲）
    s, e = text.find("{"), text.rfind("}")
    if s != -1 and e != -1 and e > s:
        strategies.append(("brace_greedy", text[s:e+1]))
    # 戦略4: 最初の { から最初の対応する } まで（非貪欲・簡易バランス）
    if s != -1:
        depth = 0
        for i in range(s, len(text)):
            if text[i] == "{": depth += 1
            elif text[i] == "}":
                depth -= 1
                if depth == 0:
                    strategies.append(("brace_balanced", text[s:i+1]))
                    break

    for name, cand in strategies:
        try:
            parsed = json.loads(cand)
            _dlog("rescue_attempt", {"context": context, "strategy": name,
                  "ok": True, "cand_len": len(cand), "result": _describe(parsed)})
            return parsed
        except Exception as pe:
            _dlog("rescue_attempt", {"context": context, "strategy": name,
                  "ok": False, "cand_len": len(cand), "cand": cand, "error": str(pe)})
    _dlog("rescue_all_failed", {"context": context, "text": text})
    return None

# ============================================================
# グローバル状態（セッション変数）
# ============================================================
# ★2改修: 毎ターン exec(f.read(), globals()) でこのファイルを丸ごと読み直すため、
# ここで無条件に {} を代入すると前ターンmain()が保存した状態が消え、深掘り(handle_turn2)が
# 「ptc_stateが空です」で必ず止まっていた。agent_platformは同一セッション内で変数を持ち越すので、
# 「既に有れば残す」に変えて前ターンの選定結果を引き継ぐ。
if "ptc_state" not in globals():
    ptc_state: dict = {}

# ============================================================
# ユーティリティ
# ============================================================

async def call_with_retry(tool_name: str, args: dict, max_attempts: int = 4,
                           base_wait: float = 1.0, max_wait: float = 8.0,
                           deadline_sec: float = None, retry_context: str = ""):
    """
    OneDrive系ツールのリトライ。診断モード: 全試行を截断ゼロで記録する。
    - 各試行の生レスポンス・例外・traceback・所要・待機・引数を全て残す
    - 500等のサーバーエラー(error_class)も成功dictも、丸ごとログ
    - 指数バックオフ（base_wait * 2^n、上限max_wait）
    - deadline_sec を指定すると、累積時間がそれを超えたら以降のリトライを打ち切る
    """
    import traceback as _tb
    t_call_start = time.time()
    last_exc = None
    last_res = None

    _dlog("retry_call_start", {"rc": retry_context, 
        "tool": tool_name,
        "args_full": _describe(args),
        "max_attempts": max_attempts,
        "base_wait": base_wait, "max_wait": max_wait,
        "deadline_sec": deadline_sec,
    })

    for attempt in range(max_attempts):
        attempt_t0 = time.time()
        elapsed_total = round(attempt_t0 - t_call_start, 3)

        # デッドライン超過チェック（リトライ前）
        if deadline_sec is not None and elapsed_total >= deadline_sec:
            _dlog("retry_deadline_exceeded", {"rc": retry_context, 
                "tool": tool_name, "attempt": attempt,
                "elapsed_total": elapsed_total, "deadline_sec": deadline_sec,
            })
            break

        _dlog("retry_attempt_start", {"rc": retry_context, 
            "tool": tool_name, "attempt": attempt + 1,
            "elapsed_total_before": elapsed_total,
        })
        try:
            res = await tool_bridge(tool_name, args)
            last_res = res
            attempt_dur = round(time.time() - attempt_t0, 3)

            # 生レスポンスを丸ごと記録（成功・失敗問わず・截断なし）
            _dlog_full("retry_raw_response", f"{tool_name}#attempt{attempt+1}", res,
                       attempt=attempt + 1, attempt_dur=attempt_dur)

            # サーバーエラー判定の材料を全部記録してから判定
            err_class = res.get("error_class") if isinstance(res, dict) else None
            err_msg   = res.get("error") if isinstance(res, dict) else None
            http_code = res.get("mercury_last_http_status_code") if isinstance(res, dict) else None
            auth_req  = res.get("auth_refresh_required") if isinstance(res, dict) else None
            log_id    = res.get("log_id") if isinstance(res, dict) else None
            successful = res.get("successful") if isinstance(res, dict) else None
            _dlog("retry_response_diag", {"rc": retry_context, 
                "tool": tool_name, "attempt": attempt + 1,
                "has_error_class": bool(err_class),
                "error_class": err_class,
                "error_full": err_msg,                 # 截断なし
                "mercury_http_code": http_code,
                "auth_refresh_required": auth_req,
                "log_id": log_id,
                "successful": successful,
                "attempt_dur": attempt_dur,
            })

            if isinstance(res, dict) and res.get("error_class"):
                # 500等。例外化してリトライへ。ただしメッセージは截断しない
                raise RuntimeError(f"server error_class={err_class} msg={err_msg}")

            # 成功
            _dlog("retry_success", {"rc": retry_context, 
                "tool": tool_name, "attempt": attempt + 1,
                "total_dur": round(time.time() - t_call_start, 3),
            })
            return res

        except Exception as e:
            last_exc = e
            attempt_dur = round(time.time() - attempt_t0, 3)
            # 例外を全文＋traceback＋型まで残す
            _dlog("retry_attempt_failed", {"rc": retry_context, 
                "tool": tool_name, "attempt": attempt + 1,
                "error_type": type(e).__name__,
                "error_full": str(e),                  # 截断なし
                "traceback": _tb.format_exc(),
                "attempt_dur": attempt_dur,
                "elapsed_total_after": round(time.time() - t_call_start, 3),
                "last_res": _describe(last_res),       # 直近レスポンスも丸ごと
            })
            if attempt == max_attempts - 1:
                _dlog("retry_exhausted", {"rc": retry_context, 
                    "tool": tool_name, "attempts": max_attempts,
                    "total_dur": round(time.time() - t_call_start, 3),
                    "final_error_full": str(e),
                })
                raise
            wait = min(base_wait * (2 ** attempt), max_wait)
            # デッドラインを考慮して待機を詰める
            if deadline_sec is not None:
                remain = deadline_sec - (time.time() - t_call_start)
                if remain <= 0:
                    _dlog("retry_deadline_no_wait", {"rc": retry_context, "tool": tool_name, "attempt": attempt + 1})
                    break
                wait = min(wait, max(0.0, remain))
            _dlog("retry_backoff", {"rc": retry_context, 
                "tool": tool_name, "attempt": attempt + 1,
                "wait_sec": round(wait, 3),
                "next_attempt": attempt + 2,
            })
            await asyncio.sleep(wait)

    _dlog("retry_giving_up", {"rc": retry_context, 
        "tool": tool_name,
        "total_dur": round(time.time() - t_call_start, 3),
        "last_error_full": str(last_exc) if last_exc else None,
        "last_res": _describe(last_res),
    })
    if last_exc:
        raise last_exc
    return last_res


def fmt_yen(n: int) -> str:
    if n >= 10000:
        return f"{n // 10000:,}万円"
    return f"{n:,}円"


# ============================================================
# score_media（score_media_56col_v2_1.py の中身）
# ============================================================

def _load_master(excel_path: str) -> pd.DataFrame:
    wb = load_workbook(excel_path, data_only=True)
    ws = wb.active
    data = list(ws.values)
    if len(data) < 2:
        raise ValueError("Excelに2行以上のデータがありません")
    headers = [str(h) if h else f"col_{i}" for i, h in enumerate(data[0])]
    df = pd.DataFrame(data[1:], columns=headers)
    if "メディア名称" in df.columns:
        df = df[df["メディア名称"].notna() & (df["メディア名称"].astype(str).str.strip() != "")]
    return df


def _apply_single_filter(df: pd.DataFrame, col: str, condition: dict) -> pd.Series:
    op = condition["op"]
    val = condition["value"]
    if col not in df.columns:
        return pd.Series(True, index=df.index)
    if op == "contains":
        return df[col].astype(str).str.contains(str(val), na=False)
    elif op == "not_contains":
        return ~df[col].astype(str).str.contains(str(val), na=False)
    elif op == "in":
        if isinstance(val, list):
            pattern = "|".join([str(v).replace("(", r"\(").replace(")", r"\)") for v in val])
            return df[col].astype(str).str.contains(pattern, na=False)
        return df[col].astype(str).str.contains(str(val), na=False)
    elif op == "==":
        return df[col].astype(str) == str(val)
    elif op == "!=":
        return df[col].astype(str) != str(val)
    elif op in ("<=", ">=", "<", ">"):
        numeric_col = pd.to_numeric(df[col], errors="coerce")
        try:
            val_num = float(val)
        except (TypeError, ValueError):
            return pd.Series(True, index=df.index)
        ops = {"<=": numeric_col.__le__, ">=": numeric_col.__ge__,
               "<": numeric_col.__lt__, ">": numeric_col.__gt__}
        return ops[op](val_num)
    return pd.Series(True, index=df.index)


def _apply_filters(df: pd.DataFrame, filters: dict) -> pd.DataFrame:
    mask = pd.Series(True, index=df.index)
    for col, condition in filters.items():
        if col not in df.columns:
            print(f"  WARN: フィルタ対象 '{col}' がマスターにありません。スキップ。")
            continue
        mask &= _apply_single_filter(df, col, condition)
    return df[mask].copy()


def _report_fallback(original_df: pd.DataFrame, filters: dict) -> str:
    """フォールバック分析テキストを返す（llm_callに渡すため文字列で返す）"""
    lines = []
    lines.append("◆ 各フィルタを1つずつ外した場合の通過件数:")
    for skip_key in filters:
        remaining = {k: v for k, v in filters.items() if k != skip_key}
        mask = pd.Series(True, index=original_df.index)
        for col, cond in remaining.items():
            if col in original_df.columns:
                mask &= _apply_single_filter(original_df, col, cond)
        cond = filters[skip_key]
        lines.append(f"  「{skip_key} {cond['op']} {cond['value']}」を外す → {int(mask.sum())}件通過")
    return "\n".join(lines)


# 規模カラムのみ実列名にカッコ付きサフィックスが付き、AIが短縮名で指定すると
# 完全一致に外れて score_media が軸を黙って落とす（旧#4のコード側）。短縮名→実列名のエイリアス。
_COL_ALIASES = {
    "規模_エンタープライズ": "規模_エンタープライズ(1000名以上)",
    "規模_SMB":            "規模_SMB(100-999名)",
    "規模_スタートアップ零細": "規模_スタートアップ零細(100名未満)",
}


def _canonical_col(col: str, columns) -> str:
    """指定列名をマスタ実列名へ正規化。完全一致→エイリアス→前方一致(一意)の順。無ければそのまま返す。"""
    if col in columns:
        return col
    alias = _COL_ALIASES.get(col)
    if alias and alias in columns:
        return alias
    # 前方一致が一意ならそれを採用（将来サフィックスが変わっても耐える保険）
    hits = [c for c in columns if c.startswith(col)]
    if len(hits) == 1:
        return hits[0]
    return col


def score_media(excel_path: str, params: dict) -> tuple[list, list, str]:
    """
    Returns:
        result_data: [{"メディア名称": ..., "file_name": ..., "score": ...}, ...]
        weight_values: [{"メディア名称": ..., "<col>": float, ...}, ...]
        fallback_analysis: str（0件時のフォールバック分析テキスト、通常は空）
    """
    filters   = params.get("filters", {})
    weights   = params.get("weights", {})
    sort_by   = params.get("sort_by")
    sort_order = params.get("sort_order", "asc")
    top_n     = params.get("top_n", 7)
    skip_top  = params.get("skip_top", 0)

    SCALE_SENSITIVE = {"最低出稿金額", "最高出稿金額", "月間PV数", "月間UU数",
                       "会員数", "PV保証_最小", "PV保証_最大", "CPL目安"}
    bad = [k for k in weights if k in SCALE_SENSITIVE]
    if bad:
        raise ValueError(f"weights に絶対値カラムは指定不可: {bad}  → sort_by を使ってください")

    _t_load = time.time()
    df_orig = _load_master(excel_path)
    _dlog("score_load_master", {"rows": len(df_orig), "dur": round(time.time() - _t_load, 4)})
    print(f"マスター読み込み: {len(df_orig)}件")

    # ★列名正規化: 短縮名(規模_SMB 等)をマスタ実列名(規模_SMB(100-999名))へ一括で揃える。
    # ここで正規化すれば以降の filters/weights/sort_by/weight_values 全てが実列名で一致する（旧#4のコード側を解消）。
    _cols = list(df_orig.columns)
    _norm_filters = {_canonical_col(c, _cols): v for c, v in filters.items()}
    _norm_weights = {_canonical_col(c, _cols): v for c, v in weights.items()}
    _remapped = ([f"{c}→{_canonical_col(c, _cols)}" for c in filters if _canonical_col(c, _cols) != c]
                 + [f"{c}→{_canonical_col(c, _cols)}" for c in weights if _canonical_col(c, _cols) != c])
    if sort_by:
        _sb_norm = _canonical_col(sort_by, _cols)
        if _sb_norm != sort_by:
            _remapped.append(f"{sort_by}→{_sb_norm}")
        sort_by = _sb_norm
    filters, weights = _norm_filters, _norm_weights
    if _remapped:
        _dlog("col_name_normalized", {"remapped": _remapped})
        print(f"  列名正規化: {', '.join(_remapped)}")

    # #6改修: 除外媒体（「実施済みなので除く」等）を最初に名前照合で行ごと弾く。
    # 正規化(NFKC+小文字+空白除去)→ ①ユーザー名⊂マスタ名 ②マスタ名⊂ユーザー文字列 ③difflib類似 の3段。
    # どれにも一致しない指定は黙って無視せずWARNで申告する（#6の本質=黙殺の再発防止）。
    exclude_media = params.get("exclude_media") or []
    if exclude_media and "メディア名称" in df_orig.columns:
        def _norm_name(s: str) -> str:
            return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(s)).lower())
        _norm_col = df_orig["メディア名称"].astype(str).map(_norm_name)
        _drop_mask = pd.Series(False, index=df_orig.index)
        _unmatched = []
        for _ex in exclude_media:
            _exn = _norm_name(_ex)
            if not _exn:
                continue
            _m = (_norm_col.str.contains(re.escape(_exn), na=False)          # ①ユーザー名がマスタ名に含まれる
                  | _norm_col.map(lambda n: bool(n) and n in _exn))          # ②マスタ名がユーザー文字列に含まれる
            if not _m.any():                                                 # ③類似照合（綴りゆれ保険）
                _close = difflib.get_close_matches(_exn, list(_norm_col), n=1, cutoff=0.8)
                if _close:
                    _m = _norm_col == _close[0]
            if _m.any():
                _drop_mask |= _m
            else:
                _unmatched.append(_ex)
        if _drop_mask.any():
            _dropped = df_orig.loc[_drop_mask, "メディア名称"].astype(str).tolist()
            print(f"  除外媒体（実施済み等）: {', '.join(_dropped)}")
            _dlog("exclude_media_applied", {"requested": exclude_media, "dropped": _dropped})
            df_orig = df_orig[~_drop_mask].copy()
        if _unmatched:
            print(f"  WARN: 除外指定 {_unmatched} に一致する媒体が見つかりません（除外されずに残ります）")
            _dlog("exclude_media_unmatched", {"unmatched": _unmatched})

    _t_filt = time.time()
    df = _apply_filters(df_orig, filters) if filters else df_orig.copy()
    _dlog("score_filter", {"dur": round(time.time() - _t_filt, 4), "rows_after": len(df)})
    print(f"フィルタ後: {len(df)}件")

    if "メディア名称" in df.columns:
        df = df.drop_duplicates(subset="メディア名称", keep="last")

    fallback_analysis = ""
    fallback_used = False
    if df.empty and filters:
        fallback_analysis = _report_fallback(df_orig, filters)
        print("⚠ フォールバック: 全フィルタを外してスコアリング")
        df = df_orig.copy()
        if "メディア名称" in df.columns:
            df = df.drop_duplicates(subset="メディア名称", keep="last")
        fallback_used = True

    # スコアリング or ソート
    # --- バリデーション: sort_by + weights 共存禁止 ---
    if sort_by and weights:
        # weightsの意図を「足切りフィルタ」に変換してからweightsを空にする
        _dlog("weights_to_filter", {
            "reason": "sort_by and weights both specified; converting weights to filters",
            "sort_by": sort_by, "weights": weights,
        })
        print(f"  WARN: sort_by='{sort_by}' と weights が同時指定。weightsをフィルタに変換します")
        for col, w in weights.items():
            if col in df.columns and col not in filters:
                filters[col] = {"op": ">=", "value": 10}  # 10%以上で足切り
                print(f"    → フィルタ追加: {col} >= 10%")
        # フィルタ変換後に再度フィルタ適用
        df = _apply_filters(df_orig, filters) if filters else df.copy()
        if "メディア名称" in df.columns:
            df = df.drop_duplicates(subset="メディア名称", keep="last")
        print(f"    → フィルタ再適用後: {len(df)}件")
        weights = {}  # weightsクリア

    if sort_by:
        if weights:
            print(f"  WARN: sort_by='{sort_by}' 指定のため weights は無視")
        ascending = (sort_order != "desc")
        if sort_by in df.columns:
            df["_score"] = pd.to_numeric(df[sort_by], errors="coerce")
            df = df.sort_values("_score", ascending=ascending, na_position="last")
        else:
            print(f"  WARN: sort_by='{sort_by}' がマスターにない")
            df["_score"] = 0
    elif weights:
        df["_score"] = 0.0
        for col, w in weights.items():
            if ":" in col:
                print(f"  WARN: '{col}' は旧形式キー。スキップ")
                continue
            if col not in df.columns:
                print(f"  WARN: '{col}' がマスターにない。スキップ")
                continue
            df["_score"] += pd.to_numeric(df[col], errors="coerce").fillna(0) * float(w)
        df = df.sort_values("_score", ascending=False)
    else:
        df["_score"] = 0

    # skip_top
    if skip_top > 0:
        df = df.iloc[skip_top:skip_top + top_n]
    else:
        df = df.head(top_n)

    # 全カラムをresult_dataに含める（remaining_dataのHTML描画に必要）
    result_data = []
    for _, row in df.iterrows():
        rec = {}
        for c in df.columns:
            if c.startswith("_"):  # _scoreなどの内部カラムはスキップ
                continue
            rec[c] = str(row[c]) if pd.notna(row[c]) else ""
        rec["score"] = float(row.get("_score", 0))
        result_data.append(rec)

    # weight_values
    weight_values = []
    if weights and not sort_by:
        weight_cols = [k for k in weights if ":" not in k and k in df.columns]
        for _, row in df.iterrows():
            rec = {"メディア名称": str(row.get("メディア名称", "?")).split(";")[0].strip()}
            for col in weight_cols:
                v = pd.to_numeric(row.get(col), errors="coerce")
                rec[col] = None if pd.isna(v) else float(v)
            weight_values.append(rec)

    print(f"結果: {len(result_data)}件  fallback={fallback_used}")
    return result_data, weight_values, fallback_analysis


# ============================================================
# extract_pdf（extract_pdf.py の中身 / pdfplumber）
# ============================================================

def extract_pdf(pdf_path: str) -> str:
    """PDFを本文+表をMarkdown形式でテキスト化して返す。診断ログ付き。"""
    output = io.StringIO()
    path = Path(pdf_path)
    output.write(f"# {path.stem}\n\n")
    _dlog("extract_pdf_start", {"path": pdf_path, "stem": path.stem})
    page_diag = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            npages = len(pdf.pages)
            output.write(f"総ページ数: {npages}\n\n")
            for page_num, page in enumerate(pdf.pages, start=1):
                output.write(f"\n{'=' * 60}\n")
                output.write(f"## ページ {page_num} / {npages}\n")
                output.write(f"{'=' * 60}\n\n")

                tables = page.find_tables()
                table_bboxes = [t.bbox for t in tables]

                def not_in_tables(obj):
                    v_mid = (obj["top"] + obj["bottom"]) / 2
                    h_mid = (obj["x0"] + obj["x1"]) / 2
                    for bbox in table_bboxes:
                        x0, top, x1, bottom = bbox
                        if x0 <= h_mid <= x1 and top <= v_mid <= bottom:
                            return False
                    return True

                text = page.filter(not_in_tables).extract_text() or ""
                if text.strip():
                    output.write("### 本文\n\n")
                    output.write(text.strip())
                    output.write("\n\n")

                if tables:
                    output.write(f"\n### 表 ({len(tables)}件)\n\n")
                    for ti, table in enumerate(tables, 1):
                        data = table.extract()
                        if data and len(data) > 0:
                            output.write(f"#### 表 {ti}\n\n")
                            header = [str(c).replace("\n", " ").strip() if c else "" for c in data[0]]
                            output.write("| " + " | ".join(header) + " |\n")
                            output.write("|" + "|".join(["---"] * len(header)) + "|\n")
                            for row in data[1:]:
                                row = [str(c).replace("\n", " ").strip() if c else "" for c in row]
                                output.write("| " + " | ".join(row) + " |\n")
                            output.write("\n")
                page_diag.append({"page": page_num, "text_chars": len(text), "tables": len(tables)})
        _dlog("extract_pdf_finish", {"path": pdf_path, "pages": page_diag,
              "total_chars": len(output.getvalue())})
    except Exception as e:
        import traceback
        output.write(f"\n[ERROR: PDF抽出失敗: {e}]\n")
        _dlog("extract_pdf_error", {"path": pdf_path, "error": str(e),
              "traceback": traceback.format_exc(), "pages_done": page_diag})

    return output.getvalue()


# ============================================================
# get_fallback_info（get_fallback_info_56col_v1.py の中身）
# ============================================================

def get_fallback_info(excel_path: str, failed_file_names: list) -> dict:
    """
    Returns: {media_name: "独自価値テキスト", ...}
    """
    df = _load_master(excel_path)
    result = {}
    for fname in failed_file_names:
        match = df[df["file_name"].astype(str) == fname]
        if match.empty:
            result[fname] = "（マスターに登録情報なし）"
            continue
        row = match.iloc[0]
        media_name = str(row.get("メディア名称", "?")).split(";")[0].strip()
        val = row.get("メディアの独自価値", "")
        result[media_name] = str(val) if pd.notna(val) else "（独自価値の登録なし）"
    return result


# ============================================================
# get_pdf_urls（get_pdf_urls.py の中身）
# ============================================================

def get_pdf_urls(excel_path: str, file_names: list) -> dict:
    """
    Returns: {file_name: url, ...}
    """
    df = _load_master(excel_path)
    result = {}
    for fname in file_names:
        match = df[df["file_name"].astype(str) == fname]
        if match.empty:
            result[fname] = "（URLなし）"
            continue
        row = match.iloc[0]
        url = row.get("PDFファイルURL", "")
        result[fname] = str(url) if pd.notna(url) else "（URLなし）"
    return result


# ============================================================
# validate_budget（validate_budget.py の中身）
# ============================================================

def validate_budget(user_budget: int, desired_count: int, attempt: int,
                    media_prices: list, score_params: dict) -> dict:
    """
    Returns:
        {
            "verdict": "ok" | "refill" | "ok_shortage",
            "within_budget": [...],
            "over_budget": [...],
            "shortage": int,
            "refill_skip_top": int,
            "refill_top_n": int,
            "summary": str
        }
    """
    within, over = [], []
    for m in media_prices:
        price = m.get("tieup_min_price", 0)
        if price <= 0:
            over.append({**m, "reason": "価格不明"})
        elif price <= user_budget:
            within.append(m)
        else:
            over.append({**m, "reason": "予算超過"})

    n_within = len(within)
    shortage = max(0, desired_count - n_within) if desired_count > 0 else 0
    prev_top_n = score_params.get("previous_top_n", len(media_prices))

    lines = [
        f"ユーザー予算: {fmt_yen(user_budget)}",
        f"希望件数: {desired_count if desired_count > 0 else '未指定'}",
        f"試行: {attempt}回目",
        f"予算内: {n_within}件  超過: {len(over)}件",
    ]

    # ケース判定
    if not over:
        if desired_count > 0 and n_within > desired_count:
            verdict = "ok"
            lines.append(f"→ 予算内{n_within}件から上位{desired_count}件を採用")
            within = within[:desired_count]
        else:
            verdict = "ok"
            lines.append("→ 全件予算内。そのまま提案へ")
    elif desired_count == 0:
        verdict = "ok"
        lines.append(f"→ 件数未指定のため補充なし。予算内{n_within}件で提案")
    elif n_within >= desired_count:
        verdict = "ok"
        lines.append(f"→ 予算内{n_within}件が目標{desired_count}件を満たす")
        within = within[:desired_count]
    elif attempt >= 2:
        verdict = "ok_shortage"
        lines.append(f"→ 補充1回上限。{n_within}件（{shortage}件不足）のまま提案")
    else:
        verdict = "refill"
        lines.append(f"→ {shortage}件不足。補充実行（skip_top={prev_top_n}, top_n={shortage}）")

    return {
        "verdict": verdict,
        "within_budget": within,
        "over_budget": over,
        "shortage": shortage,
        "refill_skip_top": prev_top_n,
        "refill_top_n": shortage,
        "summary": "\n".join(lines),
    }


# ============================================================
# OneDriveマッチング
# ============================================================

def _match_files(result_data: list, list_items: list) -> tuple[list, list]:
    """
    difflib で result_data[file_name] × list_items[name] を類似度マッチング。

    Returns:
        matched: [{"media_name": ..., "file_name": ..., "id": ..., "score": float}, ...]
        failed:  [{"media_name": ..., "file_name": ..., "candidates": [...], "reason": str}, ...]

    複数候補（score >= 0.8 が2件以上）は failed に入れてAI判断へ。
    最高スコアが0.8未満も failed へ。
    """
    matched = []
    failed = []

    for rec in result_data:
        # キーは常にsplit後の正規名で統一（sandbox_paths/failed_list/share_urlsのキーになり、
        # 下流のon_dl_complete・step3・step4はsplit後の名前で照合するため。未分割だと";"入り媒体が黙って消える）
        media_name = rec.get("メディア名称", "").split(";")[0].strip()
        file_name  = rec.get("file_name", "")

        candidates = []
        for item in list_items:
            item_name = item.get("name", "")
            ratio = difflib.SequenceMatcher(None, unicodedata.normalize("NFC", file_name), unicodedata.normalize("NFC", item_name)).ratio()
            if ratio >= 0.8:
                candidates.append({"id": item["id"], "name": item_name, "score": round(ratio, 3)})

        candidates.sort(key=lambda x: x["score"], reverse=True)

        if len(candidates) == 1:
            matched.append({
                "media_name": media_name,
                "file_name": file_name,
                "id": candidates[0]["id"],
                "match_name": candidates[0]["name"],
                "score": candidates[0]["score"],
            })
        elif len(candidates) >= 2:
            # 複数候補 → AI判断へ
            failed.append({
                "media_name": media_name,
                "file_name": file_name,
                "candidates": candidates,
                "reason": "複数候補",
            })
        else:
            # 0件 → AI判断へ（低スコア候補を参考として渡す）
            low = sorted(
                [{"id": it["id"], "name": it.get("name",""), "score": round(difflib.SequenceMatcher(None, unicodedata.normalize("NFC", file_name), unicodedata.normalize("NFC", it.get("name",""))).ratio(), 3)} for it in list_items],
                key=lambda x: x["score"], reverse=True
            )[:5]
            failed.append({
                "media_name": media_name,
                "file_name": file_name,
                "candidates": low,
                "reason": "該当なし（閾値0.8未満）",
            })

    return matched, failed


async def _resolve_failed_by_ai(failed: list, list_items: list) -> tuple[list, list]:
    """並列でAI判断を走らせ、時間を節約する。"""
    if not failed:
        return [], []

    all_names_str = "\n".join(f"- {it['name']}" for it in list_items)

    async def _resolve_single(item):
        _dlog("resolve_single_start", {"media_name": item["media_name"], "file_name": item.get("file_name","")})
        prompt = f"""探したいファイル: {item['media_name']} ({item['file_name']})
OneDrive候補: {json.dumps(item['candidates'], ensure_ascii=False)}
全ファイル一覧: {all_names_str}
明らかに同じメディアのPDFがあればidとnameを返し、なければ give_up。
{{"action": "resolve" or "give_up", "id": "...", "name": "...", "reason": "..."}}"""
        try:
            data = await _llm_call_safe(
                prompt=prompt,
                schema={
                    "type": "object",
                    "properties": {
                        "action": {"type": "string", "enum": ["resolve", "give_up"]},
                        "id": {"type": "string"}, "name": {"type": "string"}, "reason": {"type": "string"}
                    },
                    "required": ["action", "reason"],
                },
                context=f"ai_resolve:{item['media_name']}"
            )
            _dlog_full("resolve_single_received", f"ai_resolve:{item['media_name']}", data)
            _dlog("resolve_single_finish", {"media_name": item["media_name"], "action": data.get("action","?")})
            return item, data, None
        except Exception as e:
            import traceback
            _dlog("resolve_single_error", {"media_name": item["media_name"],
                  "error": str(e), "traceback": traceback.format_exc()})
            return item, None, str(e)

    tasks = [_resolve_single(item) for item in failed]
    _dlog("gather_start", {"type": "resolve", "count": len(tasks)})
    results = await asyncio.gather(*tasks)
    _dlog("gather_finish", {"type": "resolve", "count": len(results)})

    resolved, give_up = [], []
    for item, data, error in results:
        if data and data["action"] == "resolve" and data.get("id"):
            resolved.append({
                "media_name": item["media_name"], "file_name": item["file_name"],
                "id": data["id"], "match_name": data["name"], "ai_resolved": True
            })
        else:
            reason = data["reason"] if data else (error or "Failed")
            give_up.append({**item, "ai_reason": reason})
    return resolved, give_up


# ============================================================
# STEP2: OneDrive PDF取得
# ============================================================

async def step2_get_pdfs(result_data: list, on_dl_complete=None) -> tuple:
    """並列ダウンロードで高速化。"""
    print("\n[STEP2] OneDrive取得開始")
    _t_list = time.time()
    try:
        list_resp = await asyncio.wait_for(
            call_with_retry("ONE_DRIVE_LIST_FOLDER_CHILDREN", {
                "folder_path": "/メディア", "use_me_drive": True, "select": ["id", "name"],
            }, max_attempts=10, base_wait=2.0, max_wait=2.0,
               deadline_sec=15.0, retry_context="list"),
            timeout=20.0)
    except Exception as e:
        import traceback
        _dlog("LIST_failed", {"error": str(e), "traceback": traceback.format_exc(),
              "dur": round(time.time() - _t_list, 3)})
        list_resp = {}
    list_items = list_resp.get("data", {}).get("value", []) if isinstance(list_resp, dict) else []
    _dlog_full("LIST_raw_response", "step2", list_resp)
    _dlog("LIST_result", {"count": len(list_items), "names": [it.get("name") for it in list_items],
          "dur": round(time.time() - _t_list, 3)})

    _t_match = time.time()
    matched, failed = _match_files(result_data, list_items)
    _dlog("difflib_timing", {"dur": round(time.time() - _t_match, 4)})
    _dlog("difflib_match", {"matched": [{"media": m["media_name"], "file": m.get("match_name"), "score": m.get("score")} for m in matched], "failed_count": len(failed), "failed": [{"media": f["media_name"], "reason": f.get("reason")} for f in failed]})
    _dlog_full("difflib_matched_full", "step2", matched)
    _dlog_full("difflib_failed_full", "step2", failed)

    # DL設定
    DL_PER_ITEM_DEADLINE = 30.0
    DL_HARD_TIMEOUT      = 35.0

    async def _dl_single(item):
        media_name = item["media_name"]
        file_name  = item.get("match_name", item.get("file_name", ""))
        orig_fname = item.get("file_name", "")
        item_id    = item.get("id", "")
        t0 = time.time()
        _dlog("dl_single_start", {"media_name": media_name, "file_name": file_name,
              "item_id": item_id, "item_full": _describe(item)})
        try:
            dl_coro = asyncio.wait_for(
                call_with_retry("ONE_DRIVE_DOWNLOAD_FILE",
                                {"item_id": item_id, "file_name": file_name},
                                max_attempts=15, base_wait=2.0, max_wait=2.0,
                                deadline_sec=DL_PER_ITEM_DEADLINE, retry_context=f"dl:{media_name}"),
                timeout=DL_HARD_TIMEOUT)
            share_coro = call_with_retry("ONE_DRIVE_CREATE_LINK",
                                         {"item_id": item_id, "type": "view", "scope": "anonymous"},
                                         max_attempts=5, base_wait=2.0, max_wait=2.0,
                                         deadline_sec=10.0, retry_context=f"link:{media_name}")
            dl_res, share_res = await asyncio.gather(dl_coro, share_coro, return_exceptions=True)
            if isinstance(dl_res, Exception):
                raise dl_res
            _dlog_full("dl_single_raw_response", f"dl:{media_name}", dl_res)
            sp = (dl_res.get("data", {}).get("content", {}).get("s3url")
                  or dl_res.get("content", {}).get("s3url")
                  or dl_res.get("sandbox_path", "")) if isinstance(dl_res, dict) else ""
            # llm_callは相対パスのみ受け付ける
            sp = sp.replace("/opt/amazon/genesis1p-tools/var/workspace/", "")
            share_url = ""
            if isinstance(share_res, Exception):
                _dlog("share_link_error", {"media_name": media_name, "error": str(share_res)})
            elif isinstance(share_res, dict):
                _dlog_full("share_link_raw", f"share:{media_name}", share_res)
                share_url = (share_res.get("data", {}).get("link", {}).get("webUrl", "")
                             or share_res.get("link", {}).get("webUrl", "")
                             or share_res.get("data", {}).get("webUrl", "")
                             or "")
            _dlog("dl_single_finish", {"media_name": media_name, "has_sp": bool(sp),
                  "sp": sp, "share_url": share_url, "dur": round(time.time() - t0, 3),
                  "file_size_kb": round(os.path.getsize(sp)/1024, 1) if sp and os.path.exists(sp) else None})
            return media_name, sp, None, orig_fname, share_url
        except asyncio.TimeoutError:
            _dlog("dl_single_hard_timeout", {"media_name": media_name,
                  "timeout": DL_HARD_TIMEOUT, "dur": round(time.time() - t0, 3)})
            return media_name, None, f"hard timeout {DL_HARD_TIMEOUT}s", orig_fname, ""
        except Exception as e:
            import traceback
            _dlog("dl_single_error", {"media_name": media_name, "error": str(e),
                  "traceback": traceback.format_exc(), "dur": round(time.time() - t0, 3)})
            return media_name, None, str(e), orig_fname, ""

    async def _resolve_then_dl(failed_items, all_items):
        """resolve完了した媒体から即DLを開始し、結果を返す"""
        ai_resolved, give_up = await _resolve_failed_by_ai(failed_items, all_items)
        _dlog_full("ai_resolved_full", "step2", ai_resolved)
        _dlog_full("ai_giveup_full", "step2", give_up)
        # resolve成功分をDL
        dl_results = []
        if ai_resolved:
            dl_tasks_r = [_dl_single(item) for item in ai_resolved]
            dl_results = await asyncio.gather(*dl_tasks_r)
        return list(dl_results), give_up

    # --- difflib成功分を即DL開始 + resolve分を並列で走らせる ---
    _dlog("gather_start", {"type": "download", "count": len(matched) + len(failed),
          "matched_immediate": len(matched), "resolve_pending": len(failed),
          "per_item_deadline": DL_PER_ITEM_DEADLINE, "hard_timeout": DL_HARD_TIMEOUT})
    _t_dl = time.time()

    # difflib成功分のDLタスク
    dl_tasks_matched = [_dl_single(item) for item in matched]
    # resolve分のタスク（resolve→DLを内部で直列実行）
    resolve_task = asyncio.create_task(_resolve_then_dl(failed, list_items)) if failed else None

    # difflib成功分をas_completedで回収（パイプライン: DL完了次第proposal開始）
    sandbox_paths = {}
    share_urls = {}
    proposal_tasks = []
    failed_list = []

    for coro in asyncio.as_completed(dl_tasks_matched):
        media_name, sp, error, orig_fname, share_url = await coro
        if sp:
            sandbox_paths[media_name] = sp
        else:
            failed_list.append({"media_name": media_name, "file_name": orig_fname, "reason": f"DL Error: {error}"})
        if share_url:
            share_urls[media_name] = share_url
        if sp and on_dl_complete:
            task = on_dl_complete(media_name, sp, share_url, orig_fname)
            if task:
                proposal_tasks.append(task)
        _dlog("dl_pipeline_item", {"media_name": media_name, "has_sp": bool(sp),
              "proposals_started": len(proposal_tasks), "dur": round(time.time() - _t_dl, 3)})

    # resolve分の結果を回収
    if resolve_task:
        resolve_dl_results, give_up = await resolve_task
        for media_name, sp, error, orig_fname, share_url in resolve_dl_results:
            if sp:
                sandbox_paths[media_name] = sp
            else:
                failed_list.append({"media_name": media_name, "file_name": orig_fname, "reason": f"DL Error: {error}"})
            if share_url:
                share_urls[media_name] = share_url
            if sp and on_dl_complete:
                task = on_dl_complete(media_name, sp, share_url, orig_fname)
                if task:
                    proposal_tasks.append(task)
            _dlog("dl_pipeline_item", {"media_name": media_name, "has_sp": bool(sp),
                  "source": "resolve", "proposals_started": len(proposal_tasks),
                  "dur": round(time.time() - _t_dl, 3)})
        failed_list.extend([{"media_name": g["media_name"], "file_name": g["file_name"],
                            "reason": g.get("ai_reason")} for g in give_up])

    _dlog("gather_finish", {"type": "download", "count": len(matched) + len(failed),
          "phase_dur": round(time.time() - _t_dl, 3)})

    _dlog("STEP2_result", {"sandbox_paths": list(sandbox_paths.keys()), "failed_list": failed_list,
          "share_urls": share_urls, "proposal_tasks_started": len(proposal_tasks)})
    return sandbox_paths, failed_list, share_urls, list_items, proposal_tasks


# ============================================================
# STEP3: PDFテキスト抽出
# ============================================================

async def step3_extract_texts(sandbox_paths: dict, failed_list: list,
                                    excel_path: str, result_data: list) -> tuple[dict, dict, dict]:
    """非同期版抽出。"""
    print("\n[STEP3] PDFテキスト抽出開始 (Parallel)")
    
    async def _ext_single(media_name, path):
        _dlog("ext_single_start", {"media_name": media_name, "path": path})
        try:
            text = await asyncio.to_thread(extract_pdf, path)
        except Exception as e:
            import traceback
            _dlog("ext_single_error", {"media_name": media_name, "error": str(e), "traceback": traceback.format_exc()})
            return media_name, ""
        _dlog("ext_single_finish", {"media_name": media_name, "chars": len(text),
              "head": text[:1000], "tail": text[-500:] if len(text) > 500 else ""})
        return media_name, text

    tasks = [_ext_single(m, p) for m, p in sandbox_paths.items()]
    _dlog("gather_start", {"type": "extract_pdf", "count": len(tasks)})
    results = await asyncio.gather(*tasks)
    _dlog("gather_finish", {"type": "extract_pdf", "count": len(results)})
    pdf_texts = {m: t for m, t in results}
    _dlog("STEP3_pdf_texts", {"summary": [{"media": k, "chars": len(v)} for k, v in pdf_texts.items()]})

    fallback_texts = {}
    if failed_list:
        failed_fnames = [f["file_name"] for f in failed_list if "file_name" in f]
        fallback_texts = get_fallback_info(excel_path, failed_fnames)
    _dlog_full("fallback_texts_full", "step3", fallback_texts)

    file_names = [r.get("file_name", "") for r in result_data]
    pdf_urls = get_pdf_urls(excel_path, file_names)
    _dlog_full("pdf_urls_full", "step3", pdf_urls)

    return pdf_texts, fallback_texts, pdf_urls


# ============================================================
# STEP3続き: 提案文 + 最低単価 同時生成
# ============================================================

def _normalize_proposal(d: dict, mname: str = "") -> dict:
    """
    Geminiが返すキーのブレを、HTMLビルダーが読むフラットキーに吸収する。
    今回ログで観測した別名(text_advantage/summary/audience/sections入れ子等)を拾う。
    本命はプロンプトでキー固定すること。これはあくまで保険。
    """
    out = dict(d)

    # 1) もし旧来のネスト sections.* が来ていたら平坦化して吸収
    sec = d.get("sections")
    if isinstance(sec, dict):
        for k in ("overview", "readers", "menu_price", "lead_gen", "clients", "editorial"):
            if sec.get(k) not in (None, "", [], {}):
                out.setdefault("sec_" + k, sec[k])

    # 2) フラットな別名 → 正規キー（最初に見つかった非空を採用）
    aliases = {
        "sec_overview":   ["sec_overview", "overview", "summary", "text_advantage",
                           "description", "media_overview", "概要"],
        "sec_readers":    ["sec_readers", "readers", "audience", "text_readers",
                           "reader_profile", "読者プロフィール"],
        "sec_menu_price": ["sec_menu_price", "menu_price", "pricing", "text_pricing",
                           "menu", "ad_menu", "料金"],
        "sec_lead_gen":   ["sec_lead_gen", "lead_gen", "lead", "lead_generation"],
        "sec_clients":    ["sec_clients", "clients", "case_studies", "実績企業"],
        "sec_editorial":  ["sec_editorial", "editorial", "features", "compatibility",
                           "suitability", "target_match", "strength", "strengths",
                           "特記事項"],
        "reason":         ["reason", "proposal_reason", "proposal", "selection_reason",
                           "選定理由", "理由"],
    }
    for canon, names in aliases.items():
        if out.get(canon) in (None, "", [], {}):
            for nm in names:
                v = d.get(nm)
                if v not in (None, "", [], {}):
                    out[canon] = v
                    break

    # 3) 型の最低保証（配列であるべきものが文字列で来たら1要素配列に）
    for arr_key in ("tags", "sec_readers", "sec_menu_price", "ad_menus"):
        v = out.get(arr_key)
        if isinstance(v, str) and v.strip():
            out[arr_key] = [v.strip()]
        elif v in (None, "", {}):
            out[arr_key] = []
    # 文字列であるべきものが配列で来たら結合
    for str_key in ("sec_overview", "sec_lead_gen", "sec_clients", "sec_editorial", "reason", "category", "pv_ub", "guaranteed_pv", "price_disp"):
        v = out.get(str_key)
        if isinstance(v, list):
            out[str_key] = " / ".join(str(x) for x in v)
        elif v is None:
            out[str_key] = ""

    # unknown_fields のデフォルト補完
    if "unknown_fields" not in out or not isinstance(out.get("unknown_fields"), list):
        out["unknown_fields"] = []

    # 4) media_name は引数を正とする（LLMが言い換えても表示名を固定）
    if mname:
        out["media_name"] = mname

    _dlog("normalize_proposal", {
        "media_name": mname,
        "in_keys": list(d.keys()),
        "out_filled": {k: bool(out.get(k)) for k in
                       ("sec_overview","sec_readers","sec_menu_price",
                        "sec_lead_gen","sec_clients","sec_editorial","reason")},
    })
    return out


# --- フラットスキーマ（モジュールレベル定数） ---
_SECTION_SCHEMA = {
    "type": "object",
    "properties": {
        "media_name":    {"type": "string"},
        "category":      {"type": "string"},
        "pv_ub":         {"type": "string"},
        "exec_ratio":    {"type": "number"},
        "exec_ratio_note": {"type": "string"},
        "price_disp":    {"type": "string"},
        "price_int":     {"type": "integer"},
        "guaranteed_pv": {"type": "string"},
        "tags":          {"type": "array", "items": {"type": "string"}},
        "sec_overview":   {"type": "string"},
        "sec_readers":    {"type": "array", "items": {"type": "string"}},
        "sec_menu_price": {"type": "array", "items": {"type": "string"}},
        "sec_lead_gen":   {"type": "string"},
        "sec_clients":    {"type": "string"},
        "sec_editorial":  {"type": "string"},
        "reason":         {"type": "string"},
        "ad_menus":       {"type": "array", "items": {"type": "string"}},
        "unknown_fields": {"type": "array", "items": {"type": "string"}},
        "dbg_last_page": {"type": "integer"},
    },
    "required": ["media_name", "exec_ratio", "exec_ratio_note", "price_int",
                 "ad_menus", "unknown_fields", "dbg_last_page"],
}


def _shrink_pdf_for_llm(pdf_path: str, rec: dict) -> str:
    """添付上限超過PDFを、マスタの「参照ページ一覧」のページだけ抜いた縮小PDFにする。
    成功したら縮小PDFのパス、ページ情報なし・pypdf不在・失敗時は空文字を返す。"""
    try:
        pages = sorted({int(n) for n in re.findall(r"\d+", str(rec.get("参照ページ一覧", "")))})
        pages = [p for p in pages if 1 <= p <= 500][:30]   # 異常値と過剰ページ数の保険
        if not pages:
            _dlog("pdf_shrink_no_pages", {"pdf_path": pdf_path})
            return ""
        from pypdf import PdfReader, PdfWriter
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        picked = [p for p in pages if p <= len(reader.pages)]
        if not picked:
            return ""
        for p in picked:
            writer.add_page(reader.pages[p - 1])
        os.makedirs("tmp", exist_ok=True)
        out = "tmp/shrunk_" + os.path.basename(pdf_path)
        with open(out, "wb") as f:
            writer.write(f)
        _dlog("pdf_shrunk", {"orig": pdf_path, "out": out, "pages": picked,
              "size_kb": round(os.path.getsize(out) / 1024, 1)})
        return out
    except Exception as e:
        _dlog("pdf_shrink_error", {"pdf_path": pdf_path, "error": str(e)})
        return ""


async def _gen_single_proposal(rec, pdf_path, pdf_url, user_request, article_type, target_attribute):
    """1媒体の提案生成（トップレベル関数）。DL完了次第呼べる。"""
    mname = rec.get("メディア名称", "").split(";")[0].strip()
    fname = rec.get("file_name", "")
    url = pdf_url or "（URLなし）"
    has_pdf = bool(pdf_path)
    _dlog("gen_single_start", {
        "media_name": mname, "file_name": fname,
        "has_pdf": has_pdf, "pdf_path": pdf_path, "url": url,
        "pdf_size_kb": round(os.path.getsize(pdf_path)/1024, 1) if has_pdf and os.path.exists(pdf_path) else None,
        "pdf_pages": _count_pdf_pages(pdf_path) if has_pdf else None,
    })

    if not has_pdf:
        _dlog("gen_single_no_pdf", {"media_name": mname})
        # #3改修: 失敗を「要問合せ」と偽装せず、読取り失敗と明示する（正当な要問合せと区別）。
        fallback = {"media_name": mname, "category": "", "pv_ub": "—", "exec_ratio": -1,
                "price_disp": "資料取得失敗", "price_int": -1, "guaranteed_pv": "—", "tags": [],
                "sec_overview": "媒体資料PDFを取得できませんでした", "sec_readers": [], "sec_menu_price": [],
                "sec_lead_gen": "", "sec_clients": "", "sec_editorial": "",
                "reason": "媒体資料PDFを取得できず内容を読み取れませんでした（要再実行）。",
                "ad_menus": [], "unknown_fields": [], "_read_failed": True,
                "_file_name": fname, "_pdf_url": url}
        return fallback

    p = f"""あなたはメディア選定の提案作成AIです。下記1媒体について、提案レポートの各項目を埋めてください。

メディア名: {mname}
ユーザー要望: {user_request}
媒体資料URL: {url}

添付の媒体資料PDF（PDF書類として添付済み）を読んで、各項目を埋めてください。

【記入ルール（厳守）】
- 各文章フィールドは150〜300字で詳細に記述。数値・事実を省略せず、PDFに記載がある情報は全て含める。短すぎる回答は不可。
- 全てのフィールドで、記述する各文にPDFの出典ページを（NP）形式で付記する。例: "470万PV（3P）"、"製造業エンジニア向けメディア（2P）"。
- PDFに記載がない情報は創作禁止。以下のルールで記入し、フィールド名を unknown_fields に入れる:
  string → "—"
  number/integer → -1
  array → []
- price_int は「{article_type}1本の最低出稿額」を円単位の整数で。{article_type}の価格がPDFに無ければ-1（媒体全体の最安額や他メニューの単価で代用しない）。price_disp は表示用文字列（例 "¥400万"、不明は "要問合せ"）。
- ad_menus は{article_type}に関する全広告メニューの配列。各要素は1メニューにつき1文字列で、以下の情報をPDFから漏れなく含める:
  フォーマット: "メニュー名: ¥金額 | 詳細説明（NP）"
  詳細説明に含めるべき項目（PDFに記載があるもの全て）:
    ・PV保証数やリード獲得保証数
    ・掲載期間（○週間、○ヶ月）
    ・含まれるサービス（編集部取材、撮影、原稿制作、校正、画像制作等）
    ・配信チャネル（メルマガ配信付き、SNS拡散、他メディア転載等）
    ・ターゲティング条件（業種指定可、役職指定可、従業員規模指定可等）
    ・原稿仕様（文字数、ページ数、画像点数）
    ・掲載位置（トップページ、記事一覧、カテゴリページ等）
    ・オプション内容（動画埋め込み、アンケート設置、リターゲティング等）
  例: "スタンダードタイアップ: ¥1,500,000 | 5,000PV保証・4週間掲載・編集部取材＋撮影＋原稿制作込み・メルマガ1回配信付き・3,000字程度（15P）"
- exec_ratio は{target_attribute if target_attribute else '経営者・役員比率'}の数値(%)。PDFの読者プロフィールや業種分布のグラフ・表・テキストから必ず探すこと。円グラフや棒グラフの中の数値も読み取ること。見つからない場合は -1 にして、exec_ratio_note に「何ページまで確認したが該当する記載が見つからなかった」と理由を書くこと。0.0は「その属性の読者が0%」という意味であり、「見つからなかった」の意味で使ってはならない。
- pv_ub は "月間PV / UB"。片方不明 "4,720万 / —"、両方不明 "—"。
- tags は強みを表す短いラベルの配列2〜4個。
- sec_readers / sec_menu_price は配列。各要素1行。
- dbg_last_page: PDFの最後に読めたページ番号（整数）。デバッグ用。

【出力JSON（このキー名・この平坦構造で厳密に返す。入れ子オブジェクトは作らない）】
{{
  "media_name": "{mname}",
  "category": "媒体カテゴリ",
  "pv_ub": "月間PV / UB（NP）",
  "exec_ratio": 0,
  "exec_ratio_note": "見つからなかった場合の理由。見つかった場合は出典ページ番号（例: 5Pの円グラフより）",
  "price_disp": "要問合せ",
  "price_int": -1,
  "guaranteed_pv": "PV保証",
  "tags": ["強み1", "強み2"],
  "sec_overview": "メディア概要",
  "sec_readers": ["読者属性1（NP）", "読者属性2（NP）"],
  "sec_menu_price": ["広告メニュー・料金1（NP）"],
  "sec_lead_gen": "リード獲得の仕組み",
  "sec_clients": "実績企業",
  "sec_editorial": "編集企画・特記事項",
  "reason": "選定理由",
  "ad_menus": ["スタンダードタイアップ: ¥1,500,000 | 5,000PV保証・4週間掲載・編集部取材＋撮影＋原稿制作込み・メルマガ1回配信付き・3,000字程度（15P）"],
  "unknown_fields": ["sec_clients"],
  "dbg_last_page": 45
}}"""
    # 添付サイズ超過は決定論的失敗（実ログ: 13.39MBのPDFが旧実装10MB上限で5回リトライ×2コール=41秒浪費）。
    # 確実に超えるサイズは事前スキップ。実際の上限値はPdfTooLargeError（エラー文字列検知）で環境に適応。
    _pdf_too_large = False
    try:
        if has_pdf and os.path.exists(pdf_path) and os.path.getsize(pdf_path) > _LLM_ATTACH_HARD_LIMIT_BYTES:
            _dlog("pdf_over_hard_limit", {"media_name": mname,
                  "size_kb": round(os.path.getsize(pdf_path) / 1024, 1)})
            _pdf_too_large = True
    except Exception:
        pass

    try:
        if _pdf_too_large:
            raise PdfTooLargeError("precheck: over hard limit")
        d = await _llm_call_safe(p, schema=_SECTION_SCHEMA, context=f"proposal:{mname}",
                                 pdf_path=pdf_path)
        _dlog_full("gen_single_received", f"proposal:{mname}", d)
        _dlog("gen_single_output_size", {"media_name": mname,
              "json_chars": len(json.dumps(d, ensure_ascii=False)) if isinstance(d, dict) else 0})
        # LLMが参照したページ番号を解析
        all_text = json.dumps(d, ensure_ascii=False)
        import re as _re2
        page_refs = sorted(set(int(m) for m in _re2.findall(r'(\d+)P[）\)]', all_text)))
        _dlog("gen_single_page_refs", {"media_name": mname,
              "pages_referenced": page_refs,
              "max_page": max(page_refs) if page_refs else 0,
              "exec_ratio_note": d.get("exec_ratio_note", ""),
        })

        if not isinstance(d, dict):
            _dlog("gen_single_not_dict", {"media_name": mname, "got": _describe(d)})
            d = {}

        _t_norm = time.time()
        d = _normalize_proposal(d, mname)
        _dlog("normalize_timing", {"media_name": mname, "dur": round(time.time() - _t_norm, 4)})

        expected = ["media_name", "category", "pv_ub", "exec_ratio", "exec_ratio_note", "price_disp",
                    "price_int", "guaranteed_pv", "tags",
                    "sec_overview", "sec_readers", "sec_menu_price",
                    "sec_lead_gen", "sec_clients", "sec_editorial", "reason", "ad_menus"]
        missing = [k for k in expected if k not in d]
        empty   = [k for k in expected if k in d and (d[k] in (None, "", [], {}, 0))]
        _dlog("gen_single_validation", {
            "media_name": mname, "keys_present": list(d.keys()),
            "missing": missing, "empty": empty,
            "exec_ratio_value": d.get("exec_ratio"),
            "exec_ratio_type": type(d.get("exec_ratio")).__name__,
            "exec_ratio_note": d.get("exec_ratio_note", ""),
            "price_int_value": d.get("price_int"),
            "ad_menus_count": len(d.get("ad_menus", [])),
            "dbg_last_page": d.get("dbg_last_page"),
        })

        d["_file_name"] = fname
        d["_pdf_url"]   = url
        _dlog_full("gen_single_return", f"proposal:{mname}", d)
        return d
    except PdfTooLargeError as e_big:
        _pdf_too_large = True
        _dlog("proposal_pdf_too_large", {"media_name": mname, "error": str(e_big)})
    except Exception as e1:
        import traceback
        _dlog("proposal_error_attempt1", {"media_name": mname, "error": str(e1),
              "traceback": traceback.format_exc()})

    # リトライループ（最大4回）。サイズ超過は決定論的失敗なのでリトライしない
    MAX_PROPOSAL_RETRIES = 4
    for retry_i in (range(2, MAX_PROPOSAL_RETRIES + 2) if not _pdf_too_large else ()):  # attempt 2,3,4,5
        await asyncio.sleep(2)
        _dlog("proposal_retry", {"media_name": mname, "attempt": retry_i, "wait": 2})
        try:
            d = await _llm_call_safe(p, schema=_SECTION_SCHEMA,
                                     context=f"proposal_retry{retry_i}:{mname}",
                                     pdf_path=pdf_path)
            _dlog_full("gen_single_received_retry", f"proposal_retry{retry_i}:{mname}", d)
            if not isinstance(d, dict):
                d = {}
            d = _normalize_proposal(d, mname)
            d["_file_name"] = fname
            d["_pdf_url"]   = url
            _dlog("proposal_retry_success", {"media_name": mname, "attempt": retry_i})
            return d
        except PdfTooLargeError as e_big2:
            _pdf_too_large = True
            _dlog("proposal_pdf_too_large", {"media_name": mname, "attempt": retry_i,
                  "error": str(e_big2)})
            break
        except Exception as e_retry:
            import traceback
            _dlog(f"proposal_error_attempt{retry_i}", {"media_name": mname,
                  "attempt": retry_i, "error": str(e_retry),
                  "traceback": traceback.format_exc()})

    # サイズ超過なら「参照ページ一覧」だけ抜いた縮小PDFで1回だけ再試行（PDF書類の読解精度を保つ）
    if _pdf_too_large and has_pdf:
        _shrunk = _shrink_pdf_for_llm(pdf_path, rec)
        if _shrunk:
            try:
                d = await _llm_call_safe(p, schema=_SECTION_SCHEMA,
                                         context=f"proposal_shrunk:{mname}", pdf_path=_shrunk)
                if isinstance(d, dict):
                    d = _normalize_proposal(d, mname)
                    d["_file_name"] = fname
                    d["_pdf_url"]   = url
                    d["_read_mode"] = "shrunk_pdf"   # 縮小PDF読取り（参照ページ以外は見ていない印）
                    _dlog("proposal_shrunk_success", {"media_name": mname})
                    return d
            except Exception as e_shr:
                _dlog("proposal_shrunk_error", {"media_name": mname, "error": str(e_shr)})

    # 全リトライ失敗
    _dlog("proposal_all_retries_exhausted", {"media_name": mname,
          "total_attempts": MAX_PROPOSAL_RETRIES + 1})

    # ★最後の砦: Gemini+file_paths が全滅なら pdfplumber でテキスト抽出して読ませる。
    # 図・グラフ内の数値(exec_ratio等)は落ちる可能性があるが、価格・メニューは拾える。
    try:
        _dlog("proposal_text_fallback_start", {"media_name": mname, "pdf_path": pdf_path})
        pdf_text = extract_pdf(pdf_path)
        if pdf_text and len(pdf_text.strip()) > 200:
            if len(pdf_text) > _PDF_TEXT_FALLBACK_MAXCHARS:
                _dlog("proposal_text_fallback_truncate", {"media_name": mname,
                      "orig_chars": len(pdf_text), "kept_chars": _PDF_TEXT_FALLBACK_MAXCHARS})
                pdf_text = pdf_text[:_PDF_TEXT_FALLBACK_MAXCHARS]
            p_text = (p + "\n\n【PDF添付が読めなかったため、以下は同じ媒体資料PDFのテキスト抽出結果です。"
                      "これを読んで上記JSONを埋めてください。図・グラフ内の数値は欠落している場合があります。"
                      "その場合は該当フィールドを規定どおり -1 / \"—\" にして unknown_fields に入れてください】\n"
                      + pdf_text)
            # pdf_path無し=テキストのみでllm_callへ（file_paths/画像は使わない）
            d = await _llm_call_safe(p_text, schema=_SECTION_SCHEMA,
                                     context=f"proposal_textfallback:{mname}")
            if isinstance(d, dict):
                d = _normalize_proposal(d, mname)
                d["_file_name"] = fname
                d["_pdf_url"]   = url
                d["_read_mode"] = "pdfplumber_text"  # 読取り経路を記録（図中%欠落の可能性フラグ）
                _dlog("proposal_text_fallback_success", {"media_name": mname})
                return d
    except Exception as e_text:
        import traceback
        _dlog("proposal_text_fallback_error", {"media_name": mname,
              "error": str(e_text), "traceback": traceback.format_exc()})

    # #3改修: 失敗を「要問合せ」と偽装せず、読取り失敗と明示する（正当な要問合せと区別）。
    fallback = {"media_name": mname, "category": "", "pv_ub": "—", "exec_ratio": -1,
            "price_disp": "読取り失敗", "price_int": -1, "guaranteed_pv": "—", "tags": [],
            "sec_overview": f"生成失敗（{MAX_PROPOSAL_RETRIES+1}回リトライ後）", "sec_readers": [], "sec_menu_price": [],
            "sec_lead_gen": "", "sec_clients": "", "sec_editorial": "",
            "reason": "媒体資料PDFを読み取れませんでした（リトライ後も失敗・要再実行）。",
            "ad_menus": [], "unknown_fields": [], "_read_failed": True,
            "_file_name": fname, "_pdf_url": url}
    return fallback


async def step3_generate_proposals(result_data: list, weight_values: list,
                                    pdf_urls: dict, user_request: str,
                                    article_type: str = "記事タイアップ",
                                    target_attribute: str = "",
                                    sandbox_paths: dict = None) -> tuple[list, list]:
    """
    Returns:
        proposals:  [{media_name, proposal_text}]
        min_prices: [{media_name, file_name, tieup_min_price}]
    """
    print("\n[STEP3] 提案文 + 最低単価 生成（file_paths方式・PDF書類のまま読取り）")

    # _gen_single_proposal（トップレベル関数）をgatherで並列実行
    _dlog("gather_start", {"type": "proposals", "count": len(result_data)})
    results = await asyncio.gather(*[
        _gen_single_proposal(r, 
            pdf_path=(sandbox_paths or {}).get(r.get("メディア名称","").split(";")[0].strip(), ""),
            pdf_url=pdf_urls.get(r.get("file_name",""), ""),
            user_request=user_request, article_type=article_type, target_attribute=target_attribute)
        for r in result_data])
    _dlog("gather_finish", {"type": "proposals", "count": len(results)})

    # proposals = 構造化dictをそのまま保持（HTML生成・後方互換の両方に使う）
    proposals = results
    # 集約後の全proposalを丸ごと記録（HTMLに渡る直前の最終状態）
    _dlog_full("proposals_full", "step3", proposals)
    _dlog("proposals_result", {"proposals": [
        {"media_name": p.get("media_name"), "price_int": p.get("price_int", 0),
         "price_disp": p.get("price_disp"),
         "overview_len": len(p.get("sec_overview", "") or ""),
         "reason_len": len(p.get("reason", "") or ""),
         "tags_n": len(p.get("tags", []) or [])} for p in proposals]})
    # min_prices は既存の予算チェック互換のため tieup_min_price に price_int を流す
    min_prices = [{"media_name": r.get("media_name", ""),
                   "file_name": r.get("_file_name", ""),
                   "tieup_min_price": r.get("price_int", 0)} for r in results]
    _dlog("min_prices_result", {"min_prices": min_prices})
    return proposals, min_prices


# ============================================================
# STEP4: 予算チェック + 補充
# ============================================================

async def step4_budget_check(excel_path: str, user_budget: int, desired_count: int,
                              score_params: dict, result_data: list,
                              pdf_urls: dict,
                              weight_values: list, proposals: list, min_prices: list,
                              user_request: str,
                              article_type: str = "記事タイアップ",
                              target_attribute: str = "") -> tuple[list, list, list, list, list]:
    """
    補充が必要なら score_media → step2 → step3 を1回だけ再実行。
    Returns: (final_proposals, final_min_prices, final_result_data, final_weight_values, budget_excluded_proposals)
    """
    print("\n[STEP4] 予算チェック")
    verdict = validate_budget(user_budget, desired_count, 1, min_prices, score_params)
    print(verdict["summary"])

    # 予算超過で弾かれた媒体を出力
    if verdict["over_budget"]:
        print(f"\n  ⚠ 予算超過で除外された媒体:")
        for m in verdict["over_budget"]:
            price = m.get("tieup_min_price", 0)
            reason = m.get("reason", "")
            print(f"    - {m.get('media_name','?')}: ¥{price:,}（{reason}）")

    if verdict["verdict"] != "refill":
        keep_names = {m["media_name"] for m in verdict["within_budget"]}
        fp = [p for p in proposals  if p["media_name"] in keep_names]
        fm = [m for m in min_prices if m["media_name"] in keep_names]
        fd = [r for r in result_data if r.get("メディア名称", "").split(";")[0].strip() in keep_names]
        fw = [w for w in weight_values if w.get("メディア名称", "") in keep_names]
        excluded_props = [p for p in proposals if p["media_name"] not in keep_names]
        return fp, fm, fd, fw, excluded_props

    # 補充実行
    shortage   = verdict["refill_top_n"]
    skip_top   = verdict["refill_skip_top"]
    print(f"\n[STEP4] 補充実行: skip_top={skip_top}, top_n={shortage}")

    refill_params = {**score_params, "skip_top": skip_top, "top_n": shortage}
    result_data2, weight_values2, _ = score_media(excel_path, refill_params)

    if result_data2:
        sp2, fl2, su2, _, _ = await step2_get_pdfs(result_data2)
        fn2 = [r.get("file_name", "") for r in result_data2]
        pu2 = get_pdf_urls(excel_path, fn2)
        pdf_urls.update(pu2)
        p2, m2 = await step3_generate_proposals(result_data2, weight_values2, pu2, user_request,
                                                   article_type=article_type, target_attribute=target_attribute,
                                                   sandbox_paths=sp2)
        proposals.extend(p2)
        min_prices.extend(m2)
        result_data.extend(result_data2)
        weight_values.extend(weight_values2)

    # 2回目チェック
    verdict2 = validate_budget(user_budget, desired_count, 2, min_prices, score_params)
    print(verdict2["summary"])

    keep_names = {m["media_name"] for m in verdict2["within_budget"]}
    fp = [p for p in proposals  if p["media_name"] in keep_names]
    fm = [m for m in min_prices if m["media_name"] in keep_names]
    fd = [r for r in result_data if r.get("メディア名称", "").split(";")[0].strip() in keep_names]
    fw = [w for w in weight_values if w.get("メディア名称", "") in keep_names]
    # 予算超過で弾かれたproposals
    excluded_props = [p for p in proposals if p["media_name"] not in keep_names]
    return fp, fm, fd, fw, excluded_props


# ============================================================
# HTMLレポート生成（純Python・LLM不使用）
#   v4テンプレのCSS/JSをそのまま流用し、構造化proposalsを流し込む。
#   AIにHTMLを書かせない＝タグ破損・class typo・onclick破壊を構造的に排除。
# ============================================================

def _esc(s) -> str:
    """HTMLエスケープ（テキスト用）"""
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))

def _man(n: int) -> str:
    """円整数 → '¥◯万' 表示。0や端数は素直に。"""
    if not n:
        return "—"
    return f"¥{n/10000:,.0f}万" if n >= 10000 else f"¥{n:,}"

def build_html_report(user_request: str, proposals: list, budget_limit: int,
                       excluded: list, target_label: str = "経営層比率",
                       remaining_data: list = None, target_attribute: str = "",
                       score_params: dict = None, all_data: list = None) -> str:
    """
    proposals: step3_generate_proposals が返す構造化dictのリスト
    budget_limit: 円単位整数（0なら上限なし表示）
    excluded:  [{"name","category","reason"}] 除外媒体
    """
    _t_html0 = time.time()
    n = len(proposals)
    prices = [p.get("price_int", 0) for p in proposals if p.get("price_int", 0) > 0]
    total  = sum(p.get("price_int", 0) for p in proposals)
    ratios = []
    if all_data and target_attribute:
        for r in all_data[:n]:
            try: ratios.append(float(r.get(target_attribute, 0) or 0))
            except: pass

    # --- KPIストリップ ---
    kpis = []
    if prices:
        kpis.append(("最安出稿", _man(min(prices))))
    if ratios:
        kpis.append((f"最高{target_label}", f"{max(ratios)}%"))
    kpis.append((f"{n}媒体合計", _man(total)))
    kpi_html = "".join(
        f'<div class="kpi-item"><span class="kpi-label">{_esc(l)}</span>'
        f'<span class="kpi-val">{_esc(v)}</span></div>' for l, v in kpis)

    # セクション表示名辞書（sec_*キー→日本語ラベル）
    SEC_TITLES = {
        "sec_overview": "メディア概要", "sec_readers": "読者プロフィール",
        "sec_menu_price": "広告メニュー・料金", "sec_lead_gen": "リード獲得の仕組み",
        "sec_clients": "実績企業", "sec_editorial": "編集企画",
    }

    # all_dataからtarget_attribute値を引くための辞書
    _excel_lookup = {}
    if all_data:
        for r in all_data:
            mn = r.get("メディア名称", "").split(";")[0].strip()
            _excel_lookup[mn] = r

    # --- メイン行 + detail-row ---
    rows = []
    for i, p in enumerate(proposals, 1):
        # ★最重大改修: DB値(マスタ)とPDF値(exec_ratio)を突き合わせる。
        # ★1改修でPDF読取りが信頼できるようになったため、両者が矛盾(乖離>閾値)したら
        # 一次ソースのPDF値を採用し、DB値との乖離を警告表示＋ログに残す（DB修正シグナル）。
        mn = p.get("media_name", "")
        excel_row = _excel_lookup.get(mn, {})
        excel_ratio = 0
        if target_attribute and excel_row:
            try: excel_ratio = float(excel_row.get(target_attribute, 0) or 0)
            except: excel_ratio = 0
        try: pdf_ratio = float(p.get("exec_ratio", -1))
        except: pdf_ratio = -1
        disp_ratio = excel_ratio      # 既定はDB値
        ratio_note = ""
        if target_attribute:
            if pdf_ratio >= 0 and excel_ratio > 0 and abs(excel_ratio - pdf_ratio) > _DB_PDF_DIVERGENCE_PT:
                # DB値とPDF値が矛盾 → 一次ソースのPDF値を採用し、DB値との乖離を警告
                _dlog("db_pdf_contradiction", {"media_name": mn, "attribute": target_attribute,
                      "db_value": excel_ratio, "pdf_value": pdf_ratio,
                      "diff_pt": round(abs(excel_ratio - pdf_ratio), 1)})
                disp_ratio = pdf_ratio
                ratio_note = (f'<span style="color:#c0392b;font-size:11px;display:block;margin-top:2px">'
                              f'⚠️DB値{excel_ratio:g}%と乖離・要確認</span>')
            elif pdf_ratio < 0 and excel_ratio > 0:
                # PDFに該当記載が無い → DB値のまま。ただし出典が検証できていないことを明示
                ratio_note = ('<span style="color:#888;font-size:11px;display:block;margin-top:2px">'
                              'DB由来・未検証</span>')
        bar_pct = min(disp_ratio, 100)
        # tagsをExcelの上位3属性の数値から機械的に生成（LLM生成を上書き）
        excel_tags = []
        if excel_row:
            tag_candidates = []
            for col_name, col_val in excel_row.items():
                if any(col_name.startswith(prefix + "_") for prefix in ("業種","役職","職種")):
                    try:
                        fv = float(col_val)
                        if fv > 0:
                            tag_candidates.append((col_name, fv))
                    except: pass
            tag_candidates.sort(key=lambda x: -x[1])
            excel_tags = [f"{name} {val}%" for name, val in tag_candidates[:3]]
        if excel_tags:
            tags_html = "".join(f'<span class="tag">{_esc(t)}</span>' for t in excel_tags)
        else:
            tags_html = "".join(f'<span class="tag">{_esc(t)}</span>' for t in (p.get("tags") or []))
        url = p.get("_pdf_url", "") or ""
        pdf_cell = (f'<a class="pdf-link" href="{_esc(url)}" target="_blank" '
                    f'onclick="event.stopPropagation()">PDF</a>') if url and url != "（URLなし）" else "—"

        # セクションをsec_*キーから動的生成（固定順序＋未知キーも拾う）
        sec_order = ["sec_overview", "sec_readers", "sec_menu_price",
                     "sec_lead_gen", "sec_clients", "sec_editorial"]
        extra_secs = [k for k in p if k.startswith("sec_") and k not in sec_order]
        all_secs = sec_order + extra_secs

        sections_html = ""
        for sk in all_secs:
            val = p.get(sk)
            if val in (None, "", [], {}):
                continue
            title = SEC_TITLES.get(sk, sk.replace("sec_", "").replace("_", " ").title())
            if isinstance(val, list):
                items = "".join(f"<li>{_esc(x)}</li>" for x in val)
                sections_html += f'<div class="detail-section"><h4>{_esc(title)}</h4><ul>{items}</ul></div>\n'
            else:
                sections_html += f'<div class="detail-section"><h4>{_esc(title)}</h4><p>{_esc(val)}</p></div>\n'

        rows.append(f"""
        <tr class="main-row" onclick="toggleDetail(this)">
          <td><div class="rank-cell"><span class="toggle-arrow">&#9654;</span><span class="rank">{i}</span></div></td>
          <td><div class="name">{_esc(p.get('media_name',''))}</div><div class="cat">{_esc(p.get('category',''))}</div></td>
          <td class="mono">{_esc(p.get('pv_ub','—'))}</td>
          <td class="bar-wrap"><div class="bar-bg"><div class="bar-fill" style="width:{bar_pct}%"></div></div><span class="bar-num">{disp_ratio:g}%</span>{ratio_note}</td>
          <td class="price">{_esc(p.get('price_disp','—'))}</td>
          <td class="mono">{_esc(p.get('guaranteed_pv','—'))}</td>
          <td>{tags_html}</td>
          <td>{pdf_cell}</td>
        </tr>
        <tr class="detail-row">
          <td colspan="8">
            <div class="detail-inner">
              <div class="detail-grid">
                {sections_html}
              </div>
              <div class="detail-reason"><strong>選定理由：</strong>{_esc(p.get('reason',''))}</div>
            </div>
          </td>
        </tr>""")
    _dlog("html_rows_built", {"dur": round(time.time() - _t_html0, 4), "n_rows": len(rows)})
    rows_html = "".join(rows)

    # --- ランキング外媒体（more-section、メインテーブルと同じカラム構造） ---
    more_html = ""
    remaining = remaining_data or []
    if remaining:
        def _cell(v, fmt_func=None):
            """Excelの値があれば入れる、なければ—"""
            if v is None or str(v).strip() in ("", "nan", "None", "0"):
                return "—"
            if fmt_func:
                try: return fmt_func(v)
                except Exception: return str(v)
            return _esc(str(v))
        def _pv_fmt(v):
            try:
                n = int(float(v))
                return f"{n:,}" if n > 0 else "—"
            except Exception: return "—"
        def _price_cell(v):
            try:
                n = int(float(v))
                return _man(n) if n > 0 else "—"
            except Exception: return "—"
        rem_rows = ""
        for j, r in enumerate(remaining, n + 1):
            mn = r.get("メディア名称", "").split(";")[0].strip()
            pv = r.get("月間PV数", "")
            ub = r.get("月間UU数", "")
            pv_disp = _pv_fmt(pv)
            ub_disp = _pv_fmt(ub)
            ratio = r.get(target_attribute, "") if target_attribute else ""
            ratio_disp = f"{ratio}%" if ratio and str(ratio).strip() not in ("", "nan", "None", "0") else "—"
            price_raw = r.get("最低出稿金額", "")
            price_disp = _price_cell(price_raw)
            pv_g = r.get("PV保証_最小", "")
            pv_g_disp = f"{_pv_fmt(pv_g)}PV" if pv_g and str(pv_g).strip() not in ("", "nan", "None", "0") else "—"
            menu = r.get("主要メニューフォーマット", "")
            menu_disp = _esc(str(menu)) if menu and str(menu).strip() not in ("", "nan") else "—"
            pdf_url = r.get("PDFファイルURL", "")
            pdf_link = f'<a href="{_esc(str(pdf_url))}" target="_blank" style="color:#1a1a1a;font-size:12px;border-bottom:1px solid #ccc;text-decoration:none">PDF</a>' if pdf_url and str(pdf_url).strip() not in ("", "nan", "None", "（URLなし）") else "—"
            rem_rows += f"""
          <tr>
            <td><span class="rank" style="color:#ddd">{j}</span></td>
            <td><div class="name">{_esc(mn)}</div></td>
            <td class="mono">{pv_disp} / {ub_disp}</td>
            <td class="mono">{ratio_disp}</td>
            <td class="price">{price_disp}</td>
            <td class="mono">{pv_g_disp}</td>
            <td>{menu_disp}</td>
            <td>{pdf_link}</td>
          </tr>"""
        more_html = f"""
  <div class="more-section">
    <div class="more-toggle" onclick="toggleMore(this)">
      <span class="more-arrow">&#9654;</span>
      <span>その他の候補メディア（{len(remaining)}件）</span>
    </div>
    <div class="more-table">
      <table>
        <thead><tr><th>#</th><th>メディア</th><th>月間PV / UB</th><th>{_esc(target_label)}</th><th>最低価格</th><th>PV保証</th><th>特徴</th><th>資料</th></tr></thead>
        <tbody>{rem_rows}</tbody>
      </table>
    </div>
  </div>"""

    # --- 予算シミュレーター（ad_menus対応・メディア×メニュー チェックUI） ---
    import re as _re
    menus_json_data = []
    for p in proposals:
        items = []
        for am in (p.get("ad_menus") or []):
            am_str = str(am)
            price_m = _re.search(r'[¥￥]([\d,]+)', am_str)
            price_val = int(price_m.group(1).replace(',', '')) if price_m else 0
            items.append({"name": am_str, "price": price_val})
        menus_json_data.append({"media": p.get("media_name", ""), "items": items})
    menus_json = json.dumps(menus_json_data, ensure_ascii=False)
    limit_disp = f"/ 上限 {_man(budget_limit)}" if budget_limit > 0 else ""
    budget_html = f"""
  <div class="budget-sim">
    <div class="budget-sim-title">予算シミュレーション</div>
    <div id="budgetGroups"></div>
    <div class="budget-result">
      <span class="budget-total-label">選択合計</span>
      <span class="budget-total-val" id="budgetTotal">—</span>
      <span class="budget-limit" id="budgetLimit">{limit_disp}</span>
      <span class="budget-count" id="budgetCount"></span>
    </div>
    <div class="bbar-bg"><div class="bbar-fill" id="budgetBar"></div></div>
    <div class="budget-status" id="budgetStatus"></div>
  </div>
  <script>
  var BUDGET_LIMIT={budget_limit};
  var MENUS={menus_json};
  var _st={{}};var _totalB=0,_menuC=0;
  var fmt=function(n){{return n>=10000?'¥'+(n/10000).toLocaleString()+'万':'¥'+n.toLocaleString()}};
  var grp=document.getElementById('budgetGroups');
  MENUS.forEach(function(m,mi){{
    var g=document.createElement('div');g.className='media-group open';
    var h=document.createElement('div');h.className='media-header';
    h.innerHTML='<span class="media-arrow">&#9654;</span><span class="media-name">'+m.media+'</span><span class="media-subtotal" id="sub'+mi+'">—</span>';
    h.onclick=function(){{g.classList.toggle('open')}};
    var list=document.createElement('div');list.className='menu-list';
    m.items.forEach(function(item,ji){{
      var key=mi+':'+ji;_st[key]=false;
      var row=document.createElement('div');row.className='menu-item';
      var chk=document.createElement('div');chk.className='menu-check';
      chk.onclick=function(){{_st[key]=!_st[key];chk.classList.toggle('checked');_updateB()}};
      row.innerHTML='<div class="menu-name">'+item.name+'</div><div class="menu-price">'+fmt(item.price)+'</div>';
      row.prepend(chk);list.appendChild(row);
    }});
    g.appendChild(h);g.appendChild(list);grp.appendChild(g);
  }});
  function _updateB(){{
    _totalB=0;_menuC=0;
    MENUS.forEach(function(m,mi){{
      var sub=0;
      m.items.forEach(function(item,ji){{if(_st[mi+':'+ji]){{sub+=item.price;_menuC++}}}});
      var se=document.getElementById('sub'+mi);
      se.textContent=sub>0?fmt(sub):'—';se.className='media-subtotal'+(sub>0?' active':'');
      _totalB+=sub;
    }});
    document.getElementById('budgetTotal').textContent=_totalB>0?fmt(_totalB):'—';
    document.getElementById('budgetCount').textContent=_menuC>0?_menuC+'メニュー選択中':'';
    var bar=document.getElementById('budgetBar');
    var st=document.getElementById('budgetStatus');
    if(BUDGET_LIMIT>0){{
      bar.style.width=Math.min(_totalB/BUDGET_LIMIT*100,120)+'%';
      if(_totalB>BUDGET_LIMIT){{
        document.getElementById('budgetTotal').className='budget-total-val over';
        bar.className='bbar-fill over';st.textContent=fmt(_totalB-BUDGET_LIMIT)+' の超過';st.className='budget-status over';
      }}else{{
        document.getElementById('budgetTotal').className='budget-total-val';
        bar.className='bbar-fill';st.textContent=_totalB>0?'残り '+fmt(BUDGET_LIMIT-_totalB):'';st.className='budget-status';
      }}
    }}else{{bar.style.width='0%';st.textContent=''}}
  }}
  </script>"""

    sub = f"予算上限 {_man(budget_limit)}" if budget_limit > 0 else "予算上限 指定なし"

    # --- パラメータパネル + カラムトグル + 重み変更UI ---
    sp = score_params or {}
    filters_disp = ""
    for k, v in (sp.get("filters") or {}).items():
        if isinstance(v, dict):
            filters_disp += f'<span class="param-chip">{_esc(k)} {_esc(str(v.get("op","")))} {_esc(str(v.get("value","")))}</span>'
    weights_disp = ""
    for k, v in (sp.get("weights") or {}).items():
        weights_disp += f'<span class="param-chip weight">{_esc(k)}: {v}</span>'
    sort_disp = f'<span class="param-chip sort">{_esc(str(sp.get("sort_by","なし")))} {_esc(sp.get("sort_order",""))}</span>' if sp.get("sort_by") else ""

    # カラムグループ定義
    col_groups = {
        "業種": [], "役職": [], "職種": [], "規模": [], "年代": []
    }
    if all_data and len(all_data) > 0:
        for col in all_data[0].keys():
            for grp in col_groups:
                if col.startswith(grp + "_"):
                    col_groups[grp].append(col)

    # カラムグループトグルHTML
    col_toggles = ""
    for grp, cols in col_groups.items():
        if not cols: continue
        checks = "".join(f'<label class="col-check"><input type="checkbox" value="{_esc(c)}" onchange="toggleColumn(this)"><span>{_esc(c.replace(grp+"_",""))}</span></label>' for c in cols)
        col_toggles += f'<div class="col-group"><div class="col-group-title" onclick="this.parentElement.classList.toggle(\'open\')">{grp} ▶</div><div class="col-group-body">{checks}</div></div>'

    # 重み変更UI
    weight_cols = []
    for cols in col_groups.values():
        weight_cols.extend(cols)
    weight_sliders = "".join(
        f'<div class="weight-row"><label>{_esc(c)}</label><input type="range" min="0" max="5" value="{(sp.get("weights") or {}).get(c, 0)}" data-col="{_esc(c)}" oninput="updateWeightLabel(this);updateChips()"><span class="weight-val">{(sp.get("weights") or {}).get(c, 0)}</span></div>'
        for c in weight_cols[:20]  # 表示上限
    )

    # all_data JSON（JS再ソート用）
    all_data_json = json.dumps(all_data or [], ensure_ascii=False)

    # 初期パラメータJSON（JS動的更新用）
    init_filters_json = json.dumps(sp.get("filters") or {}, ensure_ascii=False)
    init_sort = sp.get("sort_by") or ""
    init_sort_order = sp.get("sort_order", "")

    # フィルタ/ソート用カラムoptions
    filter_cols = []
    sort_cols = []
    if all_data and len(all_data) > 0:
        for c in all_data[0].keys():
            if c in ("メディア名称", "file_name", "_resort_score"): continue
            filter_cols.append(c)
            # 数値カラムをソート候補に
            v = all_data[0].get(c)
            if isinstance(v, (int, float)) or (isinstance(v, str) and v.replace(".","",1).replace("-","",1).isdigit()):
                sort_cols.append(c)
    filter_options = "".join(f'<option value="{_esc(c)}">{_esc(c)}</option>' for c in filter_cols[:40])
    sort_options = "".join(f'<option value="{_esc(c)}">{_esc(c)}</option>' for c in sort_cols[:20])

    params_panel = f"""
  <div class="params-bar">
    <div class="params-chips" id="paramChips">
      {filters_disp if filters_disp else ''}
      {weights_disp if weights_disp else ''}
      {sort_disp if sort_disp else ''}
    </div>
    <button class="copy-btn" onclick="copyParams()" title="この設定をコピー（AIに貼り付けて再提案可能）">📋 コピー</button>
    <button class="settings-btn" onclick="document.getElementById('settingsPanel').classList.toggle('open')">⚙ 設定</button>
  </div>
  <div class="settings-panel" id="settingsPanel">
    <div class="settings-row">
      <div class="settings-col">
        <div class="params-label">カラム表示</div>
        <div class="col-groups">{col_toggles}</div>
      </div>
      <div class="settings-col">
        <div class="params-label">重み変更</div>
        <div class="weight-sliders">{weight_sliders}</div>
        <button class="resort-btn" onclick="resortTable()">この重みで並べ替え</button>
        <div class="resort-warn">※ テーブルの並び順のみ変更。AI提案は維持。</div>
      </div>
      <div class="settings-col">
        <div class="params-label">フィルタ</div>
        <div id="activeFilters"></div>
        <div style="display:flex;gap:4px;align-items:center;flex-wrap:wrap;margin-top:4px">
          <select id="filterCol" style="font-size:11px;padding:3px;max-width:140px">{filter_options}</select>
          <select id="filterOp" style="font-size:11px;padding:3px"><option value="contains">含む</option><option value=">=">≥</option><option value="<=">≤</option><option value="eq">一致</option></select>
          <input id="filterVal" placeholder="値" style="font-size:11px;padding:3px;width:70px">
          <button onclick="addFilter()" style="font-size:11px;padding:3px 8px;cursor:pointer">+</button>
        </div>
        <div class="params-label" style="margin-top:12px">ソート</div>
        <div style="display:flex;gap:4px;align-items:center">
          <select id="sortCol" style="font-size:11px;padding:3px;max-width:140px"><option value="">なし</option>{sort_options}</select>
          <select id="sortOrder" style="font-size:11px;padding:3px"><option value="asc">昇順</option><option value="desc">降順</option></select>
        </div>
        <button class="resort-btn" onclick="applySortFilter()" style="margin-top:8px">フィルタ/ソート適用</button>
        <div class="resort-warn">※ ランキング外テーブルの並びを変更。</div>
      </div>
    </div>
  </div>
  <script>
  var ALL_DATA={all_data_json};
  var _initFilters={init_filters_json};
  var _initSort="{_esc(init_sort)}";
  var _initSortOrder="{_esc(init_sort_order)}";
  function copyParams() {{
    var sliders = document.querySelectorAll('.weight-row input[type=range]');
    var parts = [];
    // フィルタ
    for (var k in _initFilters) {{
      var f = _initFilters[k];
      parts.push(k + ' ' + (f.op||'') + ' ' + (f.value||''));
    }}
    // 重み
    var wParts = [];
    sliders.forEach(function(s) {{
      if (parseInt(s.value) > 0) wParts.push(s.dataset.col + '×' + s.value);
    }});
    if (wParts.length) parts.push('重み: ' + wParts.join(', '));
    if (_initSort) parts.push('並び順: ' + _initSort + ' ' + _initSortOrder);
    var text = parts.join(' / ');
    navigator.clipboard.writeText(text).then(function() {{
      var btn = document.querySelector('.copy-btn');
      btn.textContent = '✅ コピー済';
      setTimeout(function() {{ btn.textContent = '📋 コピー'; }}, 2000);
    }});
  }}
  function updateChips() {{
    var sliders = document.querySelectorAll('.weight-row input[type=range]');
    var chips = document.getElementById('paramChips');
    // フィルタチップ（初期）
    var html = '';
    for (var k in _initFilters) {{
      var f = _initFilters[k];
      html += '<span class="param-chip">' + k + ' ' + (f.op||'') + ' ' + (f.value||'') + '</span>';
    }}
    // フィルタチップ（UI追加分）
    if (typeof _activeFilters !== 'undefined') {{
      _activeFilters.forEach(function(f) {{
        html += '<span class="param-chip">' + f.col + ' ' + f.op + ' ' + f.value + '</span>';
      }});
    }}
    // 重みチップ（動的）
    sliders.forEach(function(s) {{
      if (parseInt(s.value) > 0) {{
        html += '<span class="param-chip weight">' + s.dataset.col + ': ' + s.value + '</span>';
      }}
    }});
    // ソートチップ
    if (_initSort) html += '<span class="param-chip sort">' + _initSort + ' ' + _initSortOrder + '</span>';
    chips.innerHTML = html;
  }}
  var _activeFilters = [];
  function addFilter() {{
    var col = document.getElementById('filterCol').value;
    var op = document.getElementById('filterOp').value;
    var val = document.getElementById('filterVal').value;
    if (!col || !val) return;
    _activeFilters.push({{col: col, op: op, value: val}});
    renderFilters(); updateChips();
    document.getElementById('filterVal').value = '';
  }}
  function removeFilter(idx) {{
    _activeFilters.splice(idx, 1);
    renderFilters(); updateChips();
  }}
  function renderFilters() {{
    var el = document.getElementById('activeFilters');
    el.innerHTML = _activeFilters.map(function(f, i) {{
      return '<span class="param-chip" style="cursor:pointer" onclick="removeFilter('+i+')">' + f.col + ' ' + f.op + ' ' + f.value + ' ×</span>';
    }}).join(' ');
  }}
  function applySortFilter() {{
    if (typeof ALL_DATA === 'undefined' || !ALL_DATA.length) return;
    var filtered = ALL_DATA.slice();
    _activeFilters.forEach(function(f) {{
      filtered = filtered.filter(function(row) {{
        var v = row[f.col];
        if (v === undefined || v === null) v = '';
        if (f.op === 'contains') return String(v).indexOf(f.value) >= 0;
        if (f.op === '>=') return parseFloat(v) >= parseFloat(f.value);
        if (f.op === '<=') return parseFloat(v) <= parseFloat(f.value);
        if (f.op === 'eq') return String(v) === f.value;
        return true;
      }});
    }});
    var sortCol = document.getElementById('sortCol').value;
    var sortOrder = document.getElementById('sortOrder').value;
    if (sortCol) {{
      filtered.sort(function(a, b) {{
        var va = parseFloat(a[sortCol]) || 0, vb = parseFloat(b[sortCol]) || 0;
        return sortOrder === 'asc' ? va - vb : vb - va;
      }});
    }}
    // ランキング外テーブルを更新
    var moreTable = document.querySelector('.more-table tbody');
    if (!moreTable) return;
    moreTable.innerHTML = '';
    filtered.forEach(function(r, i) {{
      var name = (r['メディア名称']||'').split(';')[0].trim();
      var tr = document.createElement('tr');
      tr.innerHTML = '<td>' + (i+1) + '</td><td>' + name + '</td>';
      var heads = document.querySelector('.more-table thead tr');
      if (heads) {{
        var ths = heads.querySelectorAll('th');
        for (var j = 2; j < ths.length; j++) {{
          var col = ths[j].textContent.trim();
          var val = r[col] !== undefined ? r[col] : '—';
          var td = document.createElement('td');
          td.textContent = val; td.style.fontSize = '12px';
          tr.appendChild(td);
        }}
      }}
      moreTable.appendChild(tr);
    }});
    updateChips();
    var btn = document.querySelector('.settings-col:last-child .resort-btn');
    if (btn) {{ btn.textContent = '✅ 適用済 (' + filtered.length + '件)'; setTimeout(function(){{ btn.textContent = 'フィルタ/ソート適用'; }}, 2000); }}
  }}
  </script>"""

    _dlog("html_pre_format", {"dur_total": round(time.time() - _t_html0, 4)})
    return _HTML_SHELL.format(
        user_request=_esc(user_request), sub=_esc(sub), n=n,
        kpi=kpi_html, rows=rows_html, more=more_html, budget=budget_html,
        target_label=_esc(target_label), params_panel=params_panel,
    )


_HTML_SHELL = """<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>メディア選定レポート</title>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  body {{ font-family:'Noto Sans JP',sans-serif; background:#fff; color:#1a1a1a; line-height:1.7; -webkit-font-smoothing:antialiased; }}
  .container {{ max-width:1280px; margin:0 auto; padding:48px 32px; }}
  .header {{ margin-bottom:40px; padding-bottom:20px; border-bottom:1px solid #e5e5e5; }}
  .header h1 {{ font-size:22px; font-weight:700; letter-spacing:-0.3px; }}
  .header .sub {{ font-size:13px; color:#888; margin-top:4px; }}
  .kpi-strip {{ display:flex; gap:32px; margin-bottom:36px; flex-wrap:wrap; }}
  .kpi-item {{ display:flex; align-items:baseline; gap:8px; }}
  .kpi-label {{ font-size:12px; color:#999; }}
  .kpi-val {{ font-family:'JetBrains Mono',monospace; font-size:20px; font-weight:600; }}
  .table-wrap {{ overflow-x:auto; }}
  table {{ width:100%; border-collapse:collapse; font-size:13px; }}
  thead th {{ text-align:left; font-size:11px; font-weight:500; color:#999; text-transform:uppercase; letter-spacing:0.5px; padding:8px 12px; border-bottom:2px solid #1a1a1a; white-space:nowrap; }}
  tbody td {{ padding:12px; border-bottom:1px solid #f0f0f0; vertical-align:top; }}
  tr.main-row {{ cursor:pointer; transition:background 0.15s; }}
  tr.main-row:hover {{ background:#fafafa; }}
  .rank {{ font-family:'JetBrains Mono',monospace; font-weight:600; color:#bbb; font-size:14px; }}
  .rank-cell {{ display:flex; align-items:center; gap:6px; }}
  .toggle-arrow {{ font-size:9px; color:#bbb; transition:transform 0.2s; display:inline-block; width:12px; text-align:center; }}
  tr.main-row.open .toggle-arrow {{ transform:rotate(90deg); }}
  .name {{ font-weight:700; font-size:14px; }}
  .cat {{ font-size:11px; color:#aaa; }}
  .mono {{ font-family:'JetBrains Mono',monospace; font-weight:500; font-size:12px; }}
  .bar-wrap {{ min-width:110px; }}
  .bar-bg {{ background:#f0f0f0; height:6px; border-radius:3px; margin-bottom:3px; }}
  .bar-fill {{ height:100%; border-radius:3px; background:#1a1a1a; }}
  .bar-num {{ font-family:'JetBrains Mono',monospace; font-size:12px; font-weight:600; }}
  .price {{ font-family:'JetBrains Mono',monospace; font-weight:600; font-size:13px; }}
  .tag {{ display:inline-block; font-size:10px; padding:2px 7px; border:1px solid #ddd; border-radius:3px; margin-right:3px; margin-bottom:3px; color:#555; }}
  .pdf-link {{ color:#1a1a1a; font-size:12px; text-decoration:none; border-bottom:1px solid #ccc; }}
  .pdf-link:hover {{ border-color:#1a1a1a; }}
  tr.detail-row {{ display:none; }}
  tr.detail-row.open {{ display:table-row; }}
  tr.detail-row td {{ padding:0 12px 20px 12px; border-bottom:1px solid #e0e0e0; background:#fafafa; }}
  .detail-inner {{ padding:20px 16px 8px 16px; }}
  .detail-grid {{ display:grid; grid-template-columns:1fr 1fr; gap:20px 32px; margin-bottom:16px; }}
  .detail-section h4 {{ font-size:11px; font-weight:700; color:#999; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:8px; padding-bottom:4px; border-bottom:1px solid #eee; }}
  .detail-section p, .detail-section li {{ font-size:12px; color:#333; line-height:1.8; }}
  .detail-section ul {{ list-style:none; padding:0; }}
  .detail-section ul li {{ padding:2px 0; }}
  .detail-section ul li::before {{ content:"\\00B7"; margin-right:6px; color:#ccc; font-weight:700; }}
  .detail-reason {{ padding:12px 16px; background:#f0f0f0; border-radius:4px; font-size:13px; line-height:1.8; color:#333; }}
  .more-section {{ border-top:1px solid #f0f0f0; margin-bottom:40px; }}
  .more-toggle {{ display:flex; align-items:center; gap:6px; padding:12px; font-size:13px; color:#888; cursor:pointer; user-select:none; transition:color 0.15s; }}
  .more-toggle:hover {{ color:#555; }}
  .more-toggle .more-arrow {{ font-size:9px; transition:transform 0.2s; display:inline-block; }}
  .more-toggle.open .more-arrow {{ transform:rotate(90deg); }}
  .more-table {{ display:none; }}
  .more-table.open {{ display:block; }}
  .more-table thead th {{ border-bottom:1px solid #e5e5e5; }}
  .more-table tbody td {{ color:#888; }}
  .budget-sim {{ border-top:1px solid #e5e5e5; padding-top:24px; margin-bottom:40px; }}
  .budget-sim-title {{ font-size:14px; font-weight:700; margin-bottom:16px; }}
  .media-group {{ margin-bottom:12px; }}
  .media-header {{ display:flex; align-items:center; gap:8px; padding:8px 0; border-bottom:1px solid #f0f0f0; cursor:pointer; user-select:none; }}
  .media-header:hover {{ background:#fafafa; }}
  .media-name {{ font-weight:700; font-size:13px; flex:1; }}
  .media-subtotal {{ font-family:'JetBrains Mono',monospace; font-size:12px; font-weight:600; color:#999; transition:color .2s; }}
  .media-subtotal.active {{ color:#1a1a1a; }}
  .media-arrow {{ font-size:9px; color:#ccc; transition:transform .2s; display:inline-block; width:14px; text-align:center; }}
  .media-group.open .media-arrow {{ transform:rotate(90deg); }}
  .menu-list {{ display:none; padding:4px 0 8px 22px; }}
  .media-group.open .menu-list {{ display:block; }}
  .menu-item {{ display:flex; align-items:center; gap:10px; padding:5px 0; border-bottom:1px solid #f8f8f8; }}
  .menu-item:last-child {{ border-bottom:none; }}
  .menu-check {{ width:15px; height:15px; border:2px solid #ddd; border-radius:3px; cursor:pointer; display:flex; align-items:center; justify-content:center; transition:all .15s; flex-shrink:0; }}
  .menu-check.checked {{ background:#1a1a1a; border-color:#1a1a1a; }}
  .menu-check.checked::after {{ content:"✓"; color:#fff; font-size:9px; font-weight:700; }}
  .menu-name {{ font-size:12px; flex:1; color:#333; }}
  .menu-price {{ font-family:'JetBrains Mono',monospace; font-size:12px; font-weight:600; white-space:nowrap; }}
  .budget-result {{ display:flex; align-items:center; gap:16px; flex-wrap:wrap; margin-top:20px; padding:20px 0; border-top:1px solid #e5e5e5; }}
  .budget-total-label {{ font-size:12px; color:#999; }}
  .budget-total-val {{ font-family:'JetBrains Mono',monospace; font-size:28px; font-weight:600; transition:color 0.2s; }}
  .budget-total-val.over {{ color:#dc2626; }}
  .budget-limit {{ font-family:'JetBrains Mono',monospace; font-size:13px; color:#999; }}
  .budget-count {{ font-size:12px; color:#999; margin-left:auto; }}
  .bbar-bg {{ background:#f0f0f0; height:8px; border-radius:4px; overflow:hidden; margin-top:12px; }}
  .bbar-fill {{ height:100%; border-radius:4px; background:#1a1a1a; transition:width 0.3s,background 0.3s; }}
  .bbar-fill.over {{ background:#dc2626; }}
  .budget-status {{ font-size:12px; color:#999; margin-top:6px; min-height:18px; }}
  .budget-status.over {{ color:#dc2626; }}
  .footer {{ font-size:11px; color:#ccc; text-align:center; padding-top:16px; border-top:1px solid #f0f0f0; }}
  .params-bar {{ display:flex; align-items:center; gap:8px; padding:8px 12px; margin-bottom:4px; border:1px solid #eee; border-radius:6px; flex-wrap:wrap; }}
  .params-chips {{ display:flex; gap:4px; flex-wrap:wrap; flex:1; }}
  .param-chip {{ display:inline-block; font-size:11px; padding:3px 8px; border:1px solid #ddd; border-radius:4px; color:#333; font-family:'JetBrains Mono',monospace; }}
  .param-chip.weight {{ border-color:#7c3aed; color:#7c3aed; }}
  .param-chip.sort {{ border-color:#2563eb; color:#2563eb; }}
  .param-chip.dim {{ color:#ccc; border-color:#eee; }}
  .copy-btn {{ padding:4px 10px; font-size:11px; border:1px solid #ddd; background:#fff; border-radius:4px; cursor:pointer; white-space:nowrap; }}
  .copy-btn:hover {{ background:#f5f5f5; }}
  .settings-btn {{ padding:4px 10px; font-size:11px; border:1px solid #ddd; background:#fff; border-radius:4px; cursor:pointer; white-space:nowrap; }}
  .settings-btn:hover {{ background:#f5f5f5; }}
  .settings-panel {{ display:none; border:1px solid #eee; border-top:none; border-radius:0 0 6px 6px; padding:16px; margin-bottom:12px; }}
  .settings-panel.open {{ display:block; }}
  .settings-row {{ display:flex; gap:24px; flex-wrap:wrap; }}
  .settings-col {{ flex:1; min-width:250px; }}
  .params-label {{ font-size:11px; font-weight:600; color:#999; text-transform:uppercase; letter-spacing:.5px; margin-bottom:6px; }}
  .col-groups {{ display:flex; gap:12px; flex-wrap:wrap; }}
  .col-group {{ border:1px solid #eee; border-radius:4px; }}
  .col-group-title {{ padding:6px 10px; font-size:11px; font-weight:600; cursor:pointer; color:#666; }}
  .col-group-body {{ display:none; padding:4px 10px 8px; }}
  .col-group.open .col-group-body {{ display:block; }}
  .col-check {{ display:flex; align-items:center; gap:4px; padding:2px 0; font-size:11px; color:#555; cursor:pointer; }}
  .col-check input {{ width:13px; height:13px; }}
  .weight-sliders {{ display:grid; grid-template-columns:repeat(auto-fill, minmax(220px, 1fr)); gap:6px; margin-bottom:12px; }}
  .weight-row {{ display:flex; align-items:center; gap:8px; font-size:11px; }}
  .weight-row label {{ width:120px; color:#666; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }}
  .weight-row input[type=range] {{ flex:1; height:4px; }}
  .weight-val {{ font-family:'JetBrains Mono',monospace; width:20px; text-align:center; font-weight:600; }}
  .resort-btn {{ padding:6px 16px; font-size:12px; border:1px solid #1a1a1a; background:#1a1a1a; color:#fff; border-radius:4px; cursor:pointer; }}
  .resort-btn:hover {{ background:#333; }}
  .resort-warn {{ font-size:10px; color:#999; margin-top:6px; }}
  @media (max-width:768px) {{ .container {{ padding:24px 16px; }} .detail-grid {{ grid-template-columns:1fr; }} .kpi-strip {{ gap:16px; }} }}
</style></head>
<body><div class="container">
  <div class="header"><h1>メディア選定レポート</h1><div class="sub">{sub} ｜ {n}媒体選定</div></div>
  <div class="kpi-strip">{kpi}</div>
  {params_panel}
  <div class="table-wrap"><table>
    <thead><tr><th>#</th><th>メディア</th><th>月間PV / UB</th><th>{target_label}</th><th>最低価格</th><th>PV保証</th><th>特徴</th><th>資料</th></tr></thead>
    <tbody>{rows}</tbody>
  </table></div>
  {more}
  {budget}
  <div class="footer">MyCompany Inc. ｜ メディアマスター + 媒体資料PDFに基づく分析</div>
</div>
<script>
  function toggleDetail(row) {{ row.classList.toggle('open'); row.nextElementSibling.classList.toggle('open'); }}
  function toggleMore(el) {{ el.classList.toggle('open'); el.nextElementSibling.classList.toggle('open'); }}
  function toggleColumn(cb) {{
    var col = cb.value;
    var thead = document.querySelector('thead tr');
    var tbody = document.querySelector('tbody');
    if (!thead || !tbody) return;
    if (cb.checked) {{
      var th = document.createElement('th');
      th.className = 'dyn-col'; th.dataset.col = col;
      th.textContent = col; th.style.fontSize = '10px'; th.style.maxWidth = '80px';
      thead.querySelector('th:last-child').before(th);
      var rows = tbody.querySelectorAll('tr');
      for (var i = 0; i < rows.length; i++) {{
        var row = rows[i];
        if (row.classList.contains('main-row')) {{
          var name = row.querySelector('.name');
          var mediaName = name ? name.textContent.trim() : '';
          var data = ALL_DATA.find(function(d) {{ return (d['メディア名称']||'').split(';')[0].trim() === mediaName; }});
          var val = data ? (parseFloat(data[col]) || 0) : 0;
          var td = document.createElement('td');
          td.className = 'dyn-col mono'; td.dataset.col = col; td.style.fontSize = '12px';
          td.textContent = val > 0 ? val + '%' : '—';
          row.querySelector('td:last-child').before(td);
        }} else if (row.classList.contains('detail-row')) {{
          var td = row.querySelector('td');
          if (td) td.setAttribute('colspan', parseInt(td.getAttribute('colspan')||8) + 1);
        }}
      }}
    }} else {{
      thead.querySelectorAll('th.dyn-col').forEach(function(el) {{ if (el.dataset.col === col) el.remove(); }});
      tbody.querySelectorAll('td.dyn-col').forEach(function(el) {{ if (el.dataset.col === col) el.remove(); }});
      tbody.querySelectorAll('.detail-row td[colspan]').forEach(function(td) {{
        var c = parseInt(td.getAttribute('colspan')||8); td.setAttribute('colspan', Math.max(8, c - 1));
      }});
    }}
  }}
  function updateWeightLabel(slider) {{
    slider.nextElementSibling.textContent = slider.value;
  }}
  function resortTable() {{
    if (typeof ALL_DATA === 'undefined' || !ALL_DATA.length) return;
    var sliders = document.querySelectorAll('.weight-row input[type=range]');
    var weights = {{}};
    sliders.forEach(function(s) {{ if (parseInt(s.value) > 0) weights[s.dataset.col] = parseInt(s.value); }});
    var activeCols = Object.keys(weights);
    // --- カラム自動追加/削除 ---
    var thead = document.querySelector('thead tr');
    var tbody = document.querySelector('tbody');
    if (!thead || !tbody) return;
    // 既存の動的カラムを削除 + colspan リセット
    thead.querySelectorAll('.dyn-col').forEach(function(el) {{ el.remove(); }});
    tbody.querySelectorAll('.dyn-col').forEach(function(el) {{ el.remove(); }});
    tbody.querySelectorAll('.detail-row td[colspan]').forEach(function(td) {{ td.setAttribute('colspan', 8); }});
    // 動的カラムを追加（重み>0のもの）
    activeCols.forEach(function(col) {{
      // thead に th 追加
      var th = document.createElement('th');
      th.className = 'dyn-col';
      th.textContent = col;
      th.style.fontSize = '10px';
      th.style.maxWidth = '80px';
      thead.querySelector('th:last-child').before(th);
      // 各main-row + detail-rowペアにtd追加
      var rows = tbody.querySelectorAll('tr');
      for (var i = 0; i < rows.length; i++) {{
        var row = rows[i];
        if (row.classList.contains('main-row')) {{
          var name = row.querySelector('.name');
          var mediaName = name ? name.textContent.trim() : '';
          var data = ALL_DATA.find(function(d) {{ return (d['メディア名称']||'').split(';')[0].trim() === mediaName; }});
          var val = data ? (parseFloat(data[col]) || 0) : 0;
          var td = document.createElement('td');
          td.className = 'dyn-col mono';
          td.style.fontSize = '12px';
          td.textContent = val > 0 ? val + '%' : '—';
          row.querySelector('td:last-child').before(td);
        }} else if (row.classList.contains('detail-row')) {{
          var td = row.querySelector('td');
          if (td) td.setAttribute('colspan', 8 + activeCols.length);
        }}
      }}
    }});
    // --- スコア再計算 + 並べ替え ---
    ALL_DATA.forEach(function(r) {{
      var score = 0;
      for (var col in weights) {{
        var v = parseFloat(r[col]) || 0;
        score += v * weights[col];
      }}
      r._resort_score = score;
    }});
    ALL_DATA.sort(function(a, b) {{ return (b._resort_score || 0) - (a._resort_score || 0); }});
    var pairs = [];
    var rows = tbody.querySelectorAll('tr');
    for (var i = 0; i < rows.length; i += 2) {{
      var mainRow = rows[i]; var detailRow = rows[i + 1];
      var name = mainRow.querySelector('.name');
      pairs.push({{ el: [mainRow, detailRow], name: name ? name.textContent.trim() : '' }});
    }}
    var order = ALL_DATA.map(function(r) {{ return (r['メディア名称'] || '').split(';')[0].trim(); }});
    pairs.sort(function(a, b) {{
      var ai = order.indexOf(a.name); var bi = order.indexOf(b.name);
      return (ai < 0 ? 999 : ai) - (bi < 0 ? 999 : bi);
    }});
    pairs.forEach(function(p, idx) {{
      p.el[0].querySelector('.rank').textContent = idx + 1;
      tbody.appendChild(p.el[0]); tbody.appendChild(p.el[1]);
    }});
  }}
</script>
</body></html>"""


# ============================================================
# ユーザー要望パース
# ============================================================

async def _parse_user_request(user_request: str) -> dict:
    """
    user_requestを{filters, weights, top_n, user_budget, desired_count, sort_by, sort_order}に変換
    """
    return await _llm_call_safe(
        prompt=f"""あなたはメディア選定システムのパラメータ変換AIです。
以下のユーザー要望をscore_media用のJSONパラメータに変換してください。

【利用可能なカラム】
- filters用: 各フィルタは {{"op": "<演算子>", "value": "<値>"}} の形式で指定する
  - 主要メニューフォーマット: op=contains, value=記事タイアップ/バナー・純広/ホワイトペーパーDL/ウェビナー・イベント/メルマガ/動画/ターゲティング配信 のいずれか
  - 対応ファネル: op=contains, value=認知・ブランディング/興味喚起・理解促進/比較検討・リード獲得 のいずれか
  - 獲得リードの質: op=contains, value=名刺情報のみ/役職・部門あり/課題アンケート付与可/該当なし のいずれか
  - 最低出稿金額: op=<= または >=, value=万円単位の整数
  - 業種_◯◯: op=>=, value=%単位の整数
  - 例: {{"主要メニューフォーマット": {{"op": "contains", "value": "記事タイアップ"}}, "最低出稿金額": {{"op": "<=", "value": 300}}}}
- weights用(%系のみ): 業種_IT/製造/金融/サービス/建設不動産/医療/官公庁教育/その他, 役職_経営者役員/部長/課長/係長主任/一般社員, 職種_営業/情シス/企画/技術/製造/人事総務/経理財務/マーケ, 規模_エンタープライズ(1000名以上)/SMB(100-999名)/スタートアップ零細(100名未満), 年代_20代/30代/40代/50代/60代以上
- sort_by用(絶対値カラム): 最低出稿金額, 月間PV数, 月間UU数, 会員数

【ルール】
- フィルタは1〜2個まで
- weightsに絶対値カラム(最低出稿金額等)は入れない。並び替えならsort_byを使う
- sort_byとweightsは同時に使わない。sort_byで並べたいなら weights は空 {{}} にする。weightsで重視する属性の足切りが必要なら filters に入れる（例: "業種_製造": {{"op":">=","value":10}}）
- user_budget: 予算が明示されていれば円単位の整数。「予算上限なし」「予算は気にしない」と明示されたら 0。言及自体がなければ null
- desired_count: 件数が明示されていれば整数。「何件」「何社」等。言及がなければ null
- article_type: ユーザーが出稿したい広告種別。「タイアップ記事」「バナー」「ホワイトペーパー」「メルマガ」等。明示なければ null
- target_attribute: ユーザーが重視する読者属性。「製造業比率」「経営者役員比率」「IT比率」等。%系カラム名で返す。明示なければ null
- exclude_media: 「実施済みなので除く」「◯◯以外で」「別のメディアで」等、ユーザーが除外を指定した媒体名の配列。明示なければ []

【ユーザー要望】
{user_request}

JSON形式のみで返答:
{{
  "filters": {{}},
  "weights": {{}},
  "top_n": 7,
  "sort_by": null,
  "sort_order": "asc",
  "user_budget": null,
  "desired_count": null,
  "article_type": null,
  "target_attribute": null,
  "exclude_media": []
}}
""",
        schema={
            "type": "object",
            "properties": {
                "filters":          {"type": "object"},
                "weights":          {"type": "object"},
                "top_n":            {"type": "integer"},
                "sort_by":          {"type": ["string", "null"]},
                "sort_order":       {"type": "string"},
                "user_budget":      {"type": ["integer", "null"]},
                "desired_count":    {"type": ["integer", "null"]},
                "article_type":     {"type": ["string", "null"]},
                "target_attribute": {"type": ["string", "null"]},
                "exclude_media":    {"type": "array", "items": {"type": "string"}},
            },
            "required": ["filters", "weights", "top_n", "user_budget", "desired_count"],
        },
        context="_parse_user_request"
    )


# ============================================================
# メイン処理（ターン1）
# ============================================================

async def main(user_request: str, params: dict = None):
    global ptc_state

    # ターン別debugファイルを開始
    _turn_num = ptc_state.get("turn_count", 1) if ptc_state else 1
    _reset_debug_path(_turn_num)
    _dlog("process_start", {"user_request": user_request, "turn": _turn_num})

    print("=" * 60)
    print(f"[START] メディア選定エージェント")
    print(f"要望: {user_request}")
    print("=" * 60)

    # Excelパス取得
    import glob
    # ExcelマスターをOneDriveの/メディアフォルダから検索してDL
    _dlog("excel_dl_start", {"method": "folder_search"})
    _t_excel = time.time()
    try:
        list_resp = await call_with_retry("ONE_DRIVE_LIST_FOLDER_CHILDREN", {
            "folder_path": "/メディア", "use_me_drive": True, "select": ["id", "name"],
        }, max_attempts=10, base_wait=2.0, max_wait=2.0,
           deadline_sec=15.0, retry_context="excel_list")
        list_items = list_resp.get("data", {}).get("value", []) if isinstance(list_resp, dict) else []
        excel_item = next((it for it in list_items if it.get("name", "").endswith(".xlsx")), None)
        if not excel_item:
            raise FileNotFoundError("/メディアフォルダにExcelファイルが見つからない")
        print(f"  Excel検出: {excel_item['name']} (id={excel_item['id'][:12]}...)")
        excel_res = await call_with_retry("ONE_DRIVE_DOWNLOAD_FILE",
            {"item_id": excel_item["id"], "file_name": excel_item["name"]},
            max_attempts=10, base_wait=2.0, max_wait=2.0,
            deadline_sec=15.0, retry_context="excel_dl")
        excel_path = (excel_res.get("data", {}).get("content", {}).get("s3url")
                      or excel_res.get("content", {}).get("s3url")
                      or excel_res.get("sandbox_path", "")) if isinstance(excel_res, dict) else ""
    except Exception:
        candidates = (glob.glob("uploads/*メディア選定*.xlsx") or glob.glob("uploads/*.xlsx") or glob.glob("**/*.xlsx", recursive=True))
        excel_path = candidates[0] if candidates else ""
    if not excel_path:
        raise FileNotFoundError("ExcelマスターのDLに失敗")
    _dlog("excel_dl_done", {"path": excel_path, "dur": round(time.time() - _t_excel, 3)})
    print(f"Excelパス: {excel_path}")

    # STEP1: スコアリング
    print("\n[STEP1] スコアリング")
    if params is None:
        params = await _parse_user_request(user_request)
        print("  パラメータ: LLM変換")
    else:
        print("  パラメータ: メインAI設定（LLMコールスキップ）")
    _dlog_full("parse_user_request_full", "main", params)
    # null(None)パイプライン: 「未言及」はここまでnullのまま届き、この時点で初めて既定値に落とす。
    # .get(key, 既定値) だとキーがNone付きで存在すると既定値が効かないため or 方式（キー欠損/None/空文字を吸収）。
    desired_count = params.get("desired_count") or 7
    user_budget   = params.get("user_budget") or 0
    article_type  = params.get("article_type") or "記事タイアップ"
    target_attribute = params.get("target_attribute") or ""
    score_params  = {
        "excel_path":  excel_path,
        "filters":     params.get("filters", {}),
        "weights":     params.get("weights", {}),
        "top_n":       999,     # 全件取得してPythonでスライスする
        "sort_by":     params.get("sort_by"),
        "sort_order":  params.get("sort_order", "asc"),
        "previous_top_n": desired_count,
        "exclude_media": params.get("exclude_media", []),  # #6改修: 実施済み等の除外媒体名
    }
    _dlog_full("score_params_full", "main", score_params)
    print(f"  filters: {score_params['filters']}")
    print(f"  weights: {score_params['weights']}")
    print(f"  top_n: {desired_count}  user_budget: {fmt_yen(user_budget) if user_budget > 0 else '未指定'}")
    print(f"  exclude: {score_params['exclude_media'] if score_params['exclude_media'] else 'なし'}")

    # ★4改修: 絞り込み条件が1つも無い（filters/weights/sort_byが全て空）なら、
    # 黙って先頭7件を返す“選定したフリ”をせず、ユーザーに言い換えを聞き返して終了する。
    # （対象データに無い軸=OOH等、または要望が曖昧なケース。filtersが有れば絞れているので通す。）
    if (not score_params.get("filters") and not score_params.get("weights")
            and not score_params.get("sort_by")):
        _dlog("selection_impossible", {"reason": "no filters/weights/sort_by",
              "user_request": user_request, "params": params})
        print("\n" + "=" * 60)
        print("⚠️ この要望からは絞り込み条件を作れませんでした（対象データに無い軸か、条件が曖昧です）。")
        print("黙って上位7件を出すと誤解を招くため、選定は行いません。")
        print("次のような軸で言い換え・具体化してください：")
        print("  ・業種（例：製造業向け／情シス向け）")
        print("  ・読者の役職や規模（例：経営層比率が高い順／大企業向け）")
        print("  ・予算（例：300万円以内）")
        print("  ・並べ替え（例：掲載費用が安い順）")
        print("  ・広告種別（例：記事タイアップ／バナー／ウェビナー）")
        print("=" * 60)
        return

    _dlog("STEP1_start", {})
    all_data, all_weights, fallback_analysis = score_media(excel_path, score_params)
    _dlog_full("score_media_result_full", "main", all_data)

    # 全件取得 → 上位N件 + ランキング外
    result_data = all_data[:desired_count]
    remaining_data = all_data[desired_count:]
    weight_values = all_weights[:desired_count]
    _dlog("score_media_split", {"top_n": len(result_data), "remaining": len(remaining_data)})

    # 件数不足 → フィルタ緩和して再実行
    if len(result_data) < desired_count and fallback_analysis:
        print(f"\n  結果 {len(result_data)}件 < desired {desired_count}件 → フィルタ緩和")
        relax_res = await _llm_call_safe(
            prompt=f"""スコアリング結果が{len(result_data)}件で、目標{desired_count}件を下回りました。
フォールバック分析:
{fallback_analysis}

現在のフィルタ:
{json.dumps(score_params['filters'], ensure_ascii=False)}

最も緩和すべきフィルタを1つ外したJSONのfiltersのみを返してください。
""",
            schema={"type": "object", "properties": {"filters": {"type": "object"}}, "required": ["filters"]},
            context="filter_relax"
        )
        score_params["filters"] = relax_res["filters"]
        all_data2, all_weights2, _ = score_media(excel_path, score_params)
        result_data = all_data2[:desired_count]
        remaining_data = all_data2[desired_count:]
        weight_values = all_weights2[:desired_count]
        print(f"  緩和後: {len(result_data)}件")

    # #5改修: 条件は在るのに結果が空/重み付けが全0点なら、黙って先頭N件を出さず聞き返す。
    _top_score = max((r.get("score", 0) for r in result_data), default=0)
    if len(result_data) == 0:
        _dlog("selection_empty", {"reason": "no_match", "filters": score_params.get("filters")})
        print("\n" + "=" * 60)
        print("⚠️ 指定の条件に該当する媒体がありませんでした（フィルタを緩めても0件）。")
        print("条件を緩めるか言い換えてください（例：業種を広げる／予算上限を上げる／必須条件を減らす）。")
        print("=" * 60)
        return
    if score_params.get("weights") and not score_params.get("sort_by") and _top_score == 0:
        _dlog("selection_no_signal", {"reason": "all_zero_weighted_score",
              "weights": score_params.get("weights")})
        print("\n" + "=" * 60)
        print("⚠️ 指定の重み条件では媒体に差がつきませんでした（該当データが無く全media0点）。")
        print("別の軸で言い換えるか条件を緩めてください（例：重視する属性を変える／データのある軸にする）。")
        print("=" * 60)
        return

    print(f"\n  確定候補: {len(result_data)}件")
    _dlog("STEP1_result", {"count": len(result_data), "result_data": result_data})
    for r in result_data:
        print(f"    - {r.get('メディア名称','')}  (score: {r.get('score',0):.1f})")

    # excel_urlsをSTEP2の前に構築（DL完了次第proposal開始するため）
    all_file_names = [r.get("file_name", "") for r in result_data]
    excel_urls = get_pdf_urls(excel_path, all_file_names)
    pdf_urls = {}

    # DL完了次第proposal開始するcallback
    def on_dl_complete(media_name, sp, share_url, orig_fname):
        rec = next((r for r in result_data if r.get("メディア名称","").split(";")[0].strip() == media_name), None)
        if not rec: return None
        fn = rec.get("file_name", "")
        url = share_url or excel_urls.get(fn, "")
        pdf_urls[fn] = url
        _dlog("pipeline_proposal_start", {"media_name": media_name})
        task = asyncio.create_task(_gen_single_proposal(rec, pdf_path=sp, pdf_url=url,
            user_request=user_request, article_type=article_type, target_attribute=target_attribute))
        return task

    # STEP2+3 パイプライン: DL完了した媒体からproposal開始
    _dlog("STEP2_start", {})
    sandbox_paths, failed_list, share_urls, list_items, proposal_tasks = await step2_get_pdfs(
        result_data, on_dl_complete=on_dl_complete)

    # 最終的な提案対象メディアの絞り込み
    valid_media_names = set(sandbox_paths.keys()) | {f["media_name"] for f in failed_list}

    excluded_media = [
        {"name": r.get("メディア名称", "").split(";")[0].strip(),
         "category": "", "reason": "媒体資料PDFを取得できず、今回の提案対象外"}
        for r in result_data
        if r.get("メディア名称", "").split(";")[0].strip() not in valid_media_names
    ]

    result_data = [r for r in result_data if r.get("メディア名称", "").split(";")[0].strip() in valid_media_names]
    weight_values = [w for w in weight_values if w.get("メディア名称", "") in valid_media_names]

    if not result_data and not proposal_tasks:
        print("\n[ERROR] 提案可能なメディアが1件も見つかりませんでした")
        return

    # 全proposal完了を待つ（DL中に既に開始済み）
    _dlog("STEP4_proposal_start", {"pipeline": True, "tasks_started": len(proposal_tasks)})
    if proposal_tasks:
        proposals_raw = await asyncio.gather(*proposal_tasks)
        proposals = list(proposals_raw)
    else:
        proposals = []
    # #5/#8改修: proposalsはDL完了順で届くため、min_prices構築前に選定順（result_data順）へ並べ直す。
    # これをしないと予算チェック(validate_budget)の within[:desired_count] が
    # 「スコア上位」でなく「先にDLが終わった媒体」を採用してしまう。
    _sel_order = {r.get("メディア名称", "").split(";")[0].strip(): i for i, r in enumerate(result_data)}
    proposals.sort(key=lambda p: _sel_order.get(p.get("media_name", ""), 10**9))
    min_prices = [{"media_name": p.get("media_name", ""), "file_name": p.get("_file_name", ""),
                   "tieup_min_price": p.get("price_int", 0)} for p in proposals]
    _dlog("proposals_result", {"count": len(proposals), "proposals": [
        {"media_name": p.get("media_name"), "price_int": p.get("price_int", 0),
         "price_disp": p.get("price_disp")} for p in proposals]})

    # --- remaining分のCREATE_LINKをバックグラウンドで起動 ---
    remaining_link_tasks = {}  # file_name → asyncio.Task
    nfc = unicodedata.normalize
    list_names = [it.get("name", "") for it in list_items]
    for r in remaining_data:
        fn = r.get("file_name", "")
        if not fn: continue
        best = difflib.get_close_matches(nfc("NFC", fn), [nfc("NFC", n) for n in list_names], n=1, cutoff=0.5)
        if best:
            orig_name = [n for n in list_names if nfc("NFC", n) == best[0]][0]
            item = next((it for it in list_items if it.get("name") == orig_name), None)
            if item:
                remaining_link_tasks[fn] = asyncio.create_task(
                    call_with_retry("ONE_DRIVE_CREATE_LINK",
                                    {"item_id": item["id"], "type": "view", "scope": "anonymous"},
                                    max_attempts=5, base_wait=2.0, max_wait=2.0,
                                    deadline_sec=10.0, retry_context=f"rem_link:{fn}"))
    _dlog("remaining_link_tasks_started", {"count": len(remaining_link_tasks)})

    # --- proposal完了後 → remaining CREATE_LINKの完了分を回収、未完了をキャンセル ---
    remaining_share_urls = {}
    completed = 0; cancelled = 0
    for fn, task in remaining_link_tasks.items():
        if task.done():
            try:
                res = task.result()
                url = ""
                if isinstance(res, dict):
                    url = (res.get("data", {}).get("link", {}).get("webUrl", "")
                           or res.get("link", {}).get("webUrl", ""))
                if url:
                    remaining_share_urls[fn] = url
                    completed += 1
            except Exception:
                pass
        else:
            task.cancel()
            cancelled += 1
    _dlog("remaining_links_harvested", {"completed": completed, "cancelled": cancelled})

    # remaining_dataにshare_urlを注入（HTMLで使う）
    for r in remaining_data:
        fn = r.get("file_name", "")
        if fn in remaining_share_urls:
            r["PDFファイルURL"] = remaining_share_urls[fn]

    # STEP4: 予算チェック
    _dlog("STEP4_input", {"user_budget": user_budget, "min_prices": min_prices})
    budget_excluded_proposals = []
    if user_budget > 0:
        proposals, min_prices, result_data, weight_values, budget_excluded_proposals = await step4_budget_check(
            excel_path, user_budget, desired_count, score_params,
            result_data, pdf_urls,
            weight_values, proposals, min_prices, user_request,
            article_type=article_type, target_attribute=target_attribute)
        if budget_excluded_proposals:
            print(f"\n  予算超過の{len(budget_excluded_proposals)}件をその他の候補メディアに移動")
            for p in reversed(budget_excluded_proposals):
                _mn = p.get("media_name", "")
                _row = next((r for r in (all_data or []) if _mn in str(r.get("メディア名称",""))), {})
                _entry = dict(_row)
                _entry["_budget_excluded"] = True
                _entry["_budget_price"] = p.get("price_int", 0)
                _entry["_budget_reason"] = f"予算超過 ¥{p.get('price_int',0):,}"
                remaining_data.insert(0, _entry)
    else:
        print("\n[STEP4] 予算未指定のためスキップ")

    # STEP5: 状態保存 + 出力
    ptc_state = {
        "excel_path":     excel_path,
        "score_params":   score_params,
        "desired_count":  desired_count,
        "user_budget":    user_budget,
        "result_data":    result_data,
        "pdf_urls":        pdf_urls,
        "weight_values":  weight_values,
        "proposals":      proposals,
        "min_prices":     min_prices,
        "sandbox_paths":  sandbox_paths,
        "share_urls":     share_urls,
        "remaining_data": remaining_data,
        "all_data":       all_data,
        "failed_list":    failed_list,
        "excluded_media": excluded_media,
        "article_type":   article_type,
        "target_attribute": target_attribute,
        "turn_count":     ptc_state.get("turn_count", 1) if ptc_state else 1,
        "user_request":   user_request,
    }

    _dlog("final_output", {"adopted_media": [p.get("media_name") for p in proposals], "count": len(proposals)})

    # --- HTMLレポート生成（純Python・LLM不使用）。書き出しのみ。file_outputは最後尾で1回。 ---
    excluded_all = list(excluded_media)
    _dlog_full("html_input_proposals", "main", proposals)
    _dlog_full("html_input_excluded", "main", excluded_all)
    # --- STEP3b後の再ソート: price_intが確定したので、sort_byに応じて並べ直す ---
    # 並べ替え軸がtarget_attribute単独(%)か判定（★最重大改修A案で使用）。
    # score_params内のweights/target_attributeは未正規化の可能性があるため実列名に揃えて比較する。
    _rs_cols = list(all_data[0].keys()) if all_data else []
    _rs_ta = _canonical_col(target_attribute, _rs_cols) if target_attribute else ""
    _rs_wkeys = [_canonical_col(k, _rs_cols) for k in (score_params.get("weights") or {})]
    _single_axis_ta = bool(_rs_ta) and _rs_wkeys == [_rs_ta]

    if score_params.get("sort_by") and any(k in (score_params.get("sort_by") or "") for k in ("金額","出稿","価格")):
        _so = score_params.get("sort_order", "asc")
        # #8a改修: 価格不明(-1/0)は昇順・降順どちらでも必ず最後尾へ。
        # 旧実装 `or float('inf')` は0しか捕まえられず、-1が「最小=安い順の先頭」に化けていた。
        def _price_sort_key(p):
            v = p.get("price_int") or 0
            if v > 0:
                return (0, v if _so != "desc" else -v)
            return (1, 0)   # 価格不明は既知の後ろ（同士は元の順を維持）
        proposals.sort(key=_price_sort_key)
        _dlog("re_sort_by_price", {"sort_order": _so,
              "order": [f"{p.get('media_name','')}: {p.get('price_int',0)}" for p in proposals]})
    elif _single_axis_ta:
        # ★最重大改修(A案): 順位はSTEP1でDB値により決まるが、DB値が誤り(例Forbes57.7% vs PDF34.7%)だと
        # 誤値が媒体を不当に上位へ押し上げ、検出改修後は「1位の表示値<2位の表示値」の逆転表になる。
        # PDF値が確定した今、“検証済みの値”（DBとPDFが閾値超で乖離していればPDF値、それ以外はDB値）で
        # 選ばれたN件の順位を付け直し、表示値と順位の基準を一致させる。
        # ※応急処置: 圏外に落ちた媒体は救えない（根治はDB補完）。合成重み(複数軸)は検証不能のため対象外。
        _ex_lookup = {str(r.get("メディア名称", "")).split(";")[0].strip(): r for r in (all_data or [])}
        def _verified_ratio(p):
            mn = p.get("media_name", "")
            try: db_v = float(_ex_lookup.get(mn, {}).get(_rs_ta, 0) or 0)
            except Exception: db_v = 0.0
            try: pdf_v = float(p.get("exec_ratio", -1))
            except Exception: pdf_v = -1.0
            if pdf_v >= 0 and db_v > 0 and abs(db_v - pdf_v) > _DB_PDF_DIVERGENCE_PT:
                return pdf_v          # DBとPDFが矛盾 → 一次ソースのPDF値で順位付け
            return db_v if db_v > 0 else max(pdf_v, 0.0)
        proposals.sort(key=_verified_ratio, reverse=True)
        _dlog("re_sort_by_verified_ratio", {"attribute": _rs_ta,
              "order": [f"{p.get('media_name','')}: {_verified_ratio(p):g}" for p in proposals]})
    else:
        # #5改修: 価格ソート以外(重み付け/属性sort)はproposalsがDL完了順のままで表示順≠選定順になる。
        # score_mediaが既にresult_dataを正しい順(重み降順/sort_by昇降順)に並べているので、その順にproposalsを揃える。
        _order = {r.get("メディア名称", "").split(";")[0].strip(): i for i, r in enumerate(result_data)}
        proposals.sort(key=lambda p: _order.get(p.get("media_name", ""), 10**9))
        _dlog("re_sort_by_selection_order",
              {"order": [p.get("media_name", "") for p in proposals]})

    target_label = target_attribute if target_attribute else "経営層比率"
    html = build_html_report(user_request, proposals, user_budget, excluded_all,
                             target_label=target_label,
                             remaining_data=remaining_data,
                             target_attribute=target_attribute,
                             score_params=score_params,
                             all_data=all_data)
    _dlog("html_built", {"html_len": len(html),
          "n_main_rows": html.count('class="main-row"'),
          "n_detail_sections": html.count('class="detail-section"'),
          })
    os.makedirs("output", exist_ok=True)
    _turn = ptc_state.get("turn_count", 1)
    html_suffix = f"_turn{_turn}" if _turn > 1 else ""
    html_path = f"output/media_report{html_suffix}.html"
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    _dlog("html_written", {"path": html_path, "bytes": len(html.encode("utf-8"))})

    # stdout詳細出力（メインAIが読んでユーザーに報告する）
    print("\n" + "=" * 60)
    print(f"[OUTPUT] [セッション{_turn}] HTMLレポート生成: {len(proposals)}媒体  → {html_path}")
    print("=" * 60)
    for i, p in enumerate(proposals, 1):
        if not isinstance(p, dict):
            print(f"  [WARN] Invalid proposal format: {p}")
            continue
        ratio_disp = '不明' if p.get('exec_ratio') == -1 else f"{p.get('exec_ratio',0)}%"
        print(f"  {i}. {p.get('media_name','Unknown')}  ({p.get('category','—')})")
        _ws = score_params.get("weights", {})
        if _ws:
            _mn = p.get("media_name", "")
            _row = next((r for r in (all_data or []) if _mn in str(r.get("メディア名称",""))), {})
            _attrs = " / ".join(f"{k} {_row.get(k, '—')}%" for k in _ws)
            print(f"     PV/UB: {p.get('pv_ub','—')}  価格: {p.get('price_disp','—')}  PV保証: {p.get('guaranteed_pv','—')}")
            print(f"     重視属性: {_attrs}")
        else:
            ratio_disp = '不明' if p.get('exec_ratio') == -1 else f"{p.get('exec_ratio',0)}%"
            print(f"     PV/UB: {p.get('pv_ub','—')}  {target_label}: {ratio_disp}  価格: {p.get('price_disp','—')}  PV保証: {p.get('guaranteed_pv','—')}")
        reason = p.get('reason', '')
        if reason:
            print(f"     選定理由: {reason}")
        menus = p.get('ad_menus', [])
        if menus:
            print(f"     広告メニュー・料金:")
            for m in menus:
                print(f"       * {str(m)[:120]}")
        _su = share_urls.get(p.get("media_name",""), "") if "share_urls" in dir() else ""
        if _su:
            print(f"     PDF: {_su}")
    if excluded_all:
        print("\n  除外: " + ", ".join(e["name"] for e in excluded_all))
    if failed_list:
        print("\n  ⚠ PDF取得失敗: " + ", ".join(f["media_name"] for f in failed_list))
    if budget_excluded_proposals:
        print(f"\n{'='*60}")
        print(f"[予算超過で除外] {len(budget_excluded_proposals)}媒体（予算{fmt_yen(user_budget)}超過）")
        print("=" * 60)
        for i, p in enumerate(budget_excluded_proposals, 1):
            if not isinstance(p, dict): continue
            print(f"  {i}. {p.get('media_name','Unknown')}  ({p.get('category','—')})")
            # 以下の詳細印字はループ内（旧実装はデデントされていて最後の1件分しか出なかった）
            _ws = score_params.get("weights", {})
            if _ws:
                _mn = p.get("media_name", "")
                _row = next((r for r in (all_data or []) if _mn in str(r.get("メディア名称",""))), {})
                _attrs = " / ".join(f"{k} {_row.get(k, '—')}%" for k in _ws)
                print(f"     PV/UB: {p.get('pv_ub','—')}  価格: {p.get('price_disp','—')}  PV保証: {p.get('guaranteed_pv','—')}")
                print(f"     重視属性: {_attrs}")
            else:
                ratio_disp = '不明' if p.get('exec_ratio') == -1 else f"{p.get('exec_ratio',0)}%"
                print(f"     PV/UB: {p.get('pv_ub','—')}  {target_label}: {ratio_disp}  価格: {p.get('price_disp','—')}  PV保証: {p.get('guaranteed_pv','—')}")
                reason = p.get('reason', '')
                if reason:
                    print(f"     選定理由: {reason}")
                menus = p.get('ad_menus', [])
                if menus:
                    print(f"     広告メニュー・料金:")
                    for m in menus:
                        print(f"       * {str(m)[:120]}")
            _su = share_urls.get(p.get("media_name",""), "") if "share_urls" in dir() else ""
            if _su:
                print(f"     PDF: {_su}")


    total_sec = round(time.time() - _START_TIME, 2)
    _dlog("process_finish", {"total_duration_sec": total_sec})
    print(f"\n[完了] 処理時間: {total_sec}秒")

    # --- ファイル出力は全処理の最後に1回だけ実行する ---
    await _finalize_output(html_path)


def compute_gantt_bars(evs):
    """eventsからガントのbarsリストとtotalを計算して返す。"""
    d = {}
    for e in evs:
        t = e["elapsed"]; evt = e["event"]; ctx = e.get("context", "")
        if evt == "excel_dl_start": d["excel_s"] = t
        if evt == "excel_dl_done": d["excel_e"] = t
        if evt == "llm_call_start" and ctx == "_parse_user_request": d.setdefault("p_s", t)
        if evt == "llm_call_raw_response" and ctx == "_parse_user_request": d["p_llm"] = t
        if evt == "parse_user_request_full": d["p_e"] = t
        if evt == "STEP1_start": d["s1_s"] = t
        if evt == "STEP1_result": d["s1_e"] = t
        if evt == "STEP2_start": d["s2_s"] = t
        if evt == "retry_call_start" and "LIST" in e.get("tool", ""): d.setdefault("list_s", t)
        if evt == "retry_success" and "LIST" in e.get("tool", ""): d["list_e"] = t
        if evt == "difflib_timing": d["diff_e"] = t
        if evt == "resolve_single_start": d.setdefault("res_s", t)
        if evt in ("ai_resolved_full", "ai_giveup_full"): d["res_e"] = t
        if evt == "gather_start" and e.get("type") == "download": d["dl_s"] = t
        if evt == "gather_finish" and e.get("type") == "download": d["dl_e"] = t
        if evt == "STEP2_result": d["s2_e"] = t
        if evt == "STEP4_proposal_start": d["s3b_s"] = t
        if evt == "proposals_result": d["s3b_e"] = t
        if evt == "remaining_link_tasks_started": d["rem_link_s"] = t
        if evt == "remaining_links_harvested": d["rem_link_e"] = t
        if evt == "html_built": d["html_e"] = t
        if evt == "process_finish": d["fin"] = t
    total = d.get("fin", 1)

    bars = []
    def add(label, s, e, level=0, color="#333", extra=""):
        bars.append({"label": label, "start": round(s, 2), "end": round(e, 2),
                     "dur": round(e - s, 2), "level": level, "color": color, "extra": extra})

    # Excel DL
    if "excel_s" in d and "excel_e" in d:
        add("Excel DL(OneDrive)", d["excel_s"], d["excel_e"], 0, "#059669")
    add("STEP0 parse", d.get("p_s", 0), d.get("p_e", 0), 0, "#2563eb")
    add("LLM応答待ち", d.get("p_s", 0), d.get("p_llm", 0), 1, "#3b82f6")
    add("STEP1 score", d.get("s1_s", 0), d.get("s1_e", 0), 0, "#16a34a")
    add("STEP2 DL全体", d.get("s2_s", 0), d.get("s2_e", 0), 0, "#dc2626")
    # 全retry_attemptイベントを時系列で収集（LIST/DL/CREATE_LINK全部で使う）
    retry_events = []
    for e in evs:
        if e["event"] in ("retry_attempt_start", "retry_raw_response", "retry_response_diag",
                          "retry_success", "retry_attempt_failed", "retry_backoff",
                          "retry_giving_up", "retry_exhausted", "retry_deadline_no_wait"):
            retry_events.append(e)

    if "list_s" in d:
        add("LIST_FOLDER", d["list_s"], d.get("list_e", d["list_s"]), 1, "#ef4444")
        # LISTのリトライ
        list_retries = [e for e in retry_events if e.get("rc") == "list"]
        if list_retries:
            lok = sum(1 for r in list_retries if r["event"]=="retry_success")
            lng = sum(1 for r in list_retries if r["event"] in ("retry_attempt_failed","retry_giving_up"))
            add("LIST retry", list_retries[0]["elapsed"], list_retries[-1]["elapsed"]+0.1, 2, "#b91c1c",
                f"✅{lok} ❌{lng}")
    if "diff_e" in d: add("difflib", d.get("list_e", 0), d["diff_e"], 1, "#f87171")
    if "res_s" in d and "res_e" in d: add("resolve_by_ai", d["res_s"], d["res_e"], 1, "#fca5a5")
    if "dl_s" in d: add("DL gather", d["dl_s"], d.get("dl_e", d["dl_s"]), 1, "#b91c1c")

    # DL各媒体 + 共有リンク状態 + リトライ履歴
    dl = {}
    for e in evs:
        if e["event"] == "dl_single_start": dl.setdefault(e.get("media_name", ""), {})["s"] = e["elapsed"]
        if e["event"] == "dl_single_finish":
            mn = e.get("media_name", ""); dl.setdefault(mn, {})["f"] = e["elapsed"]
            dl[mn]["kb"] = e.get("file_size_kb", "?")
            dl[mn]["share"] = e.get("share_url", "")
            dl[mn]["status"] = "ok"
        if e["event"] in ("dl_single_error", "dl_single_hard_timeout"):
            mn = e.get("media_name", ""); dl.setdefault(mn, {}).setdefault("f", e["elapsed"])
            dl[mn]["status"] = "error" if e["event"] == "dl_single_error" else "timeout"
            dl[mn].setdefault("kb", "?")
    share_ok = set(); share_ng = set()
    for e in evs:
        if e["event"] == "share_link_raw": share_ok.add(e.get("context", "").replace("share:", ""))
        if e["event"] == "share_link_error": share_ng.add(e.get("media_name", ""))

    for mn in sorted(dl, key=lambda m: -(dl[m].get("f", 0) - dl[m].get("s", 0))):
        v = dl[mn]
        sl = "✅" if mn in share_ok else ("❌" if mn in share_ng else "—")
        st = dl[mn].get("status", "?")
        st_icon = "✅" if st == "ok" else ("❌500" if st == "error" else "⏰timeout" if st == "timeout" else "?")
        add(mn[:20], v.get("s", 0), v.get("f", v.get("s", 0) + 1), 2, "#7f1d1d",
            f"{v.get('kb', '?')}KB DL:{st_icon} link:{sl}")
        # rcフィールドで正確に紐付け
        dl_retries = [e for e in retry_events if e.get("rc") == f"dl:{mn}"]
        link_retries = [e for e in retry_events if e.get("rc") == f"link:{mn}"]
        # DL: 初回試行 vs リトライ（attempt>=2）を分離
        dl_first = [r for r in dl_retries if r["event"] == "retry_attempt_start" and r.get("attempt") == 1]
        dl_retry_only = [r for r in dl_retries if r.get("attempt", 1) >= 2]
        dl_ok = sum(1 for r in dl_retries if r["event"] == "retry_success")
        dl_ng = sum(1 for r in dl_retries if r["event"] in ("retry_attempt_failed", "retry_giving_up", "retry_exhausted"))
        dl_attempts = sum(1 for r in dl_retries if r["event"] == "retry_attempt_start")
        if dl_retries:
            first = dl_retries[0]["elapsed"]; last = dl_retries[-1]["elapsed"]
            if dl_retry_only:
                retry_first = dl_retry_only[0]["elapsed"]
                add("DL retry", retry_first, max(last, retry_first + 0.1), 3, "#991b1b",
                    f"リトライ{dl_attempts-1}回 ✅{dl_ok} ❌{dl_ng}")
            elif dl_ng > 0:
                add("DL retry", first, max(last, first + 0.1), 3, "#991b1b",
                    f"試行{dl_attempts} ❌{dl_ng}")
        # CREATE_LINK: 同様に分離
        lk_retry_only = [r for r in link_retries if r.get("attempt", 1) >= 2]
        lk_ok = sum(1 for r in link_retries if r["event"] == "retry_success")
        lk_ng = sum(1 for r in link_retries if r["event"] in ("retry_attempt_failed", "retry_giving_up", "retry_exhausted"))
        lk_attempts = sum(1 for r in link_retries if r["event"] == "retry_attempt_start")
        if link_retries:
            first = link_retries[0]["elapsed"]; last = link_retries[-1]["elapsed"]
            if lk_retry_only:
                retry_first = lk_retry_only[0]["elapsed"]
                add("LINK retry", retry_first, max(last, retry_first + 0.1), 3, "#92400e",
                    f"リトライ{lk_attempts-1}回 ✅{lk_ok} ❌{lk_ng}")
            elif lk_ng > 0:
                add("LINK retry", first, max(last, first + 0.1), 3, "#92400e",
                    f"試行{lk_attempts} ❌{lk_ng}")

    add("STEP3b proposals", d.get("s3b_s", 0), d.get("s3b_e", 0), 0, "#9333ea")
    # remaining共有リンク（STEP3bと並列）
    if "rem_link_s" in d and "rem_link_e" in d:
        completed = 0; cancelled = 0
        for e in evs:
            if e["event"] == "remaining_links_harvested":
                completed = e.get("completed", 0); cancelled = e.get("cancelled", 0)
        add("remaining CREATE_LINK", d["rem_link_s"], d["rem_link_e"], 1, "#f59e0b",
            f"完了{completed}/中止{cancelled}")
        # remainingリンクのリトライ履歴
        rem_retries = [e for e in retry_events if str(e.get("rc", "")).startswith("rem_link:")]
        if rem_retries:
            rok = sum(1 for r in rem_retries if r["event"]=="retry_success")
            rng = sum(1 for r in rem_retries if r["event"] in ("retry_attempt_failed","retry_giving_up","retry_exhausted"))
            add("rem LINK retry", rem_retries[0]["elapsed"], rem_retries[-1]["elapsed"]+0.1, 2, "#d97706",
                f"✅{rok} ❌{rng}")

    # proposal各媒体
    pr = {}
    for e in evs:
        if e["event"] == "gen_single_start":
            mn = e.get("media_name", ""); pr.setdefault(mn, {"s": e["elapsed"], "kb": e.get("pdf_size_kb", "?"), "status": "pending"})
        if e["event"] == "llm_call_start" and e.get("context", "").startswith("proposal:"):
            mn = e["context"].replace("proposal:", ""); pr.setdefault(mn, {})["llm_s"] = e["elapsed"]
        if e["event"] == "llm_call_raw_response" and e.get("context", "").startswith("proposal:"):
            mn = e["context"].replace("proposal:", "")
            if "llm_r" not in pr.get(mn, {}): pr[mn]["llm_r"] = e["elapsed"]
        if e["event"] == "gen_single_received":
            mn = e.get("context", "").replace("proposal:", "")
            if mn and "f" not in pr.get(mn, {}):
                pr[mn]["f"] = e["elapsed"]
                pr[mn]["status"] = "ok"
        if e["event"] == "gen_single_output_size":
            mn = e.get("media_name", ""); pr.setdefault(mn, {})["out_chars"] = e.get("json_chars", 0)
        # タイムアウト/エラーの終了時刻も記録
        if e["event"].startswith("proposal_error_attempt"):
            mn = e.get("media_name", "")
            if mn in pr: pr[mn]["last_error"] = e["elapsed"]
        if e["event"] == "proposal_all_retries_exhausted":
            mn = e.get("media_name", "")
            if mn in pr:
                pr[mn].setdefault("f", e["elapsed"])
                pr[mn]["status"] = "exhausted"
                pr[mn]["total_attempts"] = e.get("total_attempts", "?")
        # リトライ成功
        if e["event"] == "proposal_retry_success":
            mn = e.get("media_name", "")
            if mn in pr: pr[mn]["status"] = "retry_ok"
        # gen_single_received_retry
        if e["event"] == "gen_single_received_retry":
            mn = e.get("context", "").replace("proposal_retry:", "")
            # proposal_retryN:mediaName の形式なので番号を除去
            mn = re.sub(r'^\d+:', '', mn)
            if mn in pr and "f" not in pr[mn]:
                pr[mn]["f"] = e["elapsed"]
                pr[mn]["status"] = "retry_ok"
    # fallback: "f"がない媒体は最後のエラー時刻か process_finish で埋める
    fin_t = d.get("fin", total)
    for mn in pr:
        if "f" not in pr[mn]:
            pr[mn]["f"] = pr[mn].get("last_error", fin_t)
            if pr[mn].get("status") == "pending": pr[mn]["status"] = "timeout"

    for mn in sorted(pr, key=lambda m: -(pr[m].get("f", 0) - pr[m].get("s", 0))):
        v = pr[mn]
        if not v.get("f") or not v.get("s"): continue
        llm_dur = v.get("llm_r", 0) - v.get("llm_s", 0) if "llm_s" in v and "llm_r" in v else 0
        out_chars = v.get("out_chars", 0)
        cps = round(out_chars / llm_dur) if llm_dur > 0 else 0
        status = v.get("status", "?")
        extra = f"PDF:{v.get('kb', '?')}KB"
        if status == "ok":
            if cps > 0: extra += f" {cps}字/s"
            color = "#7c3aed"
        elif status == "retry_ok":
            if cps > 0: extra += f" {cps}字/s"
            extra += " 🔄retry→✅"
            color = "#7c3aed"
        elif status in ("timeout", "exhausted"):
            extra += f" ❌timeout×{v.get('total_attempts', '?')}"
            color = "#dc2626"
        else:
            extra += f" ❌{status}"
            color = "#dc2626"
        add(mn[:20], v["s"], v["f"], 1, color, extra)
        if "llm_s" in v and "llm_r" in v:
            add("LLM待ち", v["llm_s"], v["llm_r"], 2, "#a78bfa", f"{out_chars}字→{cps}字/s" if cps else "")

    add("STEP5 HTML", d.get("s3b_e", 0), d.get("html_e", d.get("s3b_e", 0)), 0, "#ca8a04")
    add("STEP6 finalize", d.get("html_e", 0), d.get("fin", 0), 0, "#6b7280")

    return bars, total





_VIEWER_TEMPLATE = r"""<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Debug Viewer</title>
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;600&family=Noto+Sans+JP:wght@400;500;700&display=swap" rel="stylesheet">
<style>
:root{--bg:#0c0c0c;--surface:#161616;--s2:#1e1e1e;--border:#2a2a2a;--text:#e0e0e0;--t2:#888;--accent:#7c3aed;--green:#22c55e;--red:#ef4444;--yellow:#f59e0b;--blue:#3b82f6;--mono:'JetBrains Mono',monospace;--sans:'Noto Sans JP',sans-serif;}
*{margin:0;padding:0;box-sizing:border-box;}body{font-family:var(--sans);background:var(--bg);color:var(--text);font-size:13px;line-height:1.5;}
.app{display:flex;height:100vh;}.sidebar{width:180px;background:var(--surface);border-right:1px solid var(--border);padding:16px 0;flex-shrink:0;overflow-y:auto;}
.sidebar h1{font-size:11px;font-weight:700;padding:0 16px 12px;color:var(--t2);text-transform:uppercase;letter-spacing:1px;}
.nav-item{padding:8px 16px;cursor:pointer;font-size:12px;border-left:3px solid transparent;transition:all .15s;}
.nav-item:hover{background:var(--s2);}.nav-item.active{border-left-color:var(--accent);background:var(--s2);color:#fff;font-weight:600;}
.nav-count{float:right;font-family:var(--mono);font-size:10px;color:var(--t2);background:var(--bg);padding:1px 6px;border-radius:8px;}
.main{flex:1;overflow-y:auto;padding:24px 32px;}.view{display:none;}.view.active{display:block;}
.kpi-row{display:flex;gap:14px;margin-bottom:16px;flex-wrap:wrap;}
.kpi{background:var(--surface);border:1px solid var(--border);border-radius:6px;padding:10px 14px;min-width:110px;}
.kpi-label{font-size:9px;color:var(--t2);text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px;}
.kpi-val{font-family:var(--mono);font-size:18px;font-weight:600;}
.search-bar{display:flex;gap:8px;margin-bottom:12px;}
.search-bar input{flex:1;background:var(--surface);border:1px solid var(--border);border-radius:6px;padding:7px 12px;color:var(--text);font-family:var(--mono);font-size:12px;outline:none;}
.search-bar input:focus{border-color:var(--accent);}
.search-bar select{background:var(--surface);border:1px solid var(--border);border-radius:6px;padding:7px 10px;color:var(--text);font-size:12px;outline:none;}
table{width:100%;border-collapse:collapse;font-size:11px;}
th{text-align:left;padding:5px 8px;background:var(--surface);font-size:10px;color:var(--t2);text-transform:uppercase;letter-spacing:.3px;border-bottom:1px solid var(--border);position:sticky;top:0;}
td{padding:4px 8px;border-bottom:1px solid var(--border);font-family:var(--mono);font-size:11px;vertical-align:top;max-width:500px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
tr:hover{background:var(--s2);}tr.error{color:var(--red);}tr.success{color:var(--green);}
.expand{cursor:pointer;color:var(--accent);text-decoration:underline;}
.detail-row{display:none;}.detail-row.open{display:table-row;}
.detail-row td{white-space:pre-wrap;word-break:break-all;background:var(--surface);max-width:none;padding:12px;font-size:11px;line-height:1.6;}
.card{background:var(--surface);border:1px solid var(--border);border-radius:8px;margin-bottom:12px;overflow:hidden;}
.card-header{display:flex;align-items:center;gap:10px;padding:10px 16px;cursor:pointer;border-bottom:1px solid var(--border);}
.card-header:hover{background:var(--s2);}.card-title{font-weight:700;font-size:13px;flex:1;}
.card-badge{font-family:var(--mono);font-size:10px;padding:2px 8px;border-radius:4px;}
.badge-ok{background:#052e16;color:var(--green);}.badge-err{background:#2c0b0b;color:var(--red);}.badge-retry{background:#1c1508;color:var(--yellow);}
.card-body{display:none;padding:12px 16px;}.card.open .card-body{display:block;}
.ev-line{display:flex;gap:10px;padding:3px 0;border-bottom:1px solid #1a1a1a;font-size:11px;}
.ev-time{font-family:var(--mono);color:var(--t2);width:50px;flex-shrink:0;text-align:right;}
.ev-name{font-family:var(--mono);width:200px;flex-shrink:0;overflow:hidden;text-overflow:ellipsis;}
.ev-detail{color:var(--t2);flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;cursor:pointer;}.ev-detail:hover{white-space:normal;word-break:break-all;}
.ev-name.err{color:var(--red);}.ev-name.ok{color:var(--green);}.ev-name.llm{color:var(--accent);}.ev-name.retry{color:var(--yellow);}
h2{font-size:15px;font-weight:700;margin-bottom:14px;}h3{font-size:13px;font-weight:700;margin:14px 0 6px;color:var(--t2);}
.prompt-box{background:var(--bg);border:1px solid var(--border);border-radius:6px;padding:10px;font-family:var(--mono);font-size:11px;white-space:pre-wrap;word-break:break-all;max-height:400px;overflow-y:auto;margin:6px 0;}
.tag{display:inline-block;font-size:9px;padding:1px 6px;border-radius:3px;margin-right:4px;font-family:var(--mono);}
.tag-llm{background:#2d1b69;color:#a78bfa;}.tag-mcp{background:#2c0b0b;color:#f87171;}
.param-chip{display:inline-block;font-size:11px;padding:3px 8px;border:1px solid var(--border);border-radius:4px;margin:2px;font-family:var(--mono);}
.param-chip.w{border-color:var(--accent);color:var(--accent);}.param-chip.s{border-color:var(--blue);color:var(--blue);}
.param-section{background:var(--surface);border:1px solid var(--border);border-radius:6px;padding:12px 16px;margin:8px 0;}
.param-section-title{font-size:11px;font-weight:700;color:var(--t2);text-transform:uppercase;margin-bottom:6px;}
.param-row{display:flex;align-items:center;gap:12px;padding:4px 0;font-size:12px;}
.param-key{font-family:var(--mono);color:var(--text);min-width:120px;}
.param-val{font-family:var(--mono);color:var(--accent);}
.prop-field{margin:8px 0;}.prop-field-name{font-size:10px;color:var(--t2);text-transform:uppercase;margin-bottom:2px;}
.prop-field-val{font-size:12px;line-height:1.6;}
.g-scale{display:flex;justify-content:space-between;font-family:var(--mono);font-size:10px;color:#555;padding:0 0 4px 200px;border-bottom:1px solid var(--border);}
.g-row{display:flex;align-items:center;height:24px;border-bottom:1px solid #111;cursor:pointer;}.g-row:hover{background:var(--s2);}
.g-label{width:200px;flex-shrink:0;font-size:11px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-family:var(--mono);}
.g-label.l0{color:#fff;font-weight:700;}.g-label.l1{padding-left:16px;color:#bbb;}.g-label.l2{padding-left:32px;color:#777;font-size:10px;}.g-label.l3{padding-left:48px;color:#666;font-size:10px;}
.g-track{flex:1;position:relative;height:18px;}.g-bar{position:absolute;height:14px;top:2px;border-radius:2px;min-width:2px;cursor:default;}.g-bar:hover{opacity:.8;}
.g-dur{position:absolute;top:3px;font-family:var(--mono);font-size:10px;color:#666;}
.g-tip{display:none;position:fixed;background:#1a1a1a;border:1px solid #333;border-radius:4px;padding:8px 12px;font-size:11px;z-index:100;pointer-events:none;font-family:var(--mono);}
.media-detail{background:var(--bg);border:1px solid var(--border);border-radius:6px;margin:8px 0 8px 200px;padding:12px;display:none;}
.media-detail.open{display:block;}
</style></head>
<body>
<div class="app">
<div class="sidebar">
  <h1>Debug Viewer</h1>
  <div class="nav-item active" onclick="showView('summary',this)">サマリー</div>
  <div class="nav-item" onclick="showView('timeline',this)">タイムライン</div>
  <div class="nav-item" onclick="showView('proposals',this)">提案内容 <span class="nav-count" id="navPropCount"></span></div>
  <div class="nav-item" onclick="showView('llm',this)">LLM詳細 <span class="nav-count" id="navLlmCount"></span></div>
  <div class="nav-item" onclick="showView('raw',this)">生ログ <span class="nav-count" id="navRawCount"></span></div>
</div>
<div class="main">

<!-- 1. サマリー（KPI + パラメータ + 提案結果を統合） -->
<div class="view active" id="view-summary">
  <h2>サマリー</h2>
  <div class="kpi-row" id="kpiRow"></div>
  <div id="userReq"></div>
  <div id="paramSection"></div>
  <h3>提案結果</h3>
  <div id="proposalSummary"></div>
</div>

<!-- 2. タイムライン（ガント + メディアクリック展開） -->
<div class="view" id="view-timeline">
  <h2>タイムライン <span id="ganttTotal" style="font-family:var(--mono);font-size:13px;color:var(--t2);font-weight:400"></span></h2>
  <div id="ganttChart"></div>
</div>

<!-- 3. 提案内容 -->
<div class="view" id="view-proposals"><h2>提案内容</h2><div id="proposalCards"></div></div>

<!-- 4. LLM詳細 -->
<div class="view" id="view-llm"><h2>LLMコール詳細</h2><div id="llmCards"></div></div>

<!-- 5. 生ログ -->
<div class="view" id="view-raw"><h2>生ログ</h2>
  <div class="search-bar"><input type="text" id="rawSearch" placeholder="検索..." oninput="filterRaw()"><select id="rawFilter" onchange="filterRaw()"><option value="">全て</option><option value="llm_call">LLM</option><option value="retry">リトライ</option><option value="error">エラー</option><option value="proposal">提案</option><option value="dl_">DL</option><option value="gen_single">生成</option></select></div>
  <table><thead><tr><th>#</th><th>経過</th><th>イベント</th><th>コンテキスト</th><th>詳細</th></tr></thead><tbody id="rawBody"></tbody></table>
</div>

</div></div>
<script>
var D=__DATA_JSON__;
var G=__GANTT_JSON__;

function showView(id,el){document.querySelectorAll('.view').forEach(function(v){v.classList.remove('active');});document.querySelectorAll('.nav-item').forEach(function(n){n.classList.remove('active');});document.getElementById('view-'+id).classList.add('active');el.classList.add('active');}
function esc(s){var d=document.createElement('div');d.textContent=String(s);return d.innerHTML;}
function summarize(v){if(v===null||v===undefined)return '';if(typeof v==='string')return v.length>120?v.substring(0,120)+'…':v;if(typeof v==='number'||typeof v==='boolean')return String(v);if(Array.isArray(v))return '['+v.length+']';if(typeof v==='object'){if(v.type&&v.value!==undefined)return v.type+'='+summarize(v.value);var k=Object.keys(v);return '{'+k.slice(0,3).join(',')+(k.length>3?'…':'')+'}';};return String(v);}
function findOne(evt){return D.find(function(e){return e.event===evt;});}
function evCls(evt){if(!evt)return '';if(evt.indexOf('error')>=0||evt.indexOf('timeout')>=0||evt.indexOf('exception')>=0)return 'err';if(evt.indexOf('success')>=0||evt.indexOf('finish')>=0)return 'ok';if(evt.indexOf('llm_call')>=0)return 'llm';if(evt.indexOf('retry')>=0)return 'retry';return '';}
var SKIP={seq:1,ts:1,elapsed:1,event:1,src:1};
function detailStr(e,xs){var s='';var sk=Object.assign({},SKIP,xs||{});Object.keys(e).forEach(function(k){if(!sk[k])s+=k+'='+summarize(e[k])+'  ';});return s;}
function getMediaEvents(mn){return D.filter(function(e){var m=e.media_name||'';if(!m){var c=e.context||'';if(c.indexOf('proposal')===0)m=c.replace(/^proposal(_retry\d*)?:/,'');else if(c.indexOf('dl:')===0)m=c.replace('dl:','');else if(c.indexOf('link:')===0)m=c.replace('link:','');else if(c.indexOf('share:')===0)m=c.replace('share:','');else if(e.rc)m=e.rc.replace(/^(dl|link|rem_link):/,'');}return m===mn;});}
// 値を安全に取り出す（_describe形式対応）
function val(obj,key){if(!obj)return undefined;var v=obj[key];if(!v)return undefined;if(typeof v==='object'&&v.value!==undefined)return v.value;return v;}
function valObj(obj,key){if(!obj)return {};var v=obj[key];if(!v)return {};if(typeof v==='object'&&v.items)return v.items;return v;}

// === 1. サマリー ===
(function(){
  var total=0,media=new Set(),llmC=0,errors=0,props=[],req='',turn=1;
  D.forEach(function(e){
    if(e.event==='process_finish')total=e.total_duration_sec||0;
    if(e.event==='process_start'){req=e.user_request||'';turn=e.turn||1;}
    if(e.media_name&&e.media_name!=='?')media.add(e.media_name);
    if(e.event==='llm_call_start')llmC++;
    if(e.event&&(e.event.indexOf('error')>=0||e.event.indexOf('timeout')>=0))errors++;
    if(e.event==='gen_single_validation')props.push(e);
  });
  var kpis=[['処理時間',total.toFixed(1)+'s',''],['ターン',turn,''],['メディア',media.size,''],['LLM',llmC,''],['エラー',errors,errors>0?'color:var(--red)':''],['提案',props.length+'/'+media.size,'']];
  document.getElementById('kpiRow').innerHTML=kpis.map(function(k){return '<div class="kpi"><div class="kpi-label">'+k[0]+'</div><div class="kpi-val" style="'+k[2]+'">'+k[1]+'</div></div>';}).join('');
  document.getElementById('userReq').innerHTML='<h3>ユーザー要望</h3><div class="prompt-box">'+esc(req)+'</div>';
  // パラメータセクション
  var pe=findOne('parse_user_request_full');
  var sf=findOne('score_filter');
  var sl=findOne('score_load_master');
  var sm=findOne('score_media_split');
  var ph='';
  if(pe&&pe.full&&pe.full.items){
    var it=pe.full.items;
    // フィルタ
    ph+='<div class="param-section"><div class="param-section-title">フィルタ</div>';
    var fi=valObj(it,'filters');
    if(fi&&Object.keys(fi).length>0){
      Object.keys(fi).forEach(function(k){var v=fi[k];var op=val(v,'op')||'';var vl=val(v,'value')||'';ph+='<div class="param-row"><span class="param-key">'+esc(k)+'</span><span class="param-val">'+esc(op)+' '+esc(vl)+'</span></div>';});
    } else ph+='<div style="color:var(--t2)">フィルタなし</div>';
    if(sf&&sl)ph+='<div class="param-row" style="color:var(--t2)">'+sl.rows+'件 → '+sf.rows_after+'件に絞り込み</div>';
    ph+='</div>';
    // ソート
    var sortBy=val(it,'sort_by');var sortOrd=val(it,'sort_order')||'';
    ph+='<div class="param-section"><div class="param-section-title">並べ替え</div>';
    if(sortBy){ph+='<div class="param-row"><span class="param-key">sort_by</span><span class="param-val" style="color:var(--blue)">'+esc(sortBy)+' '+esc(sortOrd)+'</span></div>';}
    else ph+='<div style="color:var(--t2)">ソートなし（スコア順）</div>';
    ph+='</div>';
    // 重み
    var wi=valObj(it,'weights');
    ph+='<div class="param-section"><div class="param-section-title">重み</div>';
    if(wi&&Object.keys(wi).length>0){
      Object.keys(wi).forEach(function(k){var w=val(wi,k)||0;ph+='<div class="param-row"><span class="param-key">'+esc(k)+'</span><span class="param-val">'+w+'</span></div>';});
    } else ph+='<div style="color:var(--t2)">重みなし'+(sortBy?' （sort_byで並べ替え）':'')+'</div>';
    ph+='</div>';
    // その他
    ph+='<div class="param-section"><div class="param-section-title">その他</div>';
    ph+='<div class="param-row"><span class="param-key">article_type</span><span class="param-val">'+esc(val(it,'article_type')||'—')+'</span></div>';
    ph+='<div class="param-row"><span class="param-key">target_attribute</span><span class="param-val">'+esc(val(it,'target_attribute')||'—')+'</span></div>';
    ph+='<div class="param-row"><span class="param-key">user_budget</span><span class="param-val">'+esc(val(it,'user_budget')||0)+'</span></div>';
    if(sm)ph+='<div class="param-row"><span class="param-key">選定</span><span class="param-val">top='+sm.top_n+' / remaining='+sm.remaining+'</span></div>';
    ph+='</div>';
  }
  document.getElementById('paramSection').innerHTML=ph;
  // 提案結果
  var rh='';
  props.forEach(function(p){
    var r=p.exec_ratio_value;var rs=(r===null||r===undefined)?'?':(r===-1?'<span style="color:var(--yellow)">-1</span>':r+'%');
    rh+='<div class="ev-line"><span class="ev-time">'+rs+'</span><span class="ev-name">'+esc(p.media_name||'?')+'</span><span class="ev-detail">'+esc(p.exec_ratio_note||'')+'</span></div>';
  });
  document.getElementById('proposalSummary').innerHTML=rh||'なし';
  document.getElementById('navPropCount').textContent=props.length;
  document.getElementById('navLlmCount').textContent=llmC;
  document.getElementById('navRawCount').textContent=D.length;
})();

// === 2. タイムライン（ガント + メディアクリック展開） ===
(function(){
  var T=G.total||1,B=G.bars||[];
  document.getElementById('ganttTotal').textContent='総処理時間: '+T.toFixed(1)+'s';
  var ch=document.getElementById('ganttChart');
  var sc=document.createElement('div');sc.className='g-scale';
  for(var i=0;i<=Math.ceil(T);i+=10){var s=document.createElement('span');s.textContent=i+'s';sc.appendChild(s);}
  ch.appendChild(sc);
  var tip=document.createElement('div');tip.className='g-tip';document.body.appendChild(tip);
  B.forEach(function(b,bi){
    var r=document.createElement('div');r.className='g-row';r.setAttribute('data-media',b.media||'');
    var l=document.createElement('div');l.className='g-label l'+b.level;l.textContent=b.label;r.appendChild(l);
    var t=document.createElement('div');t.className='g-track';
    var br=document.createElement('div');br.className='g-bar';
    br.style.left=(b.start/T*100)+'%';br.style.width=Math.max(b.dur/T*100,.3)+'%';br.style.background=b.color;
    br.onmouseenter=function(ev){tip.style.display='block';tip.innerHTML='<strong>'+b.label+'</strong><br>'+b.start.toFixed(1)+'s → '+b.end.toFixed(1)+'s<br>'+b.dur.toFixed(2)+'s ('+(b.dur/T*100).toFixed(1)+'%)'+(b.extra?'<br>'+b.extra:'');};
    br.onmousemove=function(ev){tip.style.left=(ev.clientX+12)+'px';tip.style.top=(ev.clientY-40)+'px';};
    br.onmouseleave=function(){tip.style.display='none';};
    t.appendChild(br);
    var d=document.createElement('span');d.className='g-dur';d.textContent=b.dur>=1?b.dur.toFixed(1)+'s':b.dur>=.01?(b.dur*1000).toFixed(0)+'ms':'<1ms';d.style.left=(b.end/T*100+.5)+'%';t.appendChild(d);
    // メディア行は展開マーカー付き
    if(b.media){l.textContent='▶ '+b.label;l.style.cursor='pointer';}
    r.appendChild(t);ch.appendChild(r);
    // メディアクリック展開
    if(b.media){
      var detail=document.createElement('div');detail.className='media-detail';detail.id='md-'+bi;
      var evs=getMediaEvents(b.media);
      var dh='<div style="font-weight:700;margin-bottom:8px">'+esc(b.media)+' — '+evs.length+' events</div>';
      evs.forEach(function(e){
        dh+='<div class="ev-line"><span class="ev-time">'+(e.elapsed||0).toFixed(1)+'s</span><span class="ev-name '+evCls(e.event)+'">'+esc(e.event||'')+'</span><span class="ev-detail">'+esc(detailStr(e,{media_name:1,context:1,full:1,traceback:1}))+'</span></div>';
      });
      detail.innerHTML=dh;ch.appendChild(detail);
      (function(rr,dd,lbl,bl){rr.onclick=function(ev){if(ev&&ev.stopPropagation)ev.stopPropagation();var op=dd.classList.contains('open');if(op){dd.classList.remove('open');lbl.textContent='▶ '+bl;}else{dd.classList.add('open');lbl.textContent='▼ '+bl;}};})(r,detail,l,b.label);
    }
  });
})();

// === 3. 提案内容 ===
(function(){
  var propEvs=D.filter(function(e){return e.event==='gen_single_received'||e.event==='gen_single_received_retry';});
  var startEvs=D.filter(function(e){return e.event==='gen_single_start';});
  var pageEvs=D.filter(function(e){return e.event==='gen_single_page_refs';});
  // weights + score取得
  var weights={};var sortBy='';var sortOrd='';
  var pe2=findOne('parse_user_request_full');
  if(pe2&&pe2.full&&pe2.full.items){
    var it2=pe2.full.items;var wi=valObj(it2,'weights');
    if(wi)Object.keys(wi).forEach(function(k){weights[k]=val(wi,k)||0;});
    sortBy=val(it2,'sort_by')||'';sortOrd=val(it2,'sort_order')||'';
  }
  var scoreMap={};
  var smr=findOne('score_media_result_full');
  if(smr&&smr.full&&smr.full.items){smr.full.items.forEach(function(m){var mi=m.items||{};var nm=val(mi,'メディア名称')||'';if(nm)scoreMap[nm]={score:val(mi,'score'),attrs:mi};});}
  var hasW=Object.keys(weights).length>0;
  var h='<div class="param-section"><div class="param-section-title">スコアリング</div>';
  if(hasW){Object.keys(weights).forEach(function(k){h+='<div class="param-row"><span class="param-key">'+esc(k)+'</span><span class="param-val">weight='+weights[k]+'</span></div>';});}
  else{h+='<div style="color:var(--t2)">重みなし'+(sortBy?' → '+esc(sortBy)+' '+esc(sortOrd)+' で並べ替え':'')+'</div>';}
  h+='</div>';
  propEvs.forEach(function(pe){
    try{
      var mn=(pe.context||'').replace(/^proposal(_retry\d*)?:/,'');
      var full=pe.full;
      if(!full||!full.items){h+='<div class="card"><div class="card-header"><span class="card-title">'+esc(mn)+' (no data)</span></div></div>';return;}
      var it=full.items;
      var se=startEvs.find(function(s){return s.media_name===mn;});
      var totalP=se?se.pdf_pages:'?';
      var pre=pageEvs.find(function(p){return p.media_name===mn;});
      var refs=pre?pre.pages_referenced:[];
      var maxR=refs.length>0?Math.max.apply(null,refs):0;
      var hasHal=typeof totalP==='number'&&maxR>totalP;
      var badge=hasHal?'<span class="card-badge badge-err">HAL</span>':'<span class="card-badge badge-ok">OK</span>';
      h+='<div class="card"><div class="card-header" onclick="this.parentElement.classList.toggle(\'open\')"><span class="card-title">'+esc(mn)+'</span>'+badge+'<span style="font-family:var(--mono);font-size:11px;color:var(--t2)">'+totalP+'P</span></div><div class="card-body">';
      // exec_ratio
      var erObj=it.exec_ratio||{};
      var erv=(erObj.value!==undefined)?erObj.value:'?';
      var ernObj=it.exec_ratio_note||{};
      var ernv=(ernObj.value!==undefined)?ernObj.value:'';
      var sm=scoreMap[mn];
      if(sm){
        h+='<div class="prop-field"><div class="prop-field-name">スコア = '+(sm.score!==undefined?sm.score:'?')+'</div><div class="prop-field-val">';
        if(hasW){Object.keys(weights).forEach(function(wk){var av=val(sm.attrs,wk);if(av===undefined)av=0;h+='<div class="param-row"><span class="param-key">'+esc(wk)+'</span><span class="param-val">'+av+' × '+weights[wk]+' = '+(av*weights[wk]).toFixed(2)+'</span></div>';});}
        else{h+='<span style="color:var(--t2)">'+(sortBy?esc(sortBy)+' 順':'—')+'</span>';}
        h+='</div></div>';
      }
      h+='<div class="prop-field"><div class="prop-field-name">exec_ratio</div><div class="prop-field-val">'+(erv===-1?'<span style="color:var(--yellow)">-1</span>':erv+'%')+' '+esc(ernv)+'</div></div>';
      // pages
      h+='<div class="prop-field"><div class="prop-field-name">参照ページ</div><div class="prop-field-val" style="'+(hasHal?'color:var(--red)':'')+'">refs=['+refs.join(',')+'] max='+maxR+' / total='+totalP+'</div></div>';
      // sections
      var secs=['sec_overview','sec_readers','sec_menu_price','sec_lead_gen','sec_clients','sec_editorial','reason','ad_menus','tags'];
      for(var si=0;si<secs.length;si++){
        var sk=secs[si];
        var sv=it[sk];
        if(!sv)continue;
        var txt='';
        if(sv.type==='str'){txt=sv.value||'';}
        else if(sv.type==='list'&&sv.items){txt=sv.items.map(function(x){return(x&&x.value!==undefined)?x.value:JSON.stringify(x);}).join('\n');}
        else{txt=JSON.stringify(sv);}
        var warn='';
        if(typeof totalP==='number'&&txt){
          var m2=txt.match(/\d+/g);
          if(m2){for(var mi=0;mi<m2.length;mi++){
            var idx=txt.indexOf(m2[mi]);
            if(idx>=0&&idx+m2[mi].length<txt.length&&txt[idx+m2[mi].length]==='P'){
              var pn=parseInt(m2[mi]);
              if(pn>totalP&&pn<1000)warn+=' '+pn+'P';
            }
          }}
        }
        h+='<div class="prop-field"><div class="prop-field-name">'+esc(sk)+(warn?'<span style="color:var(--red)"> ⚠'+warn+'超え</span>':'')+'</div><div class="prompt-box">'+esc(txt)+'</div></div>';
      }
      h+='</div></div>';
    }catch(err){
      console.error('Proposal error:',err);
      h+='<div class="card"><div class="card-header"><span class="card-title">'+esc(mn||'?')+' ERROR</span><span class="card-badge badge-err">JS Error</span></div></div>';
    }
  });
  document.getElementById('proposalCards').innerHTML=h||'<div style="color:var(--t2)">提案なし</div>';
})();

// === 4. LLM詳細 ===
(function(){
  var pairs=[],starts={};
  D.forEach(function(e){
    if(e.event==='llm_call_start')starts[e.context||'?']={start:e,events:[e]};
    else if(e.event&&e.event.indexOf('llm_call_')===0){var s=starts[e.context||'?'];if(s)s.events.push(e);if(e.event==='llm_call_return'||e.event==='llm_call_exception'){if(s){s.end=e;pairs.push(s);delete starts[e.context||'?'];}}}
  });
  var h='';
  pairs.forEach(function(p){
    var ctx=p.start.context||'?';
    var dur=p.end?((p.end.elapsed||0)-(p.start.elapsed||0)).toFixed(1)+'s':'?';
    var isErr=p.end&&p.end.event==='llm_call_exception';
    var badge=isErr?'<span class="card-badge badge-err">'+dur+'</span>':'<span class="card-badge badge-ok">'+dur+'</span>';
    var inputEv=p.events.find(function(e){return e.event==='llm_call_input_full';});
    var prompt=inputEv?(inputEv.prompt||''):'';
    var schema=inputEv?(inputEv.schema_cleaned||inputEv.schema_in||null):null;
    h+='<div class="card"><div class="card-header" onclick="this.parentElement.classList.toggle(\'open\')"><span class="card-title"><span class="tag tag-llm">LLM</span> '+esc(ctx)+'</span>'+badge+'</div><div class="card-body">';
    if(prompt)h+='<h3>プロンプト</h3><div class="prompt-box">'+esc(prompt)+'</div>';
    if(schema)h+='<h3>スキーマ</h3><div class="prompt-box">'+esc(JSON.stringify(schema,null,2))+'</div>';
    p.events.forEach(function(e){h+='<div class="ev-line"><span class="ev-time">'+(e.elapsed||0).toFixed(1)+'s</span><span class="ev-name '+evCls(e.event)+'">'+esc(e.event)+'</span><span class="ev-detail">'+esc(detailStr(e,{context:1,prompt:1,schema_in:1,schema_cleaned:1,full:1}))+'</span></div>';});
    h+='</div></div>';
  });
  document.getElementById('llmCards').innerHTML=h;
})();

// === 5. 生ログ ===
(function(){
  var h='';
  D.forEach(function(e){
    var cls=e.event&&(e.event.indexOf('error')>=0||e.event.indexOf('timeout')>=0||e.event.indexOf('exception')>=0)?'error':(e.event&&(e.event.indexOf('success')>=0||e.event.indexOf('finish')>=0)?'success':'');
    var detail=detailStr(e,{context:1,full:1,prompt:1,schema_in:1,schema_cleaned:1,traceback:1});
    var full=JSON.stringify(e,null,2);
    h+='<tr class="'+cls+'" data-search="'+esc(((e.event||'')+' '+(e.context||'')+' '+detail).toLowerCase())+'"><td>'+e.seq+'</td><td>'+(e.elapsed||0).toFixed(1)+'s</td><td>'+esc(e.event||'')+'</td><td>'+esc(e.context||'')+'</td><td><span class="expand" onclick="toggleDetail(this)">'+esc(detail.substring(0,80))+'</span></td></tr>';
    h+='<tr class="detail-row"><td colspan="5">'+esc(full)+'</td></tr>';
  });
  document.getElementById('rawBody').innerHTML=h;
})();

function toggleDetail(el){var row=el.closest('tr');var next=row.nextElementSibling;if(next&&next.classList.contains('detail-row'))next.classList.toggle('open');}
function filterRaw(){var q=document.getElementById('rawSearch').value.toLowerCase();var f=document.getElementById('rawFilter').value;document.getElementById('rawBody').querySelectorAll('tr:not(.detail-row)').forEach(function(row){var search=row.getAttribute('data-search')||'';var ev=row.children[2]?row.children[2].textContent:'';var show=true;if(q&&search.indexOf(q)<0)show=false;if(f&&ev.indexOf(f)<0)show=false;row.style.display=show?'':'none';var next=row.nextElementSibling;if(next&&next.classList.contains('detail-row')&&!show)next.classList.remove('open');});}
</script>
</body></html>"""

def _build_debug_viewer(jsonl_path: str, out_path: str) -> str:
    """debug.jsonl -> インタラクティブHTMLビューアー（サマリ/タイムライン/メディア別/LLM/エラー/生ログ）。"""
    events = [json.loads(l) for l in open(jsonl_path, encoding="utf-8")]
    data_json = json.dumps(events, ensure_ascii=False, default=str)
    data_json = (data_json.replace("</", "<\\/")
                 .replace("\u2028", "\\u2028").replace("\u2029", "\\u2029"))
    bars, total = compute_gantt_bars(events)
    gantt_json = json.dumps({"bars": bars, "total": total}, ensure_ascii=False)
    htm = _VIEWER_TEMPLATE.replace("__DATA_JSON__", data_json).replace("__GANTT_JSON__", gantt_json)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(htm)
    return out_path

async def _finalize_output(html_path: str):
    """全処理完了後にファイルをユーザーへ共有する。設計上ここが唯一のfile_output呼び出し。"""
    _dlog("finalize_output_start", {"path": html_path})
    from agent_sdk import file_output

    # 1) HTMLレポート
    try:
        res = await file_output(html_path)
        _dlog_full("file_output_result", "finalize", res)
    except Exception as e:
        import traceback
        _dlog("file_output_error", {"target": "html", "error": str(e), "traceback": traceback.format_exc()})

    # 2) デバッグログ（file_outputしない。tmp/に生成済み）
    dbg_out = ""   # コピー失敗時に手順3がNameErrorにならないよう先に定義
    try:
        import shutil
        os.makedirs("output", exist_ok=True)
        dbg_out = "output/" + os.path.basename(_DEBUG_PATH)
        shutil.copyfile(_DEBUG_PATH, dbg_out)
        _dlog("debug_output_done", {"path": dbg_out})
    except Exception as e:
        import traceback
        dbg_out = ""
        _dlog("file_output_error", {"target": "debug", "error": str(e), "traceback": traceback.format_exc()})

    # 3) デバッグビューアーHTML（file_outputしない。output/に生成済み）
    try:
        if not dbg_out:
            raise RuntimeError("debugログのコピーに失敗したためビューアー生成をスキップ")
        viewer_path = html_path.replace(".html", "_viewer.html")
        _build_debug_viewer(dbg_out, viewer_path)
        _dlog("viewer_output_done", {"path": viewer_path})
    except Exception as e:
        import traceback
        _dlog("viewer_output_error", {"error": str(e), "traceback": traceback.format_exc()})


# ============================================================
# ターン2以降
# ============================================================

async def handle_turn2(user_request_2: str, params: dict = None, mode: str = None, targets: list = None):
    """
    ターン2以降のエントリーポイント。
    mode="deepen"が渡されたらLLM判定をスキップ。
    再検索はmain()を直接呼ぶ方式を推奨（handle_turn2経由不要）。
    """
    global ptc_state

    if not ptc_state:
        print("ERROR: ptc_stateが空です。先にmain()を実行してください。")
        return

    # ターン数インクリメント
    ptc_state["turn_count"] = ptc_state.get("turn_count", 1) + 1
    turn = ptc_state["turn_count"]
    _reset_debug_path(turn)
    _dlog("turn2_start", {"turn": turn, "user_request_2": user_request_2})
    print(f"\n{'='*60}")
    print(f"[セッション{turn}] 追加要望: {user_request_2}")
    print(f"{'='*60}")

    # 前回のコンテキストを構築
    prev_req = ptc_state.get("user_request", "")
    prev_proposals = ptc_state.get("proposals", [])
    prev_summary = f"【前回の要望】{prev_req}\n【前回の選定結果】\n"
    for p in prev_proposals:
        if isinstance(p, dict):
            prev_summary += (f"- {p.get('media_name','')}: {p.get('price_disp','?')} "
                           f"/ {p.get('reason','')[:60]}\n")

    # 分岐判定
    if mode:
        # メインAIが判定済み
        print(f"\n  モード: {mode}（メインAI設定）")
        targets = targets or []
    else:
        # LLMでモード判定（フォールバック）
        mode_res = await _llm_call_safe(
        prompt=f"""{prev_summary}

【今回のユーザー発言】
{user_request_2}

上記を踏まえて「深掘り」か「再検索」に分類してください。
- 深掘り(deepen): 以下のいずれかに該当する場合のみ
  ① 特定の1つの媒体名を挙げて「もっと詳しく」「深掘りして」と言っている
  ② 既存の提案結果に対して「もっと詳細に」「全部の媒体を深掘り」など、今ある7件の提案をより詳しくしてほしいことが明確
  例: 「MONOistをもっと詳しく教えて」「全媒体の詳細を見たい」「提案内容をもっと詳しく」
- 再検索(re_search): それ以外は全てこちら。並べ替え、フィルタ変更、条件追加、比較、予算変更など
  例: 「製造業比率が高い順に」「決裁者重視で」「安い順に変えて」「JBpressとTECH+を比較して」

{{"mode": "deepen" or "re_search", "targets": ["メディア名"] or []}}
""",
        schema={
            "type": "object",
            "properties": {
                "mode":    {"type": "string", "enum": ["deepen", "re_search"]},
                "targets": {"type": "array", "items": {"type": "string"}},
            },
            "required": ["mode"],
        },
        context="handle_turn2_mode_classify"
    )
        mode    = mode_res["mode"]
        targets = mode_res.get("targets", [])

    # 前回コンテキスト付きのリクエストを渡す
    enriched_request = f"{prev_summary}\n【今回の追加要望】{user_request_2}"

    if mode == "deepen":
        await _handle_deepen(enriched_request, targets)
    else:
        await _handle_research(enriched_request, params)


async def _handle_deepen(user_request_2: str, targets: list):
    """深掘り: 対象メディアのPDFを再読みして詳細提案"""
    global ptc_state

    # 照合キーは常にsplit後の正規名（sandbox_paths等のキーと揃える）
    targets = [t.split(";")[0].strip() for t in (targets or []) if t]
    if not targets:
        targets = [r.get("メディア名称", "").split(";")[0].strip() for r in ptc_state["result_data"]]

    # sandbox_pathsにない媒体は再DL
    for media_name in targets:
        if media_name not in ptc_state.get("sandbox_paths", {}):
            target_data = [r for r in ptc_state["result_data"]
                           if r.get("メディア名称", "").split(";")[0].strip() == media_name]
            if target_data:
                sp, fl, su, _, _ = await step2_get_pdfs(target_data)
                ptc_state["sandbox_paths"].update(sp)

    # 詳細提案再生成（file_paths方式・PDF書類のまま読取り）
    proposals, _ = await step3_generate_proposals(
        ptc_state["result_data"],
        ptc_state["weight_values"],
        ptc_state["pdf_urls"],
        user_request_2,
        article_type=ptc_state.get("article_type", "記事タイアップ"),
        target_attribute=ptc_state.get("target_attribute", ""),
        sandbox_paths=ptc_state.get("sandbox_paths", {}),
    )

    print("\n" + "=" * 60)
    print("[深掘り] 詳細提案")
    print("=" * 60)
    for p in proposals:
        if not isinstance(p, dict):
            continue
        mname2 = p.get('media_name', 'Unknown')
        print(f"\n### {mname2}  {p.get('price_disp','—')}")
        print(p.get('sec_overview', '') or '')

    # 深掘り後もHTMLを再生成して共有
    try:
        turn = ptc_state.get("turn_count", 2)
        html_fname = f"output/media_report_turn{turn}.html"
        html2 = build_html_report(
            user_request_2, proposals,
            ptc_state.get("user_budget", 0),
            ptc_state.get("excluded_media", []),
            target_label=ptc_state.get("target_attribute") or "経営層比率",
            remaining_data=ptc_state.get("remaining_data", []),
            target_attribute=ptc_state.get("target_attribute", ""),
            score_params=ptc_state.get("score_params", {}),
            all_data=ptc_state.get("all_data", []))
        os.makedirs("output", exist_ok=True)
        with open(html_fname, "w", encoding="utf-8") as f:
            f.write(html2)
        _dlog("turn2_html_written", {"path": html_fname, "bytes": len(html2.encode("utf-8"))})
        # proposalsを更新（次のターン用）
        ptc_state["proposals"] = proposals
        await _finalize_output(html_fname)
    except Exception as e:
        import traceback
        _dlog("turn2_html_error", {"error": str(e), "traceback": traceback.format_exc()})


async def _handle_research(user_request_2: str, params: dict = None):
    """再検索: 前回設定に新条件を上書きしてSTEP1から再走"""
    global ptc_state

    print("\n[再検索] パラメータ更新")
    if params is None:
        new_params = await _parse_user_request(user_request_2)
    else:
        new_params = params

    # 前回設定をベースに新条件を上書きし、mainへ明示的に渡す。
    # 引き継ぎはnull(None)判定で行う: LLMは「未言及」をnullで返す（_parse_user_requestのルール）ため、
    # None=未言及なら前回値(ptc_state)を採用し、明示があればそれを使う（明示的な解除=0も尊重される）。
    # ※ .get(key, 前回値) 方式はキーが常に存在するため前回値が一切効かないデッドコードだった。
    def _inherit(key: str, fallback):
        """今回の発言で言及がなければ(None)前回値を引き継ぐ。明示があればそれを使う。"""
        v = new_params.get(key)
        return ptc_state.get(key, fallback) if v is None else v

    # 除外媒体は「置き換え」でなく「累積」: 前回除外した実施済み媒体が
    # 言及なしのターンで復活しないよう、前回分と和集合を取る（順序維持で重複排除）
    _prev_ex = ptc_state.get("score_params", {}).get("exclude_media", []) or []
    _new_ex  = new_params.get("exclude_media") or []
    _merged_ex = list(dict.fromkeys([*_prev_ex, *_new_ex]))

    # filters/weights/sort_byは意図的に引き継がない: 「安い順にして」はweights→sortへの置き換え意図が
    # 多く、前回weightsを機械的に継ぐとscore_mediaの「sort_by+weights共存→フィルタ変換」が誤発動する。
    # ここはenriched_request（前回要望文込み）からLLMに再構成させる。
    merged = {
        "filters":          new_params.get("filters", {}),
        "weights":          new_params.get("weights", {}),
        "sort_by":          new_params.get("sort_by"),
        "sort_order":       new_params.get("sort_order", "asc"),
        "user_budget":      _inherit("user_budget", 0),
        "desired_count":    _inherit("desired_count", 7),
        "article_type":     _inherit("article_type", "記事タイアップ"),
        "target_attribute": _inherit("target_attribute", ""),
        "exclude_media":    _merged_ex,
    }
    print(f"  新filters: {merged['filters']}")
    print(f"  新weights: {merged['weights']}")

    # STEP1から再走（mergedを渡すことで前回コンテキスト付きの再検索になる）
    await main(user_request_2, params=merged)

# ============================================================
# 自律テスト実行（直接実行された場合のみ）
# ============================================================
# 注: code_execute環境では asyncio.run() は使えないため、
# 手動で await main() を呼ぶ。
if __name__ == "__main__":
    pass # トップレベルでのawait実行を避けるためpassのみ
