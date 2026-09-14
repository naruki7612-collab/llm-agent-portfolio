# -*- coding: utf-8 -*-
"""s35 — Step 5: 伝票ごとに「起こすべき仕訳行」を決める。AI。

使い方:
    import sys; sys.path.insert(0, "notes/システム")
    import s35_split as s
    await s.main()
    await s.main(limit=3)   # 先頭3本の伝票だけ

・支払先を別名テーブルで引き当て、その支払先の「過去の行の割り方」を見せる
・**その伝票のページ画像も渡す**。明細のテキストだけでは車検一式のように
  「1枚の請求書を科目ごとに何行に割るか」が判断できない
・結果は tmp/beta/split/{伝票キー}.json にキャッシュ
"""
import asyncio, hashlib, json, os, re, unicodedata

import agent_sdk as at
from agent_sdk import ToolCallError
# Noteのパスは cwd 相対では解決されないので、自分のあるディレクトリを sys.path に通す
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))

import config as C
import prompts as P
import rules as R
import jsonlog as JL
import store as S

_IMGDIR = {}      # 元ファイル名 → ページ画像のディレクトリ


def _load_img_dirs():
    """s20 が残した「どのPDFの画像がどこにあるか」を読む。"""
    import glob
    if _IMGDIR:
        return _IMGDIR
    for p in glob.glob(os.path.join(C.D_EXTRACT, "*.json")):
        try:
            with open(p, encoding="utf-8") as f:
                m = json.load(f)
        except Exception:
            continue
        if m.get("img_dir"):
            _IMGDIR[m.get("source_pdf") or ""] = m["img_dir"]
    return _IMGDIR


def _images_of(v):
    """その伝票のページ画像を集める。多すぎるときは先頭から SPLIT_MAX_IMG 枚。"""
    d = _load_img_dirs().get(v.get("元ファイル"))
    if not d or not os.path.isdir(d):
        return []
    pages = sorted({int(x) for x in (v.get("ページ") or []) if x})
    out = []
    for p in pages:
        f = os.path.join(d, "p-%03d.jpg" % p)
        if os.path.exists(f):
            out.append(f)
    return out[:C.SPLIT_MAX_IMG]


def _fname(key):
    """伝票キーからキャッシュ名を作る。日本語を含むのでハッシュを付ける。"""
    safe = re.sub(r"[^0-9A-Za-z]", "_", unicodedata.normalize("NFKC", key))[:40]
    return "%s_%s" % (safe, hashlib.md5(key.encode("utf-8")).hexdigest()[:8])


def _ym_of(voucher):
    """伝票の年月（YYYYMM）。取れなければ空（そのときは過去実績を渡さない）。"""
    t = str(voucher.get("締日") or "")
    m = re.search(r"(20\d{2})\D?(\d{1,2})", t)
    if not m:
        return ""
    return "%04d%02d" % (int(m.group(1)), int(m.group(2)))


async def _one(v, sem):
    try:
        return await _one_inner(v, sem)
    except Exception as e:                                # noqa: BLE001 — 1件で全体を落とさない
        return v["伝票キー"], None, "%s: %s" % (type(e).__name__, e)


async def _one_inner(v, sem):
    key = v["伝票キー"]
    out = os.path.join(C.D_SPLIT, _fname(key) + ".json")
    if os.path.exists(out):
        with open(out, encoding="utf-8") as f:
            return key, json.load(f), None

    ym = _ym_of(v)
    ctx = {"支払先": v.get("発行者"), "支払先コード": None, "対象年月": ym,
           "voucher_amount": v.get("金額"),
           "source_lines": [{
               "line_id": l.get("line_id"), "品名": l.get("item_name"),
               "金額": l.get("amount_incl_tax"), "車番": l.get("vehicle_no"),
               "手書きコード": l.get("handwritten_account_code"),
               "この帳票の手書きコード": l.get("帳票の手書きコード") or None,
               # 車番は帳票の見出しにあることが多い。行に無ければ帳票のを渡す
               "この帳票の車番": l.get("帳票の車番") or None,
               "用途区分": l.get("purpose_category"),
               # 事業所と経費区分は「どこの部門・どのセグメントか」そのもの。
               # 抽出はしていたのに Step 5 に渡していなかった
               "事業所": l.get("office") or None,
               "経費区分": l.get("cost_segment") or None,
           } for l in v.get("明細", [])]}
    # 社内配分表があれば渡す。事業所・区分ごとの小計は仕訳行の割り方そのもの
    if v.get("配分表"):
        ctx["社内配分表"] = v["配分表"]
    # 鑑1枚＋内訳N枚の伝票。内訳の1枚1枚が仕訳行の単位（1台1枚の整備請求書など）
    if v.get("内訳帳票"):
        ctx["内訳帳票"] = v["内訳帳票"]
    async with sem:
        payee = await S.payee_code_of(v.get("発行者"))
        ctx["支払先コード"] = payee
        ctx["過去の伝票数"] = 0
        if payee and ym:
            pc = await S.payee_context(payee, ym)
            ctx["直近5伝票の行テンプレート"] = pc["直近5伝票の行テンプレート"]
            ctx["行数分布"] = pc["行数分布"]
            ctx["過去の伝票数"] = len(pc["直近5伝票の行テンプレート"])
        elif payee:
            ctx["注意"] = "締日が読めないので過去実績は渡していません（未来の実績を混ぜないため）"
        imgs = _images_of(v)
        note = ("\n\n## 渡した画像\nこの伝票のページ画像を%d枚渡しています。"
                "明細のテキストと画像の両方を見て判断してください。"
                "手書きの書き込み・赤チェック・欄外のメモも手がかりです。\n" % len(imgs)) if imgs else ""
        try:
            r = await at.llm_call(
                P.SPLIT_PROMPT + note + "\n\n## ctx\n" + json.dumps(ctx, ensure_ascii=False),
                file_paths=imgs or None, schema=P.SPLIT_SCHEMA, model=C.MODEL_SPLIT)
        except ToolCallError as e:
            return key, None, "bridge: %s" % e
    if "error" in r:
        return key, None, r["error"]

    data = r.get("data")
    if isinstance(data, dict) and "error" in data:
        return key, None, data["error"]
    if not isinstance(data, dict) or "lines" not in data:
        return key, None, "lines が返ってこなかった"
    # 車番を機械で埋める。AIが label に書き忘れるとセグメントが当てられなくなる
    R.fill_line_vehicle(data["lines"], {str(l.get("line_id")): l for l in v.get("明細", [])})
    nfill = sum(1 for l in data["lines"] if l.get("vehicle_no"))
    JL.log("s35 仕訳行に割る", 伝票=key, 支払先=ctx.get("支払先"),
           入力明細=len(ctx["source_lines"]), 配分表=len(ctx.get("社内配分表") or []),
           内訳帳票=len(ctx.get("内訳帳票") or []), 画像=len(imgs),
           過去伝票=ctx.get("過去の伝票数"), 出力行=len(data["lines"]),
           車番あり=nfill, 合計=sum(int(l.get("amount") or 0) for l in data["lines"]),
           伝票金額=v.get("金額"), sum_matches=data.get("sum_matches"),
           理由=(data.get("reasoning") or "")[:120])
    data["伝票キー"] = key
    data["支払先コード"] = payee
    data["対象年月"] = ym
    with open(out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    return key, data, None


async def main(limit=None):
    os.makedirs(C.D_SPLIT, exist_ok=True)
    path = os.path.join(C.D_VOUCH, "vouchers.json")
    if not os.path.exists(path):
        print("伝票がありません。先に s30_vouchers.main() を実行してください")
        return
    with open(path, encoding="utf-8") as f:
        vouchers = json.load(f)
    if limit:
        vouchers = vouchers[:limit]

    sem = asyncio.Semaphore(C.CONCURRENCY)
    results = await asyncio.gather(*[_one(v, sem) for v in vouchers])

    V = {v["伝票キー"]: v for v in vouchers}
    ok, ng, nline, mismatch, nopayee = 0, [], 0, [], 0
    for key, data, err in results:
        if err:
            ng.append((key, err))
            continue
        ok += 1
        nline += len(data.get("lines", []))
        if not data.get("支払先コード"):
            nopayee += 1
        total = int(V[key].get("金額") or 0)
        s = sum(int(l.get("amount") or 0) for l in data.get("lines", []))
        if total and abs(s - total) > 3:
            mismatch.append((key, s, total))

    nimg = sum(1 for v in vouchers if _images_of(v))
    print("--- 伝票 %d / 成功 %d / 失敗 %d / 仕訳行 %d ---" % (len(vouchers), ok, len(ng), nline))
    print("ページ画像を渡せた伝票: %d / %d" % (nimg, len(vouchers)))
    print("支払先が引けなかった伝票: %d 本（初出の可能性）" % nopayee)
    print("行合計と伝票金額が違う: %d 本" % len(mismatch))
    for key, s, t in mismatch[:5]:
        print("  %s: 行合計%d ≠ 伝票%d" % (key, s, t))
    for key, why in ng[:5]:
        print("NG %s: %s" % (key, why))
