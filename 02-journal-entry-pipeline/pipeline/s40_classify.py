# -*- coding: utf-8 -*-
"""s40 — Step 6: 各行の4項目（部門・科目・補助・セグメント）と摘要を決める。

機械がやること（AIにさせない）:
  * 辞書（車両マスタ・手書きコード規則・全社品目・管理事業対応・支払先実績）を渡す
  * 請求書の「経費区分」のうち、当たらない「管理」「事業」を落とす
  * **AIの答えを機械の見立てと突き合わせる**（手書きコード・事業所・経費区分）
  * セグ449なら45x系という切替を機械でもう一度当てる

AIがやること: どの手がかりを採るかの判断（事業所「富山」でも検修なら0007、など）

使い方:
    import sys; sys.path.insert(0, "notes/システム")
    import s40_classify as s
    await s.main()
    await s.main(limit=3)

結果は tmp/beta/class/{伝票キー}.json にキャッシュ。
"""
import asyncio, glob, hashlib, json, os, re, unicodedata

import agent_sdk as at
from agent_sdk import ToolCallError
# Noteのパスは cwd 相対では解決されないので、自分のあるディレクトリを sys.path に通す
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))

import config as C
import prompts as P
import rules as R
import store as S
import jsonlog as JL


def _fname(key):
    safe = re.sub(r"[^0-9A-Za-z]", "_", unicodedata.normalize("NFKC", key))[:40]
    return "%s_%s" % (safe, hashlib.md5(key.encode("utf-8")).hexdigest()[:8])


def _invoice_note(src_lines, line):
    """その行の元になった明細の「請求書の記載」を集める。

    経費区分は的中率の高い 通運/営業/検修 だけ渡す（管理・事業は渡すと悪くなる）。
    """
    ids = [str(x) for x in (line.get("source_line_ids") or [])]
    srcs = [src_lines[i] for i in ids if i in src_lines]
    note = {}

    def pick(key):
        vals = sorted({str(s.get(key) or "").strip() for s in srcs if str(s.get(key) or "").strip()})
        return vals[0] if vals else None

    # 車番も拾う。Step 5 が vehicle_no を落としても、元の明細に残っていれば
    # 車両マスタが引ける。セグメントは車で決まるのでここが効く
    for label, key in (("事業所", "office"), ("用途区分", "purpose_category"),
                       ("実際の供給元", "actual_supplier_name"),
                       ("車番", "vehicle_no"),
                       ("明細の手書きコード", "handwritten_account_code")):
        v = pick(key)
        if v:
            note[label] = v
    seg = R.filter_cost_segment(pick("cost_segment") or "")
    if seg:
        note["経費区分"] = seg

    # 手書きコードは帳票ごと1つだけのことが多い。行に付いていなくても帳票にあれば渡す
    # 車番も帳票側から拾う（1台1枚の請求書では見出しにしかない）
    if not note.get("車番"):
        dv = sorted({str(s_.get("帳票の車番") or "").strip() for s_ in srcs} - {""})
        if len(dv) == 1:
            note["この帳票の車番"] = dv[0]
    doc_codes = sorted({c for s_ in srcs for c in (s_.get("帳票の手書きコード") or []) if c})
    if doc_codes and not note.get("明細の手書きコード"):
        if len(doc_codes) == 1:
            note["この請求書の手書きコード"] = doc_codes[0]
        else:
            note["この請求書の手書きコード（複数。どれがこの行か判断すること）"] = doc_codes
    return note or None


async def _one(v, split, dicts, sem):
    try:
        return await _one_inner(v, split, dicts, sem)
    except Exception as e:                                # noqa: BLE001 — 1件で全体を落とさない
        return v["伝票キー"], None, "%s: %s" % (type(e).__name__, e)


async def _one_inner(v, split, dicts, sem):
    key = v["伝票キー"]
    out = os.path.join(C.D_CLASS, _fname(key) + ".json")
    if os.path.exists(out):
        with open(out, encoding="utf-8") as f:
            return key, json.load(f), None

    src_lines = {str(l.get("line_id")): l for l in v.get("明細", [])}
    payee = split.get("支払先コード")
    ym = split.get("対象年月") or ""

    lines = []
    for i, l in enumerate(split.get("lines", []), 1):
        lines.append({
            "no": i, "label": l.get("label"), "item_label": l.get("item_label"),
            "vehicle_no": l.get("vehicle_no"), "amount": l.get("amount"),
            "請求書の記載": _invoice_note(src_lines, l),
            "参照元明細ID": [str(x) for x in (l.get("source_line_ids") or [])],
        })

    veh_ctx, veh_hist = {}, {}
    facts = await S.load_journal_facts(ym or None)
    for l in lines:
        vn = R.norm_vehicle(l.get("vehicle_no") or "")
        if not vn:
            vn = R.vehicle_in_text(l.get("label"))       # label に「金沢4821」と書かれていることがある
        if not vn:
            continue
        if vn in dicts["vehicle"]:
            veh_ctx[vn] = dicts["vehicle"][vn]
        if vn in facts["車番の実績"]:
            veh_hist[vn] = facts["車番の実績"][vn][:3]

    ctx = {"伝票": key, "支払先": v.get("発行者"), "支払先コード": payee, "対象年月": ym,
           "lines": lines,
           "車両マスタ": veh_ctx,
           # 車両マスタに無い車番でも、過去の摘要から部門/セグを引ける
           # （元の L3_SPEC_v5e「マスタに無くても過去実績で必ず埋めにいく」）
           "車番の実績": veh_hist,
           # 配分表の小計行の名前には事業所と経費区分が書いてある
           # （例「配賦集計 営業/備消品費/福井」）。部門とセグメントの直接の根拠になる
           "社内配分表": v.get("配分表") or None,
           "手書きコードの変換規則": dicts["code_rule"],
           "管理と事業の科目対応": dicts["acct_mgmt"],
           "全社の品目と科目の対応": dicts["global_item"][:120],
           "注意": "セグメント449（管理）に付ける行は45x系の科目、それ以外は42x系。"
                   "同じ品目でも切り替わる（対応表参照）"}

    async with sem:
        ctx["この支払先の過去の仕訳実績"] = []
        ctx["この支払先の請求書品名と4項目の対応"] = []
        hist = []
        if payee and ym:
            pc = await S.payee_context(payee, ym)
            hist = pc["過去の仕訳実績"]
            ctx["この支払先の過去の仕訳実績"] = hist
            ctx["この支払先の請求書品名と4項目の対応"] = pc["請求書品名と4項目の対応"]

        try:
            r = await at.llm_call(
                P.CLASSIFY_PROMPT + "\n\n## ctx\n" + json.dumps(ctx, ensure_ascii=False),
                schema=P.CLASSIFY_SCHEMA, model=C.MODEL_CLASS)
        except ToolCallError as e:
            return key, None, "bridge: %s" % e
        if "error" in r:
            return key, None, r["error"]
        data_in = r.get("data")
        if isinstance(data_in, dict) and "error" in data_in:
            return key, None, data_in["error"]
        if not isinstance(data_in, dict):
            return key, None, "data が返ってこなかった"
        got = data_in.get("lines") or []
        if len(got) != len(lines):
            return key, None, "行数が違う（入力%d / 出力%d）" % (len(lines), len(got))

        # --- S7 摘要（同じセマフォの中で。並列スロットを食いすぎないため）
        summary = {}
        try:
            rs = await at.llm_call(
                P.SUMMARY_PROMPT + "\n\n## 行\n" + json.dumps(
                    [{"no": l["no"], "label": l["label"], "車番": l.get("vehicle_no")} for l in lines],
                    ensure_ascii=False) +
                "\n\n## この支払先の過去の摘要\n" + json.dumps(
                    [h.get("摘要") for h in hist[:15]], ensure_ascii=False),
                schema=P.SUMMARY_SCHEMA, model=C.MODEL_SUMMARY)
            if "error" not in rs and isinstance(rs.get("data"), dict):
                for x in rs["data"].get("lines") or []:
                    try:
                        summary[int(x.get("no"))] = x.get("摘要")
                    except (TypeError, ValueError):
                        continue
        except ToolCallError:
            pass

    # --- 機械の検算 ---------------------------------------------------------
    fixed, flagged = 0, 0
    merged = []
    for l, g in zip(lines, got):
        g = dict(g)

        # 1) セグ449なら45x系、それ以外は42x系。ずれていたら対応表で入れ替える
        g2, changed = R.fix_mgmt_business(g, dicts["acct_mgmt"])
        if changed:
            g.update(g2)
            g["根拠"] = (g.get("根拠") or "") + " ／ 機械が管理・事業の科目を入れ替えた"
            fixed += 1

        # 2) 手書きコード・事業所・経費区分から機械で言えることと突き合わせる
        hint, why = R.machine_hint(l, dicts["code_rule"], dicts["vehicle"])
        diff = R.compare_hint(g, hint)
        if diff:
            flagged += 1
            g["機械との食い違い"] = diff
            g["機械の見立て"] = why
            # 食い違いは即誤りではないので、直さずに確信度だけ下げて人の確認に回す
            g["confidence"] = min(float(g.get("confidence") or 0), C.CONF_THRESHOLD - 0.05)

        merged.append({**l, **g, "摘要": summary.get(l["no"]) or l.get("label")})

    data = {"伝票キー": key, "支払先コード": payee, "対象年月": ym,
            "機械が直した行数": fixed, "機械と食い違った行数": flagged, "lines": merged}
    import collections as _c
    JL.log("s40 4項目を決める", 伝票=key, 支払先=v.get("発行者"), 行数=len(merged),
           機械が直した=fixed, 食い違い=flagged,
           根拠種別=dict(_c.Counter(str(x.get("根拠種別") or "") for x in merged)),
           確信度の最小=min([float(x.get("確信度") or 0) for x in merged] or [0]))
    with open(out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    return key, data, None


async def main(limit=None):
    os.makedirs(C.D_CLASS, exist_ok=True)
    vpath = os.path.join(C.D_VOUCH, "vouchers.json")
    if not os.path.exists(vpath):
        print("伝票がありません。先に s30_vouchers.main() を実行してください")
        return
    with open(vpath, encoding="utf-8") as f:
        vouchers = json.load(f)

    splits = {}
    if os.path.isdir(C.D_SPLIT):
        for p in sorted(glob.glob(os.path.join(C.D_SPLIT, "*.json"))):
            with open(p, encoding="utf-8") as f:
                d = json.load(f)
            splits[d["伝票キー"]] = d
    if not splits:
        print("行分割の結果がありません。先に s35_split.main() を実行してください")
        return

    targets = [v for v in vouchers if v["伝票キー"] in splits]
    if limit:
        targets = targets[:limit]

    dicts = await S.load_small_dicts()
    print("辞書: 車両%d / コード規則%d / 全社品目%d / 管理事業%d / 名称%d" % (
        len(dicts["vehicle"]), len(dicts["code_rule"]),
        len(dicts["global_item"]), len(dicts["acct_mgmt"]), len(dicts["master"])))

    sem = asyncio.Semaphore(C.CONCURRENCY)
    results = await asyncio.gather(*[_one(v, splits[v["伝票キー"]], dicts, sem) for v in targets])

    ok, ng, nline, low, fixed, flagged = 0, [], 0, 0, 0, 0
    for key, data, err in results:
        if err:
            ng.append((key, err))
            continue
        ok += 1
        fixed += data.get("機械が直した行数", 0)
        flagged += data.get("機械と食い違った行数", 0)
        for l in data["lines"]:
            nline += 1
            if float(l.get("confidence") or 0) < C.CONF_THRESHOLD:
                low += 1
    print("--- 伝票 %d / 成功 %d / 失敗 %d / 行 %d ---" % (len(targets), ok, len(ng), nline))
    print("確信度 %.1f 未満: %d 行 (%.0f%%)" % (
        C.CONF_THRESHOLD, low, 100 * low / nline if nline else 0))
    print("管理・事業の切替を機械が直した行: %d" % fixed)
    print("機械の見立てと食い違った行（要確認に回す）: %d" % flagged)
    for key, why in ng[:5]:
        print("NG %s: %s" % (key, why))
