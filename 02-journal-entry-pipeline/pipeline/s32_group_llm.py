# -*- coding: utf-8 -*-
"""s32 — Step 4: 伝票の括りをAIに見直させる。s30 の直後に走らせる。

使い方:
    import sys; sys.path.insert(0, "notes/システム")
    import s32_group_llm as s
    await s.main()

機械の括り案をそのまま見せて「前提と矛盾するところだけ直せ」と頼む。
ゼロから括らせると悪くなる。金額と明細は s30 と同じ機械ルールで作り直す。

入出力:
  入力  tmp/beta/voucher/vouchers.json（s30 の結果）＋ tmp/beta/extract/*.json
  出力  同じ vouchers.json を上書き（元は vouchers_機械.json に残す）
"""
import asyncio, glob, json, os

import agent_sdk as at
from agent_sdk import ToolCallError
# Noteのパスは cwd 相対では解決されないので、自分のあるディレクトリを sys.path に通す
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))

import config as C
import rules as R

SPEC = """綴じられた証憑の束（1ファイル）を、**奉行に入力する伝票の単位**に区切ってください。
仕訳データは見ません。綴りの並びと帳票の情報だけで判断します。

## 前提（実データで確認済み）
- 綴りは**伝票の順（債務No順）**に綴じられている。1伝票は連続するページのまとまり
- 1伝票 ＝ **1つの支払先への1回の支払**。典型は「請求書1枚（＋添付）」
- 同じ発行者の請求書が**複数枚続く場合、ほとんど（91件中87件）は1伝票**にまとめられている。
  別伝票になるのは、**それぞれに社内集計表が付いている**か、
  **明らかに種類の違う支払**（鉄道運賃の精算書と別件の請求）のときだけ
- 社内集計表（配分表・集計表）は、その伝票の請求額と同額。**請求書の前にも後ろにも綴じられる**。
  集計表の総額が「後続の複数の請求書の合計」に一致するなら、その請求書群で1伝票
- 添付（納品書・見積書・明細書・受領書・領収書）は、**同じ発行者の請求書に付く**。
  前後どちらに付くかは発行者で判断する。発行者が読めない添付は、
  金額がその請求書の明細に含まれる側に付ける
- 同じ請求書の写し（同じ発行者・同じ金額）が2枚綴じられていることがある → **同じ伝票**
- 総額が無く明細だけの帳票は、直前の請求書の続きページ

## 入力
- `帳票`: ページ順の帳票一覧
- `機械の括り案`: ルールで作った括り。**たいてい合っています。**
  上の前提と矛盾するところ**だけ**直してください。ゼロから括り直さないこと

## 出力
- **全帳票をどれか1つの伝票に入れる**（漏れも重複もなし）
- `金額` は伝票金額。**写しや添付は足さない**。社内集計表があればその総額
- 直したところは `理由` に「機械案では〜だったが、〜なので分けた／まとめた」と書く
"""

SCHEMA = {
    "type": "object",
    "properties": {
        "伝票": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "docs": {"type": "array", "items": {"type": "string"}},
                    "支払先": {"type": "string"},
                    "金額": {"type": "integer"},
                    "理由": {"type": "string"},
                },
                "required": ["docs"],
            },
        },
        "迷った箇所": {"type": "array", "items": {"type": "string"}},
    },
    "required": ["伝票"],
}


def _ctx(src, docs, lines_by_doc, groups):
    """1ファイル分の文脈。明細は「例」を3件だけ（全部渡すと長くなりすぎる）。"""
    return {
        "file_id": src,
        "総ページ": max([d.get("page_to") or 0 for d in docs] or [0]),
        "帳票": [{
            "doc_id": d["doc_id"],
            "page": "%s-%s" % (d.get("page_from"), d.get("page_to")),
            "帳票種別": d.get("doc_type"),
            "社内集計表": bool(d.get("is_internal_worksheet")),
            "添付": bool(d.get("is_supporting_doc")),
            "発行者": d.get("vendor_name"),
            "宛先": d.get("buyer_office"),
            "締日": d.get("closing_date"),
            "総額": d.get("total_incl_tax"),
            "前月繰越": d.get("prev_balance"),
            "明細件数": len(lines_by_doc.get(d["doc_id"], [])),
            "明細の例": [str(l.get("item_name") or "")[:16]
                       for l in lines_by_doc.get(d["doc_id"], [])[:3]],
            "手書きコード": d.get("handwritten_account_codes") or [],
        } for d in docs],
        "機械の括り案": groups,
    }


async def _one(src, ctx, sem):
    try:
        async with sem:
            try:
                r = await at.llm_call(
                    SPEC + "\n\n## ctx\n" + json.dumps(ctx, ensure_ascii=False),
                    schema=SCHEMA, model=C.MODEL_SPLIT)
            except ToolCallError as e:
                return src, None, "bridge: %s" % e
        if "error" in r:
            return src, None, r["error"]
        data = r.get("data")
        if isinstance(data, dict) and "error" in data:
            return src, None, data["error"]
        if not isinstance(data, dict) or not data.get("伝票"):
            return src, None, "伝票 が返ってこなかった"

        # 漏れ・重複が無いか機械で検算する。壊れていたら機械案のまま使う
        want = {d["doc_id"] for d in ctx["帳票"]}
        got, dup = set(), []
        for g in data["伝票"]:
            for x in g.get("docs") or []:
                if x in got:
                    dup.append(x)
                got.add(x)
        if dup or got != want:
            return src, None, "帳票の漏れ%d件 / 重複%d件（機械案を使います）" % (
                len(want - got), len(dup))
        return src, data, None
    except Exception as e:                                # noqa: BLE001
        return src, None, "%s: %s" % (type(e).__name__, e)


async def main():
    vp = os.path.join(C.D_VOUCH, "vouchers.json")
    if not os.path.exists(vp):
        print("伝票がありません。先に s30_vouchers.main() を実行してください")
        return
    with open(vp, encoding="utf-8") as f:
        vouchers = json.load(f)

    # 機械の結果を残しておく（あとで比べられるように）
    mech = os.path.join(C.D_VOUCH, "vouchers_機械.json")
    if not os.path.exists(mech):
        with open(mech, "w", encoding="utf-8") as f:
            json.dump(vouchers, f, ensure_ascii=False)

    # ファイル単位に組み直す
    docs_by_src, lines_by_doc, D = {}, {}, {}
    for p in sorted(glob.glob(os.path.join(C.D_EXTRACT, "*.json"))):
        with open(p, encoding="utf-8") as f:
            m = json.load(f)
        src = m.get("source_pdf") or os.path.splitext(os.path.basename(p))[0]
        docs_by_src[src] = sorted(m["documents"], key=lambda d: (d.get("page_from") or 0, d["doc_id"]))
        for d in m["documents"]:
            D[d["doc_id"]] = d
        for l in m["line_items"]:
            lines_by_doc.setdefault(l.get("doc_id"), []).append(l)

    groups_by_src = {}
    for v in vouchers:
        groups_by_src.setdefault(v["元ファイル"], []).append(
            {"docs": v["帳票"], "支払先": v["発行者"], "金額": v["金額"]})

    sem = asyncio.Semaphore(C.CONCURRENCY)
    tasks = [_one(src, _ctx(src, docs, lines_by_doc, groups_by_src.get(src, [])), sem)
             for src, docs in docs_by_src.items()]
    results = await asyncio.gather(*tasks)

    # AIが直した括りで伝票を作り直す（金額と明細の決め方は s30 と同じ機械ルール）
    import s30_vouchers as s30
    new, ng, changed = [], [], 0
    for src, data, err in results:
        if err:
            ng.append((src, err))
            new += [v for v in vouchers if v["元ファイル"] == src]
            continue
        before = {frozenset(g["docs"]) for g in groups_by_src.get(src, [])}
        after = {frozenset(g["docs"]) for g in data["伝票"]}
        changed += len(after - before)
        by_doc = lines_by_doc
        for k, g in enumerate(data["伝票"], 1):
            ids = [x for x in g["docs"] if x in D]
            bodies = [D[x] for x in ids if R.is_body(D[x])]
            if not bodies:
                cands = [D[x] for x in ids if (D[x].get("total_incl_tax") or 0) > 0]
                bodies = [max(cands, key=lambda x: x.get("total_incl_tax") or 0)] if cands else []

            # 金額に数える本体を選ぶ。**s30 と同じ関数を使う**（別に書くと食い違う）
            bodies, copies, flags = R.fold_bodies(bodies)

            amount = 0
            for b in bodies:
                a, fl = R.doc_amount(b)
                amount += a
                flags += fl
            # 金額は**機械の合算を正**とする。AIが返した金額は照合にだけ使う
            try:
                ai_amt = int(g.get("金額") or 0)
            except (TypeError, ValueError):
                ai_amt = 0
            if ai_amt and amount and abs(ai_amt - amount) > 3:
                # **どの帳票を足してその金額になったかも書く。**
                # 「AIの金額の2倍」になる伝票が実データにあり、写しを畳めていない
                # 疑いがあるが、内訳が出ないと原因が絞れなかった
                naiyaku = " ".join(
                    "%s(%s:%s)" % (b.get("doc_id"), (b.get("doc_type") or "?")[:6],
                                   format(int(b.get("total_incl_tax") or 0), ","))
                    for b in bodies)
                flags.append("要確認:AIの伝票金額%d と 帳票の総額合計%d が違う（内訳 %s）"
                             % (ai_amt, amount, naiyaku))
            if not amount and ai_amt:
                amount = ai_amt                    # 帳票から総額が取れなかったときだけ採る
            # 明細の集め方も **s30 と同じ関数を使う**（手書きコードの引き継ぎを含む）。
            # 写しは外す。請求内訳書は金額には数えないが明細の出どころとしては使う
            dropped = {x for x in copies
                       if (D.get(x, {}).get("doc_type") or "") not in R.BREAKDOWN_TYPES}
            body_ids = {b["doc_id"] for b in bodies}
            only_body = R.collect_lines([x for x in ids if x in body_ids], D, by_doc)
            wide = R.collect_lines(
                [x for x in ids if x not in dropped and not D[x].get("is_internal_worksheet")],
                D, by_doc)

            def total_of(ls):
                return sum(int(l.get("amount_incl_tax") or 0) for l in ls)

            # 社内集計表が伝票金額に合うなら、それが仕訳の行そのもの。最優先で採る
            lines = R.worksheet_lines(ids, D, by_doc, amount)
            if lines:
                pass
            elif amount and only_body and abs(total_of(only_body) - amount) <= 3:
                lines = only_body                    # 本体だけで合う（添付は写し）
            elif amount and wide and abs(total_of(wide) - amount) <= 3:
                lines = wide                         # 鑑に明細が無く、内訳側に載っている
            else:
                lines = only_body or wide
            R.fill_voucher_codes(lines, ids, D)
            dsum = total_of(lines)
            if amount and dsum and abs(dsum - amount) > 3:
                flags.append("要確認:明細合計%d と 伝票金額%d が違う" % (dsum, amount))
            new.append({
                "伝票キー": "%s#%02d" % (src, k), "元ファイル": src,
                "発行者": (bodies[0].get("vendor_name") or "") if bodies else (g.get("支払先") or ""),
                "締日": (bodies[0].get("closing_date") or bodies[0].get("issue_date") or "") if bodies else "",
                "金額": int(amount), "帳票": ids,
                "ページ": [int(D[x].get("page_from") or 0) for x in ids],
                "明細": lines, "配分表": R.worksheet_hint(ids, D, by_doc),
                "内訳帳票": R.breakdown_hint(ids, D, amount),
                "フラグ": sorted(set(flags)),
                "括りの理由": g.get("理由"),
            })

    new = s30._merge_across_files(new)
    with open(vp, "w", encoding="utf-8") as f:
        json.dump(new, f, ensure_ascii=False)

    print("--- ファイル %d / 伝票 %d本（機械案は %d本）---" % (
        len(docs_by_src), len(new), len(vouchers)))
    print("AIが機械案から変えた括り: %d 件" % changed)
    print("要確認フラグつき: %d 本" % sum(1 for v in new if v["フラグ"]))
    for src, why in ng[:5]:
        print("NG %s: %s（このファイルは機械案のまま）" % (src, why))
    print("生成: %s（機械案は %s に残しました）" % (vp, mech))
