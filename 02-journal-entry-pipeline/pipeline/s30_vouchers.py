# -*- coding: utf-8 -*-
"""s30 — Step 3: 綴りの帳票を「伝票」に括り、伝票の金額を確定する。AIは呼ばない。

使い方:
    import sys; sys.path.insert(0, "notes/システム")
    import s30_vouchers as s
    await s.main()

前提: 綴りは債務No順に綴じられている。
ルール:
  1. 「本体」（総額があり、社内集計表でも参考添付でもない帳票）が出たら新しい伝票の候補
  2. 直前の本体と同じ発行者・同額なら写し → 同じ伝票に畳む
  3. 添付（納品書・見積書・受領書・明細書）と社内集計表は、直近の本体の伝票に付ける
  4. 社内集計表が本体より前にあり、その合計が後続の本体（連続する数枚）の合計と一致するなら、
     それらをひとつの伝票にまとめる（鉄道運賃のように複数請求書で1伝票のケース）
  5. 鑑＋内訳（本体の総額 ＝ 直後に続く本体の合計）はひとつの伝票にまとめる
  6. 総額が無く明細だけの帳票は、直前の本体の続きとして付ける（ページ割れ対策）

出力: tmp/beta/voucher/vouchers.json
"""
import glob, json, os

# Noteのパスは cwd 相対では解決されないので、自分のあるディレクトリを sys.path に通す
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))

import config as C
import rules as R
import jsonlog as JL


def _lookahead(docs, i, t):
    """docs[i] の総額 t に、後続の本体の合計がぴったり一致するなら、その最後の添字を返す。"""
    acc, nb = 0, 0
    for j in range(i + 1, min(i + 14, len(docs))):
        x = docs[j]
        if R.is_body(x):
            acc += (x.get("total_incl_tax") or 0)
            nb += 1
            if abs(acc - t) <= 3 and nb >= 2:
                return j
            if acc > t + 3:
                return None
        elif x.get("is_internal_worksheet"):
            return None
    return None


def group_file(docs):
    """1ファイル分の帳票を伝票に括る。戻り: [[doc_id, ...], ...]"""
    docs = sorted(docs, key=lambda d: (d.get("page_from") or 0, d.get("doc_id") or ""))
    groups, cur, prev = [], None, None
    pending, pending_vendor = [], None
    i = 0
    while i < len(docs):
        d = docs[i]

        # 4. 先行する社内集計表
        if d.get("is_internal_worksheet") and (d.get("total_incl_tax") or 0) > 0:
            t = d["total_incl_tax"]
            acc, span = 0, []
            for j in range(i + 1, min(i + 8, len(docs))):
                x = docs[j]
                if R.is_body(x):
                    acc += (x.get("total_incl_tax") or 0)
                    span.append(j)
                    if abs(acc - t) <= 3:
                        break
                    if acc > t:
                        span = []
                        break
                elif x.get("is_internal_worksheet"):
                    break
            if span and abs(acc - t) <= 3:
                g = {"docs": [d["doc_id"]], "vendor": None, "total": t, "bodies": [], "has_ws": True}
                for j in range(i + 1, span[-1] + 1):
                    g["docs"].append(docs[j]["doc_id"])
                    if R.is_body(docs[j]):
                        g["bodies"].append(docs[j])
                        g["vendor"] = g["vendor"] or R.norm_payee(docs[j].get("vendor_name"))
                groups.append(g)
                cur = prev = g
                i = span[-1] + 1
                continue

        if R.is_body(d):
            v = R.norm_payee(d.get("vendor_name"))
            t = d.get("total_incl_tax") or 0
            same_vendor = bool(cur and cur["bodies"] and cur["vendor"] == v)
            span = _lookahead(docs, i, t)

            if same_vendor and any(abs((b.get("total_incl_tax") or 0) - t) <= 3 for b in cur["bodies"]):
                # 2. 写し。紙としては伝票に付けるが、金額は足さない（bodies に入れない）
                cur["docs"].append(d["doc_id"])
                if (d.get("doc_type") or "") in R.BREAKDOWN_TYPES:
                    # 請求内訳書は写しではなく請求書の内訳。写し扱いにすると明細ごと
                    # 捨てるので、鑑が総額だけのとき伝票の明細が空になる
                    cur["breakdown"] = cur.get("breakdown", []) + [d["doc_id"]]
                else:
                    cur["copies"] = cur.get("copies", []) + [d["doc_id"]]
            elif (same_vendor and not cur.get("has_ws")
                  and all(R.same_closing_month(x, d) for x in cur["bodies"])):
                # 同じ発行者の請求書が続く。**ただし締月が同じときだけ。**
                # 月が違えば経理は別伝票にする（実データ: ハシモトの5月分と6月分）
                cur["docs"].append(d["doc_id"])
                cur["bodies"].append(d)
            elif span:
                # 5. 鑑＋内訳。金額は鑑の総額だけ。内訳を bodies に入れると二重になる
                cur = {"docs": [d["doc_id"]], "vendor": v, "total": t, "bodies": [d], "has_ws": False}
                for j in range(i + 1, span + 1):
                    cur["docs"].append(docs[j]["doc_id"])
                    if docs[j].get("is_internal_worksheet"):
                        cur["has_ws"] = True
                groups.append(cur)
                prev, pending, pending_vendor = cur, [], None
                i = span + 1
                continue
            else:
                cur = {"docs": [d["doc_id"]], "vendor": v, "total": t, "bodies": [d], "has_ws": False}
                if pending and pending_vendor in (None, "", v):
                    cur["docs"] = list(pending) + cur["docs"]
                    if prev is not None:
                        for x in pending:
                            if x in prev["docs"]:
                                prev["docs"].remove(x)
                pending, pending_vendor = [], None
                groups.append(cur)
            prev = cur
        else:
            # 3./6. 添付・集計表・総額なし
            if cur is None:
                cur = {"docs": [d["doc_id"]], "vendor": R.norm_payee(d.get("vendor_name")),
                       "total": d.get("total_incl_tax") or 0, "bodies": [], "has_ws": False}
                groups.append(cur)
                prev = cur
            else:
                cur["docs"].append(d["doc_id"])
                if d.get("is_internal_worksheet"):
                    cur["has_ws"] = True
                elif not (d.get("total_incl_tax") or 0) or d.get("is_supporting_doc"):
                    pending.append(d["doc_id"])
                    pending_vendor = pending_vendor or R.norm_payee(d.get("vendor_name")) or None
        i += 1
    return groups


def _merge_across_files(vouchers):
    """別のPDFに同じ請求書（発行者・締月・金額が同じ）が綴じられていたら1伝票に畳む。"""
    seen, out = {}, []
    for v in vouchers:
        key = R.dup_key(v.get("発行者"), v.get("締日"), v.get("金額"))
        if R.norm_payee(v.get("発行者")) and int(v.get("金額") or 0) and key in seen:
            base = seen[key]
            base["帳票"] += v["帳票"]
            base["ページ"] += v["ページ"]
            # 明細は足さない（同じ請求書の写しなので、足すと二重に行が立つ）
            base["フラグ"] = sorted(set(base["フラグ"]) | set(v["フラグ"]) | {"別ファイルの写しを畳んだ"})
            base["畳んだ伝票"] = base.get("畳んだ伝票", []) + [v["伝票キー"]]
            continue
        seen[key] = v
        out.append(v)
    return out


async def main():
    os.makedirs(C.D_VOUCH, exist_ok=True)
    # 「再送」は後ろに回す。_merge_across_files は先に来たほうを残すので、本編を残す
    files = R.resend_last(glob.glob(os.path.join(C.D_EXTRACT, "*.json")))
    if not files:
        print("抽出結果がありません。先に s20_extract.main() を実行してください")
        return

    vouchers, ndoc = [], 0
    for f in files:
        with open(f, encoding="utf-8") as fh:
            m = json.load(fh)
        src = m.get("source_pdf") or os.path.splitext(os.path.basename(f))[0]
        docs = m["documents"]
        ndoc += len(docs)
        by_doc = {}
        for l in m["line_items"]:
            by_doc.setdefault(l.get("doc_id"), []).append(l)
        D = {d["doc_id"]: d for d in docs}

        for k, g in enumerate(group_file(docs), 1):
            bodies = g["bodies"]
            if not bodies:
                # 本体が1枚も無い群（添付や集計表だけ）。合算すると二重になるので最大額の1枚を使う
                cands = [D[x] for x in g["docs"] if x in D and (D[x].get("total_incl_tax") or 0) > 0]
                bodies = [max(cands, key=lambda x: x.get("total_incl_tax") or 0)] if cands else []
            # 金額に数える本体を選ぶ。**s32 と同じ関数を使う**（別に書くと食い違う）
            bodies, folded, flags = R.fold_bodies(bodies)
            amount = 0
            for b in bodies:
                a, fl = R.doc_amount(b)
                amount += a
                flags += fl
            if g.get("copies"):
                flags.append("同じ請求書の写しを%d枚畳んだ" % len(g["copies"]))
            # 明細は「本体だけ」と「写しを除く全部」の2通り作り、伝票金額に合うほうを採る。
            # 同じ明細が請求書と納品書の両方に載っていて、全部足すと二重になるため。
            # 写しは明細からも外す（外すと合計が2倍にならない）。
            # **請求内訳書は外さない。** 金額には数えないが明細の出どころとしては使う
            dropped = set(g.get("copies") or []) | {
                x for x in folded
                if (D.get(x, {}).get("doc_type") or "") not in R.BREAKDOWN_TYPES}
            body_ids = {b["doc_id"] for b in bodies}

            only_body = R.collect_lines([x for x in g["docs"] if x in body_ids], D, by_doc)
            wide = R.collect_lines(
                [x for x in g["docs"]
                 if x not in dropped and x in D and not D[x].get("is_internal_worksheet")],
                D, by_doc)

            def total_of(ls):
                return sum(int(l.get("amount_incl_tax") or 0) for l in ls)

            # 社内集計表が伝票金額に合うなら、それが仕訳の行そのもの。最優先で採る
            lines = R.worksheet_lines(g["docs"], D, by_doc, amount)
            if lines:
                pass
            elif amount and only_body and abs(total_of(only_body) - amount) <= 3:
                lines = only_body                    # 本体だけで合う（添付は写し）
            elif amount and wide and abs(total_of(wide) - amount) <= 3:
                lines = wide                         # 鑑に明細が無く、内訳側に載っている
            else:
                lines = only_body or wide            # どちらも合わない → 下でフラグを立てる
            R.fill_voucher_codes(lines, g["docs"], D)
            dsum = total_of(lines)
            if amount and dsum and abs(dsum - amount) > 3:
                flags.append("要確認:明細合計%d と 伝票金額%d が違う" % (dsum, amount))
            vouchers.append({
                "伝票キー": "%s#%02d" % (src, k),
                "元ファイル": src,
                "発行者": (bodies[0].get("vendor_name") or "") if bodies else "",
                "締日": (bodies[0].get("closing_date") or bodies[0].get("issue_date") or "") if bodies else "",
                "金額": int(amount),
                "帳票": g["docs"],
                "ページ": [int(D[x].get("page_from") or 0) for x in g["docs"] if x in D],
                "明細": lines,
                "配分表": R.worksheet_hint(g["docs"], D, by_doc),
                "内訳帳票": R.breakdown_hint(g["docs"], D, amount),
                "フラグ": sorted(set(flags)),
            })
            v = vouchers[-1]
            JL.log("s30 伝票に括る", 伝票=v["伝票キー"], 発行者=v["発行者"], 金額=v["金額"],
                   帳票数=len(g["docs"]), 本体数=len(bodies), 明細数=len(lines),
                   配分表=len(v["配分表"]), 内訳帳票=len(v["内訳帳票"]),
                   明細合計=dsum, フラグ=v["フラグ"])

    vouchers = _merge_across_files(vouchers)

    with open(os.path.join(C.D_VOUCH, "vouchers.json"), "w", encoding="utf-8") as f:
        json.dump(vouchers, f, ensure_ascii=False)

    nflag = sum(1 for v in vouchers if v["フラグ"])
    print("--- ファイル %d / 帳票 %d → 伝票 %d 本 ---" % (len(files), ndoc, len(vouchers)))
    print("金額0の伝票: %d 本" % sum(1 for v in vouchers if not v["金額"]))
    print("要確認フラグつき: %d 本" % nflag)
    for v in [x for x in vouchers if x["フラグ"]][:5]:
        print("  %s %s: %s" % (v["伝票キー"], v["発行者"], " / ".join(v["フラグ"])))
    print("生成: %s/vouchers.json" % C.D_VOUCH)
