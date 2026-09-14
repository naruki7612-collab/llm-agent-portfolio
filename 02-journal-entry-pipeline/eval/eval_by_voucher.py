# -*- coding: utf-8 -*-
"""予測仕訳CSVを伝票単位で正解と突き合わせる。

    python3 CSV評価スクリプト_伝票単位.py <予測仕訳CSV>

`評価スクリプト.py` はサンドボックスの tmp/beta/*.json が要る。こちらはCSVだけで
「会計上正しい」「行データ完全一致」「行単位の一致率」を出す。

対応づけ: 正解の伝票と予測の伝票を、同じPDF内でページの重なりが大きい順に
1対1で割り当てる。CSVの参照元ページは先頭6ページで切られているので、
Jaccard ではなく重なりの実数で順位づける。

母数: 予測CSVに出てきたPDFに紐づく正解の伝票だけ。入れていない証憑のぶんを
取りこぼしに数えないため。買掛金のPDFは両側から外す（未払金の仕組みなので）。
"""
import collections
import csv
import io
import json
import os
import pathlib
import re
import sys

HERE = pathlib.Path(__file__).resolve().parent
GEN = HERE.parent / "4_元データ"
F4 = ["CDJS300", "CDJS301", "CDJS302", "CDJS303"]
LABEL = {"CDJS300": "購入部門", "CDJS301": "勘定科目",
         "CDJS302": "補助科目", "CDJS303": "セグメント"}
YM = "202607"
SKIP_WORD = "買掛金"


def _truth_vouchers():
    """正解の伝票 → {"file": PDF名, "pages": set, "rows": [仕訳行]}"""
    text = open(GEN / "journal.csv", "rb").read().decode("cp932")
    journal = [r for r in csv.DictReader(io.StringIO(text))
               if r["CDJS005"][:7].replace("/", "") == YM]
    assign = json.load(open(GEN / "voucher_assign_v26.json"))
    docs = {d["doc_id"]: d for d in json.load(open(GEN / "all_extraction.json"))["documents"]}
    manifest = {x["file_id"]: x["file_name"] for x in json.load(open(GEN / "manifest_files.json"))}
    stem = {fid: name.rsplit(".", 1)[0] for fid, name in manifest.items()}

    def vk(r):
        y, m = int(r["CDJS005"][:4]), int(r["CDJS005"][5:7])
        return "%d-%d" % (y - 1 if m <= 3 else y, int(r["CDJS008"]))

    by_v = collections.defaultdict(list)
    for r in journal:
        by_v[vk(r)].append(r)

    out = {}
    for v, ds in assign.items():
        if v not in by_v:
            continue
        fids, pages = collections.Counter(), set()
        for x in ds:
            d = docs.get(x)
            if not d:
                continue
            fids[d["file_id"]] += 1
            pages.update(range(int(d.get("page_from") or 0), int(d.get("page_to") or 0) + 1))
        if not fids:
            continue
        fid = fids.most_common(1)[0][0]
        if SKIP_WORD in manifest.get(fid, ""):
            continue
        out[v] = {"file": stem[fid], "pages": pages - {0}, "rows": by_v[v]}
    return out, len(journal)


def _pred_vouchers(path):
    """予測の伝票 → {"file": PDF名, "pages": set, "lines": [(金額, 4項目)], "auto": bool}"""
    with open(path, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    out = collections.OrderedDict()
    for r in rows:
        r = {(k or "").strip().lstrip("﻿"): v for k, v in r.items()}
        if str(r.get("行No") or "0") == "0":
            continue                                  # 金額だけの説明行
        # ローカル実行版は "F46_債務446_..." と file_id を頭に付けている。落として揃える
        src = re.sub(r"^F\d+_", "", r.get("参照元ファイル") or "")
        if SKIP_WORD in src:
            continue
        key = r.get("伝票キー") or ""
        v = out.setdefault(key, {"file": src, "pages": set(), "lines": [], "auto": True})
        v["pages"].update(int(x) for x in (r.get("参照元ページ") or "").split(",") if x.strip())
        try:
            amt = int(float((r.get("金額") or "0").replace(",", "")))
        except ValueError:
            amt = 0
        v["lines"].append((amt, tuple(str(r.get(k + "_" + LABEL[k]) or "").strip() for k in F4)))
        if (r.get("要確認") or "").strip():
            v["auto"] = False
    return out


def _assign(truth, pred):
    """同じPDFの中で、ページの重なりが大きい順に1対1で割り当てる。"""
    cand = []
    for tv, t in truth.items():
        for pv, p in pred.items():
            if t["file"] != p["file"]:
                continue
            ov = len(t["pages"] & p["pages"])
            if ov:
                cand.append((ov, tv, pv))
    cand.sort(reverse=True)
    t_used, p_used, pair = set(), set(), {}
    for ov, tv, pv in cand:
        if tv in t_used or pv in p_used:
            continue
        pair[tv] = pv
        t_used.add(tv)
        p_used.add(pv)
    return pair, [pv for pv in pred if pv not in p_used]


def _by4(lines):
    """4項目ごとに金額を足したもの。行の割り方が違っても会計上同じならこれが一致する。"""
    g = collections.defaultdict(int)
    for amt, key in lines:
        g[key] += amt
    return collections.Counter({k: v for k, v in g.items()})


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    path = sys.argv[1]
    truth_all, n_journal = _truth_vouchers()
    pred = _pred_vouchers(path)

    files = {p["file"] for p in pred.values()}
    truth = {v: t for v, t in truth_all.items() if t["file"] in files}
    pair, extra = _assign(truth, pred)

    print("■ %s" % os.path.basename(path))
    print("  PDF %d本 / 正解の伝票 %d本・%d行 / 予測の伝票 %d本・%d行"
          % (len(files), len(truth), sum(len(t["rows"]) for t in truth.values()),
             len(pred), sum(len(p["lines"]) for p in pred.values())))
    print("  （7月の仕訳は全体で %d行。PDFに紐づく伝票だけを母数にしている）" % n_journal)

    c = collections.Counter()
    ng = []
    for tv, t in truth.items():
        T = t["rows"]
        c["伝票"] += 1
        c["行"] += len(T)
        tk = collections.Counter((int(x["CDJS312"] or 0),) + tuple(x[k].strip() for k in F4)
                                 for x in T)
        t4 = _by4([(int(x["CDJS312"] or 0), tuple(x[k].strip() for k in F4)) for x in T])
        tamt = sorted(int(x["CDJS312"] or 0) for x in T)

        pv = pair.get(tv)
        if not pv:
            ng.append((tv, T, "対応する予測なし（取りこぼし）"))
            continue
        p = pred[pv]
        qk = collections.Counter((amt,) + key for amt, key in p["lines"])
        q4 = _by4(p["lines"])
        qamt = sorted(amt for amt, _ in p["lines"])

        c["金額一致"] += int(sum(tamt) == sum(qamt))
        c["行の割り方一致"] += int(tamt == qamt)
        ok = int(t4 == q4)
        c["会計上正しい"] += ok
        c["行データ完全一致"] += int(tk == qk)
        c["完全一致行"] += sum((tk & qk).values())
        c["4項目一致行"] += sum((_by4_rows(t4) & _by4_rows(q4)).values())
        for k in F4:
            tvc = collections.Counter(x[k].strip() for x in T)
            qvc = collections.Counter(key[F4.index(k)] for _, key in p["lines"])
            c["一致_" + k] += sum((tvc & qvc).values())
        if not ok:
            why = "金額が違う（%s円 → %s円）" % (format(sum(tamt), ","), format(sum(qamt), ",")) \
                if sum(tamt) != sum(qamt) else "金額は合うが4項目が違う"
            ng.append((tv, T, why))

    m, nl = c["伝票"], c["行"]
    print()
    print("  --- 伝票単位（%d本）---" % m)
    for k in ("金額一致", "行の割り方一致", "会計上正しい", "行データ完全一致"):
        mark = "**" if k == "会計上正しい" else "  "
        print("  %s%-16s %5.1f%%  (%d/%d)" % (mark, k, 100 * c[k] / m, c[k], m))
    print()
    print("  --- 行単位（正解 %d行）---" % nl)
    print("  金額も4項目も一致  %5.1f%%  (%d/%d)" % (100 * c["完全一致行"] / nl, c["完全一致行"], nl))
    print("  項目ごとの一致     " + " ".join(
        "%s%.0f%%" % (LABEL[k], 100 * c["一致_" + k] / nl) for k in F4))

    # --- PDF単位。伝票の対応づけを介さないので、括りのずれの影響を受けない ---
    # 「元帳に載る数字が合っているか」はこちらのほうが素直に測れる。
    tf, qf = collections.defaultdict(list), collections.defaultdict(list)
    for t in truth.values():
        tf[t["file"]] += [(int(x["CDJS312"] or 0), tuple(x[k].strip() for k in F4))
                          for x in t["rows"]]
    for p in pred.values():
        qf[p["file"]] += p["lines"]
    n_file_ok = n_amt_ok = 0
    hit_rows = 0
    for f in tf:
        if _by4(tf[f]) == _by4(qf.get(f, [])):
            n_file_ok += 1
        if sum(a for a, _ in tf[f]) == sum(a for a, _ in qf.get(f, [])):
            n_amt_ok += 1
        hit_rows += sum((collections.Counter(tf[f]) & collections.Counter(qf.get(f, []))).values())
    print()
    print("  --- PDF単位（%d本・伝票の対応づけを介さない）---" % len(tf))
    print("    金額の合計が一致    %5.1f%%  (%d/%d)" % (100 * n_amt_ok / len(tf), n_amt_ok, len(tf)))
    print("  **会計上正しい        %5.1f%%  (%d/%d)" % (100 * n_file_ok / len(tf), n_file_ok, len(tf)))
    print("    行が一致（行単位）  %5.1f%%  (%d/%d)" % (100 * hit_rows / nl, hit_rows, nl))

    if extra:
        print()
        print("  正解に対応しない予測の伝票 %d本（二重計上の疑い）" % len(extra))
        for pv in extra[:10]:
            p = pred[pv]
            print("    %-40s %s円 %d行" % (
                pv[-40:], format(sum(a for a, _ in p["lines"]), ","), len(p["lines"])))

    if ng:
        print()
        print("  会計上ちがった伝票 %d本" % len(ng))
        for tv, T, why in ng[:20]:
            print("    %-9s 支払先%s %s円 正解%d行  %s"
                  % (tv, T[0]["CDJS009"],
                     format(sum(int(x["CDJS312"] or 0) for x in T), ","), len(T), why))
    return 0


def _by4_rows(counter):
    """{4項目: 金額} を「金額つきの行」の多重集合に直す（& が取れるように）。"""
    return collections.Counter({(v,) + k: 1 for k, v in counter.items()})


if __name__ == "__main__":
    sys.exit(main())
