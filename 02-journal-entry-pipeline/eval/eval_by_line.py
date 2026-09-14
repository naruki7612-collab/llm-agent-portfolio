# -*- coding: utf-8 -*-
"""AgentPlatformが出した予測仕訳CSVを、正解（journal.csv）と突き合わせる。

    python3 CSV評価スクリプト.py <予測仕訳CSV> [要確認CSV]

`評価スクリプト.py` はサンドボックスの tmp/beta/*.json が要るので、
**CSVしか手元に無いとき**はこちらを使う。

突き合わせ方: 支払先コードと金額が同じ行を対応づけ、4項目（CDJS300〜303）が
全部合っていれば「一致」、対応はついたが1つでも違えば「誤り」、
対応する正解が無ければ「不能」とする。行の割り方がずれると金額が合わないので
「不能」に入る。

**予測側から数える。** 入れた証憑の本数で正解の側の母数が変わるため
（11本入れて232行の正解と比べると、入れていない証憑のぶんが取りこぼしに見える）。
買掛金のファイル（債務578）の予測行は落とす（未払金の仕組みなので対象外）。
"""
import collections
import csv
import io
import os
import pathlib
import sys

HERE = pathlib.Path(__file__).resolve().parent
GEN = HERE.parent / "4_元データ"
F4 = ["CDJS300", "CDJS301", "CDJS302", "CDJS303"]
YM = "202607"
# 買掛金。README の「F56 は使わない」に対応する
SKIP_WORD = "買掛金"


def _truth():
    """正解の行。戻り: [(支払先コード, 金額, {4項目}, 摘要)]

    買掛金の伝票は voucher_assign に載っていないので、ファイル単位では外せない。
    **予測側から数える**ので外さなくてよい（予測に買掛金が無ければ突き合わない）。
    """
    text = open(GEN / "journal.csv", "rb").read().decode("cp932")
    rows = [r for r in csv.DictReader(io.StringIO(text))
            if r["CDJS005"][:7].replace("/", "") == YM]
    return [(str(r["CDJS009"]).strip(), int(r["CDJS312"] or 0),
             {k: str(r[k]).strip() for k in F4}, r["CDJS103"].strip())
            for r in rows]


def _pred(path):
    """予測の行。説明だけの行（行No 0）と買掛金のファイルを除く。"""
    with open(path, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    out, n_skip, n_note = [], 0, 0
    for r in rows:
        r = {(k or "").strip().lstrip("﻿"): v for k, v in r.items()}
        if str(r.get("行No") or "0") == "0":
            n_note += 1
            continue
        if SKIP_WORD in (r.get("参照元ファイル") or ""):
            n_skip += 1
            continue
        out.append(r)
    return out, n_skip, n_note


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    pred_path = sys.argv[1]
    truth = _truth()
    pred, p_skip, p_note = _pred(pred_path)

    print("■ %s" % os.path.basename(pred_path))
    print("  正解 %d行（2026/07 全件）/ 予測 %d行（買掛金 %d行・説明行 %d行を除外）"
          % (len(truth), len(pred), p_skip, p_note))
    print("  ※ **予測側から数える。** 入れた証憑の本数で正解の側が変わるため")

    # 支払先コード＋金額で対応づける
    pool = collections.defaultdict(list)
    for payee, amt, want, memo in truth:
        pool[(payee, amt)].append((want, memo))

    label = {"CDJS300": "購入部門", "CDJS301": "勘定科目",
             "CDJS302": "補助科目", "CDJS303": "セグメント"}
    n_hit = n_wrong = n_none = 0
    auto_hit = auto_wrong = auto_none = 0
    wrong_kind = collections.Counter()
    examples = []
    for r in pred:
        auto = not (r.get("要確認") or "").strip()
        try:
            amt = int(r.get("金額") or 0)
        except ValueError:
            amt = 0
        cand = pool.get((str(r.get("支払先コード") or "").strip(), amt))
        if not cand:
            n_none += 1
            auto_none += 1 if auto else 0
            continue
        want, memo = cand.pop(0)
        diff = [k for k in F4
                if str(r.get(k + "_" + label[k]) or "").strip() != want[k]]
        if diff:
            n_wrong += 1
            auto_wrong += 1 if auto else 0
            wrong_kind[",".join(label[k] for k in diff)] += 1
            if len(examples) < 8:
                examples.append((amt, memo[:22], [(label[k], want[k],
                                 str(r.get(k + "_" + label[k]) or "")) for k in diff],
                                 "自動確定" if auto else "要確認"))
        else:
            n_hit += 1
            auto_hit += 1 if auto else 0

    auto = sum(1 for r in pred if not (r.get("要確認") or "").strip())
    print()
    print("  %-30s %3d 行" % ("正解と一致（4項目すべて）", n_hit))
    print("  %-30s %3d 行" % ("対応はついたが4項目が違う", n_wrong))
    print("  %-30s %3d 行" % ("正解に対応が無い（不能）", n_none))
    print("  %-30s %3d 行" % ("突き合わなかった正解", sum(len(v) for v in pool.values())))
    print()
    print("  %-30s %3d 行 (%.0f%%)" % ("自動確定", auto,
                                       100 * auto / len(pred) if pred else 0))
    print("  %-30s %3d 行 (%.0f%%)" % ("要確認", len(pred) - auto,
                                       100 * (len(pred) - auto) / len(pred) if pred else 0))
    print("  %-30s %3d / %d = %.0f%%" % ("★自動確定のうち誤り＋不能",
                                         auto_wrong + auto_none, auto,
                                         100 * (auto_wrong + auto_none) / auto if auto else 0))
    print("  %-30s %3d / %d = %.0f%%" % ("　うち4項目の誤りだけ", auto_wrong,
                                         auto_hit + auto_wrong,
                                         100 * auto_wrong / (auto_hit + auto_wrong)
                                         if (auto_hit + auto_wrong) else 0))

    if wrong_kind:
        print()
        print("  誤りの内訳（どの項目が違ったか）")
        for k, n in wrong_kind.most_common():
            print("    %-34s %d" % (k, n))
    if examples:
        print()
        print("  誤りの例")
        for amt, memo, d, tag in examples:
            print("    [%s] %10s円 %s" % (tag, format(amt, ","), memo))
            for name, w, g in d:
                print("        %s: 予測 %s → 正解 %s" % (name, g or "（空）", w))
    return 0


if __name__ == "__main__":
    sys.exit(main())
