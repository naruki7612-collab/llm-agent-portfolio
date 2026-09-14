# -*- coding: utf-8 -*-
"""s60 — 辞書を更新する。人が確認した仕訳を学習データに戻す。【db 書き込み】

使い方:
    import sys; sys.path.insert(0, "notes/システム")
    import s60_feedback as s
    await s.preview()      # 何が増えるかだけ見る（書き込まない）
    await s.main()         # 実際に辞書へ入れる（承認カードが出ます）

入力: Note の 03_確認済み仕訳/ に置いた「確認済みCSV」（uploads/ でも可）。
      02_仕訳データ/要確認.csv をそのまま直したものでよい。
      直した行は 要確認 列を空にするか、確認結果 列に「OK」または直した値を入れる。

更新するもの:
  1. beta_journal        … 確認済みの仕訳を実績として追加（次回からこれを見て判断する）
  2. beta_item_account   … 支払先 × 品名 → 4項目 の対応を追加・件数を+1

確認済みとして入れる行は ルール期='新' で登録する（値が割れたとき優先される）。
"""
import csv, glob, os

import agent_sdk as at
# Noteのパスは cwd 相対では解決されないので、自分のあるディレクトリを sys.path に通す
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))

import config as C
import rules as R
import store as S

NEEDED = ["伝票キー", "行No", "CDJS300_購入部門", "CDJS301_勘定科目",
          "CDJS302_補助科目", "CDJS303_セグメント"]
# 「請求書の品名」列があれば beta_item_account も更新する（無くても仕訳実績は更新される）


def _read(path):
    for enc in ("utf-8-sig", "cp932"):
        try:
            with open(path, encoding=enc, newline="") as f:
                return list(csv.DictReader(f))
        except UnicodeDecodeError:
            continue
    raise ValueError("文字コードが読めません: %s" % path)


def _collect(src=None):
    """確認済みCSVを集める。src を省略すると Note の 03_確認済み仕訳 と uploads/ を見る。"""
    dirs = [src] if src else [C.D_FEEDBACK, "uploads"]
    paths = sorted(p for d in dirs for p in glob.glob(os.path.join(d, "*.csv")))
    rows, used = [], []
    for p in paths:
        try:
            r = _read(p)
        except Exception:
            continue
        if r and all(c in r[0] for c in NEEDED):
            rows += r
            used.append(os.path.basename(p))
    return rows, used


def _build(rows):
    """CSV → (仕訳レコード, 品名対応レコード)。空欄の行は捨てる。"""
    journal, items, skipped = [], [], 0
    for r in rows:
        dept = (r.get("CDJS300_購入部門") or "").strip()
        acct = (r.get("CDJS301_勘定科目") or "").strip()
        sub = (r.get("CDJS302_補助科目") or "").strip()
        seg = (r.get("CDJS303_セグメント") or "").strip()
        amt = (r.get("金額") or "0").replace(",", "").strip()
        if not (dept and acct and seg):
            skipped += 1
            continue
        try:
            amt = int(float(amt))
        except ValueError:
            amt = 0
        ym = (r.get("対象年月") or "").strip()
        payee = (r.get("支払先コード") or "").strip()
        rec = {
            "伝票キー": (r.get("伝票キー") or "").strip(),
            "支払先コード": payee, "取引先名": (r.get("支払先名") or "").strip(),
            "摘要": (r.get("CDJS103_摘要") or "").strip(),
            "購入部門": dept, "勘定科目": acct, "補助科目": sub or "0000", "セグメント": seg,
            "本体金額": amt, "ルール期": "新", "登録元": "確認済み",
        }
        if ym:
            rec["年月"] = ym          # 空値は入れない（列の型が崩れる）
        journal.append(rec)

        # beta_item_account は「請求書の品名 → 4項目」。摘要ではなく品名を入れる。
        # 品名が無い（列が消されている）行は足さない
        name = (r.get("請求書の品名") or "").strip()
        if payee and name:
            it = {
                "支払先コード": payee, "取引先名": (r.get("支払先名") or "").strip(),
                "品名": name, "品名正規化": R.norm_item(name),
                "購入部門": dept, "勘定科目": acct, "補助科目": sub or "0000", "セグメント": seg,
                "件数": 1, "ルール期": "新",
            }
            if ym:
                it["最終年月"] = ym
            items.append(it)
    return journal, items, skipped


def _dedup_items(items):
    """同じ（支払先×品名×4項目）を1行にまとめ、件数を足す。"""
    agg = {}
    for it in items:
        k = (it["支払先コード"], it["品名正規化"], it["購入部門"],
             it["勘定科目"], it["補助科目"], it["セグメント"])
        if k in agg:
            agg[k]["件数"] += 1
            if it.get("最終年月", "") > agg[k].get("最終年月", ""):
                agg[k]["最終年月"] = it["最終年月"]
        else:
            agg[k] = dict(it)
    return list(agg.values())


async def preview(src=None):
    rows, used = _collect(src)
    if not rows:
        print("確認済みCSVが見つかりません。%s に置いてください（必要な列: %s）"
              % (C.D_FEEDBACK, ", ".join(NEEDED)))
        return
    journal, items, skipped = _build(rows)
    items = _dedup_items(items)
    print("--- 読み込み %d 行（%s）---" % (len(rows), ", ".join(used)))
    print("辞書に足す仕訳: %d 行" % len(journal))
    print("品名の対応に足す: %d 行" % len(items))
    print("4項目が空で飛ばした行: %d" % skipped)
    for j in journal[:5]:
        print("  %s %s %s/%s/%s/%s %s" % (j["伝票キー"], j["摘要"][:20], j["購入部門"],
                                          j["勘定科目"], j["補助科目"], j["セグメント"], j["本体金額"]))
    return journal, items


async def main(src=None):
    got = await preview(src)
    if not got:
        return
    journal, items = got

    n1 = await S.insert_rows(C.T_JOURNAL, journal)
    n2 = await S.insert_rows(C.T_ITEM_ACCOUNT, items)
    print("beta_journal に %d 行 / beta_item_account に %d 行 追加しました" % (n1, n2))

    total = await at.db.table(C.T_JOURNAL).count_documents({})
    print("--- 辞書の仕訳実績は合計 %d 行になりました ---" % total)
