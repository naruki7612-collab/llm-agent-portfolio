# -*- coding: utf-8 -*-
"""s50 — Step 7: 検算して、予測仕訳一覧のCSVを出す。AIは呼ばない。

使い方:
    import sys; sys.path.insert(0, "notes/システム")
    import s50_export as s
    await s.main()

出力:
    Noteの 02_仕訳データ/奉行取込用.csv・要確認.csv（結果の本体。積み増す）
    output/奉行取込用_YYYYMMDD.csv / 要確認_YYYYMMDD.csv（このセッション分の控え）

要確認にする条件: 確信度 < CONF_THRESHOLD、検算に引っかかった、伝票にフラグ、
機械の見立てと食い違い、のどれか。
"""
import csv, datetime, glob, io, json, os, re

# Noteのパスは cwd 相対では解決されないので、自分のあるディレクトリを sys.path に通す
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))

import config as C
import rules as R
import jsonlog as JL
import store as S
import notefs as NF

COLS = ["伝票キー", "行No", "対象年月", "支払先コード", "支払先名", "金額",
        "CDJS300_購入部門", "CDJS301_勘定科目", "CDJS302_補助科目", "CDJS303_セグメント",
        "CDJS103_摘要", "請求書の品名", "確信度", "根拠種別", "根拠", "機械との食い違い",
        "要確認", "要確認の理由", "参照元ファイル", "参照元ページ", "参照元明細ID"]


def _load():
    vp = os.path.join(C.D_VOUCH, "vouchers.json")
    if not os.path.exists(vp):
        return None, {}
    with open(vp, encoding="utf-8") as f:
        vouchers = {v["伝票キー"]: v for v in json.load(f)}
    cls = {}
    for p in sorted(glob.glob(os.path.join(C.D_CLASS, "*.json"))):
        with open(p, encoding="utf-8") as f:
            d = json.load(f)
        cls[d["伝票キー"]] = d
    return vouchers, cls


def _voucher_keys_in_csv(path):
    """すでに結果CSVに出ている伝票を {(支払先, 対象年月, 伝票金額): {伝票キー}} で返す。

    s30 の _merge_across_files はそのバッチで処理したPDFの中だけを見るので、
    バッチ・ルームをまたいだ「一部再送」は畳めない。それをここで落とすためのキー。

    **どの伝票キーが出したキーかも持つ。** 集合だけにすると、同じルームで
    Step 7 をやり直したときに **自分がさっき書いた行を「別ファイルの重複」と
    判定して自分を消す**（そして「重複で出しませんでした」の説明行まで足す）。
    """
    text = NF.read_text(path)
    if not text:
        return {}
    try:
        rows = list(csv.DictReader(io.StringIO(text)))
    except Exception:                                     # noqa: BLE001
        return {}
    tot, info = {}, {}
    for r in rows:
        k = r.get("伝票キー")
        if str(r.get("行No") or "0") == "0":
            continue                  # 説明だけの行（下の「伝票を作れなかった」行）は数えない
        try:
            tot[k] = tot.get(k, 0) + int(r.get("金額") or 0)
        except (TypeError, ValueError):
            continue
        info[k] = (R.norm_payee(r.get("支払先名")), str(r.get("対象年月") or ""))
    # 支払先名も金額も無いキーは落とす。これを残すと ("", "", 0) が入り、
    # あとから来た別のPDFの説明行まで「重複」と判定して消してしまう。
    out = {}
    for k, v in tot.items():
        if not (info.get(k) and info[k][0] and v):
            continue
        out.setdefault((info[k][0], info[k][1], v), set()).add(k)
    return out


async def _write_csv_note(path, rows, keys):
    """Note上のCSVに積み増す。file_create は上書きしかできないので全部書き直す。

    `keys` は今回作り直した伝票キーの集合。**その伝票の古い行はまず全部消す。**
    伝票キー＋行Noで置き換えるだけだと、作り直して行数が減ったとき（5行→3行）に
    古い行が残る。
    """
    text = NF.read_text(path)
    old = _read_rows(text)
    if old is None:
        # 読めたのに中身が解釈できない。ここで書くと積み上げてきた行を
        # このバッチの分で置き換えてしまうので、**書かずに止める**
        raise RuntimeError(
            "%s の見出し行に「伝票キー」がありません。上書きするとこれまでの"
            "結果が消えるので中断しました。ファイルを確認してください" % path)

    # 今回の伝票が既に入っているなら、その古い行を消さないといけないので
    # 全部書き直す。入っていなければ**追記で足す**（読めなくても前の分が消えない）
    stale = [r for r in old if r.get("伝票キー") in keys]
    if text and not stale:
        add = _to_csv(rows, header=False)
        how = await NF.append_text(path, add) if add else "何もしない"
        print("  %s に %d行 追記（%s）" % (path, len(rows), how))
        return NF.read_text(path) or (text + add)

    keep = [r for r in old if r.get("伝票キー") not in keys]
    # Excelで開いても文字化けしないようBOMを付ける
    full = "\ufeff" + _to_csv(keep + rows, header=True)
    await NF.write_text(path, full)
    if stale:
        print("  %s を書き直し（作り直した伝票の古い %d行 を消した）" % (path, len(stale)))
    return full


def _to_csv(rows, header=True):
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=COLS, lineterminator="\n")
    if header:
        w.writeheader()
    w.writerows(rows)
    return buf.getvalue()


def _read_rows(text):
    """積み増しの元になる既存CSVを読む。空なら []、解釈できなければ None。

    列は COLS だけに揃える。古い版が書いたCSVに知らない列が混ざっていると
    DictWriter が「fieldnames に無い列がある」で落ち、**そのバッチの結果ごと
    失う**。見出しの前後の空白とBOMも落とす（BOMが2個付いた例があった）。
    """
    if not text:
        return []
    try:
        rd = csv.DictReader(io.StringIO(text))
        names = [(n or "").strip().lstrip("\ufeff") for n in (rd.fieldnames or [])]
        if "伝票キー" not in names:
            return None
        out = []
        for raw in rd:
            r = dict(zip(names, list(raw.values())))
            out.append({c: (r.get(c) if r.get(c) is not None else "") for c in COLS})
        return out
    except Exception:                                     # noqa: BLE001
        return None


async def main(stamp=None):
    """stamp は output/ のファイル名に使う（省略すると今日の日付）。"""
    os.makedirs(C.D_OUTPUT, exist_ok=True)
    vouchers, cls = _load()
    if vouchers is None:
        print("伝票がありません。先に s30_vouchers.main() を実行してください")
        return
    if not cls:
        print("判定結果がありません。先に s40_classify.main() を実行してください")
        return
    dicts = await S.load_small_dicts()
    master_set = dicts["master_set"]
    acct_sub, dept_seg = await S.load_combos()   # C-4 / C-5 の検算に使う
    stamp = stamp or datetime.date.today().strftime("%Y%m%d")

    rows, review = [], []
    ng_voucher = 0
    n_zero = 0
    for key, d in sorted(cls.items()):
        v = vouchers.get(key, {})

        # 金額0の行は出さない（奉行に0円行は入れられない）。
        # 番号を振る前に落とす（後で落とすと 行No に穴が空く）
        kept = [l for l in d["lines"] if int(l.get("amount") or 0) != 0]
        n_zero += len(d["lines"]) - len(kept)
        if not kept:
            if not int(v.get("金額") or 0):
                continue                  # 伝票金額も0。出すものが無いので落とす
            # 伝票金額はあるのに行が全部0円。黙って落とすとその伝票が結果CSVから
            # 消えて完了判定も付かないので、金額だけの1行を要確認に出す
            row = {c: "" for c in COLS}
            row.update({
                "伝票キー": key, "行No": 1, "対象年月": d.get("対象年月"),
                "支払先コード": d.get("支払先コード") or "",
                "支払先名": v.get("発行者") or "", "金額": int(v.get("金額") or 0),
                "確信度": 0, "根拠種別": "—",
                "CDJS103_摘要": "（仕訳行に割れませんでした。4項目を入れてください）",
                "要確認": "○",
                "要確認の理由": "S-5 明細の金額が全部0円で仕訳行に割れなかった",
                "参照元ファイル": v.get("元ファイル") or "",
                "参照元ページ": ",".join(str(x) for x in (v.get("ページ") or [])[:6]),
            })
            rows.append(row)
            review.append(row)
            continue

        lines = [{
            "金額": int(l.get("amount") or 0),
            "CDJS300": l.get("CDJS300"), "CDJS301": l.get("CDJS301"),
            "CDJS302": l.get("CDJS302"), "CDJS303": l.get("CDJS303"),
            "摘要": l.get("摘要"),
        } for l in kept]
        v_ng = R.check_voucher({"金額": v.get("金額")}, lines)
        if v_ng:
            ng_voucher += 1

        for i, l in enumerate(kept, 1):
            ng = list(v_ng) + R.check_master(l, master_set) \
                + R.check_combos(l, acct_sub, dept_seg)
            conf = float(l.get("confidence") or 0)
            vflag = [x for x in v.get("フラグ", []) if str(x).startswith("要確認")]
            mdiff = l.get("機械との食い違い") or []
            need = bool(ng) or bool(vflag) or bool(mdiff) or conf < C.CONF_THRESHOLD
            why = []
            if conf < C.CONF_THRESHOLD:
                why.append("確信度%.2f" % conf)
            why += ng
            why += vflag
            why += ["機械と食い違い: " + x for x in mdiff]
            row = {
                "伝票キー": key, "行No": i, "対象年月": d.get("対象年月"),
                "支払先コード": d.get("支払先コード") or "", "支払先名": v.get("発行者") or "",
                "金額": int(l.get("amount") or 0),
                "CDJS300_購入部門": l.get("CDJS300") or "",
                "CDJS301_勘定科目": l.get("CDJS301") or "",
                "CDJS302_補助科目": l.get("CDJS302") or "",
                "CDJS303_セグメント": l.get("CDJS303") or "",
                "CDJS103_摘要": l.get("摘要") or "",
                "請求書の品名": l.get("item_label") or "",
                "確信度": round(conf, 2),
                "根拠種別": l.get("根拠種別") or "", "根拠": l.get("根拠") or "",
                "機械との食い違い": " / ".join(mdiff),
                "要確認": "○" if need else "",
                "要確認の理由": " / ".join(dict.fromkeys(why)),
                "参照元ファイル": v.get("元ファイル") or "",
                "参照元ページ": ",".join(str(x) for x in (v.get("ページ") or [])[:6]),
                "参照元明細ID": ",".join(str(x) for x in (l.get("参照元明細ID") or [])),
            }
            rows.append(row)
            if need:
                review.append(row)

    # 別のバッチ・ルームで処理したPDFに同じ請求書があったら、その伝票の行は出さない
    seen = _voucher_keys_in_csv(C.OUT1_NOTE)
    n_dup = 0
    got_before = {r["参照元ファイル"] for r in rows} - {""}
    if seen:
        tot, ym_of = {}, {}
        for r in rows:
            tot[r["伝票キー"]] = tot.get(r["伝票キー"], 0) + int(r["金額"] or 0)
            ym_of.setdefault(r["伝票キー"], str(r.get("対象年月") or ""))
        dropped = set()
        for key, amount in tot.items():
            if not amount:
                continue              # 金額0の伝票は畳む判断ができない（キーが総取りになる）
            k = (R.norm_payee(vouchers.get(key, {}).get("発行者")), ym_of[key], amount)
            # 自分と同じ伝票キーしか持っていないなら、それは作り直しなので落とさない
            owners = seen.get(k) or set()
            if owners - {key}:
                dropped.add(key)
        if dropped:
            rows = [r for r in rows if r["伝票キー"] not in dropped]
            review = [r for r in review if r["伝票キー"] not in dropped]
            n_dup = len(dropped)
            for key in sorted(dropped):
                print("  重複のため除外: %s（別のバッチ／ルームで処理した同じ請求書）" % key)

    # 1行も出なかったPDFを、要確認の1行として出す。**これを消してはいけない。**
    # 完了の判定は結果CSVの「参照元ファイル」列で行うので、1行も出ないPDFは
    # いつまでも未処理のままになり毎回読み直される。起きるのは2通り。
    #   ・伝票が1本も立たなかった（仕訳の要らない資料か、読み取り失敗か）
    #   ・立った伝票が全部「別ファイルにある同じ請求書」で畳まれた（一部再送のPDF）
    # 重複の除外より **後** に数える。先に数えると、除外で空になったPDFを取りこぼす
    got_files = {r["参照元ファイル"] for r in rows} - {""}
    read_files = {os.path.splitext(os.path.basename(p))[0]
                  for p in glob.glob(os.path.join(C.D_EXTRACT, "*.json"))}
    for stem in sorted(read_files - got_files):
        dup = stem in got_before             # 行は出ていたが、重複で全部落ちた
        row = {c: "" for c in COLS}
        row.update({
            "伝票キー": stem + "#00", "行No": 0, "金額": 0,
            "参照元ファイル": stem, "確信度": 0, "根拠種別": "—",
            "CDJS103_摘要": "（このPDFからは出す行がありません）",
            "要確認": "○",
            "要確認の理由": ("同じ請求書が別のファイルにあったので、こちらは出していません。"
                            "この行は消してください" if dup else
                            "S-4 読めたが伝票が1本も作れなかった。"
                            "仕訳の必要が無い資料ならこの行を消してください"),
        })
        rows.append(row)
        review.append(row)
        print("  行が出なかった証憑: %s（%s・要確認に1行出しました）"
              % (stem, "重複" if dup else "伝票なし"))

    # 消す範囲は予測仕訳のほう（＝今回作った全伝票）から取る。
    # 出力
    #   Note  … 一覧(21列)と要確認(21列)。**伝票キー／参照元ファイルを持つので、
    #           中断からの再開と重複除外がここで効く。** 形は変えないこと
    #   output… 上と同じもの＋ 奉行49列(cp932)。ダウンロードして使う
    # 要確認CSVには要確認の行だけを渡すので、消す対象は全行(rows)から取る
    touched = {r["伝票キー"] for r in rows}
    out1 = "%s/一覧_%s.csv" % (C.D_OUTPUT, stamp)
    out2 = "%s/要確認_%s.csv" % (C.D_OUTPUT, stamp)
    for note_path, local_path, data in ((C.OUT1_NOTE, out1, rows),
                                        (C.OUT2_NOTE, out2, review)):
        text = await _write_csv_note(note_path, data, touched)
        with open(local_path, "w", newline="", encoding="utf-8") as f:
            f.write(text)

    # 奉行49列。**要確認の行は混ぜない**（人がまだ見ていないものを取り込める形にしない）
    try:
        auto = [r for r in rows if not (r.get("要確認") or "").strip()]
        obc_rows, obc_ng = rows_to_obc(auto)
        _obc_write("%s/奉行取込_%s.csv" % (C.D_OUTPUT, stamp), obc_rows)
        print("奉行取込（49列・Shift-JIS）: %d行 / %d伝票（要確認の%d行は入れていません）"
              % (len(obc_rows), len({r["CDJS008"] for r in obc_rows}), len(rows) - len(auto)))
        for x in obc_ng[:5]:
            print("  ※ %s" % x)
    except Exception as e:                                # noqa: BLE001
        print("奉行49列への変換に失敗しました（%s）。一覧CSVは出ています。" % e)

    JL.log("s50 検算してCSVを出す", 伝票=len(cls), 行=len(rows), 要確認=len(review),
           検算NGの伝票=ng_voucher, ゼロ円で落とした行=n_zero)
    print("--- 伝票 %d / 行 %d ---" % (len(cls), len(rows)))
    print("要確認 %d 行 (%.0f%%) / 自動確定 %d 行" % (
        len(review), 100 * len(review) / len(rows) if rows else 0, len(rows) - len(review)))
    print("検算に引っかかった伝票: %d 本" % ng_voucher)
    if n_zero:
        print("金額0の行を %d 行落としました（奉行に入れられないため）" % n_zero)
    if n_dup:
        print("重複で除外した伝票: %d 本（別のバッチ／ルームにある同じ請求書）" % n_dup)
    for x in (out1, out2, C.OUT1_NOTE, C.OUT2_NOTE):
        print("生成: %s（これまでの分もすべて）" % x)


# =============================================================== 奉行49列への変換
# obc — 予測した仕訳を、奉行に取り込める49列のCSVに直す。
# 
# 御社の `未払金.csv` と**同じ49列・同じ並び・Shift-JIS(cp932)**で出す。
# そのままインポートできる形が目的。
# 
# 列の埋め方は3通り:
#   ① こちらが決める   … 日付・取引先・部門・金額・借方4項目・摘要
#   ② 実績から引く     … 支払条件・貸方科目・税区分など。下の OBC_DEFAULT と
#                         OBC_BY_PAYEE に、実績1,759行から作った値を持っている
#   ③ 固定             … OBCD001 や CDJS001〜003 など、全行で同じ値
# 
# ②は支払先ごとに 85〜100% 当たる（実績で検証）。**外れることがあるので、
# 人が最後に目を通す前提**。未知の支払先は全体の最頻値を使う。
# 
# 消費税（CDJS313）は税込金額から計算する（税区分0010 なら 税込 − 税込/1.1）。
# 実績1,413行のうち1,404行が1円以内で一致した。

# 奉行49列の見出し（未払金.csv と同じ並び）
OBC_COLS = [
    "OBCD001", "CDJS005", "CDJS009", "CDJS100", "CDJS312", "CDJS400", "CDJS001", "CDJS002",
    "CDJS003", "CDJS006", "CDJS007", "CDJS008", "CDJS010", "CDJS011", "CDJS012", "CDJS013",
    "CDJS014", "CDJS101", "CDJS102", "CDJS103", "CDJS104", "CDJS105", "CDJS106", "CDJS109",
    "CDJS110", "CDJS111", "CDJS107", "CDJS200", "CDJS201", "CDJS202", "CDJS203", "CDJS300",
    "CDJS301", "CDJS302", "CDJS303", "CDJS307", "CDJS317", "CDJS314", "CDJS309", "CDJS310",
    "CDJS311", "CDJS318", "CDJS319", "CDJS313", "CDJS401", "CDJS402", "CDJS403", "CDJS404",
    "CDJS405",
]

# 支払先ごとの列の値。未払金.csv の実績1,759行から作った。
# **既定値と違う支払先だけ** 持つ（187支払先のうち141件が何かしら違う）。
#   CDJS200 貸方部門 94件 / CDJS403 91件 / CDJS202 貸方補助 83件 が既定と違う。
#   消すと半分以上の行で貸方が間違うので、減らせない。
OBC_DEFAULT = {
    "CDJS011": "0002",
    "CDJS200": "0000",
    "CDJS201": "212102",
    "CDJS202": "0000",
    "CDJS307": "0010",
    "CDJS310": "2",
    "CDJS311": "1",
    "CDJS314": "10",
    "CDJS317": "0",
    "CDJS318": "0",
    "CDJS401": "スポット",
    "CDJS403": "0",
}

OBC_BY_PAYEE = {
    "20000000": {"CDJS011": "0001", "CDJS307": "0012"},
    "20000001": {"CDJS011": "0001", "CDJS201": "211302"},
    "20000002": {"CDJS011": "0001", "CDJS202": "3241", "CDJS307": "0000"},
    "20110000": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "20120000": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "20260000": {"CDJS011": "0001"},
    "30050000": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "30130000": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "30180000": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "30220000": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50000100": {"CDJS307": "0000", "CDJS310": "0", "CDJS314": "0", "CDJS317": "", "CDJS318": "2"},
    "50000700": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50000800": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50000900": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50002300": {"CDJS307": "0000", "CDJS310": "0", "CDJS314": "0", "CDJS317": "", "CDJS318": "2"},
    "50002400": {"CDJS011": "0022"},
    "50003100": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50003200": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50003500": {"CDJS011": "0001"},
    "50003600": {"CDJS011": "0001"},
    "50003700": {"CDJS011": "0001"},
    "50003900": {"CDJS011": "0001"},
    "50004100": {"CDJS314": "8"},
    "50005000": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50005100": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50005500": {"CDJS200": "0002", "CDJS403": "1"},
    "50005800": {"CDJS011": "0021"},
    "50007500": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50008100": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50008200": {"CDJS200": "0004", "CDJS202": "3220", "CDJS403": "1"},
    "50009300": {"CDJS011": "0001"},
    "50012000": {"CDJS200": "0001"},
    "50012200": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50012400": {"CDJS200": "0001", "CDJS307": "0012"},
    "50012900": {"CDJS307": "0012"},
    "50013200": {"CDJS307": "0012"},
    "50014500": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50014700": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50015000": {"CDJS011": "0001"},
    "50015200": {"CDJS011": "0001"},
    "50015500": {"CDJS011": "0001"},
    "50016000": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50016100": {"CDJS307": "0012", "CDJS314": "8"},
    "50017000": {"CDJS011": "0001", "CDJS310": "0", "CDJS314": "0", "CDJS317": "", "CDJS318": "2"},
    "50017500": {"CDJS307": "0012"},
    "50018200": {"CDJS314": "8", "CDJS317": "1"},
    "50018500": {"CDJS011": "0001"},
    "50019800": {"CDJS011": "0021"},
    "50020200": {"CDJS200": "0001"},
    "50020400": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50020800": {"CDJS200": "0002", "CDJS403": "1"},
    "50021000": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50021100": {"CDJS200": "0004", "CDJS403": "1"},
    "50021200": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50021700": {"CDJS200": "0002", "CDJS403": "1"},
    "50022100": {"CDJS307": "0002", "CDJS310": "0", "CDJS314": "0", "CDJS317": "", "CDJS318": "2"},
    "50022300": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50022600": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50022700": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50023100": {"CDJS200": "0004", "CDJS202": "3220", "CDJS403": "1"},
    "50023700": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50023900": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50024000": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50024200": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50024600": {"CDJS200": "0001", "CDJS403": "1"},
    "50024603": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50024700": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50025000": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50025100": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50025500": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50025600": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50025700": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50026000": {"CDJS307": "0012"},
    "50026200": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50026600": {"CDJS200": "0004", "CDJS202": "3220", "CDJS403": "1"},
    "50026700": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50026900": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50027500": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50027600": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50027700": {"CDJS011": "0021"},
    "50028000": {"CDJS200": "0004", "CDJS202": "3220", "CDJS403": "1"},
    "50028100": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50028200": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50028300": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50029500": {"CDJS011": "0001"},
    "50029800": {"CDJS307": "0012"},
    "50029901": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50030100": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50030800": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50031000": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50031100": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50031400": {"CDJS200": "0004", "CDJS202": "3220", "CDJS403": "1"},
    "50031800": {"CDJS011": "0001"},
    "50032000": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50032400": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50033200": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50033400": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50033500": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50033800": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50034100": {"CDJS011": "0001"},
    "50034300": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50034400": {"CDJS200": "0004", "CDJS202": "3220", "CDJS403": "1"},
    "50034500": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50034900": {"CDJS011": "0001", "CDJS307": "0001", "CDJS310": "0", "CDJS314": "0", "CDJS317": "", "CDJS318": "2"},
    "50035700": {"CDJS200": "0001", "CDJS403": "1"},
    "50035900": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50037101": {"CDJS200": "0004", "CDJS403": "1"},
    "50037300": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50037302": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50037306": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50037400": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50037500": {"CDJS200": "0002", "CDJS403": "1"},
    "50037501": {"CDJS200": "0004", "CDJS202": "3220", "CDJS403": "1"},
    "50037600": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50037700": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50037900": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50038000": {"CDJS011": "0001", "CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50038100": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50038200": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50038500": {"CDJS200": "0001", "CDJS403": "1"},
    "50038600": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50039200": {"CDJS307": "0012"},
    "50039300": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50039500": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50040100": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50040500": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50040600": {"CDJS200": "0001", "CDJS202": "3220", "CDJS403": "1"},
    "50041300": {"CDJS011": "0001"},
    "50041700": {"CDJS011": "0022"},
    "50041900": {"CDJS200": "0002", "CDJS202": "3220", "CDJS403": "1"},
    "50043100": {"CDJS307": "0012"},
    "50043200": {"CDJS011": "0001"},
    "50043300": {"CDJS011": "0001"},
    "50045900": {"CDJS307": "0000", "CDJS310": "0", "CDJS311": "0", "CDJS314": "0", "CDJS317": "", "CDJS318": "2"},
    "50048900": {"CDJS307": "0012"},
    "50049000": {"CDJS011": "0022"},
    "50049200": {"CDJS011": "0021"},
    "50049900": {"CDJS011": "0021", "CDJS307": "0001", "CDJS310": "0", "CDJS314": "0", "CDJS317": "", "CDJS318": "2"},
    "90000300": {"CDJS307": "0000", "CDJS310": "0", "CDJS314": "0", "CDJS317": "", "CDJS318": "2"},
    "90000400": {"CDJS307": "0310", "CDJS318": "1"},
    "90000700": {"CDJS307": "0000", "CDJS310": "0", "CDJS314": "0", "CDJS317": "", "CDJS318": "2"},
}

# 実績に出てくる支払先。ここに無い＝初めての支払先なので、既定値を使ったと警告する。
OBC_KNOWN = set("""
00000000 20000000 20000001 20000002 20110000 20120000 20260000 30050000 30130000 30180000
30220000 50000100 50000400 50000601 50000700 50000800 50000900 50001000 50001100 50001200
50002300 50002400 50002700 50003100 50003200 50003500 50003600 50003700 50003900 50004100
50005000 50005100 50005500 50005600 50005700 50005800 50006500 50007500 50008000 50008100
50008200 50008900 50009300 50010100 50010800 50010900 50011001 50012000 50012200 50012400
50012900 50013200 50013600 50013700 50014200 50014500 50014700 50014900 50015000 50015200
50015500 50016000 50016100 50016900 50017000 50017400 50017500 50018200 50018400 50018500
50019000 50019600 50019700 50019800 50019900 50020200 50020400 50020800 50021000 50021100
50021200 50021700 50021800 50022100 50022300 50022600 50022700 50023100 50023700 50023900
50024000 50024200 50024401 50024500 50024600 50024603 50024700 50025000 50025100 50025300
50025500 50025600 50025700 50026000 50026200 50026600 50026700 50026900 50027500 50027600
50027700 50028000 50028100 50028200 50028300 50028800 50029500 50029800 50029901 50030100
50030700 50030800 50031000 50031100 50031400 50031700 50031800 50032000 50032400 50032600
50033200 50033400 50033500 50033800 50034100 50034300 50034400 50034500 50034600 50034900
50035700 50035900 50036400 50037101 50037300 50037302 50037306 50037400 50037500 50037501
50037600 50037700 50037900 50038000 50038100 50038200 50038500 50038600 50038800 50039100
50039200 50039300 50039500 50040100 50040500 50040600 50040800 50041300 50041700 50041900
50043100 50043200 50043300 50044200 50045300 50045900 50048600 50048700 50048800 50048900
50049000 50049100 50049200 50049900 90000300 90000400 90000700
""".split())


def columns():
    return list(OBC_COLS)


def _ym_last_day(ymd):
    """締日の翌月末。'2026/07/31' → '2026/08/31'。読めなければ空。"""
    m = re.search(r"(\d{4})\D(\d{1,2})", str(ymd or ""))
    if not m:
        return ""
    y, mo = int(m.group(1)), int(m.group(2))
    y, mo = (y + 1, 1) if mo == 12 else (y, mo + 1)
    import calendar
    return "%04d/%02d/%02d" % (y, mo, calendar.monthrange(y, mo)[1])


def _closing(ym):
    """対象年月（202607）→ 締日 '2026/07/31'。"""
    s = str(ym or "")
    m = re.fullmatch(r"(\d{4})(\d{2})", s)
    if not m:
        return ""
    import calendar
    y, mo = int(m.group(1)), int(m.group(2))
    return "%04d/%02d/%02d" % (y, mo, calendar.monthrange(y, mo)[1])


def _tax(amount, kubun):
    """税込金額から消費税を出す。課税10%(0010)以外は0。"""
    a = int(amount or 0)
    if not a or str(kubun) not in ("0010", "0012", "0310"):
        return 0
    return a - int(round(a / 1.1))


def rows_to_obc(rows):
    """予測仕訳の行（s50 の COLS 形式）→ 奉行49列の行。

    rows は伝票キーの順に並んでいること。**伝票の1行目だけ OBCD001 に '*' を立てる。**
    実績でも「'*' の数 = 伝票の数」で、1行目だけに付いていた。
    """
    H = columns()
    dflt, ex = OBC_DEFAULT, OBC_BY_PAYEE
    out, ng, seen = [], [], {}

    for r in rows:
        if str(r.get("行No") or "0") == "0":
            continue                                      # 金額だけの説明行は出さない
        payee = str(r.get("支払先コード") or "").strip()
        # 既定値に、その支払先の例外だけ上書きする
        d = dict(dflt); d.update(ex.get(payee) or {})
        if payee and payee not in OBC_KNOWN:
            ng.append("支払先 %s は実績に無いので既定値を使いました（要確認）" % payee)

        key = str(r.get("伝票キー") or "")
        first = key not in seen
        if first:
            # 伝票Noは奉行側で採番する想定。ここでは**通し番号**を入れる。
            # ハッシュのようなでたらめな値だと、既存の伝票と衝突したときに気づけない。
            seen[key] = len(seen) + 1

        ym = str(r.get("対象年月") or "")
        closing = _closing(ym)
        amount = int(float(str(r.get("金額") or 0).replace(",", "") or 0))
        seg = str(r.get("CDJS303_セグメント") or "")
        kubun = d.get("CDJS307", "0010")

        v = {c: "" for c in H}
        v["OBCD001"] = "*" if first else ""               # 伝票の1行目だけ
        v["CDJS005"] = closing                            # 伝票日付＝締日
        v["CDJS006"] = closing
        v["CDJS012"] = _ym_last_day(closing)              # 支払期日＝翌月末
        v["CDJS008"] = "%06d" % seen[key]                 # 1から通し。奉行側で振り直す前提
        v["CDJS009"] = payee
        v["CDJS100"] = str(r.get("CDJS300_購入部門") or "")
        v["CDJS103"] = str(r.get("CDJS103_摘要") or "")
        v["CDJS312"] = str(amount)
        v["CDJS313"] = str(_tax(amount, kubun))
        # 借方4項目 — ここが自動仕訳の本体
        v["CDJS300"] = str(r.get("CDJS300_購入部門") or "")
        v["CDJS301"] = str(r.get("CDJS301_勘定科目") or "")
        v["CDJS302"] = str(r.get("CDJS302_補助科目") or "")
        v["CDJS303"] = seg
        # 貸方 — 実績から。セグは借方に合わせる（実績で1,726/1,759が一致）
        v["CDJS200"] = d.get("CDJS200", "0000")
        v["CDJS201"] = d.get("CDJS201", "212102")
        v["CDJS202"] = d.get("CDJS202", "0000")
        v["CDJS203"] = seg
        # 実績から引く残り
        for k in ("CDJS011", "CDJS307", "CDJS310", "CDJS311",
                  "CDJS314", "CDJS317", "CDJS318", "CDJS401", "CDJS403"):
            if k in H:
                v[k] = d.get(k, "")
        v["CDJS405"] = str(amount) if d.get("CDJS403") == "1" else "0"
        # 全行で同じ値
        for k, val in (("CDJS400", "0000"), ("CDJS001", "00"), ("CDJS002", "0"),
                       ("CDJS003", "1"), ("CDJS007", "0"), ("CDJS010", "0000"),
                       ("CDJS013", "0"), ("CDJS101", "0000"), ("CDJS106", "0"),
                       ("CDJS107", "0"), ("CDJS309", "0001")):
            if k in H:
                v[k] = val
        out.append(v)
    return out, sorted(set(ng))


def _obc_to_csv(rows):
    """奉行49列のCSVの中身（文字列）を返す。書き出しは呼び側で cp932 にする。"""
    H = columns()
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=H, lineterminator="\r\n", quoting=csv.QUOTE_ALL)
    w.writeheader()
    for v in rows:
        w.writerow(v)
    return buf.getvalue()


def _obc_write(path, rows):
    """cp932 で書き出す。読めない文字は「?」にせず例外にする（気づけるように）。"""
    text = _obc_to_csv(rows)
    with open(path, "wb") as f:
        f.write(text.encode("cp932"))
    return len(rows)


