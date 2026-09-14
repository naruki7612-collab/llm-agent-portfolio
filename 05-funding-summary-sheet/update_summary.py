"""
財源状況集計表 日次更新スクリプト

2026-08-20 の変更: 千円値の丸めを廃止した。
  原本は 検収額(円)/1000 を小数のまま保持している (例: 450.74 = 450,740円)。
  セル書式が '#,##0' で整数表示なので画面では見えないが、旧版は
  int(round(x/1000)) で保存しており最大 ±0.5 千円を失っていた
  (実測: 13部店中12部店が原本の流儀と食い違い)。

  前日比較も round(today) <= round(prev) と両辺を丸めていたため、
  「増えているのに前日以下」と誤って色を付けることがあった。
  実測1件: 本社1 が 450.740 → 451.327 (+587円の増加) なのに
  round が両方 451 になり黄色になっていた。
  round() は単調なので取りこぼし (前日以下なのに色が付かない) は起きず、
  過剰検知だけが起きる。加えて Python の round() は銀行家丸めで
  round(450.5)=450 / round(451.5)=452 と基準が揺れるため比較には使わない。

  集計条件 (除外進捗・財源区分・部店名寄せ・千円化そのもの) は一切変更していない

入力 (uploads/ から自動特定):
    月度リース割賦情報管理表*.csv  (BOM付き Shift-JIS 可)
    財源状況集計表*.xlsx
出力:
    output/財源状況集計表_更新版.xlsx  最終成果物
    output/debug.xlsx                  色付け前のスナップショット。
                                       環境変数 ZAIGEN_DEBUG を付けたときだけ出す
"""

import os
import sys
import glob
from datetime import datetime, date
from pathlib import Path
from collections import defaultdict

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


# ========== 設定 ==========
INPUT_DIR = "uploads"
CSV_PATTERNS = ["*リース割賦*csv", "*月度*csv", "*リース*csv"]  # 上位ほど優先
XLSX_PATTERNS = ["*財源状況集計*xlsx", "*財源*xlsx"]
OUTPUT_XLSX = "output/財源状況集計表_更新版.xlsx"
DEBUG_XLSX = "output/debug.xlsx"

SHEET_NAME = "集計表"
HOLDING_SHEET = "財源保有状況"   # 実物Excelにのみ存在する2枚目
EXCLUDE_PROGRESS = {"競合負け", "自主謝絶", "否決", "没"}
ABC_KEYS = {"A", "B", "C"}
D_KEYS = {"D"}
YELLOW = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

SPECIAL_BRANCH_ALIASES = {
    "本社部": "本社店口",
}


import re
import unicodedata


# v8: 除去対象の空白・区切り記号
_STRIP_CHARS = str.maketrans("", "", " 　\t\r\n-_./\\・")


def normalize_branch(name: str) -> str:
    """
    部店名を比較用に正規化する (v8: NFKC + 空白/記号除去 + 業務語除去)。

    処理順:
      1. NFKC 正規化
           - 全角英数 → 半角 (「本社１」→「本社1」)
           - 丸数字 → 数字 (「本社①」→「本社(1)」)
           - ローマ数字 → アルファベット (「本社Ⅰ」→「本社I」)
           - 全角スペース → 半角スペース
      2. 空白 / 区切り記号を除去
           (スペース・タブ・ハイフン・アンダースコア・ドット・スラッシュ・中黒 等)
      3. 業務語 (支店・チーム・営業) を除去
      4. 小文字化 (英字混在時の吸収)

    例:
      「本社 １ 支店」→「本社1」
      「本社_1」    →「本社1」
      「本社-1」    →「本社1」
      「本社１」    →「本社1」
      「本社1営業」  →「本社1」
    """
    if not name:
        return ""
    # 1. NFKC 正規化 (全角英数・記号を半角に)
    s = unicodedata.normalize("NFKC", str(name))
    # 2. 空白 / 区切り記号を除去
    s = s.translate(_STRIP_CHARS)
    # 3. 業務語を除去
    s = re.sub(r"(支店|チーム|営業)", "", s)
    # 4. 小文字化 (英字混在時)
    s = s.lower()
    return s.strip()


# ========== 入力ファイル解決 (案F: 部分一致 + 中身判定 ハイブリッド) ==========

CSV_REQUIRED_COLS = ("部店", "時期", "確度", "進捗", "検収額")


_BOM = b"\xef\xbb\xbf"


def strip_boms(raw: bytes) -> tuple[bytes, int]:
    """先頭の UTF-8 BOM を「あるだけ」剥がす。

    2026-08-25 の実物 CSV は BOM が 2 つ続いていた (efbbbf efbbbf + cp932 本体)。
    1 つだけ剥がす実装だと 2 つ目が残り、cp932 デコードが position 0 で失敗して
    「中身が想定と異なります」で弾かれる。書き出し側の事情は読めないので、
    何個続いていても剥がす。
    """
    n = 0
    while raw.startswith(_BOM):
        raw = raw[len(_BOM):]
        n += 1
    return raw, n


def _looks_like_target_csv(path: str) -> bool:
    """
    CSV の中身を軽く覗いて「これは月度リース割賦情報管理表だ」と判定する。
    ヘッダ行に必須列 (部店/時期/確度/進捗/検収額) のうち 4 つ以上が
    含まれていれば真とみなす (ヘッダ表記の微差にはある程度寛容にする)。
    """
    try:
        with open(path, "rb") as f:
            raw = f.read(4096)  # 先頭 4KB で十分 (ヘッダ行だけ見れれば良い)
        raw, _ = strip_boms(raw)
        header_line = raw.split(b"\n", 1)[0]
        # 複数エンコーディングで試す (エンコーディング判定と同じ順)
        for enc in ("utf-8", "cp932", "shift_jis", "utf-8-sig"):
            try:
                decoded = header_line.decode(enc)
            except UnicodeDecodeError:
                continue
            hits = sum(1 for k in CSV_REQUIRED_COLS if k in decoded)
            if hits >= 4:
                return True
        return False
    except Exception:
        return False


MONTH_MARK = "月初"        # ヘッダ行の E 列に必ず入っている目印
MONTH_MARK_COL = 5


def _looks_like_target_xlsx(path: str) -> bool:
    """
    Excel の中身を軽く覗いて「これは財源状況集計表だ」と判定する。
    条件: シート「集計表」があり、E 列に「月初」の行が 1 つ以上ある。

    以前は「A 列に日付型セルがある」で判定していたが、お客様の原本は
    月ラベルが文字列 ('8月' / '53期下期') なのでこれでは却下されてしまう。
    ヘッダ行の E="月初" は書式に依存しない目印で、実物・旧フォーマットとも
    4/39/74 行を誤検知なく拾えることを実測で確認している。
    """
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        if SHEET_NAME not in wb.sheetnames:
            wb.close()
            return False
        ws = wb[SHEET_NAME]
        found = False
        for i, row in enumerate(
                ws.iter_rows(min_col=MONTH_MARK_COL, max_col=MONTH_MARK_COL,
                             values_only=True)):
            if str(row[0]).strip() == MONTH_MARK:
                found = True
                break
            if i > 200:
                break
        wb.close()
        return found
    except Exception:
        return False


def find_input_file(patterns: list, kind: str, content_check) -> str:
    """
    uploads/ 配下から patterns のいずれかにマッチするファイルを 1 つ選ぶ (v9改)。

    1. パターン部分一致で候補を集める
    2. パターンで見つからない場合、拡張子ベースで全走査 (v9 追加のフォールバック)
       これによりファイル名が想定と異なっていても中身が正しければ採用可能
    3. 候補に対し content_check() で中身が本物かを確認して絞り込む
    4. 中身判定通過が 1 つなら採用
    5. 中身判定通過が複数なら最新 mtime を採用 (警告ログ)
    6. 中身判定通過が 0 件なら FileNotFoundError (「名前は似ているが中身が違う」)
    """
    # STEP1: パターン部分一致
    seen = []
    for pat in patterns:
        for p in glob.glob(f"{INPUT_DIR}/{pat}"):
            if p not in seen:
                seen.append(p)

    # STEP1.5 (v9): パターン不一致時は拡張子で全走査してフォールバック
    if not seen:
        # patterns から対象拡張子を推定 (例: "*リース割賦*csv" → ".csv")
        exts = set()
        for pat in patterns:
            low = pat.lower()
            if low.endswith("csv"):
                exts.add(".csv")
            elif low.endswith(("xlsx", "xlsm")):
                exts.update([".xlsx", ".xlsm"])
        if not exts:
            # 想定外のパターン。従来通りエラー
            raise FileNotFoundError(
                f"{kind} が {INPUT_DIR}/ に見つかりません (試したパターン: {patterns})"
            )
        print(
            f"[入力] {kind} 名前パターン一致なし。"
            f"{INPUT_DIR}/ 内の {sorted(exts)} を中身で走査します (v9 フォールバック)。"
        )
        for p in sorted(glob.glob(f"{INPUT_DIR}/*")):
            if Path(p).suffix.lower() in exts and p not in seen:
                seen.append(p)
        if not seen:
            raise FileNotFoundError(
                f"{kind} が {INPUT_DIR}/ に見つかりません "
                f"(名前パターン: {patterns} / 拡張子フォールバック: {sorted(exts)} も 0 件)"
            )

    # STEP2: 中身判定 (候補が 1 件でも実施 — 名前だけ似た別ファイルを弾くため)
    if len(seen) == 1:
        print(f"[入力] {kind} 候補 1 件。中身を確認します。")
    else:
        print(f"[入力] {kind} 候補が複数 ({len(seen)} 件)。中身を確認して絞り込みます。")
    verified = []
    for p in seen:
        ok = content_check(p)
        mark = "✓" if ok else "×"
        print(f"        {mark} {p}")
        if ok:
            verified.append(p)

    # STEP3: 中身判定通過が 1 つ → 採用
    if len(verified) == 1:
        chosen = verified[0]
        print(f"[入力] {kind} = {chosen} (中身判定で確定)")
        return chosen

    # STEP4: 通過が複数 → 最新 mtime を採用
    if len(verified) >= 2:
        verified.sort(key=lambda p: Path(p).stat().st_mtime, reverse=True)
        chosen = verified[0]
        print(f"[入力] {kind} = {chosen} (中身判定通過が複数のため最新を採用)")
        for p in verified[1:]:
            print(f"        (未採用) {p}")
        return chosen

    # STEP5: 通過が 0 件 → エラー (名前は似ているが中身が違う)
    raise FileNotFoundError(
        f"{kind} 候補は見つかりましたが、いずれも中身が想定と異なります。"
        f" 対象ファイルが正しく uploads/ に置かれているか確認してください。"
        f" 候補: {seen}"
    )


# ========== CSV 読込 (エンコーディング自動判定) ==========
def read_csv_auto_encoding(csv_path: str) -> pd.DataFrame:
    """
    実運用で来る 3 パターンを堅牢に読む:
      (a) 純 UTF-8
      (b) UTF-8 BOM 付き (utf-8-sig)
      (c) Shift_JIS / cp932
      (d) BOM だけ UTF-8 で中身が cp932 (アルファリース様の実物)
          → BOM の 3 バイトを剥がしてから cp932 デコードする

    方針: 先頭の BOM を除去したバイト列を、候補エンコーディングで順にデコード試行。
    最初に日本語ヘッダが正しく現れたものを採用する。
    """
    with open(csv_path, "rb") as f:
        raw = f.read()

    body, n_bom = strip_boms(raw)
    had_bom = n_bom > 0

    # ヘッダ 1 行分だけデコードして日本語列名が読めるかで判定
    header_line = body.split(b"\n", 1)[0]

    candidates = ["utf-8", "cp932", "shift_jis", "utf-8-sig"]
    chosen_enc = None
    for enc in candidates:
        try:
            decoded = header_line.decode(enc)
        except UnicodeDecodeError:
            continue
        # 必須列のいずれかが読めていれば採用
        if any(k in decoded for k in ("部店", "時期", "確度", "進捗", "検収額")):
            chosen_enc = enc
            break

    if chosen_enc is None:
        # 最終手段: cp932 で強制読み込み (BOM を剥がしてある想定)
        chosen_enc = "cp932"

    print(f"[入力] CSV エンコーディング判定: BOM={f'あり×{n_bom}' if had_bom else 'なし'} / body={chosen_enc}")

    # BOM を剥がしたバイト列を StringIO 経由で読ませる (元ファイルは触らない)
    import io
    text = body.decode(chosen_enc, errors="replace")
    return pd.read_csv(io.StringIO(text))


# ========== STEP1: CSV読込 & クレンジング ==========
def load_and_clean_csv(csv_path: str) -> pd.DataFrame:
    df = read_csv_auto_encoding(csv_path)
    n_total = len(df)

    required = ["部店", "時期", "確度", "進捗", "検収額"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"CSV に必須列が不足: {missing} / 実際の列={list(df.columns)}")

    df = df[~df["進捗"].isin(EXCLUDE_PROGRESS)].copy()

    df["時期_dt"] = pd.to_datetime(df["時期"], errors="coerce")
    df = df.dropna(subset=["時期_dt"]).copy()
    df["月キー"] = df["時期_dt"].dt.strftime("%Y-%m")

    df["検収額"] = pd.to_numeric(df["検収額"], errors="coerce").fillna(0)

    def _norm_csv_branch(name):
        n = normalize_branch(name)
        return SPECIAL_BRANCH_ALIASES.get(n, n)

    df["部店_norm"] = df["部店"].map(_norm_csv_branch)
    df = df[df["部店_norm"] != ""].copy()

    df["確度"] = df["確度"].astype(str).str.strip().str.upper()
    df["財源区分"] = df["確度"].map(
        lambda x: "ABC" if x in ABC_KEYS else ("D" if x in D_KEYS else None)
    )
    df = df.dropna(subset=["財源区分"]).copy()

    # 財源保有状況シートは A / BC / D の3分割で持つ (集計表は A~C / D の2分割)。
    # 同じ確度から切り方を変えるだけなので、集計表側の値は一切変わらない。
    df["財源区分3"] = df["確度"].map(
        lambda x: "A" if x == "A" else ("BC" if x in ("B", "C") else "D")
    )

    print(f"[STEP1] CSV読込: 全{n_total}件 → 有効{len(df)}件（除外{n_total - len(df)}件）")
    return df


# ========== STEP2: 集計 ==========
def aggregate(df: pd.DataFrame) -> dict:
    agg = defaultdict(float)
    for _, row in df.iterrows():
        key = (row["部店_norm"], row["月キー"], row["財源区分"])
        agg[key] += row["検収額"]
    print(f"[STEP2] 集計グループ数: {len(agg)}")
    return dict(agg)


def aggregate3(df: pd.DataFrame) -> dict:
    """財源保有状況シート用に A / BC / D の3分割で集計する。

    集計対象・除外条件は aggregate() と同一。確度の束ね方だけが違う。
    A + BC = 集計表の A~C財源 になることを実測で確認済み。
    """
    agg = defaultdict(float)
    for _, row in df.iterrows():
        key = (row["部店_norm"], row["月キー"], row["財源区分3"])
        agg[key] += row["検収額"]
    return dict(agg)


# ========== STEP3: Excel表構造検出 ==========
def _find_date_col_in_header(ws, header_row: int, today: date) -> int | None:
    """指定ヘッダ行の中から today に一致する列を返す。無ければ None。"""
    for c in range(5, ws.max_column + 1):
        v = resolve_cell(ws, ws.cell(header_row, c).value)
        if isinstance(v, datetime) and v.date() == today:
            return c
    return None


def _col_label(ws, header_row: int, col) -> str:
    """ログ用に「21列(8/25)」の形にする。日付が読めなければ列番号だけ。"""
    if col is None:
        return "なし"
    v = resolve_cell(ws, ws.cell(header_row, col).value)
    if isinstance(v, datetime):
        return f"{col}列({v.month}/{v.day})"
    if isinstance(v, str) and v.strip():
        return f"{col}列({v.strip()})"
    return f"{col}列"


def _find_prev_col(ws, header_row: int, today_col: int, probe_rows: list) -> int | None:
    """比較対象の列を返す。**当日列の左側で、実際に値が入っている最も近い列**。

    以前は「日付ヘッダがある左隣」を無条件に選んでいた。しかし実行が 1 日でも
    飛ぶとその列は空のままなので、翌日の比較が「空欄スキップ」になり
    色付けが黙って効かなくなっていた（本番で一度も色が付いていなかった）。

    日付を指定して更新する運用（実行が毎営業日連続するとは限らない）に合わせ、
    「前回値が入っている列」と比べる。見つからなければ E 列（月初）を使う。
    """
    for c in range(today_col - 1, 5, -1):
        v = resolve_cell(ws, ws.cell(header_row, c).value)
        if not isinstance(v, datetime):
            continue
        if any(isinstance(ws.cell(r, c).value, (int, float)) for r in probe_rows):
            return c
    return 5 if today_col > 5 else None      # E列 = 月初


DATE_COL_FROM = 6          # F 列から日付が並ぶ
_CELLREF = re.compile(r"^=\s*\$?([A-Z]{1,2})\$?(\d+)\s*$")
_MONTH_LABEL = re.compile(r"^\s*(\d{1,2})\s*月(度)?\s*$")



def resolve_cell(ws, value, depth: int = 3):
    """'=F4' のような同一シートの単純参照を辿って実際の値を返す。

    3 つの表が同じ日付列を共有する設計のため、9月表・下期表のヘッダが
    '=F4' になっているのは正しい。openpyxl はキャッシュ値を持たないので
    参照を自分で辿る。連鎖していても depth 回まで追う。
    四則演算や関数を含む数式は対象外で、その場合は None を返す。
    """
    for _ in range(depth):
        if not (isinstance(value, str) and value.startswith("=")):
            return value
        m = _CELLREF.match(value)
        if not m:
            return None
        value = ws[f"{m.group(1)}{m.group(2)}"].value
    return None


def find_header_rows(ws) -> list:
    """E 列が「月初」の行＝ヘッダ行。その 1 つ上が月ラベル行。"""
    return [r for r in range(2, ws.max_row + 1)
            if str(ws.cell(r, MONTH_MARK_COL).value).strip() == MONTH_MARK]


def header_anchor(ws, header_row: int):
    """ヘッダ行の日付列から (年, 月) を 1 つ取る。年を決める手がかりに使う。"""
    for c in range(DATE_COL_FROM, ws.max_column + 1):
        v = resolve_cell(ws, ws.cell(header_row, c).value)
        if isinstance(v, datetime):
            return v.year, v.month
    return None


def months_ahead(base: date, n: int) -> str:
    """base から n か月後の月キー ('YYYY-MM')。年跨ぎも正しく進む。"""
    m = base.month + n
    return f"{base.year + (m - 1) // 12}-{(m - 1) % 12 + 1:02d}"


def label_month(value, anchor):
    """月ラベルから月キーを読む。読めなければ None。**照合と警告にだけ使う。**

    集計対象の月はラベルではなく「表の位置」で決める（下記 detect_tables 参照）。
    原本のラベルは実態と食い違っていることがあり
    （2026-08 のファイルの3つ目が '53期下期' のまま残っていた）、
    ラベルを根拠にすると誤った月を集計してしまう。
    """
    if isinstance(value, datetime):
        return value.strftime("%Y-%m")
    m = _MONTH_LABEL.match(str(value or "").strip())
    if m and anchor:
        mon = int(m.group(1))
        if 1 <= mon <= 12:
            ay, am = anchor
            return f"{ay + 1 if mon < am else ay}-{mon:02d}"
    return None


def detect_tables(ws, today: date) -> list:
    """
    A列を走査して月ラベル(日付型)を検出し、各表の構造を返す。

    v6 変更: 日付列は各表のヘッダ行を個別に走査して求める。
             実行日に一致する日付列が無い表は today_col=None のまま返す
             (後続の書込・色付けは None の表をスキップするため他月表を汚さない)。
    """
    # 月として読めないラベルも「表の区切り」としては必ず登録する。
    # 区切りから外すと直前の表がその表の行まで範囲を広げ、部店名が同じため
    # 後勝ちで上書きされ、別の月のデータを書き込んでしまう (実測で再現済み)。
    # 表の集計対象月は「表の位置」で決める。上から 当月 / 翌月 / 翌々月。
    # ファイルは月ごとに切り替わる運用（先方確定・2026-09-08）なので、
    # 実行日の月が常に 1 つ目の表になる。
    #
    # ラベルを根拠にしないのは、原本のラベルが実態と食い違うことがあるため
    # （2026-08 のファイルの 3 つ目が '53期下期' のまま残っていた）。
    # ラベルは照合にだけ使い、食い違えば警告を出す。
    month_label_rows = []
    for i, hdr in enumerate(find_header_rows(ws)):
        label_row = hdr - 1
        raw = ws.cell(label_row, 1).value
        mk = months_ahead(today, i)
        shown = label_month(raw, header_anchor(ws, hdr))
        if shown and shown != mk:
            print(f"[警告] 行{label_row} のラベルは {raw!r} ({shown}) ですが、"
                  f"表の位置から {mk} として集計します。"
                  f"別の月のファイルを添付していないか確認してください。",
                  file=sys.stderr)
        elif shown is None:
            print(f"[STEP3] 行{label_row} のラベル {raw!r} からは月を読めないため"
                  f"照合を省略します（{mk} として集計します）。")
        month_label_rows.append((label_row, mk, [mk]))

    tables = []
    for i, (label_row, label_name, month_keys) in enumerate(month_label_rows):
        header_row = label_row + 1
        next_label_row = (
            month_label_rows[i + 1][0]
            if i + 1 < len(month_label_rows)
            else ws.max_row + 2
        )

        # ★ 表ごとに日付列を特定。集計対象の月が無い表は書き込まない
        today_col = (_find_date_col_in_header(ws, header_row, today)
                     if month_keys else None)

        branch_rows = {}
        subtotal_rows = []
        grand_total_row = None
        pending_branches = []
        all_branches_in_table = []

        for r in range(header_row + 1, next_label_row):
            name = ws.cell(r, 1).value
            if not name:
                continue
            if name == "合計":
                grand_total_row = (r, r + 1)
                continue
            if name in ("本社計", "北九州計"):
                subtotal_rows.append((name, r, r + 1, list(pending_branches)))
                pending_branches = []
                continue
            norm = normalize_branch(name)
            branch_rows[norm] = (r, r + 1)
            pending_branches.append(norm)
            all_branches_in_table.append(norm)

        # 比較対象の列は、その表の部店行に値が入っている最も近い左の列
        probe = [r for pair in branch_rows.values() for r in pair]
        prev_col = (_find_prev_col(ws, header_row, today_col, probe)
                    if today_col else None)

        tables.append({
            "month_key": label_name,      # 表示用
            "month_keys": month_keys,     # 集計対象の月 (複数の場合あり)
            "label_row": label_row,
            "header_row": header_row,
            "today_col": today_col,
            "prev_col": prev_col,
            "branch_rows": branch_rows,
            "subtotal_rows": subtotal_rows,
            "grand_total_row": grand_total_row,
            "all_branches": all_branches_in_table,
        })

    return tables


# ========== 財源保有状況シート ==========
HOLD_BRANCH_COL = 18   # R列: 部店名 (3行グループの先頭のみ)
HOLD_KIND_COL = 19     # S列: 'A' / 'BC' / 'D'
HOLD_DATE_ROW = 1      # 1行目: 各データ列の対象月 (日付型)
HOLD_FIRST_DATA_COL = 20  # T列から


def detect_holding(ws) -> dict:
    """財源保有状況シートの構造を読む。

    行: R列の部店名 + S列の区分(A/BC/D)。部店ごとに3行の並び。
    列: 1行目の日付セルで対象月を判別する (集計表の日付ヘッダと同じ方式)。

    戻り値 {"rows": {(部店norm, 区分): 行}, "cols": {月キー: 列},
            "subtotals": [(名前, {区分: 行}, [構成部店])], "total": {区分: 行}}
    """
    cols = {}
    for c in range(HOLD_FIRST_DATA_COL, ws.max_column + 1):
        v = ws.cell(HOLD_DATE_ROW, c).value
        if isinstance(v, datetime):
            cols[v.strftime("%Y-%m")] = c

    rows, subtotals, total = {}, [], {}
    cur_name, cur_norm, cur = None, None, {}
    pending, groups = [], []

    def _flush():
        if cur_name is None:
            return
        if cur_name == "合計":
            total.update(cur)
        elif cur_name in ("本社計", "北九州計"):
            subtotals.append((cur_name, dict(cur), list(pending)))
            pending.clear()
        else:
            for k, r in cur.items():
                rows[(cur_norm, k)] = r
            pending.append(cur_norm)
            groups.append(cur_norm)

    for r in range(1, ws.max_row + 1):
        name = ws.cell(r, HOLD_BRANCH_COL).value
        kind = ws.cell(r, HOLD_KIND_COL).value
        if name:
            _flush()
            cur_name = str(name)
            cur_norm = normalize_branch(cur_name)
            cur = {}
        if kind and cur_name is not None:
            cur[str(kind).strip()] = r
    _flush()

    return {"rows": rows, "cols": cols, "subtotals": subtotals,
            "total": total, "branches": groups}


def write_holding(ws, hold: dict, agg3: dict, months: set) -> int:
    """財源保有状況シートに A / BC / D を書き込む。

    months は「CSV に実在した月」。ここに無い月の列は触らない。
    (今日の CSV に9月案件が無いのに9月列を 0 で潰す事故を防ぐ)
    """
    written = 0
    for mk, col in sorted(hold["cols"].items()):
        if mk not in months:
            print(f"[STEP6] 財源保有状況 {mk} 列: CSVに当月のデータが無いため触りません。")
            continue

        vals = {}
        for (branch, kind), row in hold["rows"].items():
            v = round(agg3.get((branch, mk, kind), 0.0) / 1000, 3)
            ws.cell(row, col).value = v
            vals[(branch, kind)] = v
            written += 1

        for name, kinds, members in hold["subtotals"]:
            for kind, row in kinds.items():
                ws.cell(row, col).value = round(
                    sum(vals.get((b, kind), 0.0) for b in members), 3)
                written += 1

        for kind, row in hold["total"].items():
            ws.cell(row, col).value = round(
                sum(vals.get((b, kind), 0.0) for b in hold["branches"]), 3)
            written += 1

        print(f"[STEP6] 財源保有状況 {mk} 列 ({CL_(col)}列) を更新しました。")
    return written


def CL_(idx: int) -> str:
    """列番号 → 列記号 (openpyxl の get_column_letter 相当の簡易版)。"""
    s = ""
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


TARGET_COL = 3       # C列: 目標 (A～C財源の行にだけ入る。小計・合計は数式)
GETSUSHO_COL = 5     # E列: 月初 (A～C財源・D財源の両方に入る)


def warn_missing_manual_inputs(ws, tables: list) -> int:
    """目標(C列)・月初(E列) の空欄を警告する。書き込みは止めない。

    この 2 つはお客様が手で入れる項目で、CSV にもシートのどこにも根拠が無い。
    月ごとに新しいファイルへ切り替える運用なので、**新しい月のファイルで
    入れ忘れが起きやすい**。空のままでも当日列への書き込みは成功するため、
    気づかないまま「要増加」「目標比」「月初比」が計算できない状態で
    運用が進んでしまう。それを防ぐために名指しで出す。
    """
    total = 0
    for t in tables:
        if t["today_col"] is None:
            continue

        no_target, no_getsusho = [], []
        for branch, (abc_row, d_row) in t["branch_rows"].items():
            v = ws.cell(abc_row, TARGET_COL).value
            if v is None or (isinstance(v, str) and not v.strip()):
                no_target.append(branch)
            if all(ws.cell(r, GETSUSHO_COL).value is None for r in (abc_row, d_row)):
                no_getsusho.append(branch)

        n = len(t["branch_rows"])
        if no_target:
            print(f"[警告] {t['month_key']} の表: 目標(C列)が空の部店が "
                  f"{len(no_target)}/{n} 件あります {sorted(no_target)}。"
                  f"「要増加」「目標比」が計算できません。"
                  f"当日列への書き込みは行いますが、Excel側の入力をご確認ください。",
                  file=sys.stderr)
            total += 1
        if no_getsusho:
            print(f"[警告] {t['month_key']} の表: 月初(E列)が空の部店が "
                  f"{len(no_getsusho)}/{n} 件あります {sorted(no_getsusho)}。"
                  f"「月初比」が計算できず、月内で初回の実行では比較相手も"
                  f"無くなります。",
                  file=sys.stderr)
            total += 1
    return total


# ========== STEP4: 書込 ==========
def write_values(ws, tables: list, agg: dict, today: date | None = None) -> tuple[int, list]:
    written = 0
    errors = []
    excel_branches = set()
    for t in tables:
        excel_branches.update(t["branch_rows"].keys())
    csv_branches = {k[0] for k in agg.keys()}
    orphan = csv_branches - excel_branches
    if orphan:
        # v8: 未マッピング部店を強調表示 (数字が集計から漏れている可能性)
        print(f"[警告] Excelに対応行が無いCSV部店 ({len(orphan)}件): {sorted(orphan)}", file=sys.stderr)
        print(f"[警告] 上記部店の検収額は集計に反映されていません。normalize_branch のエイリアス追加を検討してください。", file=sys.stderr)
    # v8: Excelにあるが今日CSVに無い部店 (情報レベル: 単に該当月案件がゼロの可能性もある)
    unused_excel = excel_branches - csv_branches
    if unused_excel:
        print(f"[情報] CSVに出現しなかったExcel部店 ({len(unused_excel)}件): {sorted(unused_excel)}")

    # 全表で一致列が無かった場合の事後判定用
    tables_with_today = [t for t in tables if t["today_col"] is not None]

    for t in tables:
        if t["today_col"] is None:
            # v6: 一致する日付列が無い表はスキップ (エラーではなく情報)
            # v9.1: "列が無いため明示できていない" 事実を伝わる文言に強化
            print(
                f"[STEP4] 表 {t['month_key']}: シート内に該当日付の列が見つかりません。"
                f"当該表は雛形に該当日付列が存在しないため、書き込みをスキップします。"
            )
            continue
        branch_raw = {}
        for branch, (abc_row, d_row) in t["branch_rows"].items():
            abc_val = sum(agg.get((branch, mk, "ABC"), 0) for mk in t["month_keys"])
            d_val = sum(agg.get((branch, mk, "D"), 0) for mk in t["month_keys"])
            # 原本は 検収額(円)/1000 を小数のまま保持しているため丸めない。
            #      検収額は整数(円)なので小数第3位までで割り切れる。
            abc_i = round(abc_val / 1000, 3)
            d_i = round(d_val / 1000, 3)
            ws.cell(abc_row, t["today_col"]).value = abc_i
            ws.cell(d_row, t["today_col"]).value = d_i
            branch_raw[branch] = (abc_val, d_val)
            written += 2

        for name, abc_row, d_row, members in t["subtotal_rows"]:
            abc_sum_yen = sum(branch_raw.get(b, (0, 0))[0] for b in members)
            d_sum_yen = sum(branch_raw.get(b, (0, 0))[1] for b in members)
            ws.cell(abc_row, t["today_col"]).value = round(abc_sum_yen / 1000, 3)
            ws.cell(d_row, t["today_col"]).value = round(d_sum_yen / 1000, 3)    
            written += 2

        if t["grand_total_row"]:
            abc_row, d_row = t["grand_total_row"]
            abc_sum_yen = sum(branch_raw.get(b, (0, 0))[0] for b in t["all_branches"])
            d_sum_yen = sum(branch_raw.get(b, (0, 0))[1] for b in t["all_branches"])
            ws.cell(abc_row, t["today_col"]).value = round(abc_sum_yen / 1000, 3)
            ws.cell(d_row, t["today_col"]).value = round(d_sum_yen / 1000, 3)    
            written += 2

    # v9.1: どの表にも該当日付列が無かった場合の全体警告
    if not tables_with_today:
        today_label = today.strftime("%Y-%m-%d") if today else "実行日"
        print(
            f"[警告] どの表にも {today_label} の列が見つかりませんでした。"
            f"雛形の日付列を確認してください。書き込みは 0 件で終了します。",
            file=sys.stderr,
        )

    return written, errors


# ========== STEP5: 前日比較 & 色付け ==========
def highlight_decreased(ws, tables: list) -> dict:
    stats = {"executed": 0, "colored": 0, "skipped_no_prev": 0, "skipped_empty": 0,
             "targets": []}
    for t in tables:
        if t["today_col"] is None:
            continue
        abc_targets = [(branch, abc_row) for branch, (abc_row, _) in t["branch_rows"].items()]
        for name, abc_row, _, _ in t["subtotal_rows"]:
            abc_targets.append((name, abc_row))
        if t["grand_total_row"]:
            abc_targets.append(("合計", t["grand_total_row"][0]))

        for branch, abc_row in abc_targets:
            today_val = ws.cell(abc_row, t["today_col"]).value
            if t["prev_col"] is None:
                stats["skipped_no_prev"] += 1
                continue
            prev_val = ws.cell(abc_row, t["prev_col"]).value
            if not isinstance(prev_val, (int, float)) or not isinstance(today_val, (int, float)):
                stats["skipped_empty"] += 1
                continue
            stats["executed"] += 1
            # round() は銀行家丸め (round(450.5)=450 / round(451.5)=452) で
            #      基準が揺れるため使わない。仕様どおり実値で「前日以下」を判定する。
            #      EPS は浮動小数の誤差吸収のみが目的 (千円単位で 1e-9 = 0.000001円)。
            if today_val <= prev_val + 1e-9:
                ws.cell(abc_row, t["today_col"]).fill = YELLOW
                stats["colored"] += 1
                stats["targets"].append(f"{t['month_key']}/{branch}")
    return stats


# ========== サマリ (変更なし) ==========
def build_summary_for_table(ws, target: dict, today: date) -> dict:
    if target is None or target["today_col"] is None:
        return {"available": False, "reason": "当日列なし",
                "month_key": target["month_key"] if target else None,
                "today": today}

    today_col = target["today_col"]
    prev_col = target["prev_col"]

    def _num(r, c):
        v = ws.cell(r, c).value if (r and c) else None
        return v if isinstance(v, (int, float)) else 0

    abc_total = d_total = abc_prev = d_prev = 0
    if target["grand_total_row"]:
        gt_abc, gt_d = target["grand_total_row"]
        abc_total = _num(gt_abc, today_col)
        d_total = _num(gt_d, today_col)
        abc_prev = _num(gt_abc, prev_col)
        d_prev = _num(gt_d, prev_col)

    branch_diffs = []
    for branch, (abc_row, d_row) in target["branch_rows"].items():
        abc_now = _num(abc_row, today_col)
        abc_bef = _num(abc_row, prev_col)
        d_now = _num(d_row, today_col)
        d_bef = _num(d_row, prev_col)
        abc_diff = abc_now - abc_bef      # 丸めてから引かない
        d_diff = d_now - d_bef            # 表示側 (_fmt_signed) で丸める
        branch_diffs.append({
            "branch": branch,
            "abc_now": abc_now, "abc_diff": abc_diff,
            "d_now": d_now, "d_diff": d_diff,
        })

    abc_diff_total = abc_total - abc_prev   # 丸めてから引かない
    d_diff_total = d_total - d_prev       
    abc_rate = (abc_diff_total / abc_prev * 100) if abc_prev else None
    d_rate = (d_diff_total / d_prev * 100) if d_prev else None

    return {
        "available": True,
        "month_key": target["month_key"],
        "abc_total": abc_total, "abc_prev": abc_prev, "abc_diff": abc_diff_total, "abc_rate": abc_rate,
        "d_total": d_total, "d_prev": d_prev, "d_diff": d_diff_total, "d_rate": d_rate,
        "branch_diffs": branch_diffs,
        "today": today,
    }


def _fmt_signed(v):
    if v is None:
        return "-"
    sign = "＋" if v > 0 else ("△" if v < 0 else "±")
    return f"{sign}{abs(int(round(v))):,}"


def build_summaries(ws, tables: list, today: date) -> list:
    return [build_summary_for_table(ws, t, today) for t in tables]


def print_summary(s: dict) -> None:
    print("\n" + "=" * 60)
    print(f"[集計サマリ] 対象月: {s.get('month_key')}")
    print("=" * 60)
    if not s.get("available"):
        print(f"  集計不可: {s.get('reason')}")
        return

    print(f"■ 対象月: {s['month_key']}  (実行日 {s['today']})")
    print("\n■ 全体概要（千円単位）")
    abc_rate_s = f"  ({s['abc_rate']:+.2f}%)" if s['abc_rate'] is not None else ""
    d_rate_s = f"  ({s['d_rate']:+.2f}%)" if s['d_rate'] is not None else ""
    abc_total_s = f"{int(round(s.get('abc_total') or 0)):,}"
    d_total_s = f"{int(round(s.get('d_total') or 0)):,}"
    print(f"  ABC 検収額合計: {abc_total_s:>12s} 千円  前日比 {_fmt_signed(s['abc_diff'])} 千円{abc_rate_s}")
    print(f"  D   検収額合計: {d_total_s:>12s} 千円  前日比 {_fmt_signed(s['d_diff'])} 千円{d_rate_s}")

    diffs = s["branch_diffs"]
    increased = [b for b in diffs if (b["abc_diff"] or 0) > 0 or (b["d_diff"] or 0) > 0]
    decreased = [b for b in diffs if (b["abc_diff"] or 0) < 0 or (b["d_diff"] or 0) < 0]

    print("\n■ 部店別の増減")
    if not diffs:
        print("  （変動なし）")
    for b in sorted(diffs, key=lambda x: -abs((x["abc_diff"] or 0) + (x["d_diff"] or 0))):
        tot = (b["abc_diff"] or 0) + (b["d_diff"] or 0)
        note = "増加" if tot > 0 else "減少" if tot < 0 else "横ばい"
        print(f"  {b['branch']:<10s} ABC:{_fmt_signed(b['abc_diff']):>8s} 千円 / "
              f"D:{_fmt_signed(b['d_diff']):>8s} 千円  [{note}]")

    print("\n■ 特記事項")
    if decreased:
        print(f"  減少部店: {'、'.join(b['branch'] for b in decreased)}")
    if increased:
        top = sorted(increased, key=lambda x: -((x['abc_diff'] or 0) + (x['d_diff'] or 0)))[:3]
        parts = [f"{b['branch']}({_fmt_signed((b['abc_diff'] or 0)+(b['d_diff'] or 0))}千円)" for b in top]
        print(f"  増加額 上位: {'、'.join(parts)}")


def print_mail_body(summaries: list) -> None:
    valids = [s for s in summaries if s.get("available")]
    print("\n" + "-" * 60)
    print("[メール本文テンプレート]")
    print("-" * 60)
    if not valids:
        print("集計可能な月がありません")
        return

    today = valids[0]["today"]
    ymd = today.strftime("%Y/%m/%d")
    print(f"件名：財源状況集計表 更新結果（{ymd}）\n")
    print("本文：")
    print("本日の財源状況集計表を更新しました。\n")

    for s in valids:
        abc_total_s = f"{int(round(s.get('abc_total') or 0)):,}"
        d_total_s = f"{int(round(s.get('d_total') or 0)):,}"
        diffs = s["branch_diffs"]
        decreased = [b for b in diffs if ((b["abc_diff"] or 0) + (b["d_diff"] or 0)) < 0]

        print(f"■{s['month_key']} 集計結果")
        print(f"・ABC 検収額合計：{abc_total_s} 千円（前日比：{_fmt_signed(s['abc_diff'])} 千円）")
        print(f"・D   検収額合計：{d_total_s} 千円（前日比：{_fmt_signed(s['d_diff'])} 千円）")
        print(f"\n■{s['month_key']} 主な増減")
        printed = False
        for b in sorted(diffs, key=lambda x: -abs((x["abc_diff"] or 0) + (x["d_diff"] or 0)))[:5]:
            if b["abc_diff"]:
                print(f"・{b['branch']}　ABC：{_fmt_signed(b['abc_diff'])} 千円")
                printed = True
            if b["d_diff"]:
                print(f"・{b['branch']}　D：{_fmt_signed(b['d_diff'])} 千円")
                printed = True
        if not printed:
            print("・前日から変動なし")
        if decreased:
            print(f"\n■{s['month_key']} 注意事項")
            print(f"・前日以下となった部店：{'、'.join(b['branch'] for b in decreased)}")
        print()

    print("以上、ご確認をお願いいたします。")


# ========== メイン ==========
def main(today_str: str | None = None):
    if today_str:
        today = datetime.strptime(today_str, "%Y-%m-%d").date()
    else:
        today = date.today()
    print(f"[開始] 実行日 = {today}")

    Path("output").mkdir(exist_ok=True)

    # v7: 入力ファイルを「部分一致 + 中身判定」で解決 (案F)
    input_csv = find_input_file(CSV_PATTERNS, "CSV", _looks_like_target_csv)
    input_xlsx = find_input_file(XLSX_PATTERNS, "Excel", _looks_like_target_xlsx)

    df = load_and_clean_csv(input_csv)
    agg = aggregate(df)
    agg3 = aggregate3(df)
    csv_months = set(df["月キー"].unique())

    wb = load_workbook(input_xlsx)
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"シート '{SHEET_NAME}' が見つかりません / 実際={wb.sheetnames}")
    ws = wb[SHEET_NAME]
    tables = detect_tables(ws, today)
    print(f"[STEP3] 検出表数: {len(tables)}")
    for t in tables:
        print(f"        - {t['month_key']}: 部店{len(t['branch_rows'])} "
              f"当日列={_col_label(ws, t['header_row'], t['today_col'])} "
              f"比較列={_col_label(ws, t['header_row'], t['prev_col'])}")

    # v6: 当月表 (実行日と同じ月) に当日列が無ければエラー
    # v9.1: "列が無いため明示できていない" 文脈を伝える文言に統一
    current_month_key = today.strftime("%Y-%m")
    # 実行日の月を集計対象に含む表を「当月表」とする
    current = next((t for t in tables if current_month_key in t["month_keys"]), None)
    if current is None:
        raise ValueError(
            f"実行日 {today} と同じ月 ({current_month_key}) の表がシートに存在しません。"
            f"雛形に該当月の表が用意されていないため、書き込み先が明示されておらず処理を中止します。"
        )
    if current["today_col"] is None:
        raise ValueError(
            f"当月表 ({current_month_key}) のヘッダに実行日 {today} の列が見つかりません。"
            f"雛形に該当日付列が存在しないため、書き込み先が明示されておらず処理を中止します。"
        )

    # CSV に当月ぶんが 1 件も無ければ、別の期の CSV を渡された可能性が高い。
    # そのまま進めると 96 セルすべてに 0 を書き、「財源ゼロ」という
    # 事実と違う数字がシートに残る（実際に 2026-09-10 の実行で発生）。
    # 「その月に案件が無い」のか「CSV がその月を含んでいない」のかは
    # 区別できないため、当月が入っていないなら書かずに止める。
    if current_month_key not in csv_months:
        raise ValueError(
            f"CSV に当月 ({current_month_key}) のデータが 1 件もありません。"
            f"CSV に含まれる月: {sorted(csv_months)}。"
            f"別の期のCSVが渡されている可能性が高いため、処理を中止します。"
            f"（このまま進めると全部店に 0 を書き込み、"
            f"「財源ゼロ」という誤った数字が残ります）"
        )

    # 当月は入っているので CSV の断面自体は正しいと判断する。
    # 将来月がCSVに無いのは「まだ案件が無い」という意味なので 0 を書く。
    for t in tables:
        if t["today_col"] is None:
            continue
        for mk in t["month_keys"]:
            if mk not in csv_months:
                print(f"[情報] {t['month_key']} の表: CSVに {mk} の案件がありません。"
                      f"0 として書き込みます（まだ受注が無い月という扱い）。")

    warn_missing_manual_inputs(ws, tables)

    written, errors = write_values(ws, tables, agg, today)
    expected = sum(
        (len(t["branch_rows"]) + len(t["subtotal_rows"]) + (1 if t["grand_total_row"] else 0)) * 2
        for t in tables if t["today_col"]
    )
    print(f"[STEP4] 書込セル数: {written} (想定 {expected})")
    if errors:
        for e in errors:
            print(f"        [エラー] {e}")

    # 値を書いた直後・色付け前のスナップショット。開発時の切り分け用。
    # 通常運用では出さない (output/ に xlsx が 2 つ並び、お客様にそのまま見せる
    # 処理ログにも debug の行が出てしまうため)。
    # 必要なときだけ環境変数 ZAIGEN_DEBUG=1 を付けて実行する。
    if os.environ.get("ZAIGEN_DEBUG"):
        wb.save(DEBUG_XLSX)
        print(f"[保存] {DEBUG_XLSX} (ZAIGEN_DEBUG 指定時のみ)")

    stats = highlight_decreased(ws, tables)
    print(f"[STEP5] 前日比較実行={stats['executed']} 色付け={stats['colored']} "
          f"前日列なしスキップ={stats['skipped_no_prev']} 空欄スキップ={stats['skipped_empty']}")
    if stats["targets"]:
        print(f"        色付け対象: {stats['targets']}")

    # 財源保有状況シート (実物Excelにのみ存在)。無ければ何もしない。
    if HOLDING_SHEET in wb.sheetnames:
        hold = detect_holding(wb[HOLDING_SHEET])
        if hold["rows"] and hold["cols"]:
            n = write_holding(wb[HOLDING_SHEET], hold, agg3, csv_months)
            print(f"[STEP6] 財源保有状況 書込セル数: {n}"
                  f" (部店{len(hold['branches'])} × 区分3 + 小計・合計)")
        else:
            print(f"[STEP6] {HOLDING_SHEET} の構造を読めませんでした。"
                  f"このシートは更新していません。")
    else:
        print(f"[STEP6] {HOLDING_SHEET} シートがありません。集計表のみ更新しました。")

    wb.save(OUTPUT_XLSX)
    print(f"[保存] {OUTPUT_XLSX}")

    summaries = build_summaries(ws, tables, today)
    for s in summaries:
        print_summary(s)
    print_mail_body(summaries)

    ok = (written == expected) and not errors
    print(f"[結果] {'OK ✓' if ok else 'NG ✗'}  書込={written}/{expected}  エラー={len(errors)}")
    return 0 if ok else 1


if __name__ == "__main__":
    today_arg = sys.argv[1] if len(sys.argv) > 1 else None
    sys.exit(main(today_arg))
