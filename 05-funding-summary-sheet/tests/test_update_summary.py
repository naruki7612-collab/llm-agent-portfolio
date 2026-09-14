"""update_zaigen.py の受け入れ検証。

前提（2026-09-08 に先方から確定した運用）
  ・「集計表」シートに同じ形の表が 3 つ縦に並ぶ
  ・3 表は上から **当月 / 翌月 / 翌々月**（8月ファイル = 8月・9月・10月）
  ・**日付列は 3 表とも共通（＝実行日）**。1 回の実行で 3 表すべての同じ列に書く
  ・**ファイルは月ごとに新しく切り替える**（年を跨いでも同じ）
  ・お客様の原本を**無加工**で受け取れる
      - 月ラベルは文字列（'8月' / '53期下期'）
      - 2つ目・3つ目の日付ヘッダは '=F4' 等の参照

受け入れ条件
  ① 無加工の原本を対象ファイルとして認識する
  ② 3 表すべての当日列に書き込む（32セル × 3表 = 96セル）
  ③ ラベルを書き換えない（'53期下期' が残る）
  ④ 集計対象の月は**表の位置**で決める（ラベルが実態と違っても正しく割り当てる）
  ⑤ 前日以下のセルが黄色になる
  ⑥ 書込先以外は 1 セルも変わらない
  ⑦ 財源保有状況シートが同期される（CSVに無い月の列は触らない）
  ⑧ 該当日付列が無い日は安全に停止する
  ⑨ 月として読めないラベルの表は、区切りとして扱い書き込まない
  ⑩ 実行が飛んでも「直近で値が入っている列」と比較して色付けする
  ⑪ CSV に当月ぶんが 1 件も無ければ、0 を書かずに停止する

注意: ローカルの Python は 3.9 で `int | None` 記法が通らないため、
      テスト時のみ `from __future__ import annotations` を挿入したコピーを実行する。
      update_zaigen.py 本体は一切変更しない（実行環境 側は 3.11+ で素のまま動く）。
"""

import csv
import io
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter as CL

HERE = Path(__file__).parent
SCRIPT = HERE / "update_zaigen.py"
# お客様の原本（無加工）
XLSX = HERE / "入力_実物" / "財源集計表_v1.xlsx"
# make_month_file.py で作った 9月ファイル（目標・月初は空）
XLSX_SEPT = HERE / "財源集計表_2026年9月.xlsx"
# 本番CSV（2026-08-25 に先方が使用。774件・全月ぶん・BOMが2つ）
CSV_REAL = HERE / "入力_実物" / "月度リース割賦情報管理表 (1).csv"
# 共有ドライブ版。時期が 2026/08 だけの部分集合
CSV_AUG_ONLY = HERE / "入力_実物" / "月度リース割賦情報管理表.csv"

YELLOW = "00FFFF00"          # openpyxl は 6桁指定を "00" 詰めで保持する
RUN = "2026-08-25"
TABLE_TOPS = (5, 40, 75)     # 各表の 1 部店目の行
EXPECT_MONTHS = ("2026-08", "2026-09", "2026-10")   # 8月ファイルを 8/25 に実行

EXCLUDE = {"競合負け", "自主謝絶", "否決", "没"}
_STRIP = str.maketrans("", "", " 　\t\r\n-_./\\・")


def check(label: str, ok: bool, detail: str = "") -> bool:
    print(f"  [{'OK ' if ok else 'NG '}] {label}" + (f" — {detail}" if detail else ""))
    return bool(ok)


def norm_branch(name: str) -> str:
    s = unicodedata.normalize("NFKC", str(name)).translate(_STRIP)
    s = re.sub(r"(支店|チーム|営業)", "", s).lower().strip()
    return {"本社部": "本社店口"}.get(s, s)


def csv_agg(path: Path) -> dict:
    """CSV から (部店, 月キー, 区分) → 千円 を作る。スクリプトと同じ条件。"""
    b = path.read_bytes()
    while b.startswith(b"\xef\xbb\xbf"):
        b = b[3:]
    agg = defaultdict(float)
    for r in csv.DictReader(io.StringIO(b.decode("cp932"))):
        if r["進捗"] in EXCLUDE:
            continue
        k = r["確度"].strip().upper()
        if k not in ("A", "B", "C", "D"):
            continue
        t = r["時期"].replace("/", "-")
        dt = None
        for f in ("%Y-%m-%d", "%Y-%m"):
            try:
                dt = datetime.strptime(t, f)
                break
            except ValueError:
                pass
        if dt is None:
            continue
        kind = "ABC" if k != "D" else "D"
        agg[(norm_branch(r["部店"]), dt.strftime("%Y-%m"), kind)] += float(r["検収額"] or 0)
    return agg


def make_runnable(dst: Path) -> Path:
    lines = SCRIPT.read_text().split("\n")
    end = next(i for i in range(1, len(lines)) if lines[i].rstrip() == '"""')
    lines.insert(end + 1, "from __future__ import annotations")
    out = dst / "run_local.py"
    out.write_text("\n".join(lines))
    return out


def run_case(name: str, xlsx: Path, csv_path: Path, run_date: str) -> dict:
    work = Path(tempfile.mkdtemp(prefix="zaigen_"))
    (work / "uploads").mkdir()
    (work / "output").mkdir()
    shutil.copy(xlsx, work / "uploads" / xlsx.name)
    shutil.copy(csv_path, work / "uploads" / csv_path.name)
    proc = subprocess.run([sys.executable, str(make_runnable(work)), run_date],
                          cwd=work, capture_output=True, text=True)
    log = proc.stdout + proc.stderr
    print(f"\n{'=' * 72}\n【{name}】 実行日={run_date}  終了コード={proc.returncode}\n{'=' * 72}")
    for line in log.split("\n"):
        if any(k in line for k in ("[入力]", "[STEP", "[保存]", "[警告]", "[情報]", "Error")):
            print("  " + line.rstrip())
    return {"rc": proc.returncode, "log": log, "work": work,
            "produced": work / "output" / "財源状況集計表_更新版.xlsx"}


def find_col(xlsx: Path, header_row: int, date_str: str):
    """日付ヘッダの列を探す。'=F4' のような参照は辿る。"""
    y, m, d = map(int, date_str.split("-"))
    ws = openpyxl.load_workbook(xlsx)["集計表"]
    for c in range(5, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        for _ in range(3):
            if isinstance(v, str) and v.startswith("="):
                mm = re.match(r"^=\s*\$?([A-Z]{1,2})\$?(\d+)\s*$", v)
                v = ws[f"{mm.group(1)}{mm.group(2)}"].value if mm else None
            else:
                break
        if isinstance(v, datetime) and (v.year, v.month, v.day) == (y, m, d):
            return c
    return None


def diff_except(src: Path, produced: Path, allowed: set) -> list:
    a, b = openpyxl.load_workbook(src), openpyxl.load_workbook(produced)
    if a.sheetnames != b.sheetnames:
        return [f"シート構成が変化: {a.sheetnames} → {b.sheetnames}"]
    out = []
    for sn in a.sheetnames:
        wa, wb_ = a[sn], b[sn]
        for r in range(1, max(wa.max_row, wb_.max_row) + 1):
            for c in range(1, max(wa.max_column, wb_.max_column) + 1):
                if (sn, r, c) in allowed:
                    continue
                if wa.cell(r, c).value != wb_.cell(r, c).value:
                    out.append(f"{sn}!{CL(c)}{r}")
    return out


def main() -> int:
    ok = True
    branches = ["本社1", "本社2", "輸送機", "本社店口", "本社計", "北九州1", "北九州2",
                "北九州計", "久留米", "熊本", "大分", "長崎", "東京", "大阪",
                "kli不動産", "合計"]

    # ---------- ケース1: 無加工の原本に 3 表とも書く ----------
    r = run_case("ケース1 無加工の原本に3表とも書く", XLSX, CSV_REAL, RUN)
    print()
    ok &= check("終了コード 0", r["rc"] == 0)
    ok &= check("無加工の原本を対象ファイルとして認識する",
                "Excel = uploads/財源集計表_v1.xlsx (中身判定で確定)" in r["log"])
    ok &= check("BOMが2つ付いたCSVを読める", "BOM=あり×2 / body=cp932" in r["log"])
    ok &= check("3表すべて検出", "[STEP3] 検出表数: 3" in r["log"])
    ok &= check("ラベルが読めない表も位置で月を割り当てる",
                "'53期下期' からは月を読めないため照合を省略します（2026-10 として集計します）"
                in r["log"])

    cols = re.findall(r"- 2026-\d\d: 部店13 当日列=(\d+)", r["log"])
    ok &= check("3表とも同じ当日列を指す",
                len(cols) == 3 and len(set(cols)) == 1,
                f"当日列={set(cols)}" if cols else "検出できず")
    ok &= check("書込セル数 96 (32セル × 3表)",
                "[STEP4] 書込セル数: 96 (想定 96)" in r["log"])

    col = find_col(XLSX, 4, RUN)
    out = openpyxl.load_workbook(r["produced"])["集計表"]
    per = [sum(1 for row in range(t, t + 32)
               if isinstance(out.cell(row, col).value, (int, float))) for t in TABLE_TOPS]
    ok &= check(f"{CL(col)}列({RUN})に 3表 × 32セル", per == [32, 32, 32], str(per))

    # ③ ラベルを書き換えていない
    labels = [out.cell(row, 1).value for row in (3, 38, 73)]
    ok &= check("月ラベルを書き換えていない", labels == ["8月", "9月", "53期下期"],
                str(labels))

    # ④ 3表が 当月 / 翌月 / 翌々月 として集計されているか
    ok &= check("3表が 当月/翌月/翌々月 として検出される",
                all(f"- {mk}: 部店13" in r["log"] for mk in EXPECT_MONTHS),
                " / ".join(EXPECT_MONTHS))
    agg = csv_agg(CSV_REAL)
    ng = []
    for t, mk in zip(TABLE_TOPS, EXPECT_MONTHS):
        for i, n in enumerate(branches):
            if n in ("本社計", "北九州計", "合計"):
                continue
            want = round(agg.get((n, mk, "ABC"), 0) / 1000, 3)
            got = out.cell(t + 2 * i, col).value
            if abs(got - want) > 1e-6:
                ng.append(f"{mk}/{n}: {got} != {want}")
    ok &= check("3表とも該当月の単月集計と一致 (13部店 × 3表)", not ng,
                "全一致" if not ng else str(ng[:3]))

    # ⑥ 書込先以外が無傷
    allowed = {("集計表", row, col) for t in TABLE_TOPS for row in range(t, t + 32)}
    allowed |= {("財源保有状況", row, c) for row in range(3, 51) for c in (20, 21)}
    d = diff_except(XLSX, r["produced"], allowed)
    ok &= check("書込先以外の全セルが原本と一致 (2シート全域)", not d,
                "差分なし" if not d else f"{len(d)}件: {d[:3]}")

    # ⑦ 財源保有状況
    hold = openpyxl.load_workbook(r["produced"])["財源保有状況"]
    ng = []
    for i, n in enumerate(branches):
        hr, sr = 3 + 3 * i, 5 + 2 * i
        a, bc, dd = (hold.cell(hr + k, 20).value for k in (0, 1, 2))
        if abs((a + bc) - out.cell(sr, col).value) > 1e-6:
            ng.append(f"{n}:A+BC")
        if abs(dd - out.cell(sr + 1, col).value) > 1e-6:
            ng.append(f"{n}:D")
    ok &= check("財源保有状況の A+BC / D が8月表の当日列と一致 (16部店)", not ng,
                "全一致" if not ng else str(ng[:3]))
    ok &= check("財源保有状況に96セル書き込まれた",
                "[STEP6] 財源保有状況 書込セル数: 96" in r["log"])
    ok &= check("サマリ・メール本文を出力",
                "[集計サマリ]" in r["log"] and "[メール本文テンプレート]" in r["log"])

    # ---------- ケース2: 前日以下の色付け ----------
    a = run_case("ケース2-1 前日ぶんを作る (8/24)", XLSX, CSV_REAL, "2026-08-24")
    chained = a["work"] / "chained.xlsx"
    shutil.copy(a["produced"], chained)
    b = run_case("ケース2-2 前日以下の色付け (8/25)", chained, CSV_REAL, RUN)
    print()
    ws2 = openpyxl.load_workbook(b["produced"])["集計表"]
    yellow = {t: [row for row in range(t, t + 32)
                  if ws2.cell(row, col).fill
                  and ws2.cell(row, col).fill.start_color.rgb == YELLOW]
              for t in TABLE_TOPS}
    ok &= check("前日比較が3表ぶん実行された（48件）",
                "[STEP5] 前日比較実行=48" in b["log"])
    ok &= check("比較列が直近の値ありの列（8/24）になる",
                "比較列=20列(8/24)" in b["log"])
    ok &= check("空欄スキップが発生しない", "空欄スキップ=0" in b["log"])
    ok &= check("色付けが発生した", sum(len(v) for v in yellow.values()) > 0,
                " / ".join(f"表{t}:{len(v)}件" for t, v in yellow.items()))
    src = openpyxl.load_workbook(XLSX)["集計表"]
    ok &= check("色付けはA~C財源の行だけ",
                all(src.cell(row, 4).value == "A～C財源"
                    for v in yellow.values() for row in v))

    # ---------- ケース3: CSVに該当月が無い表 ----------
    c = run_case("ケース3 8月のみのCSV", XLSX, CSV_AUG_ONLY, RUN)
    print()
    ok &= check("終了コード 0", c["rc"] == 0)
    ok &= check("3表とも書き込まれる (96セル)",
                "[STEP4] 書込セル数: 96 (想定 96)" in c["log"])
    ws3 = openpyxl.load_workbook(c["produced"])["集計表"]
    ok &= check("8月表には値が入る", ws3.cell(5, col).value not in (None, 0),
                f"本社1 A~C={ws3.cell(5, col).value}")
    ok &= check("翌月表・翌々月表は全て 0（CSVに該当月が無い）",
                all(ws3.cell(row, col).value == 0
                    for t in (40, 75) for row in range(t, t + 32)))
    ok &= check("財源保有状況は9月列を触らない",
                "財源保有状況 2026-09 列: CSVに当月のデータが無いため触りません"
                in c["log"])

    # ---------- ケース3.5: 実行が飛んだ日の比較 ----------
    # 日付を指定して更新する運用（毎営業日連続とは限らない）に合わせ、
    # 比較対象は「当日列の左で、実際に値が入っている最も近い列」にしている。
    # 以前は日付ヘッダの左隣を無条件に選んでいたため、1 日飛ぶと
    # その列が空で「空欄スキップ」になり色付けが効かなかった。
    f1 = run_case("ケース3.5-1 8/24を入れる", XLSX, CSV_REAL, "2026-08-24")
    chain2 = f1["work"] / "c2.xlsx"
    shutil.copy(f1["produced"], chain2)
    f2 = run_case("ケース3.5-2 8/25を飛ばして8/26", chain2, CSV_REAL, "2026-08-26")
    print()
    ok &= check("8/25を飛ばしても比較列が 8/24 になる",
                "比較列=20列(8/24)" in f2["log"],
                re.search(r"- 2026-08:.*", f2["log"]).group(0)
                if re.search(r"- 2026-08:.*", f2["log"]) else "")
    ok &= check("空欄スキップが発生しない", "空欄スキップ=0" in f2["log"])
    ok &= check("色付けが発生する（従来はここで0件だった）",
                re.search(r"色付け=(\d+)", f2["log"])
                and int(re.search(r"色付け=(\d+)", f2["log"]).group(1)) > 0,
                re.search(r"\[STEP5\].*", f2["log"]).group(0)
                if re.search(r"\[STEP5\].*", f2["log"]) else "")

    # ---------- ケース4: 該当日付列が無い日 ----------
    e = run_case("ケース4 該当列なし (8/11 山の日)", XLSX, CSV_REAL, "2026-08-11")
    print()
    ok &= check("エラー終了する (終了コード 1)", e["rc"] == 1)
    ok &= check("原因を明示して中止する",
                "ヘッダに実行日 2026-08-11 の列が見つかりません" in e["log"])
    ok &= check("更新版ファイルを一切出力しない", not e["produced"].exists())

    # ---------- ケース5: 別の期のCSVを渡したら止まる ----------
    # 9月ファイル + 8月ぶんだけのCSV。当月(2026-09)が CSV に無いので、
    # そのまま進めると 96 セルすべてに 0 を書き「財源ゼロ」という
    # 事実と違う数字が残る（2026-09-10 の実行で実際に発生した）。
    if XLSX_SEPT.exists():
        g = run_case("ケース5 9月ファイル + 8月のみCSV", XLSX_SEPT, CSV_AUG_ONLY,
                     "2026-09-10")
        print()
        ok &= check("エラー終了する (終了コード 1)", g["rc"] == 1)
        ok &= check("CSVに含まれる月を挙げて中止する",
                    "CSV に当月 (2026-09) のデータが 1 件もありません" in g["log"]
                    and "['2026-08']" in g["log"])
        ok &= check("0を1セルも書かない", not g["produced"].exists())

        # 正しいCSVなら通ること
        h = run_case("ケース5-2 9月ファイル + 全月ぶんCSV", XLSX_SEPT, CSV_REAL,
                     "2026-09-10")
        print()
        ok &= check("終了コード 0", h["rc"] == 0)
        ok &= check("3表とも96セル書き込む",
                    "[STEP4] 書込セル数: 96 (想定 96)" in h["log"])
        ok &= check("当月列は9/10、比較列は月初",
                    "当日列=13列(9/10) 比較列=5列(月初)" in h["log"])
    else:
        print(f"\n（{XLSX_SEPT.name} が無いためケース5をスキップ）")

    print(f"\n{'=' * 72}")
    print("総合判定:", "全項目 OK" if ok else "NG あり")
    print("=" * 72)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
