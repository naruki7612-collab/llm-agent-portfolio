# -*- coding: utf-8 -*-
"""selftest — AgentPlatform でも普通の Python でも動く自己テスト。AI も db も使わない。

使い方（AgentPlatform）:
    import sys; sys.path.insert(0, "notes/システム")
    import selftest; selftest.run()

使い方（手元）:
    python3 selftest.py

機械で決めている部分（手書きコードの変換 / 事業所→部門 / 管理事業の切替 / 源泉の4月切替 /
伝票の括り / 検算）が壊れていないかを確かめる。rules.py を直したら必ず流すこと。

**最初に全ファイルの静的チェックをする。** ここを飛ばすと、import 漏れのような
実行するまで分からない誤りが本番の Step 6 で初めて出る（実際に起きた）。
"""
import sys, os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import rules as R
# group_file は写しを持たず **本体を import する**。写しにすると本体を直したときに
# こちらが古いまま通り、テストが通るのに本番が壊れる（実際に起きた）。
# s30_vouchers.py は agent_sdk を使わないので、手元の Python でも import できる。
from s30_vouchers import group_file

CODE_RULE = [
    {"段": "2", "入力": "624", "出力項目": "CDJS301 購入勘定科目", "出力値": "425000"},
    {"段": "2", "入力": "639", "出力項目": "CDJS301 購入勘定科目", "出力値": "424600"},
    {"段": "3", "入力": "18", "出力項目": "CDJS302 補助科目", "出力値": "0000"},
    {"段": "3", "入力": "11", "出力項目": "CDJS302 補助科目", "出力値": "6391"},
]
ACCT_MGMT = [
    {"事業側_科目": "425200", "事業側_補助": "0000", "管理側_科目": "451015", "管理側_補助": "0000"},
    {"事業側_科目": "427000", "事業側_補助": "0000", "管理側_科目": "451175", "管理側_補助": "0000"},
]

_ng = []


# ---------------------------------------------------------------- テスト本体
def _amount(group):
    """群の金額（bodies の合計）。金額の二重計上を検出するために使う。"""
    return sum(R.doc_amount(b)[0] for b in group["bodies"])


def eq(got, want, label):
    if got == want:
        print("  ok   %s" % label)
    else:
        print("  NG   %s: %r ≠ %r" % (label, got, want))
        _ng.append(label)


def static_check():
    """全ファイルの構文と未定義の名前を調べる。指摘の一覧を返す（空なら合格）。

    import 漏れは compile() では見つからない。その関数が呼ばれた瞬間に NameError に
    なるので、本番だと Step 6 まで進んでから落ちる。pyflakes があればそれで見つける。
    """
    import glob as _g, io, py_compile, tempfile
    here = os.path.dirname(os.path.abspath(__file__))
    files = sorted(_g.glob(os.path.join(here, "*.py")))
    files = [f for f in files if not os.path.basename(f).startswith("test_")]
    ng = []
    with tempfile.TemporaryDirectory() as td:
        for f in files:
            try:
                py_compile.compile(f, cfile=os.path.join(td, "x.pyc"), doraise=True)
            except Exception as e:                        # noqa: BLE001
                ng.append("%s: %s" % (os.path.basename(f), e))
    try:
        from pyflakes.api import checkPath
        from pyflakes.reporter import Reporter
        out = io.StringIO()
        rep = Reporter(out, out)
        for f in files:
            checkPath(f, rep)
        for line in out.getvalue().splitlines():
            if "undefined name" in line:
                ng.append(line.replace(here + os.sep, ""))
        how = "pyflakes"
    except ImportError:
        for f in files:
            ng += _undefined_names(f)
        how = "簡易（pyflakes なし）"
    print("  %d ファイルを検査（構文・未定義の名前 / %s）" % (len(files), how))
    return ng


def _undefined_names(path):
    """pyflakes が無いときの代わり。**どこにも束縛されていない名前**を探す。

    スコープは見ない（別の関数で束縛された名前は見逃す）。狙いは import 漏れで、
    誤検出しないことを優先している。
    """
    import ast, builtins
    with open(path, encoding="utf-8") as f:
        tree = ast.parse(f.read(), path)
    bound = set(dir(builtins)) | {"__file__", "__name__", "__doc__"}
    for n in ast.walk(tree):
        if isinstance(n, ast.Name) and isinstance(n.ctx, (ast.Store, ast.Del)):
            bound.add(n.id)
        elif isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            bound.add(n.name)
        elif isinstance(n, ast.arg):
            bound.add(n.arg)
        elif isinstance(n, (ast.Import, ast.ImportFrom)):
            for al in n.names:
                bound.add((al.asname or al.name).split(".")[0])
        elif isinstance(n, ast.ExceptHandler) and n.name:
            bound.add(n.name)
        elif isinstance(n, ast.Global):
            bound.update(n.names)
    ng = []
    for n in ast.walk(tree):
        if isinstance(n, ast.Name) and isinstance(n.ctx, ast.Load) and n.id not in bound:
            ng.append("%s:%d: 未定義の名前 %r（import 漏れ？）"
                      % (os.path.basename(path), n.lineno, n.id))
    return sorted(set(ng))


def run():
    global _ng
    _ng = []

    print("--- 静的チェック ---")
    for x in static_check():
        _ng.append(x)
        print("  NG   %s" % x)

    print("--- 手書き3段コード ---")
    eq(R.from_code("3-624-18", CODE_RULE),
       {"CDJS303": "349", "CDJS301": "425000", "CDJS302": "0000", "CDJS300": "0007"},
       "3-624-18 → 通運349 / 備消品425000 / 富山検修0007")
    eq(R.from_code("5-639-11", CODE_RULE),
       {"CDJS303": "199", "CDJS301": "424600", "CDJS302": "6391", "CDJS300": "0001"},
       "5-639-11 → 検修199 / 燃料424600 / ガソリン6391 / 本社0001")
    eq(R.from_code("", CODE_RULE), {}, "空のコードは何も返さない")
    eq(R.from_code("よみとれず", CODE_RULE), {}, "読めないコードは何も返さない")

    print("--- 事業所 → 部門 ---")
    eq(R.office_to_dept("富山"), "0004", "富山 → 0004")
    eq(R.office_to_dept("金沢"), "0001", "金沢 → 0003ではなく0001")
    eq(R.office_to_dept("高岡営業所"), "0005", "高岡 → 0005")
    eq(R.office_to_dept("東京"), None, "知らない事業所は None")

    print("--- 経費区分のふるい分け ---")
    eq(R.filter_cost_segment("通運"), "通運", "通運は渡す（的中98%）")
    eq(R.filter_cost_segment("検修"), "検修", "検修は渡す（的中100%）")
    eq(R.filter_cost_segment("管理"), None, "管理は渡さない（的中91%で害が出た）")
    eq(R.filter_cost_segment("事業"), None, "事業は渡さない（的中72%）")

    print("--- 管理／事業の科目切替 ---")
    eq(R.fix_mgmt_business({"CDJS301": "425200", "CDJS302": "0000", "CDJS303": "449"}, ACCT_MGMT)[0]["CDJS301"],
       "451015", "セグ449なら制服は451015へ")
    eq(R.fix_mgmt_business({"CDJS301": "451175", "CDJS302": "0000", "CDJS303": "349"}, ACCT_MGMT)[0]["CDJS301"],
       "427000", "セグ349なら車検整備は427000へ")
    eq(R.fix_mgmt_business({"CDJS301": "425200", "CDJS302": "0000", "CDJS303": "349"}, ACCT_MGMT)[1],
       False, "既に整合していれば触らない")

    print("--- 源泉徴収の4月切替 ---")
    eq(R.doc_amount({"total_incl_tax": 100000, "withholding_tax": 10210, "closing_date": "2026-03-31"})[0],
       110210, "3月まで: 源泉を戻して総額にする")
    eq(R.doc_amount({"total_incl_tax": 100000, "withholding_tax": 10210, "closing_date": "2026-04-30"})[0],
       100000, "4月から: 差引請求額をそのまま使う")
    eq(R.doc_amount({"total_incl_tax": 50000})[0], 50000, "源泉が無ければ総額のまま")

    print("--- 支払先名の正規化 ---")
    eq(R.norm_payee("㈲園田商会"), R.norm_payee("有限会社園田商会"), "㈲ と 有限会社 は同じキーになる")
    eq(R.norm_payee("㈱ﾍﾞｰﾀ商事"), "ベータ商事", "半角カナは全角にそろえる")
    # 異体字。NFKC では揃わないので別会社になり、支払先コードが引けなくなる
    eq(R.norm_payee("株式会社カモメ重車輛工業"), R.norm_payee("カモメ重車輌工業"),
       "輛 と 輌 は同じキーになる")
    eq(R.norm_payee("株式会社カモメ重車両工業"), R.norm_payee("カモメ重車輌工業"),
       "両 と 輌 も同じキーになる")
    eq(R.norm_payee("有限会社高畠商会"), R.norm_payee("髙畠商会"), "高 と 髙 は同じキー")
    eq(R.norm_payee("梅沢礦油"), R.norm_payee("梅澤礦油"), "沢 と 澤 は同じキー")
    eq(R.norm_payee("ハヤテ急便") != R.norm_payee("北陸白洋舎"), True,
       "別会社が同じキーに潰れない")

    print("--- 機械の見立てと突き合わせ ---")
    hint, why = R.machine_hint(
        {"手書きコード": "3-624-18", "請求書の記載": {"事業所": "金沢", "経費区分": "通運"}}, CODE_RULE)
    eq(hint.get("CDJS300"), "0007", "3段目18 → 富山検修0007（事業所「金沢」より手書きコードが優先）")
    eq(R.compare_hint({"CDJS300": "0004", "CDJS303": "349"}, hint),
       ["CDJS300: AI=0004 機械=0007"], "食い違いを1件だけ挙げる")
    eq(R.compare_hint({"CDJS300": "0007", "CDJS303": "349"}, hint), [], "一致していれば挙げない")

    print("--- 帳票に付いた手書きコードを行に引き継ぐ ---")
    hint, _ = R.machine_hint(
        {"請求書の記載": {"この請求書の手書きコード": "5-624-8"}}, CODE_RULE)
    eq(hint.get("CDJS303"), "199", "合計欄の横のコードでもセグメントが決まる")
    eq(hint.get("CDJS300"), "0007", "同じく部門も決まる")
    hint2, _ = R.machine_hint(
        {"請求書の記載": {"この請求書の手書きコード（複数。どれがこの行か判断すること）":
                        ["3-639-21", "3-639-31"]}}, CODE_RULE)
    eq(hint2, {}, "帳票に複数あるときは機械では決めない（AIに任せる）")
    hint3, _ = R.machine_hint(
        {"手書きコード": "3-624-18",
         "請求書の記載": {"この請求書の手書きコード": "5-624-8"}}, CODE_RULE)
    eq(hint3.get("CDJS303"), "349", "行に付いたコードのほうが優先される")

    print("--- 伝票の括り ---")
    docs = [
        {"doc_id": "D01", "page_from": 1, "doc_type": "請求書", "vendor_name": "○○商事", "total_incl_tax": 30000},
        {"doc_id": "D02", "page_from": 2, "doc_type": "納品書", "vendor_name": "○○商事"},
        {"doc_id": "D03", "page_from": 3, "doc_type": "請求書", "vendor_name": "△△運輸", "total_incl_tax": 50000},
    ]
    g = group_file(docs)
    eq(len(g), 2, "本体2枚 + 添付1枚 → 2伝票")
    eq(g[0]["docs"], ["D01", "D02"], "納品書は直前の請求書に付く")
    eq(_amount(g[0]), 30000, "添付は金額に足さない")

    docs2 = [
        {"doc_id": "D01", "page_from": 1, "doc_type": "請求書", "vendor_name": "□□社", "total_incl_tax": 80000},
        {"doc_id": "D02", "page_from": 2, "doc_type": "請求書", "vendor_name": "□□社", "total_incl_tax": 30000},
        {"doc_id": "D03", "page_from": 3, "doc_type": "請求書", "vendor_name": "□□社", "total_incl_tax": 50000},
    ]
    g2 = group_file(docs2)
    eq(len(g2), 1, "鑑80,000 = 内訳30,000+50,000 は1伝票")
    eq(_amount(g2[0]), 80000, "鑑＋内訳の金額は80,000（160,000にしない）")

    docs3 = [
        {"doc_id": "D01", "page_from": 1, "doc_type": "請求書", "vendor_name": "◇◇社", "total_incl_tax": 30000},
        {"doc_id": "D02", "page_from": 2, "doc_type": "請求書", "vendor_name": "◇◇社", "total_incl_tax": 30000},
    ]
    g3 = group_file(docs3)
    eq(len(g3), 1, "同じ発行者・同額の写しは1伝票")
    eq(_amount(g3[0]), 30000, "写しの金額は30,000（60,000にしない）")
    eq(g3[0].get("copies"), ["D02"], "写しとして記録される")

    print("--- 同じ発行者でも締月が違えば別伝票 ---")
    # 実データ: ソラチの5月分2,704円と6月分2,879円が1伝票にまとまり、
    # 伝票金額が5,583円になった。正解は月ごとに別伝票
    M = [{"doc_id": "D1", "doc_type": "請求書", "vendor_name": "ソラチエネルギー",
          "total_incl_tax": 2704, "closing_date": "2026-05-31"},
         {"doc_id": "D2", "doc_type": "請求書", "vendor_name": "ソラチエネルギー",
          "total_incl_tax": 2879, "closing_date": "2026-06-30"}]
    g = group_file(M)
    eq(len(g), 2, "同じ発行者でも締月が違えば2伝票")
    eq([_amount(x) for x in g], [2704, 2879], "金額は足し合わせない")
    M2 = [dict(M[0]), dict(M[1])]
    M2[1]["closing_date"] = "2026-05-31"
    eq(len(group_file(M2)), 1, "締月が同じなら1伝票にまとめる")
    M3 = [dict(M[0]), dict(M[1])]
    M3[1]["closing_date"] = ""
    eq(len(group_file(M3)), 1, "締日が読めなければ分けない（迷ったらまとめる）")
    eq(R.same_closing_month({"closing_date": "2026-06-30"}, {"issue_date": "2026/6/5"}), True,
       "締日が無ければ発行日で代える")

    print("--- 別ファイルの写しを見分けるキー ---")
    k = R.dup_key("○○商事", "2026-05-31", 30000)
    eq(R.dup_key("○○商事（株）", "2026-05-31", 30000), k, "法人格の表記が違っても同じキー")
    eq(R.dup_key("○○商事", "2026-05-20", 30000), k, "同じ月なら締日が違っても同じキー")
    eq(R.dup_key("○○商事", "2026-06-30", 30000) != k, True, "締月が違えば別のキー")
    eq(R.dup_key("△△商事", "2026-05-31", 30000) != k, True, "発行者が違えば別のキー")
    # 実データにある取りこぼし。写しの側で金額が3円ずれていて、これは捕まらない
    eq(R.dup_key("○○商事", "2026-05-31", 30003) != k, True, "金額が少しでも違えば別のキー（既知の取りこぼし）")

    print("--- 金額に数える本体を選ぶ（fold_bodies）---")
    # 実データで壊れた形。Step 3 と Step 4 の両方から同じ関数を呼ぶことで直した。
    def amt(docs):
        b, _, _ = R.fold_bodies(docs)
        return sum(R.doc_amount(x)[0] for x in b)

    inv = {"doc_id": "D1", "doc_type": "請求書", "vendor_name": "ナギサ株式会社", "total_incl_tax": 80864}
    eq(amt([inv, {"doc_id": "D2", "doc_type": "請求内訳書", "vendor_name": "ナギサ株式会社",
                  "total_incl_tax": 80864}]), 80864, "請求書＋内訳書（同額）は請求書だけ数える")
    eq(amt([inv, {"doc_id": "D2", "doc_type": "請求内訳書", "vendor_name": "",
                  "total_incl_tax": 80864}]), 80864, "内訳書に発行者名が無くても数えない")
    eq(amt([inv, {"doc_id": "D2", "doc_type": "請求内訳書", "vendor_name": "ナギサ株式会社",
                  "total_incl_tax": 32053}]), 80864, "内訳書が一部の小計でも数えない")
    eq(amt([{"doc_id": "D1", "doc_type": "請求書", "vendor_name": "ハヤテ急便", "total_incl_tax": 12164},
            {"doc_id": "D2", "doc_type": "請求書", "vendor_name": "北陸白洋舎", "total_incl_tax": 12164}]),
       24328, "別会社で偶然同額なら畳まない")
    eq(amt([{"doc_id": "D1", "doc_type": "請求内訳書", "vendor_name": "A社", "total_incl_tax": 5000}]),
       5000, "内訳書だけなら（請求書が無いので）数える")

    # 鑑＋内訳。doc_type が全部「請求書」でも畳む
    # 実データ: 鑑565,868 と 内訳37,400+28,468+500,000 を両方数えて2倍になった
    ud = [{"doc_id": "D06", "doc_type": "請求書", "vendor_name": "UDトラックス", "total_incl_tax": 565868},
          {"doc_id": "D01", "doc_type": "請求書", "vendor_name": "UDトラックス", "total_incl_tax": 37400},
          {"doc_id": "D02", "doc_type": "請求書", "vendor_name": "UDトラックス", "total_incl_tax": 28468},
          {"doc_id": "D03", "doc_type": "請求書", "vendor_name": "UDトラックス", "total_incl_tax": 500000}]
    eq(amt(ud), 565868, "鑑の総額＝内訳3枚の合計なら鑑だけ数える")
    b_, folded_, _ = R.fold_bodies(ud)
    eq(sorted(folded_), ["D01", "D02", "D03"], "畳んだのは内訳3枚")
    eq(amt(list(reversed(ud))), 565868, "並び順が変わっても畳める")
    # 安全側: 内訳が1枚しか無いときは畳まない（同額の別の請求と区別できない）
    eq(amt([{"doc_id": "A", "doc_type": "請求書", "vendor_name": "甲", "total_incl_tax": 1000},
            {"doc_id": "B", "doc_type": "請求書", "vendor_name": "乙", "total_incl_tax": 1000}]),
       2000, "同額2枚だけなら畳まない（別会社の偶然と区別できない）")
    eq(amt([{"doc_id": "A", "doc_type": "請求書", "vendor_name": "甲", "total_incl_tax": 300},
            {"doc_id": "B", "doc_type": "請求書", "vendor_name": "乙", "total_incl_tax": 400},
            {"doc_id": "C", "doc_type": "請求書", "vendor_name": "丙", "total_incl_tax": 500}]),
       1200, "合計が一致しない3枚はそのまま足す")

    print("--- 締月の表記ゆれ ---")
    eq(R.ym6("2026-07-31"), "202607", "ハイフン区切り")
    eq(R.ym6("2026/7/31"), "202607", "スラッシュ・1桁の月")
    eq(R.ym6("20260731"), "202607", "区切りなし")
    eq(R.ym6(""), "", "空なら空")
    # 表記が違うだけで別伝票扱いになると、別ファイルの写しを畳めない
    eq(R.dup_key("A社", "2026/07/31", 100), R.dup_key("A社", "20260731", 100),
       "表記が違っても同じ締月なら同じキー")
    # 源泉の新旧判定も締月で分かれる。スラッシュ表記で旧ルールに落ちていた
    wh = {"total_incl_tax": 90000, "withholding_tax": 10000}
    eq(R.doc_amount(dict(wh, closing_date="2026/07/31"))[0], 90000,
       "2026年4月以降はスラッシュ表記でも差引額のまま")
    eq(R.doc_amount(dict(wh, closing_date="2026/03/31"))[0], 100000,
       "2026年3月までは源泉を戻して総額にする")

    print("--- 明細を集める（collect_lines）---")
    D = {"D1": {"doc_id": "D1", "handwritten_account_codes": ["3-624-1"]},
         "D2": {"doc_id": "D2", "handwritten_account_codes": []}}
    BY = {"D1": [{"line_id": "1", "row_type": "detail"},
                 {"line_id": "2", "row_type": "total"}],
          "D2": [{"line_id": "3", "row_type": "detail"}]}
    got = R.collect_lines(["D1", "D2", "D9"], D, BY)
    eq([l["line_id"] for l in got], ["1", "3"], "detail行だけ集める（無い帳票は飛ばす）")
    eq(got[0]["帳票の手書きコード"], ["3-624-1"], "帳票の手書きコードを行に引き継ぐ")
    eq(BY["D1"][0].get("帳票の手書きコード"), None, "元の明細は書き換えない")

    # 鑑（請求書）の合計欄の横にコードがあり、明細は請求内訳書側にある形。
    # 補わないと、いちばん当たる手がかりが Step 5・Step 6 に届かない
    ls = R.collect_lines(["D2"], D, BY)
    eq(R.fill_voucher_codes(ls, ["D1", "D2"], D), ["3-624-1"], "伝票内のコードを明細に補う")
    eq(ls[0]["帳票の手書きコード"], ["3-624-1"], "補ったコードが行に入る")
    ls2 = R.collect_lines(["D1"], D, BY)
    eq(R.fill_voucher_codes(ls2, ["D1", "D2"], D), [], "すでに付いている行があれば補わない")

    print("--- 車番から部門・セグメントを引く（車両マスタ）---")
    # セグメントは**車で決まる**。実データで間違えた2件はどちらもマスタと正反対だった
    VEH = {"4821": {"購入部門": "0001", "セグメント": "249"},
           "7304": {"購入部門": "0001", "セグメント": "349"},
           "0532": {"購入部門": "0004", "セグメント": "149"}}
    eq(R.norm_vehicle("金沢4821"), "4821", "事業所つきでも車番を取れる")
    eq(R.norm_vehicle("富山 100 き 0532"), "0532", "登録番号から4桁を取れる")
    eq(R.norm_vehicle("867"), "0532", "3桁は頭に0を足して4桁にする")
    eq(R.norm_vehicle("整備一式"), "", "車番が無ければ空")
    eq(R.from_vehicle("金沢4821", VEH), {"CDJS300": "0001", "CDJS303": "249"},
       "車番4821 → 事業249（マスタどおり）")
    eq(R.from_vehicle("金沢7304", VEH), {"CDJS300": "0001", "CDJS303": "349"},
       "車番7304 → 通運349（マスタどおり）")
    eq(R.from_vehicle("867", VEH), {"CDJS300": "0004", "CDJS303": "149"},
       "3桁で書かれていても引ける")
    eq(R.from_vehicle("金沢9999", VEH), {}, "マスタに無い車番は空")
    # 機械の見立てに入っているか（ここが無いとAIの誤りを検出できない）
    hint, why = R.machine_hint({"vehicle_no": "金沢4821", "請求書の記載": {}}, CODE_RULE, VEH)
    eq(hint.get("CDJS303"), "249", "machine_hint が車両マスタを引く")
    eq("車両マスタ" in why, True, "根拠に車両マスタが出る")
    hint2, _ = R.machine_hint(
        {"請求書の記載": {"この帳票の車番": "7304"}}, CODE_RULE, VEH)
    eq(hint2.get("CDJS303"), "349", "帳票側の車番でも引ける")
    hint3, _ = R.machine_hint(
        {"vehicle_no": "金沢4821", "手書きコード": "3-624-18",
         "請求書の記載": {}}, CODE_RULE, VEH)
    eq(hint3.get("CDJS303"), "349", "手書きコードのほうが車両マスタより優先")

    print("--- 帳票の車番を明細に引き継ぐ ---")
    VD = {"D1": {"doc_id": "D1", "vehicle_no": "富山 100 き 5926"}}
    VL = {"D1": [{"row_type": "detail", "line_id": "1"},
                 {"row_type": "detail", "line_id": "2", "vehicle_no": "9999"}]}
    got = R.collect_lines(["D1"], VD, VL)
    eq(got[0].get("帳票の車番"), "5926", "行に車番が無ければ帳票の車番を付ける")
    eq(got[1].get("帳票の車番"), None, "行に車番があれば上書きしない")

    print("--- 社内集計表を明細に使う（worksheet_lines）---")
    # 実データで壊れた形。鑑（カード利用明細）に合計1行だけの明細があり、それで
    # 金額が合ってしまうので本当の明細を捨てていた（北陸カード 明細1行・正解13行）。
    # 集計表は経理担当者の配分表で、金額・事業所・区分まで正解と一致する
    WD = {"D1": {"doc_id": "D1"},                                   # 鑑
          "D2": {"doc_id": "D2", "is_internal_worksheet": True},    # 配分表
          "D3": {"doc_id": "D3", "is_internal_worksheet": True}}    # 金額が合わない集計表
    WL = {"D1": [{"row_type": "detail", "amount_incl_tax": 489770}],
          "D2": [{"row_type": "detail", "amount_incl_tax": 300000},
                 {"row_type": "detail", "amount_incl_tax": 189770},
                 {"row_type": "detail", "amount_incl_tax": 0}],
          "D3": [{"row_type": "detail", "amount_incl_tax": 999}]}
    eq(len(R.worksheet_lines(["D1", "D2"], WD, WL, 489770)), 3,
       "集計表の合計が伝票金額に合えばその明細を使う")
    eq(R.worksheet_lines(["D1", "D3"], WD, WL, 489770), [],
       "合わない集計表は使わない")
    eq(R.worksheet_lines(["D1"], WD, WL, 489770), [], "集計表が無ければ空")
    eq(R.worksheet_lines(["D1", "D2"], WD, WL, 0), [], "伝票金額が0なら使わない")

    print("--- 摘要から車番を抜く（vehicle_in_text / 元の車番の実績）---")
    # 過去の仕訳の摘要は「車番＋作業名＋取引先」。ここから車番ごとの部門/セグを作る
    eq(R.vehicle_in_text("金沢4821 ﾁｬｰｼﾞﾗﾝﾌﾟ点灯 カナタ自動車㈱"), "4821", "摘要の頭から車番")
    eq(R.vehicle_in_text("富山 100 き 8137 高速料金"), "8137", "間に空白があっても拾う")
    eq(R.vehicle_in_text("福井0418 車検整備"), "0418", "頭の0を落とさない")
    eq(R.vehicle_in_text("7月分 軽油代 ﾍﾞｰﾀ商事㈱"), "", "車番が無ければ空")
    eq(R.vehicle_in_text("金沢6150 高速料金"), "6150", "地名＋4桁")

    print("--- 車番を機械で埋める（fill_line_vehicle）---")
    # AIが label に車番を書き忘れるとセグメントが当てられなくなる
    # （実データ: 金沢4821が249→349、金沢7304が349→249 に反転した）
    SRC = {"L001": {"line_id": "L001", "vehicle_no": "4821"},
           "L002": {"line_id": "L002", "帳票の車番": "7304"},
           "L003": {"line_id": "L003", "vehicle_no": "867"},
           "L004": {"line_id": "L004"}}
    LN = [{"label": "ﾁｬｰｼﾞﾗﾝﾌﾟ点灯修理", "source_line_ids": ["L001"]},
          {"label": "3ｹ月点検", "source_line_ids": ["L002"]},
          {"label": "車検整備", "source_line_ids": ["L001", "L002"]},
          {"label": "ﾁｬｰﾄ紙", "source_line_ids": ["L004"]},
          {"label": "既にある", "vehicle_no": "9999", "source_line_ids": ["L003"]}]
    eq(R.fill_line_vehicle(LN, SRC), 2, "車番が1つに定まる行だけ埋める")
    eq(LN[0]["vehicle_no"], "4821", "明細の車番を引き継ぐ")
    eq(LN[1]["vehicle_no"], "7304", "帳票の車番も引き継ぐ")
    eq(LN[2].get("vehicle_no"), None, "2台ぶんが混ざる行には入れない")
    eq(LN[4]["vehicle_no"], "9999", "すでに入っている行は触らない")

    print("--- 科目×補助・部門×セグの組合せ（check_combos / 元のC-4・C-5）---")
    AS = {("426200", "0000"), ("425000", "6251")}
    DS = {("0002", "349"), ("0007", "199")}
    eq(R.check_combos({"CDJS301": "426200", "CDJS302": "6351",
                       "CDJS300": "0002", "CDJS303": "349"}, AS, DS),
       ["C-4 科目426200×補助6351 の組合せが実績に無い"], "実績に無い科目×補助を挙げる")
    eq(R.check_combos({"CDJS301": "426200", "CDJS302": "0000",
                       "CDJS300": "0004", "CDJS303": "149"}, AS, DS),
       ["C-5 部門0004×セグ149 の組合せが実績に無い"], "実績に無い部門×セグを挙げる")
    eq(R.check_combos({"CDJS301": "426200", "CDJS302": "0000",
                       "CDJS300": "0002", "CDJS303": "349"}, AS, DS), [], "実績にある組合せは通す")
    eq(R.check_combos({"CDJS301": "999999", "CDJS302": "9999"}, set(), set()), [],
       "実績が読めていなければ何も言わない")

    print("--- 鑑＋内訳を割り方の手がかりに（breakdown_hint）---")
    # 実データのカモメ重車輌 391,952円。鑑1枚と内訳3枚で、その3枚が正解の3行そのもの。
    # fold_bodies は金額が2倍にならないよう内訳を落とすので、ここで別に拾う。
    BD = {"D1": {"doc_id": "D1", "vendor_name": "カモメ重車輌工業", "total_incl_tax": 391952},
          "D2": {"doc_id": "D2", "vendor_name": "カモメ重車輌工業", "total_incl_tax": 299596,
                 "vehicle_no": "4023", "page_from": 3},
          "D3": {"doc_id": "D3", "vendor_name": "カモメ重車輌工業", "total_incl_tax": 68695,
                 "vehicle_no": "5926", "page_from": 4},
          "D4": {"doc_id": "D4", "vendor_name": "カモメ重車輌工業", "total_incl_tax": 23661,
                 "vehicle_no": "867", "page_from": 5}}
    got = R.breakdown_hint(["D1", "D2", "D3", "D4"], BD, 391952)
    eq([x["金額"] for x in got], [299596, 68695, 23661], "内訳3枚の総額を返す（鑑は返さない）")
    eq(got[2]["車番"], "0532", "車番は4桁にそろえる")
    eq(R.breakdown_hint(["D1"], BD, 391952), [], "内訳が無ければ空")
    eq(R.breakdown_hint(["D1", "D2"], BD, 391952), [], "内訳1枚だけでは使わない（同額の別請求と区別できない）")
    eq(R.breakdown_hint(["D1", "D2", "D3", "D4"], BD, 0), [], "伝票金額が0なら使わない")

    print("--- 配分表として渡す小計行（worksheet_hint）---")
    # 実データの債務579。集計表に軽油とアドブルーの配分はあるが軽油引取税が無く、
    # 合計が伝票金額に届かないので worksheet_lines では採れない。それでも配分の
    # 手がかりとしては効くので、小計行だけを Step 5 に渡す。
    HD = {"D1": {"doc_id": "D1"},
          "D2": {"doc_id": "D2", "is_internal_worksheet": True, "total_incl_tax": 2485423}}
    HL = {"D1": [{"row_type": "detail", "item_name": "軽油代", "amount_incl_tax": 2173098}],
          "D2": [{"row_type": "detail", "item_name": "軽油", "amount_incl_tax": 60921},
                 {"row_type": "subtotal", "item_name": "計（本社 事業 軽油）",
                  "amount_incl_tax": 747725},
                 {"row_type": "subtotal", "item_name": "計（富山 営業 軽油）",
                  "amount_incl_tax": 552748},
                 {"row_type": "subtotal", "item_name": "軽油 合計",
                  "amount_incl_tax": 1300473},          # 上2本のまとめ。落とす
                 {"row_type": "total", "item_name": "総計", "amount_incl_tax": 2485423}]}
    got = R.worksheet_hint(["D1", "D2"], HD, HL)
    eq([x["金額"] for x in got], [747725, 552748], "小計行だけを返す（明細・まとめ・総計は外す）")
    eq("line_id" in got[0], True, "参照元をたどれるよう line_id を付ける")
    eq(R.worksheet_hint(["D1"], HD, HL), [], "集計表が無ければ空")

    print("--- dbの型に合わせる（rules.coerce_types）---")
    # テーブルの型は最初に入れた値で決まる。文字列で入れると全件まとめて弾かれる
    # （実データ: field '年月' expects a number で139件すべて拒否された）
    T = {"年月": int, "勘定科目": int, "本体金額": int,
         "補助科目": str, "購入部門": str, "件数": int}
    got = R.coerce_types({"年月": "202607", "勘定科目": "451020", "補助科目": "0000",
                    "購入部門": "0001", "本体金額": 2244, "登録元": "確認済み"}, T)
    eq(got["年月"], 202607, "数値の列は int にする")
    eq(got["勘定科目"], 451020, "勘定科目も int にする")
    eq(got["補助科目"], "0000", "文字列の列は頭の0を保つ")
    eq(got["購入部門"], "0001", "購入部門も文字列のまま")
    eq(got["本体金額"], 2244, "すでに int ならそのまま")
    eq(got["登録元"], "確認済み", "型が分からない列はそのまま")
    eq(R.coerce_types({"年月": "よみとれず"}, T)["年月"], "よみとれず",
       "数値にできない値は変えない（弾かれて原因が分かるほうがよい）")
    eq(R.coerce_types({"件数": 3}, {"件数": str})["件数"], "3", "逆に文字列の列なら str にする")
    eq(R.coerce_types({"年月": "202607"}, {}), {"年月": "202607"},
       "型が取れなければ（空テーブル）何もしない")

    print("--- 検算 ---")
    ng = R.check_voucher({"金額": 10000}, [{"金額": 6000, "CDJS300": "0001", "CDJS301": "425000",
                                            "CDJS302": "0000", "CDJS303": "349"}])
    eq(bool([x for x in ng if x.startswith("V-1")]), True, "行合計 ≠ 伝票金額 を検出")
    ng = R.check_voucher({"金額": 6000}, [{"金額": 6000, "CDJS300": "0001", "CDJS301": "425000",
                                           "CDJS302": "0000", "CDJS303": "449"}])
    eq(bool([x for x in ng if x.startswith("C-9")]), True, "セグ449なのに42x系 を検出")
    ng = R.check_voucher({"金額": 6000}, [{"金額": 6000, "CDJS300": "0001", "CDJS301": "451015",
                                           "CDJS302": "0000", "CDJS303": "449"}])
    eq(ng, [], "整合していれば指摘なし")

    print()
    if _ng:
        print("=== NG %d 件: %s ===" % (len(_ng), ", ".join(_ng)))
    else:
        print("=== 全部 ok ===")
    return not _ng


if __name__ == "__main__":
    sys.exit(0 if run() else 1)
