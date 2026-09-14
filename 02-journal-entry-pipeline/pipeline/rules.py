# -*- coding: utf-8 -*-
"""機械で決められることだけを書いたモジュール。AIは呼ばない。

入れてよいのは例外なく成り立つ規則だけ。判断が要るものは s40 に渡す。
"""
import re, unicodedata

# ---------------------------------------------------------------- 正規化
def nz(s):
    """全角半角・空白をならす。"""
    return re.sub(r'[\s　]+', '', unicodedata.normalize('NFKC', s or ''))

def norm_item(s):
    """品名の照合キー。数字は # に潰す（「3ヶ月点検」と「6ヶ月点検」を同じ穴に入れない用途では使わない）。"""
    return re.sub(r'[0-9０-９]+', '#', nz(s))

# 会社名に混ざる異体字。**片方に寄せないと別会社になる。**
# 実データ: AIは「カモメ重車輛工業」「カモメ重車両工業」と読み、別名テーブルは
# 「カモメ重車輌工業」。輛(U+8F1B)/両/輌(U+8F0C) は別の文字なので NFKC では揃わず、
# 支払先コードが引けずに過去実績が渡らなくなった（19行）。
KANJI_VARIANT = {
    '輛': '輌', '両': '輌', '濱': '浜', '邊': '辺', '邉': '辺',
    '齋': '斎', '齊': '斎', '斉': '斎', '﨑': '崎', '曾': '曽',
    '冨': '富', '廣': '広', '澤': '沢', '瀨': '瀬', '龍': '竜',
    '嶋': '島', '髙': '高', '圓': '円', '國': '国', '學': '学',
    '眞': '真', '榮': '栄', '壽': '寿', '德': '徳', '籐': '藤', '檜': '桧',
}

def norm_payee(s):
    """会社名の照合キー。法人格を落とし、異体字を片方に寄せる。

    NFKC で ㈱→(株) になるので括弧つきの表記もまとめて落とす。
    異体字は NFKC では揃わないので KANJI_VARIANT で置き換える。
    """
    t = nz(s)
    t = re.sub(r'(株式会社|有限会社|合同会社|一般社団法人|公益社団法人|公益財団法人'
               r'|\(株\)|\(有\)|\(同\)|㈱|㈲)', '', t)
    t = re.sub(r'[・･\-‐]', '', t)
    return ''.join(KANJI_VARIANT.get(c, c) for c in t)

def ym6(s):
    """日付から年月（YYYYMM）だけ取る。取れなければ空。

    2026-07-31 / 2026/7/31 / 20260731 が混ざるので、先頭を切る作りでは比べられない。
    """
    m = re.search(r'(20\d{2})\D?(\d{1,2})', str(s or ''))
    return '%04d%02d' % (int(m.group(1)), int(m.group(2))) if m else ''

def dup_key(payee, closing, amount):
    """別ファイルの同じ請求書を見つけるキー。発行者・締月・金額の3点。

    金額の完全一致で見るので、写しの側で金額がずれた請求書は捕まらない（検算で拾う）。
    """
    return "%s|%s|%d" % (norm_payee(payee), ym6(closing), int(amount or 0))

def closing_ym(doc):
    """帳票の締月（YYYYMM）。締日が無ければ発行日で代える。読めなければ空。"""
    return ym6(doc.get('closing_date') or doc.get('issue_date') or '')


def same_closing_month(a, b):
    """2枚の帳票が同じ締月か。**どちらかが読めなければ True**（分けない）。

    同じ支払先の請求書が続いていても、締月が違えば別の伝票になる。
    実データ: ソラチの5月分2,704円と6月分2,879円が1伝票にまとまり、
    伝票金額が5,583円になった（正解は月ごとに別伝票）。

    読めないときに分けてしまうと、1枚の請求書が複数伝票に割れて金額が壊れる。
    **迷ったらまとめる**側に倒す。
    """
    x, y = closing_ym(a), closing_ym(b)
    return not (x and y) or x == y


def resend_last(paths):
    """「再送」を名前に含むファイルを後ろに回した並び。

    重複は**先に来たほうが残る**ので、本編を残すために再送を後ろにする。
    """
    import os
    return sorted(paths, key=lambda p: ("再送" in os.path.basename(p), p))

# ---------------------------------------------------------------- 手書き3段コード
# 例: 3-624-1 → 第1段=3(通運→セグ349) / 第2段=624(→勘定科目) / 第3段=1(→補助科目と箇所)
SEG1 = {'1': '449', '2': '249', '3': '349', '4': '149', '5': '199'}
# 第3段の2桁目 = 箇所
PLACE = {'1': '0001', '3': '0002', '5': '0005', '6': '0004', '7': '0006', '8': '0007'}

def parse_code(code):
    """手書きコードを段に割る。読めなければ空リスト。"""
    if not code:
        return []
    t = nz(str(code)).replace('ー', '-').replace('−', '-').replace('―', '-')
    parts = [p for p in re.split(r'[-‐]', t) if p != '']
    if not parts or not re.fullmatch(r'[0-9A-Za-z]+', parts[0]):
        return []
    return parts

def from_code(code, code_rule):
    """手書きコード → {部門/科目/補助/セグ} の分かる範囲。

    code_rule: [{段, 入力, 出力項目, 出力値}, ...]（beta_code_rule テーブルの中身）
    """
    p = parse_code(code)
    if not p:
        return {}
    idx = {}
    for r in code_rule:
        idx[(str(r.get('段')), str(r.get('入力')))] = (r.get('出力項目') or '', str(r.get('出力値') or ''))
    out = {}
    if len(p) >= 1:
        seg = SEG1.get(p[0][0]) or idx.get(('1', p[0]), ('', ''))[1]
        if seg:
            out['CDJS303'] = seg
    if len(p) >= 2:
        hit = idx.get(('2', p[1]))
        if hit and hit[0].startswith('CDJS301'):
            out['CDJS301'] = hit[1]
    if len(p) >= 3:
        hit = idx.get(('3', p[2]))
        if hit and hit[0].startswith('CDJS302'):
            out['CDJS302'] = hit[1]
        if len(p[2]) >= 2 and p[2][1] in PLACE:
            out['CDJS300'] = PLACE[p[2][1]]
        elif len(p[2]) == 1 and p[2] in PLACE:
            out['CDJS300'] = PLACE[p[2]]
    return out

# ---------------------------------------------------------------- 事業所 → 部門
OFFICE = {'本社': '0001', '福井': '0002', '富山': '0004', '魚津': '0004', '伏木': '0004',
          '高岡': '0005', '青海': '0006'}

def office_to_dept(office):
    """事業所名 → 部門コード。決まらなければ None。例外の判断は s40 にさせる。"""
    t = nz(office)
    for k, v in OFFICE.items():
        if k in t:
            return v
    if '金沢' in t:
        return '0001'   # 金沢営業所コード0003はほとんど使われていない（実データ12件）
    return None

# 請求書の「経費区分」は的中率が高いものだけ s40 に渡す（管理・事業は渡さない）
TRUSTED_COST_SEGMENT = {'通運': '349', '営業': '149', '検修': '199'}

def filter_cost_segment(v):
    t = nz(v)
    return t if t in TRUSTED_COST_SEGMENT else None

# ---------------------------------------------------------------- 管理⇔事業の切替
def is_mgmt(seg):
    return str(seg or '') == '449'

def fix_mgmt_business(row, acct_mgmt):
    """セグ449なら科目は45x系、それ以外なら42x系。ずれていたら対応表で入れ替える。

    acct_mgmt: [{事業側_科目, 事業側_補助, 管理側_科目, 管理側_補助}, ...]
    戻り: (直した行, 直したかどうか)
    """
    acct = str(row.get('CDJS301') or '')
    sub  = str(row.get('CDJS302') or '')
    seg  = str(row.get('CDJS303') or '')
    if not acct or not seg:
        return row, False
    want_mgmt = is_mgmt(seg)
    now_mgmt  = acct.startswith('45')
    if want_mgmt == now_mgmt:
        return row, False
    for e in acct_mgmt:
        if want_mgmt and str(e.get('事業側_科目')) == acct and str(e.get('事業側_補助')) in (sub, ''):
            r = dict(row); r['CDJS301'] = str(e['管理側_科目']); r['CDJS302'] = str(e['管理側_補助'])
            return r, True
        if (not want_mgmt) and str(e.get('管理側_科目')) == acct and str(e.get('管理側_補助')) in (sub, ''):
            r = dict(row); r['CDJS301'] = str(e['事業側_科目']); r['CDJS302'] = str(e['事業側_補助'])
            return r, True
    return row, False

# ---------------------------------------------------------------- 伝票の金額
def doc_amount(doc):
    """1つの請求書から、仕訳に立てる金額を出す。

    源泉徴収は2026年4月に扱いが変わった：
      2026年3月まで … 差引請求額に源泉分を戻して総額にする
      2026年4月から … 差引請求額をそのまま使う（源泉は別伝票で212800-3402へ）
    """
    total = int(doc.get('total_incl_tax') or 0)
    wh    = int(doc.get('withholding_tax') or 0)
    flags = []
    if wh:
        flags.append('要確認:源泉徴収あり')
        ymd = ym6(doc.get('closing_date') or doc.get('issue_date') or '')
        if ymd and ymd >= '202604':
            return total, flags
        return total + wh, flags
    return total, flags

# ---------------------------------------------------------------- 帳票の種別
BODY_TYPES = {'請求書', '請求内訳書', 'ご請求書', '振込依頼書', 'その他'}
ATTACH_TYPES = {'納品書', '見積書', '御買上明細書', '部品請求明細書', '整備請求明細書',
                '請求明細書', '受領書', '領収書'}

# 「請求書の内訳」を書いた帳票。請求書と同じ伝票に入っていたら金額に数えない
# （内訳なので、足すと請求書のぶんと二重になる）。
BREAKDOWN_TYPES = {'請求内訳書'}
# その内訳の親になれる帳票。
INVOICE_TYPES = {'請求書', 'ご請求書'}

def is_body(doc):
    """その帳票が「1伝票の本体」になりうるか（＝表紙・鑑か）。添付の明細書は False。"""
    if doc.get('is_internal_worksheet') or doc.get('is_supporting_doc'):
        return False
    if not (doc.get('total_incl_tax') or 0) > 0:
        return False
    if doc.get('doc_type') in BODY_TYPES:
        return True
    if doc.get('doc_type') in ATTACH_TYPES:
        return False
    return bool(doc.get('vendor_name'))

def fold_bodies(bodies):
    """1伝票の本体から、**金額に数えるもの**だけを選ぶ。s30 と s32 の両方から呼ぶ。

    戻り: (数える本体のリスト, 畳んだ doc_id のリスト, フラグのリスト)

    ① 請求内訳書は、同じ伝票に請求書があるなら数えない（足すと二重になる）。
       **総額が一致するかは見ない。** 内訳書が一部の小計しか載せていないことがあり、
       総額一致で畳む方式だと畳めない。
    ② 写し。発行者と総額が同じものは片方だけ数える。
    """
    bodies = list(bodies)
    folded, flags = [], []

    # ⓪ 鑑＋内訳。**1枚の総額が残りの合計と一致したら、その1枚だけ数える。**
    # doc_type が全部「請求書」でも起きるので、①の請求内訳書ルールでは畳めない
    # （実データ: 鑑565,868 と 内訳37,400+28,468+500,000 を両方数えて2倍になった）。
    # 残りが2枚以上のときだけ。2枚だと「同額の別の請求」と区別がつかない
    for i, b in enumerate(bodies):
        t = int(b.get('total_incl_tax') or 0)
        rest = [x for j, x in enumerate(bodies) if j != i]
        if not t or len(rest) < 2:
            continue
        if abs(sum(int(x.get('total_incl_tax') or 0) for x in rest) - t) <= 3:
            flags.append('鑑1枚と内訳%d枚を見つけたので鑑の総額だけ数えた' % len(rest))
            return [b], [x['doc_id'] for x in rest], flags

    # ① 請求内訳書を落とす（請求書が同じ伝票にあるときだけ）
    has_invoice = any((b.get('doc_type') or '') in INVOICE_TYPES for b in bodies)
    if has_invoice:
        keep = []
        for b in bodies:
            if (b.get('doc_type') or '') in BREAKDOWN_TYPES:
                folded.append(b['doc_id'])
            else:
                keep.append(b)
        if folded:
            flags.append('請求内訳書%d枚を金額に数えなかった（請求書の内訳）' % len(folded))
        bodies = keep

    # ② 写し（発行者と総額が同じ）
    seen, keep, copies = set(), [], []
    for b in bodies:
        sig = (norm_payee(b.get('vendor_name')), int(b.get('total_incl_tax') or 0))
        if sig[1] and sig in seen:
            copies.append(b['doc_id'])
            continue
        seen.add(sig)
        keep.append(b)
    if copies:
        flags.append('同じ請求書の写しを%d枚畳んだ' % len(copies))
    return keep, folded + copies, flags


def collect_lines(doc_ids, docs_by_id, lines_by_doc):
    """指定した帳票の明細（detail行）を集める。s30 と s32 の両方から呼ぶ。

    **帳票に付いた手書きコードを各行に引き継ぐ。** 3段コードは合計欄の横に帳票ごと
    1つ書かれていて明細行には付いていないので、ここで引き継がないと下流に届かない。
    """
    out = []
    for did in doc_ids:
        d = docs_by_id.get(did)
        if d is None:
            continue
        codes = [c for c in (d.get('handwritten_account_codes') or []) if c]
        # 車番は明細行ではなく**帳票の見出し**に書かれている（整備の請求書は1台1枚）。
        # ここで引き継がないと車両マスタが引けず、セグメントを当てられない
        veh = norm_vehicle(d.get('vehicle_no') or d.get('vehicle_plate') or '')
        for l in lines_by_doc.get(did, []):
            if l.get('row_type') != 'detail':
                continue
            l = dict(l)
            l['帳票の手書きコード'] = codes
            if veh and not norm_vehicle(l.get('vehicle_no') or ''):
                l['帳票の車番'] = veh
            out.append(l)
    return out


def worksheet_lines(doc_ids, docs_by_id, lines_by_doc, amount):
    """社内集計表の明細が伝票金額に合えばそれを返す。合うものが無ければ空リスト。

    集計表は経理担当者が自分で書いた**配分表**で、仕訳の行そのものであることが多い。
    実データの北陸カード489,770円は、集計表の22行から0円を落とすと13行になり、
    正解13行と金額・事業所・経費区分まで一致した。

    これを候補に入れないと、**鑑に合計1行だけの明細があるとそれで金額が合ってしまい、
    本当の明細を全部捨てる**（実データで 明細1行・正解13行 になった）。
    明細を集めるときは集計表を外しているので、ここだけ別に見る。

    集計表が「品目の一覧」で事業所も区分も入っていないこともある。その場合も
    行数は変わるが、Step 5 は明細をまとめて仕訳行にするので害は小さい。
    """
    if not amount:
        return []
    for did in doc_ids:
        d = docs_by_id.get(did)
        if not (d and d.get('is_internal_worksheet')):
            continue
        got = collect_lines([did], docs_by_id, lines_by_doc)
        if got and abs(sum(int(l.get('amount_incl_tax') or 0) for l in got) - amount) <= 3:
            return got
    return []


# 摘要に書かれた車番。元の make_l3_ctx_v5e.py と同じ拾い方
#   例「金沢9999 ○○点検 ○○自動車㈱」→ 4821
PLATE_AREA = r'金沢|富山|福井|石川|高岡|魚津|伏木'

def vehicle_in_text(s):
    """摘要や label から車番（4桁）を抜く。無ければ空。

    過去の仕訳の摘要は「車番＋作業名＋取引先」の形なので、ここから
    「その車番が過去どの部門/セグで計上されたか」を作れる。

    元の make_l3_ctx_v5e.py は地名の直後の数字を採っていたが、それだと
    「富山 100 き 8137」のような登録番号の形で分類番号(100)を拾ってしまう。
    仕訳の摘要は短い形しか出ないので元の実装では表に出なかったが、ここでは
    Step 5 の label にも使うため、**地名のあと10文字までの最後の数字**を採る。
    """
    t = nz(s)
    m = re.search(r'(?:%s)' % PLATE_AREA, t)
    if not m:
        return ''
    nums = re.findall(r'[0-9]{2,4}', t[m.end():m.end() + 10])
    return norm_vehicle(nums[-1]) if nums else ''


def fill_line_vehicle(lines, src_lines):
    """Step 5 が返した仕訳行に、車番を機械で埋める。戻り: 埋めた行数。

    車番があるとセグメントが車両マスタで確定する。**AIが label に車番を書き忘れると
    そこだけセグメントが当てずっぽうになる**（実データ: 金沢9999と金沢7304で
    249と349が入れ替わった）。元になった明細の車番が1つに定まるなら機械で入れる。

    src_lines: {line_id: 明細}。明細には collect_lines が帳票の車番も引き継いでいる。
    """
    n = 0
    for l in lines:
        if norm_vehicle(l.get('vehicle_no') or ''):
            continue
        got = set()
        for lid in (l.get('source_line_ids') or []):
            x = src_lines.get(str(lid)) or {}
            v = norm_vehicle(x.get('vehicle_no') or '') or norm_vehicle(x.get('帳票の車番') or '')
            if v:
                got.add(v)
        if len(got) == 1:
            l['vehicle_no'] = got.pop()
            n += 1
    return n


def breakdown_hint(doc_ids, docs_by_id, amount):
    """1伝票が「鑑1枚＋内訳N枚」でできているとき、内訳N枚の総額を返す。
    戻り: [{'名前','金額','車番','ページ'}]（そういう形でなければ空）

    **1台1枚の整備請求書がこれ。** 実データのカモメ重車輌 391,952円は
    鑑1枚と内訳3枚（299,596 / 68,695 / 23,661）で、その3枚が正解の3行そのもの。

    fold_bodies は金額が2倍にならないよう内訳側を落とすので、**割り方の情報が
    そこで消える**。金額に数えるのは鑑だけのまま、割り方の手がかりとして別に渡す。
    """
    amount = int(amount or 0)
    if not amount:
        return []
    docs = [docs_by_id[d] for d in doc_ids if d in docs_by_id]
    kept = [d for d in docs if abs(int(d.get('total_incl_tax') or 0) - amount) <= 3]
    rest = [d for d in docs if int(d.get('total_incl_tax') or 0)
            and abs(int(d.get('total_incl_tax') or 0) - amount) > 3]
    if not kept or len(rest) < 2:
        return []
    if abs(sum(int(d.get('total_incl_tax') or 0) for d in rest) - amount) > 3:
        return []
    return [{'名前': nz(d.get('vendor_name')), '金額': int(d.get('total_incl_tax') or 0),
             '車番': norm_vehicle(d.get('vehicle_no') or d.get('vehicle_plate') or '') or None,
             'ページ': d.get('page_from')} for d in rest]


def worksheet_hint(doc_ids, docs_by_id, lines_by_doc):
    """社内集計表の小計行を「配分表」として返す。戻り: [{'名前','金額'}]

    worksheet_lines は集計表の合計が伝票金額と合うときだけ使える。実データの
    債務579は、集計表に軽油とアドブルーの配分（10本の小計）はあるのに軽油引取税の
    配分が無く、合計が伝票金額に届かないので採用されない。**それでも配分の
    手がかりとしては効く**（引取税も同じ比率で配ればよい）ので、明細とは別に渡す。

    小計行だけを返す。個別の給油行まで渡すと本数が多すぎて Step 5 が読み切れない。
    いちばん外側の総計行（伝票金額と一致する行）は配分ではないので外す。
    """
    out = []
    for did in doc_ids:
        d = docs_by_id.get(did)
        if not (d and d.get('is_internal_worksheet')):
            continue
        top = int(d.get('total_incl_tax') or 0)
        for l in lines_by_doc.get(did) or []:
            name = nz(l.get('item_name'))
            amt = int(l.get('amount_incl_tax') or 0)
            rt = nz(l.get('row_type'))
            if rt in ('total', 'carryover', 'payment', 'note', 'tax_summary'):
                continue
            if not amt or not (rt == 'subtotal' or '計' in name):
                continue
            if top and abs(amt - top) <= 3:
                continue                   # その帳票の総計。配分ではない
            out.append({'line_id': l.get('line_id'), '名前': name, '金額': amt})
    if not out:
        return []
    # 「軽油 合計」のような上位のまとめは落とす。直前の未消化の小計の和と一致する行がそれ。
    keep, run = [], []
    for x in out:
        if len(run) >= 2 and x['金額'] == sum(y['金額'] for y in run):
            run = []                       # ここまでを1つの束として締める
            continue
        keep.append(x)
        run.append(x)
    return keep or out


def fill_voucher_codes(lines, doc_ids, docs_by_id):
    """明細のどれにもコードが付いていないとき、伝票内のコードを全行に補う。
    戻り: 補ったコードの一覧（補わなければ空）。

    鑑にコードがあり明細は請求内訳書側にある形で必要になる。
    どの行のものかは分からないので、複数あればそのまま渡して s40 に選ばせる。
    """
    if not lines or any(l.get('帳票の手書きコード') for l in lines):
        return []
    codes = sorted({c for did in doc_ids
                    for c in (docs_by_id.get(did, {}).get('handwritten_account_codes') or []) if c})
    if not codes:
        return []
    for l in lines:
        l['帳票の手書きコード'] = codes
    return codes


# ---------------------------------------------------------------- dbに入れる前の型合わせ
def coerce_types(rec, types):
    """1行を、テーブルの型に合わせる。合わせられない値はそのまま返す。

    テーブルの型は最初に入れた値で決まっていて、あとから違う型で入れると
    **全件まとめて弾かれる**（実データ: field '年月' expects a number で139件拒否）。
    どの列が数値かを宣言している場所が無いので、既存の1行から読んで合わせる。

    数値にできない値はそのまま返す。黙って0にすると誤りが混ざるので、
    弾かれて原因が分かるほうがよい。
    """
    out = {}
    for k, v in rec.items():
        t = types.get(k)
        if t in (int, float) and isinstance(v, str):
            try:
                v = int(float(v)) if t is int else float(v)
            except ValueError:
                pass
        elif t is str and isinstance(v, (int, float)) and not isinstance(v, bool):
            v = str(v)
        out[k] = v
    return out


# ---------------------------------------------------------------- 検算（S8）
def check_voucher(voucher, lines):
    """伝票1本の検算。壊れていたら理由の一覧を返す。空リストなら合格。

    ここで見るのは「正解を知らなくても分かること」だけ。
    4項目の中身が正しいかは分からないので、ここでは判定しない。
    """
    ng = []
    total = int(voucher.get('金額') or 0)
    s = sum(int(l.get('金額') or 0) for l in lines)
    if total and s != total:
        ng.append('V-1 行の合計%d が伝票金額%d と違う' % (s, total))
    if any(int(l.get('金額') or 0) == 0 for l in lines):
        ng.append('S-2 金額0の行がある')
    for i, l in enumerate(lines, 1):
        for k, name in (('CDJS300', '部門'), ('CDJS301', '科目'), ('CDJS302', '補助'), ('CDJS303', 'セグ')):
            if not l.get(k):
                ng.append('C-6 %d行目の%sが空' % (i, name))
        acct = str(l.get('CDJS301') or ''); seg = str(l.get('CDJS303') or '')
        if acct and seg:
            if is_mgmt(seg) and not acct.startswith('45'):
                ng.append('C-9 %d行目 セグ449なのに科目%s（45x系のはず）' % (i, acct))
            if (not is_mgmt(seg)) and acct.startswith('45'):
                ng.append('C-9 %d行目 セグ%sなのに科目%s（42x系のはず）' % (i, seg, acct))
        if int(l.get('金額') or 0) < 0 and '値引' not in (l.get('摘要') or '') and '相殺' not in (l.get('摘要') or ''):
            ng.append('S-3 %d行目 マイナス金額（値引・相殺の記載なし）' % i)
    return ng

def check_combos(line, acct_sub, dept_seg):
    """C-4 / C-5。実績に出てこない組合せを挙げる。元の invariants.py からの移植。

    C-4 補助科目は勘定科目ごとに決まっている。無い組合せは入力エラーか誤り
    C-5 部門とセグメントには対応関係がある（例 0007↔199）

    実績が読めていない（空集合）ときは何も言わない。**辞書が空のときに
    全行へフラグを立てても、人が見る行が増えるだけで役に立たない。**
    """
    ng = []
    a, b = str(line.get('CDJS301') or ''), str(line.get('CDJS302') or '')
    d, g = str(line.get('CDJS300') or ''), str(line.get('CDJS303') or '')
    if acct_sub and a and b and (a, b) not in acct_sub:
        ng.append('C-4 科目%s×補助%s の組合せが実績に無い' % (a, b))
    if dept_seg and d and g and (d, g) not in dept_seg:
        ng.append('C-5 部門%s×セグ%s の組合せが実績に無い' % (d, g))
    return ng


def check_master(line, master_set):
    """コードがマスタに存在するか。"""
    ng = []
    if line.get('CDJS300') and ('部門', str(line['CDJS300'])) not in master_set:
        ng.append('C-1 部門%s がマスタに無い' % line['CDJS300'])
    if line.get('CDJS301') and ('勘定科目', str(line['CDJS301'])) not in master_set:
        ng.append('C-2 科目%s がマスタに無い' % line['CDJS301'])
    if line.get('CDJS303') and ('セグメント', str(line['CDJS303'])) not in master_set:
        ng.append('C-3 セグメント%s がマスタに無い' % line['CDJS303'])
    return ng


# ---------------------------------------------------------------- 機械側の裏取り（s40 が使う）
def norm_vehicle(s):
    """車番の照合キー。**4桁の数字に揃える。**

    車両マスタは「4821」「0532」の4桁。請求書には「金沢9999」「富山 100 き 0532」
    「867」など色々な形で書かれる。桁が欠けるとマスタが引けず、いちばん確実な
    手がかりを落とす（実データ: 北信タイヤの「867」でマスタが引けなかった）。
    """
    t = nz(s)
    m = re.findall(r'\d{3,4}', t)
    if not m:
        return ''
    return m[-1].zfill(4)


def from_vehicle(vehicle_no, vehicle):
    """車番 → {部門, セグメント}。車両マスタを引く。引けなければ空。

    セグメントは**車で決まる**（同じ整備でも車が違えば通運349と事業249に分かれる）。
    実データで間違えた2件はどちらもマスタの値と正反対だった
    （車番4821はマスタ249なのに349、車番7304はマスタ349なのに249）。
    """
    key = norm_vehicle(vehicle_no)
    if not key or not vehicle:
        return {}
    v = vehicle.get(key) or vehicle.get(key.lstrip('0'))
    if not v:
        return {}
    out = {}
    for col, dst in (('購入部門', 'CDJS300'), ('セグメント', 'CDJS303')):
        x = str(v.get(col) or '').strip()
        if x:
            out[dst] = x
    return out


def machine_hint(line_ctx, code_rule, vehicle=None):
    """1行について、機械だけで言えることを返す。AIの答えを検算するのに使う。

    line_ctx: s40 が作る行（手書きコード / 請求書の記載 / vehicle_no を持つ）
    vehicle: 車両マスタ {車番: 行}。渡すと車番から部門・セグメントを引く
    戻り: {"CDJS300": ..., ...} の分かる範囲。根拠は hint_source に入れる。
    """
    out, why = {}, []
    note = line_ctx.get("請求書の記載") or {}
    # 行に付いたコード → 無ければ帳票に1つだけ付いたコード の順に見る。
    # 帳票に複数あるときは、どれがこの行か機械には決められないので使わない
    code = (line_ctx.get("手書きコード")
            or note.get("明細の手書きコード")
            or note.get("この請求書の手書きコード"))
    if code:
        got = from_code(code, code_rule)
        if got:
            out.update(got)
            why.append("手書きコード %s" % code)
    seg_word = note.get("経費区分")
    if seg_word and seg_word in TRUSTED_COST_SEGMENT:
        out.setdefault("CDJS303", TRUSTED_COST_SEGMENT[seg_word])
        why.append("経費区分 %s" % seg_word)
    # 車両マスタ。**手書きコードの次に強い**。セグメントは車で決まるので、
    # 事業所や経費区分より先に見る（事業所は「富山」でも検修なら別セグになる）
    vno = (line_ctx.get("vehicle_no") or note.get("車番")
           or note.get("この帳票の車番") or "")
    veh = from_vehicle(vno, vehicle)
    if veh:
        for k, v in veh.items():
            out.setdefault(k, v)
        why.append("車両マスタ %s" % norm_vehicle(vno))
    dept = office_to_dept(note.get("事業所") or "")
    if dept:
        out.setdefault("CDJS300", dept)
        why.append("事業所 %s" % note.get("事業所"))
    return out, " / ".join(why)


def compare_hint(answer, hint):
    """AIの答えと機械の見立てが食い違っている項目を返す。

    食い違い＝即誤りではない（事業所「富山」でも検修なら0007が正しい 等）。
    要確認に回すための材料として使う。
    """
    diff = []
    for k, v in (hint or {}).items():
        a = str(answer.get(k) or "")
        if a and str(v) and a != str(v):
            diff.append("%s: AI=%s 機械=%s" % (k, a, v))
    return diff
