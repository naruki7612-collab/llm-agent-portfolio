"""現在の実装が、思いつく限りのケースを正しく扱えるか検証する。

実カルテを用意しなくても確かめられるよう、_calculate_gyokyo_bunseki と
縦軸・横軸の判定ロジックに直接データを流して確認する。
"""
import sys

sys.path.insert(0, '<PROJECT_ROOT>')
import saimusha_kubun_hantei_ptc as m

PASS, FAIL = [], []


def check(no, name, cond, detail=''):
    (PASS if cond else FAIL).append(no)
    mark = '✓' if cond else '✗'
    print(f'  {mark} {no:2} {name}')
    if detail:
        print(f'        {detail}')


def zaimu(**kw):
    """3期分の財務データを作る。値は (前々期, 前期, 当期) のタプルで渡す。"""
    base = {
        '売上高': (1000, 1000, 1000), '月商': (83, 83, 83), '売上総利益': (200, 200, 200),
        '営業利益': (50, 50, 50), '経常利益': (50, 50, 50),
        '税引後利益': (30, 30, 30), '減価償却費': (20, 20, 20), '支払利息': (5, 5, 5),
        '総資産': (1000, 1000, 1000), '自己資本': (300, 300, 300), '資本金': (50, 50, 50),
        '繰越利益剰余金': (100, 100, 100),
        '流動資産': (400, 400, 400), '流動負債': (300, 300, 300),
        '受取手形': (10, 10, 10), '売掛金': (100, 100, 100), '棚卸資産': (50, 50, 50),
        '流動性支払手形': (10, 10, 10), '買掛金': (50, 50, 50), '割引・譲渡手形': (0, 0, 0),
        '有利子負債合計': (500, 500, 500), '劣後債務': (0, 0, 0),
    }
    base.update(kw)
    return {k: dict(zip(('前々期', '前期', '当期'), v)) for k, v in base.items()}


def shihyo(**kw):
    return {k: dict(zip(('前々期', '前期', '当期'), v)) for k, v in kw.items()}


print('■ 償還年数の算出不能パターン')

# 1 正常に算出できる
z = zaimu()
g = m._calculate_gyokyo_bunseki(z, shihyo(実質債務償還年数=(9.0, 9.0, 9.0)))
check(1, '正常に算出できる → 値がそのまま出る',
      g['実質債務償還年数']['当期'] == 9.0 and not g['実質債務償還年数']['確認事項'])

# 2 赤字で分母が0以下（カルテは0と記載）
z = zaimu(税引後利益=(30, -40, -50), 減価償却費=(20, 20, 20))
g = m._calculate_gyokyo_bunseki(z, shihyo(実質債務償還年数=(9.0, 0, 0)))
v = g['実質債務償還年数']
check(2, '赤字で分母0以下 → 算出不能・確認事項あり',
      v['当期'] is None and v['フラグ'] == m.FLAG_SANTEI_FUNO and '赤字により' in v['確認事項'],
      f"当期={v['当期']} 確認事項={v['確認事項'][:40]}")

# 3 前期も当期も算出不能（前期の0が残らないか）★今回直したところ
check(3, '前期の0も算出不能に変換される（当期だけでない）',
      g['実質債務償還年数']['前期'] is None,
      f"前期={g['実質債務償還年数']['前期']}")

# 4 分母がちょうど0
z = zaimu(税引後利益=(30, 30, -20), 減価償却費=(20, 20, 20))
g = m._calculate_gyokyo_bunseki(z, shihyo(実質債務償還年数=(9.0, 9.0, 0)))
check(4, '分母がちょうど0 → 算出不能',
      g['実質債務償還年数']['当期'] is None and g['実質債務償還年数']['フラグ'] == m.FLAG_SANTEI_FUNO)

# 5 実質債務がマイナス（実質無借金）→ 確認事項は出さない
z = zaimu(有利子負債合計=(100, 100, 100), 棚卸資産=(400, 400, 400))
g = m._calculate_gyokyo_bunseki(z, shihyo(実質債務償還年数=(0, 0, 0), 実質債務=(0, 0, 0)))
v = g['実質債務償還年数']
check(5, '実質無借金 → 算出不能だが確認事項なし',
      v['当期'] is None and v['確認事項'] == '' and '実質無借金' in v['注記'],
      f"注記={v['注記'][:50]}")

# 6 理由が特定できない0（分母>0かつ実質債務>0なのにカルテが0）★今回直したところ
z = zaimu()
g = m._calculate_gyokyo_bunseki(z, shihyo(実質債務償還年数=(9.0, 0, 9.0)))
v = g['実質債務償還年数']
check(6, '理由不明の0 → 算出不能として扱い理由を確認させる',
      v['前期'] is None,
      f"前期={v['前期']}")

# 7 カルテが空欄（None）
z = zaimu()
g = m._calculate_gyokyo_bunseki(z, shihyo(実質債務償還年数=(9.0, 9.0, None)))
v = g['実質債務償還年数']
check(7, 'カルテが空欄 → 自前計算で補完される（データ欠損にしない）',
      v['当期'] is not None, f"当期={v['当期']}（自前計算＝実質債務÷分母）")

print()
print('■ 実質債務の表示')

# 8 実質債務がカルテ上0で実際はマイナス → 注記で補足
z = zaimu(有利子負債合計=(100, 100, 100), 棚卸資産=(400, 400, 400))
g = m._calculate_gyokyo_bunseki(z, shihyo(実質債務=(0, 0, 0)))
check(8, '実質債務0（実際はマイナス）→ 注記で補足される',
      '実質無借金' in (g['実質債務']['注記'] or ''),
      f"注記={g['実質債務']['注記'][:60]}")

print()
print('■ 縦軸の判定（別表1）')


def row_of(zaimu_kw, shokan):
    """縦軸の候補を返す。extract_and_calculateの判定部分と同じ条件で再現する。"""
    z = zaimu(**zaimu_kw)
    cur = {k: v['当期'] for k, v in z.items()}
    saimu_choka = (cur.get('自己資本') or 0) < 0
    akaji = any((cur.get(k) or 0) < 0 for k in m.CRITICAL_PL_ITEMS[:3])
    base = 1 if not saimu_choka and not akaji else 2 if not saimu_choka else 3 if not akaji else 4
    rows = [base]
    if shokan is not None:
        if 10 < shokan <= 15:
            rows.append(5)
        elif shokan > 15:
            rows.append(6)
    return rows, max(rows, key=lambda r: m._kubun_rank(r, 'A'))


for no, name, kw, shokan, exp_rows, exp_pick in [
    (9,  '黒字・債務超過なし → 縦軸1', {}, 9.0, [1], 1),
    (10, '赤字 → 縦軸2', {'営業利益': (50, 50, -10)}, 9.0, [2], 2),
    (11, '債務超過 → 縦軸3', {'自己資本': (300, 300, -50)}, 9.0, [3], 3),
    (12, '債務超過＋赤字 → 縦軸4', {'自己資本': (300, 300, -50), '営業利益': (50, 50, -10)}, 9.0, [4], 4),
    (13, '償還年数10年超15年以内 → 縦軸1と5に該当し低位の5', {}, 12.5, [1, 5], 5),
    (14, '償還年数15年超 → 縦軸1と6に該当し低位の6', {}, 20.0, [1, 6], 6),
    (15, '償還年数ちょうど10年 → 縦軸5に落ちない', {}, 10.0, [1], 1),
    (16, '償還年数ちょうど15年 → 縦軸5', {}, 15.0, [1, 5], 5),
    (17, '償還年数が算出不能 → 縦軸5・6の判定を行わない', {}, None, [1], 1),
]:
    rows, pick = row_of(kw, shokan)
    check(no, name, rows == exp_rows and pick == exp_pick, f'候補={rows} 採用={pick}')

print()
print('■ 交点と注①（低位区分優先）')
for no, name, r, col, exp in [
    (18, '縦軸1×横軸A → 正常', 1, 'A', '正常'),
    (19, '縦軸5×横軸A → 正常2', 5, 'A', '正常2'),
    (20, '縦軸2×横軸A → 要注意（正常2）', 2, 'A', '要注意（正常2）'),
    (21, '縦軸4×横軸A → 要管理', 4, 'A', '要管理'),
]:
    got = m.MATRIX_TABLE[r][col]
    check(no, name, got == exp, f'実際={got}')

# 22 低位区分優先が正しい向きに働くか
check(22, '注①：縦軸1と5なら悪い方（5）を採る',
      max([1, 5], key=lambda r: m._kubun_rank(r, 'A')) == 5)

print()
print('■ 確認事項の出力先')
z = zaimu(繰越利益剰余金=(100, 100, 40))
g = m._calculate_gyokyo_bunseki(z, shihyo())
check(23, '繰越利益剰余金は確認事項の対象外（フラグは立つが宿題は出さない）',
      not g['繰越利益剰余金']['確認事項'] and g['繰越利益剰余金']['フラグ'] != '',
      f"フラグ={g['繰越利益剰余金']['フラグ']!r} 確認事項={g['繰越利益剰余金']['確認事項']!r}")
check(24, '対象外にした理由は様式に出力行がないこと',
      '繰越利益剰余金' not in m.GYOKYO_BUNSEKI_ROWS,
      '→ 縦軸判定には使用する')

print()
print('■ 数値書式の振り分け')
for no, name, item, exp in [
    (25, '金額はカンマ区切り整数', '売上高', m.NUMFMT_KINGAKU),
    (26, '比率は小数を残す', '自己資本比率', m.NUMFMT_SHOSU),
    (27, '回転期間は小数を残す', '売上債権回転期間', m.NUMFMT_SHOSU),
    (28, '償還年数は小数を残す', '実質債務償還年数', m.NUMFMT_SHOSU),
    (29, '正常運転資金は金額扱い', '正常運転資金', m.NUMFMT_KINGAKU),
]:
    got = m.NUMFMT_SHOSU if item in m.SHOSU_ITEMS else m.NUMFMT_KINGAKU
    check(no, name, got == exp, f'{item} → {got}')

print()
print('■ 検索クエリの生成')
for no, name, gyoshu, hinmoku, must in [
    (30, '通常の業種はそのまま', '一般貨物自動車運送業', ['リース'], f'一般貨物自動車運送業 {m.SEARCH_JUTSUGO}'),
    (31, '「その他」始まりは取扱品目1語に置き換える（業種名と連結しない）',
     'その他不動産', ['リース', '太陽光売電'], f'リース {m.SEARCH_JUTSUGO}'),
    (32, '業種が空ならクエリを作らない', None, None, None),
]:
    got = m._build_search_query(gyoshu, hinmoku)
    check(no, name, got == must, f'→ {got!r}')

print()
print('=' * 62)
print(f'合格 {len(PASS)}件 / 不合格 {len(FAIL)}件' + (f'　不合格: {FAIL}' if FAIL else '　（全ケース通過）'))
