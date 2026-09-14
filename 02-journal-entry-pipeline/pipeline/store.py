# -*- coding: utf-8 -*-
"""辞書（db テーブル）の読み書き。

AgentPlatform の db は演算子が8つしか無く（$eq/$ne/$in/$contains/$gt/$gte/$lt/$lte）、
$or も集約も使えない。なので「必要な範囲を find で引いて Python 側で束ねる」書き方に統一する。

注意:
  * すべて await
  * find は limit で切れるので、件数は count_documents で取る
  * 書き込みはシステムが承認カードを出す。approve() を先回りで呼ばない
"""
import asyncio, collections
import agent_sdk as at

import config as C
import rules as R


# ---------------------------------------------------------------- 読み出し
async def fetch_all(table, query=None, limit=5000, max_pages=20):
    """1テーブルを（必要なら複数ページに分けて）全部読む。"""
    t = at.db.table(table)
    out, after = [], None
    for _ in range(max_pages):
        page = await t.find(query or {}, limit=limit, after=after) if after \
            else await t.find(query or {}, limit=limit)
        recs = list(page.records)
        out.extend(recs)
        after = getattr(page, "next_after", None)
        if not after or not recs:
            break
    return out


_DICTS_CACHE = None


async def load_small_dicts():
    """毎回まるごと読んでよい小さい辞書をまとめて取る。

    s40 と s50 が別々に呼ぶので、プロセス内でキャッシュする（実行中に辞書は変わらない）。
    """
    global _DICTS_CACHE
    if _DICTS_CACHE is not None:
        return _DICTS_CACHE
    tables = [C.T_VEHICLE, C.T_CODE_RULE, C.T_GLOBAL_ITEM, C.T_ACCT_MGMT, C.T_MASTER]
    got = await asyncio.gather(*[fetch_all(t) for t in tables])
    veh, code_rule, global_item, acct_mgmt, master = got

    _DICTS_CACHE = {
        "vehicle":    {str(r.get("車番")): r for r in veh if r.get("車番")},
        "code_rule":  code_rule,
        "global_item": global_item,
        "acct_mgmt":  acct_mgmt,
        "master_set": {(str(r.get("種別")), str(r.get("コード"))) for r in master},
        "master":     master,
    }
    return _DICTS_CACHE


_JOURNAL_ROWS = None
_FACTS_CACHE = {}


async def _journal_rows():
    """beta_journal を1回だけまるごと読む。プロセス内でキャッシュ（1バッチ1回）。"""
    global _JOURNAL_ROWS
    if _JOURNAL_ROWS is None:
        try:
            _JOURNAL_ROWS = await fetch_all(C.T_JOURNAL)
        except Exception:                                 # noqa: BLE001
            _JOURNAL_ROWS = []
    return _JOURNAL_ROWS


async def load_journal_facts(before_ym=None):
    """仕訳実績から、辞書を引くだけでは分からない事実を作る。

    元の実装（5_検証用の実装）にはあって、AgentPlatform版に移植されていなかったもの：
      * 組合せ  … impl/invariants.py の C-4 / C-5
      * 車番の実績 … make_l3_ctx_v5e.py 110行目。**摘要から車番を正規表現で抜き**、
                    その車番が過去どの部門/セグで計上されたかを数える。
                    L3_SPEC_v5e は「車両マスタに無くても過去実績で必ず埋めにいく」と書いている

    before_ym を渡すと、その月より前の実績だけを使う（前向き）。
    """
    ck = str(before_ym or "")
    if ck in _FACTS_CACHE:
        return _FACTS_CACHE[ck]
    rows = await _journal_rows()
    acct_sub, dept_seg = set(), set()
    veh = {}
    for r in rows:
        ym = str(r.get("年月") or "")
        if before_ym and ym and ym >= str(before_ym):
            continue                                      # 前向き。未来の実績は使わない
        a, b = str(r.get("勘定科目") or ""), str(r.get("補助科目") or "")
        d, g = str(r.get("購入部門") or ""), str(r.get("セグメント") or "")
        if a and b:
            acct_sub.add((a, b))
        if d and g:
            dept_seg.add((d, g))
        v = R.vehicle_in_text(r.get("摘要"))
        if v and d and g:
            veh.setdefault(v, {}).setdefault((d, g), 0)
            veh[v][(d, g)] += 1
    got = {"acct_sub": acct_sub, "dept_seg": dept_seg,
           "車番の実績": {k: [{"部門": d, "セグ": g, "件数": n}
                          for (d, g), n in sorted(v.items(), key=lambda kv: -kv[1])]
                       for k, v in veh.items()}}
    _FACTS_CACHE[ck] = got
    return got


async def load_combos():
    """C-4 / C-5 に使う組合せ。load_journal_facts の一部。"""
    f = await load_journal_facts()
    return f["acct_sub"], f["dept_seg"]


_ALIAS_CACHE = None


async def load_alias():
    """別名テーブルを1回だけ全部読んで、Python 側で照合する。

    db の $contains は向きが逆（マスタ側が引数を含む判定）で、この照合ができない。
    """
    global _ALIAS_CACHE
    if _ALIAS_CACHE is None:
        rows = await fetch_all(C.T_PAYEE_ALIAS)
        _ALIAS_CACHE = [{
            "code": str(r.get("支払先コード") or ""),
            "key": R.norm_payee(r.get("表記") or ""),
            "n": int(r.get("根拠伝票数") or 0),
        } for r in rows if r.get("支払先コード") and r.get("表記")]
    return _ALIAS_CACHE


async def payee_code_of(name):
    """請求書の発行者名 → 支払先コード。

    ① 完全一致 → ② 部分一致（マスタ側が3文字以上のときだけ）の順で引く。
    違う支払先を引くと実績がまるごと誤って渡るので、短い一致は採らない。
    """
    key = R.norm_payee(name)
    if not key or len(key) < 2:
        return None
    alias = await load_alias()

    exact = [a for a in alias if a["key"] == key]
    if exact:
        return max(exact, key=lambda a: a["n"])["code"]

    part = [a for a in alias if len(a["key"]) >= 3 and (a["key"] in key or key in a["key"])]
    if not part:
        return None
    # 一致した文字数が長いものを優先し、同じなら実績の多いほう
    part.sort(key=lambda a: (-len(a["key"]), -a["n"]))
    best = part[0]
    if len({a["code"] for a in part if len(a["key"]) == len(best["key"])}) > 1:
        return None            # 同じ長さで複数の支払先に当たるなら引かない
    return best["code"]


_PC_CACHE = {}


async def payee_context(payee_code, before_ym, hist_limit=400):
    """この支払先の過去実績を集める。(支払先コード, 対象年月) ごとにキャッシュする。

    s35 と s40 が同じキーで別々に呼ぶので、キャッシュしないと2回ずつ引くことになる。
    """
    ck = (str(payee_code), str(before_ym))
    if ck in _PC_CACHE:
        return _PC_CACHE[ck]
    out = await _payee_context_uncached(payee_code, before_ym, hist_limit)
    _PC_CACHE[ck] = out
    return out


async def _payee_context_uncached(payee_code, before_ym, hist_limit=400):
    """1支払先の「過去の実績」を作る。before_ym より前のものだけを使う。"""
    if not before_ym:
        # 伝票の年月が読めなかったとき。未来の実績を混ぜないために、渡さない。
        return {"過去の仕訳実績": [], "請求書品名と4項目の対応": [],
                "直近5伝票の行テンプレート": [], "行数分布": []}

    jt = at.db.table(C.T_JOURNAL)
    it = at.db.table(C.T_ITEM_ACCOUNT)
    hist_page, item_page = await asyncio.gather(
        # 年月はdb側で数値型のため、$lt には文字列(before_ym)ではなく数値を渡す。
        # Python側の比較（下の str(...) < before_ym）は文字列同士のままでよい。
        jt.find({"支払先コード": payee_code, "年月": {"$lt": int(before_ym)}},
                sort=[{"field": "年月", "direction": -1}], limit=hist_limit),
        it.find({"支払先コード": payee_code}, sort=[{"field": "件数", "direction": -1}], limit=200),
    )
    hist = [r for r in hist_page.records]
    items = [r for r in item_page.records if str(r.get("最終年月") or "") < before_ym]

    # 摘要の重複を落として直近30行
    seen, rows = set(), []
    for r in sorted(hist, key=lambda r: str(r.get("年月") or ""), reverse=True):
        k = (R.norm_item(r.get("摘要")), r.get("購入部門"), r.get("勘定科目"),
             r.get("補助科目"), r.get("セグメント"))
        if k in seen:
            continue
        seen.add(k)
        rows.append({
            "年月": r.get("年月"), "ルール期": r.get("ルール期"),
            "摘要": (r.get("摘要") or "").strip(), "金額": r.get("本体金額"),
            "CDJS300": r.get("購入部門"), "CDJS301": r.get("勘定科目"),
            "CDJS302": r.get("補助科目"), "CDJS303": r.get("セグメント"),
            "伝票No": r.get("伝票No"),
        })
        if len(rows) >= 30:
            break

    # 直近5伝票の「行の並び」= L2 に見せる行テンプレート
    by_v = collections.defaultdict(list)
    for r in hist:
        by_v[str(r.get("伝票キー"))].append(r)
    recent = sorted(by_v.items(), key=lambda kv: str(kv[1][0].get("年月") or ""), reverse=True)[:5]
    templates = [{
        "年月": v[0].get("年月"), "行数": len(v),
        # 割る軸は部門×セグメントなので、4項目まで見せないと割り方を真似できない
        "行": [{"摘要": (x.get("摘要") or "").strip(), "金額": x.get("本体金額"),
                "部門": x.get("購入部門"), "科目": x.get("勘定科目"),
                "補助": x.get("補助科目"), "セグ": x.get("セグメント")} for x in v],
    } for _, v in recent]

    return {
        "過去の仕訳実績": rows,
        "請求書品名と4項目の対応": [{
            "品名": r.get("品名"), "用途区分": r.get("用途区分"), "車番": r.get("車番"),
            "CDJS300": r.get("購入部門"), "CDJS301": r.get("勘定科目"),
            "CDJS302": r.get("補助科目"), "CDJS303": r.get("セグメント"),
            "件数": r.get("件数"), "ルール期": r.get("ルール期"),
        } for r in items[:60]],
        "直近5伝票の行テンプレート": templates,
        "行数分布": sorted(collections.Counter(len(v) for v in by_v.values()).items()),
    }


# ---------------------------------------------------------------- 書き込み
_TYPES_CACHE = {}


async def field_types(table, keys=None):
    """既存の1行を見て、各列が数値なのか文字列なのかを返す。空なら {}。

    テーブルの型は最初に入れた値から決まっていて、あとから文字列で入れると
    `field '年月' expects a number` で**全件まとめて弾かれる**（実データで発生）。
    型を宣言している場所が他に無いので、入っている値から読む。

    keys を渡すとその列だけ `r.get(k)` で見る。**レコードが辞書とは限らない**ので、
    dict() に頼らない。ここで型が取れないと黙って変換しなくなり、
    同じエラーに戻ってしまう。
    """
    ck = (table, tuple(sorted(keys)) if keys else None)
    if ck in _TYPES_CACHE:
        return _TYPES_CACHE[ck]
    out = {}
    try:
        page = await at.db.table(table).find({}, limit=1)
        recs = list(page.records)
        if recs:
            r = recs[0]
            names = list(keys) if keys else None
            if names is None:
                try:
                    names = list(dict(r).keys())
                except Exception:                         # noqa: BLE001
                    names = []
            for k in names:
                try:
                    v = r.get(k) if hasattr(r, "get") else getattr(r, k, None)
                except Exception:                         # noqa: BLE001
                    continue
                if v is not None:
                    out[k] = type(v)
    except Exception:                                     # noqa: BLE001
        pass
    _TYPES_CACHE[ck] = out
    return out


async def insert_rows(table, rows, chunk=None):
    """insert_many を分割して投げる。1件でも配列で渡すこと。

    **入れる前にテーブルの型へ合わせる。** insert は全件まとめて検証され、
    1件でも型が違うと何も入らない。

    書き込みは AgentPlatform が承認カードを出す。approve() をここで呼んではいけない。
    """
    if not rows:
        return 0
    keys = sorted({k for r in rows for k in r})
    types = await field_types(table, keys)
    if types:
        rows = [R.coerce_types(r, types) for r in rows]
    n = chunk or C.DB_CHUNK
    done = 0
    for i in range(0, len(rows), n):
        await at.db.table(table).insert_many(rows[i:i + n])
        done += len(rows[i:i + n])
    return done


async def ensure_tables():
    """必要なテーブルが在るか見て、足りないものの名前を返す。"""
    have = set()
    try:
        for t in await at.db.list_tables():
            have.add(t if isinstance(t, str) else str(t.get("name") or t.get("table") or ""))
    except Exception:
        pass
    need = C.DICT_TABLES
    return [t for t in need if t not in have], sorted(have)
