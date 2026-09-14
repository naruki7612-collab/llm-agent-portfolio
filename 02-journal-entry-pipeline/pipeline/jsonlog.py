# -*- coding: utf-8 -*-
"""jsonlog — 各段が「何を受け取って何を返したか」を1行1件のJSONで残す。

なぜ要るか:
  段の途中結果は tmp/beta/ に出ているが、**ルームが変わると消える**。
  結果CSVには最終の4項目しか残らないので、あとから
  「配分表は渡っていたのか」「明細は何本あったのか」「画像は何枚か」が分からない。
  実際にこれで原因の特定に何往復もかかった（配分表を捨てていた件・車番が落ちた件）。

出力先:
  notes/システム/ログ.jsonl（JSON Lines・追記）
  **お客様が見る 02_仕訳データ/ には置かない。** デバッグ用なので システム/ に置く。

使い方:
  import jsonlog as JL
  JL.log("s35", 伝票=key, 明細=12, 配分表=10, 出力行=15)
  await JL.flush()        # バッチの終わりに1回。Noteへ追記する

  溜めてから1回で書く。1件ずつ書くと file_edit の往復で遅くなる。

止め方:
  config.DEBUG_LOG = False にすると log() は何もしない（flush も書かない）。
"""
import json, time

import config as C
# notefs は agent_sdk を要る。selftest は agent_sdk 無しで動かすので、
# **モジュールの先頭では import しない。** 実際に書き出すときだけ読む。

_BUF = []


def on():
    """ログを取る設定か。config に無ければ取らない。"""
    return bool(getattr(C, "DEBUG_LOG", False))


def log(step, **fields):
    """1件記録する。**例外を投げない。** ログのせいで本体を落とさない。"""
    if not on():
        return
    try:
        rec = {"時刻": time.strftime("%H:%M:%S"), "段": step}
        rec.update(fields)
        _BUF.append(rec)
    except Exception:                                     # noqa: BLE001
        pass


def count():
    return len(_BUF)


async def flush():
    """溜めたぶんを Note に追記して、バッファを空にする。戻り: 書いた件数。

    書けなくても例外にしない。**ログが原因で処理を止めない。**
    """
    if not on() or not _BUF:
        return 0
    n = len(_BUF)
    try:
        text = "".join(json.dumps(r, ensure_ascii=False) + "\n" for r in _BUF)
        import notefs as NF
        await NF.append_text(getattr(C, "LOG_PATH", C.SCRIPT_DIR + "/ログ.jsonl"), text)
    except Exception as e:                                # noqa: BLE001
        print("（ログの書き出しに失敗しました: %s。処理は続けます）" % e)
        return 0
    finally:
        _BUF.clear()
    return n


def dump_local(path="output/ログ.jsonl"):
    """サンドボックス側にも同じものを落とす（ダウンロードして見る用）。"""
    if not _BUF:
        return 0
    import os
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "a", encoding="utf-8") as f:
        for r in _BUF:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")
    return len(_BUF)
