# -*- coding: utf-8 -*-
"""notefs — Noteに残す3ファイルだけを読み書きする。

書くのは2段階で、1段目だけでは保存されない。

    1段目 … at.file_create(sandbox_path=..., content=...) を通す。
        素の open(path, "w") はホスト側が書き込みを認識できない。
    2段目 … **notes_sync ツールを呼ぶ。ここまでやらないとNoteに保存されない。**
        agent_sdk に無いのでこのモジュールからは呼べない。アシスタントが
        code_execute の外で呼ぶ。

読むのは open() でよい（Noteのパスは触れば実体化される）。ただし
os.path.exists() / ls / find は stat ベースで実体化前のファイルを見落とすので、
**Noteのファイルの有無は exists() で判定しないこと**（read_text が None を
返すかどうかで見る）。
"""
import os

import agent_sdk as at


async def write_text(path, text):
    """Note上のパスにテキストを書く（上書き）。"""
    d = os.path.dirname(path)
    if d:
        os.makedirs(d, exist_ok=True)
    return await at.file_create(sandbox_path=path, content=text)


async def append_text(path, text):
    """Note上のファイルの**末尾に足す**。無ければ作る。戻り: 使った手段の名前。

    上書き（file_create）で積み増すには、まず既存を読めないといけない。読めないと
    そのとき書いた分だけになり、**前の実行の結果が消える**。追記なら読めなくても
    前の分は残る（最悪、同じ行が二重に入るだけで、消えることは無い）。

    使うのは **file_edit だけ**。file_insert は agent_sdk に無い（実機で確認）。
    末尾の数行を目印にして「その数行＋足したいもの」に差し替える。
    引数名は版によって違い、**知らない引数は黙って無視される**ので、
    呼べたかどうかでは判定しない。毎回読み返して確かめ、駄目なら全部書き直す。
    """
    if not text:
        return "何もしない"
    cur = read_text(path)
    if not cur:
        await write_text(path, text)
        return "file_create（新規）"
    # 末尾に改行が無いファイルにそのまま足すと、最後の行とくっついてしまう
    add = text if cur.endswith("\n") else "\n" + text

    edit = getattr(at, "file_edit", None)
    if edit is not None:
        for anchor in _tail_anchors(cur):
            for kw in ({"old_string": anchor, "new_string": anchor + add},
                       {"old_str": anchor, "new_str": anchor + add}):
                try:
                    await edit(sandbox_path=path, **kw)
                except Exception:                         # noqa: BLE001
                    continue
                got = read_text(path)
                if got and got.startswith(cur) and got.endswith(add):
                    return "file_edit(%s)" % ",".join(kw)

    await write_text(path, cur + add)
    return "file_create（全部書き直し）"


def _tail_anchors(cur):
    """末尾を指す目印の候補。**ファイル内で1回しか出てこない**ものだけ返す。

    file_edit は差し替えなので、目印が複数あるとどこを直すか決まらない。
    末尾1行から順に伸ばして、一意になったものを使う。改行も含める
    （含めないと、見出しだけのCSVで見出しの改行より前に足してしまう）。
    """
    lines = cur.splitlines(keepends=True)
    out = []
    for n in (1, 2, 3, 5, 10):
        if n > len(lines):
            break
        a = "".join(lines[-n:])
        if a and cur.count(a) == 1:
            out.append(a)
    return out


def read_text(path, default=None):
    """Note上のテキストを読む。無ければ default。

    os.path.exists() は実体化前のNoteファイルを見落とすので、
    存在確認はせずに開いて FileNotFoundError を拾う。

    **先頭のBOMは何個あっても全部落とす。** utf-8-sig が落とすのは1個だけで、
    2個付いたCSVを読むと1列目の見出しが "﻿伝票キー" になる。そうなると
    伝票キーで引く処理（Step 7 の積み増しと重複判定）が丸ごと外れ、
    DictWriter は「fieldnames に無い列がある」で落ちる。
    実際に出力ファイルでBOMが2個になっている例があった。
    """
    try:
        with open(path, encoding="utf-8-sig") as f:
            return f.read().lstrip("﻿")
    except FileNotFoundError:
        return default
