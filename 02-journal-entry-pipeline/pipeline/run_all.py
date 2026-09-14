# -*- coding: utf-8 -*-
"""run_all — 取り込みから予測仕訳CSVまでを通しで実行する。

使い方:
    import sys; sys.path.insert(0, "notes/システム")
    import run_all
    await run_all.main(limit=3)        # 未処理の先頭3本だけ試す
    await run_all.main()               # 全件

未処理のPDFを BATCH_FILES 本ずつに分け、バッチごとに Step 2〜7 を通す。
Step 7 まで来たバッチの結果は結果CSVに残るので、次のバッチで15分の実行上限に
当たっても失われない。

すでに結果CSVの「参照元ファイル」列に出ているPDFは Step 2 に渡さない。
tmp/beta/ はルームが変わると消えるが、それを頼りにしていないので新しいルームでも
続きから進む。全部処理済みなら何もせずに戻る。やり直すときは reset()。

file_create で書いただけではNoteに保存されない。アシスタントが notes_sync を
呼ぶまでバックエンドに届かない（このスクリプトからは呼べない）。
"""
import glob
import os
import time

# Noteのパスは cwd 相対では解決されないので、自分のあるディレクトリを sys.path に通す
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))

import config as C
import notefs as NF
import rules as R
import jsonlog as JL

import s10_ingest, s20_extract, s30_vouchers, s32_group_llm, s35_split, s40_classify, s50_export


def _input_pdfs(src_dir=None):
    """入力PDFの名前（拡張子なし）の集合。証憑フォルダ・uploads/・取り込み済みを見る。"""
    dirs = [src_dir] if src_dir else [C.D_NOTE_SRC, "uploads"]
    out = set()
    for d in dirs + [C.D_SRC]:
        for p in glob.glob(os.path.join(d, "*.pdf")):
            out.add(os.path.splitext(os.path.basename(p))[0])
    return out


def _done_files():
    """すでに結果が出ているPDFの名前。予測仕訳CSVの「参照元ファイル」列から求める。

    完了記録の専用ファイルは作らない（結果そのものが記録なのでずれない）。
    ルームをまたいだ再開はこれが土台。
    """
    import csv, io
    text = NF.read_text(C.OUT1_NOTE)
    if not text:
        return set()
    try:
        return {(r.get("参照元ファイル") or "").strip()
                for r in csv.DictReader(io.StringIO(text))} - {""}
    except Exception:                                     # noqa: BLE001
        return set()


def _clear_stage_dirs():
    """Step 2〜6 の途中結果を消す。バッチの切り替えで使う。ページ画像は消さない。"""
    import shutil
    for d in (C.D_EXTRACT, C.D_VOUCH, C.D_SPLIT, C.D_CLASS):
        shutil.rmtree(d, ignore_errors=True)


async def _write_status(text):
    """今どこまで進んだかをNoteに書く（毎回上書き）。"""
    await NF.write_text(C.STATUS_PATH, text)


async def main(limit=None, src_dir=None, stamp=None):
    t0 = time.time()

    # **走り出す前に置き場所を確かめる。** NOTE_ROOT を間違えると全段が
    # 「ファイルがありません」で落ちる。15分かけてから気づくのを防ぐ関門。
    ng = C.check_paths()
    if ng and not src_dir:
        print("★ Note のフォルダが見つかりません。config.NOTE_ROOT を確かめてください。")
        print("   いまの設定: NOTE_ROOT = %r" % C.NOTE_ROOT)
        for x in ng:
            print("   見つからない: %s" % x)
        found = C.where_is_note()
        if found:
            print("   実際に在りそうな場所: %s" % ", ".join(found))
            print("   → config.py の NOTE_ROOT をそこに合わせてください")
        return

    # 全部出ていれば何もせずに戻る（定期実行の入口）
    todo = _input_pdfs(src_dir)
    finished = _done_files()
    if todo and finished and not (todo - finished):
        print("★ 入力PDF %d件はすべて処理済みです。実行することはありません。" % len(todo))
        print("  結果:   %s" % C.OUT1_NOTE)
        print("  要確認: %s" % C.OUT2_NOTE)
        print("  やり直したい場合は await run_all.reset() を実行してください。")
        return
    if finished:
        print("--- 未処理 %d件 / 処理済み %d件 ---" % (len(todo - finished), len(finished)))

    stages = ["1/7 取り込み", "2/7 請求書を読む（L1）", "3/7 伝票に括る（S3・機械）",
              "4/7 伝票の括りをAIが見直す（S3-L）", "5/7 仕訳行に割る（L2）",
              "6/7 4項目と摘要を決める（L3/S7）", "7/7 検算してCSVを出す（S8）"]

    async def stage(name, bi=0, bn=0):
        print("\n===== %s （経過 %.0f 秒）=====" % (name, time.time() - t0))
        rest = sorted(todo - finished)
        await _write_status(
            "=== 請求書自動仕訳 進捗状況 ===\n"
            "最終更新: %s\n\n"
            "いま %s を実行中（%s経過 %.0f 秒）\n"
            "入力PDF %d件 / 処理済み %d件 / 残り %d件\n\n"
            "--- 処理済み（結果CSVに出ているPDF）---\n%s\n"
            "--- 残り ---\n%s\n\n"
            "→ ここで止まっていたら、**新しいルームで** もう一度\n"
            "  await run_all.main() を実行してください。\n"
            "  処理済みのPDFは読み直さず、残りだけを処理します。\n"
            % (time.strftime("%Y-%m-%d %H:%M:%S"), name,
               ("バッチ %d/%d・" % (bi, bn)) if bn else "", time.time() - t0,
               len(todo), len(finished), len(rest),
               "\n".join("  ○ " + x for x in sorted(finished)) or "  （まだ無し）",
               "\n".join("  ・" + x for x in rest) or "  （無し）"))

    # Step 1 は毎回・全件やる。**飛ばしてはいけない。**
    # 新しいルームでは tmp/beta/src が空で、下の pending が0件になり何も処理されない。
    await stage(stages[0])
    await s10_ingest.main(src_dir=src_dir)

    # 残りのPDFを BATCH_FILES 件ずつに割る。
    # 並びは「再送」を最後に回す。重複を畳むとき **先に来たほうが残る** ので、
    # 再送が先のバッチに入ると本編のほうが落ちてしまう。
    pending = R.resend_last(p for p in glob.glob(os.path.join(C.D_SRC, "*.pdf"))
                            if os.path.splitext(os.path.basename(p))[0] not in finished)
    if limit:
        pending = pending[:limit]
    if not pending:
        print("読み取る対象がありません（すべて結果CSVに出ています）。")
        return
    batches = [pending[i:i + C.BATCH_FILES] for i in range(0, len(pending), C.BATCH_FILES)]
    print("--- 未処理 %d件 を %d件ずつ %dバッチで処理します（処理済み %d件は飛ばす）---"
          % (len(pending), C.BATCH_FILES, len(batches), len(finished)))

    for bi, batch in enumerate(batches, 1):
        print("\n########## バッチ %d/%d（%d件）##########" % (bi, len(batches), len(batch)))
        for p in batch:
            print("    %s" % os.path.splitext(os.path.basename(p))[0])

        # 前のバッチの途中結果を消す。残すと Step 3〜6 が前のバッチのぶんまで見てしまう
        _clear_stage_dirs()

        await stage(stages[1], bi, len(batches))
        await s20_extract.main(files=batch)

        await stage(stages[2], bi, len(batches))
        await s30_vouchers.main()

        await stage(stages[3], bi, len(batches))
        await s32_group_llm.main()

        await stage(stages[4], bi, len(batches))
        await s35_split.main()

        await stage(stages[5], bi, len(batches))
        await s40_classify.main()

        # ここで結果CSVに書かれる。**このバッチはここまで来れば残る。**
        # 次のバッチで15分に当たっても、このバッチの結果は失われない。
        await stage(stages[6], bi, len(batches))
        await s50_export.main(stamp=stamp)

        n = await JL.flush()          # このバッチのログをNoteへ追記する
        if n:
            print("  （デバッグログ %d件を %s に追記しました）" % (n, C.LOG_PATH))

        finished = _done_files()
        print("--- バッチ %d/%d 完了。結果CSVに出たPDF %d件 / 残り %d件 ---"
              % (bi, len(batches), len(finished), len(todo - finished)))

    rest = todo - finished

    print("\n===== 完了（合計 %.0f 秒）=====" % (time.time() - t0))
    if rest:
        print("→ まだ %d件 残っています。**新しいルームで** await run_all.main() を"
              "実行してください。" % len(rest))
        for s in sorted(rest)[:5]:
            print("    %s" % s)
    else:
        print("★ 全%d件の処理が完了しました。続きの実行は不要です。" % len(todo))

    await _write_status(
        "=== 請求書自動仕訳 進捗状況 ===\n"
        "最終更新: %s\n\n"
        "入力PDF %d件 / 処理済み %d件 / 残り %d件\n\n"
        "--- 処理済み（結果CSVに出ているPDF）---\n%s\n"
        "--- 残り ---\n%s\n\n%s\n"
        "結果:   %s\n要確認: %s\n"
        % (time.strftime("%Y-%m-%d %H:%M:%S"), len(todo), len(finished), len(rest),
           "\n".join("  ○ " + x for x in sorted(finished)) or "  （まだ無し）",
           "\n".join("  ・" + x for x in sorted(rest)) or "  （無し）",
           "★ 全ファイルの処理が完了しました。続きの実行は不要です。" if not rest and not limit
           else "→ 残りがあります。**新しいルームで** await run_all.main() を実行してください。\n"
                "  処理済みのPDFは読み直さず、残りだけを処理します。",
           C.OUT1_NOTE, C.OUT2_NOTE))

    print("次にやること:")
    print("  1. %s を人が見て直す" % C.OUT2_NOTE)
    print("  2. 直したCSVを %s に置いて s60_feedback.main() を実行（辞書が更新されます）"
          % C.D_FEEDBACK)


async def reset():
    """全部やり直す。途中結果と結果CSV2本の**両方**を消す。

    片方だけでは噛み合わない。結果CSVだけ消すと段のキャッシュが残って結果が出ず、
    段の結果だけ消すとCSVの参照元ファイルを見て「処理済み」と判定される。
    """
    import csv, io, shutil
    from s50_export import COLS

    n = 0
    for d in (C.D_EXTRACT, C.D_VOUCH, C.D_SPLIT, C.D_CLASS):
        if os.path.isdir(d):
            n += sum(len(fs) for _, _, fs in os.walk(d))
            shutil.rmtree(d, ignore_errors=True)
    header = io.StringIO()
    csv.DictWriter(header, fieldnames=COLS, lineterminator="\n").writeheader()
    await NF.write_text(C.OUT1_NOTE, "﻿" + header.getvalue())
    await NF.write_text(C.OUT2_NOTE, "﻿" + header.getvalue())
    await NF.write_text(C.STATUS_PATH, "（reset しました）\n")
    print("途中結果 %d件を消し、Noteの結果CSV2本を空にしました。" % n)
    print("次に await run_all.main() を実行すると最初からやり直します。")
