# -*- coding: utf-8 -*-
"""s10 — Step 1: 新しい請求書PDFを tmp/beta/src/ に取り込む。

使い方:
    import sys; sys.path.insert(0, "notes/システム")
    import s10_ingest as s
    await s.show_tools()      # OneDriveを使う場合。最初に1回だけ引数を確認する
    await s.main()

取り込み元は、config.ONEDRIVE を埋めていれば OneDrive、埋めていなければ
Noteの証憑フォルダ（config.D_NOTE_SRC）と uploads/ の両方。
**証憑はNoteに置くのが本筋**（uploads/ はサンドボックスのリセットで消える）。
"""
import asyncio, glob, json, os, re, shutil

import agent_sdk as at
from agent_sdk import ToolCallError
# Noteのパスは cwd 相対では解決されないので、自分のあるディレクトリを sys.path に通す
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))

import config as C


def _safe(name):
    """ファイル名を安全な形にする（日本語は残す）。"""
    return re.sub(r'[\\/:*?"<>|]+', "_", name).strip() or "file"


async def show_tools(keyword="DRIVE"):
    """使える Composio ツールと引数の形を出す。

    ツール名・引数はテナントによって違うので、確認して config.ONEDRIVE に書き写す。
    """
    try:
        names = await at.list_composio_tools()
    except ToolCallError as e:
        print("Composio ツールが使えません:", e)
        return []
    hit = sorted(n for n in names if keyword.upper() in n.upper())
    print("候補 %d 件:" % len(hit), hit)
    if not hit:
        print("（見つからないときは keyword='SHAREPOINT' や 'FILE' でも試す）")
        return []
    try:
        specs = await at.list_composio_tool_specs()
    except ToolCallError as e:
        print("spec 取得に失敗:", e)
        return hit
    for sp in specs:
        if sp.get("name") in hit:
            props = ((sp.get("inputSchema") or {}).get("properties") or {})
            print("- %s: %s" % (sp["name"], list(props)[:12]))
    return hit


async def _from_onedrive(dst):
    cfg = C.ONEDRIVE
    list_tool = getattr(at, cfg["list_tool"])
    res = await list_tool(**(cfg.get("list_args") or {}))

    # Composio の戻りは実装差があるので、配列らしいものを総当りで探す
    items = None
    for cand in (res, res.get("data") if isinstance(res, dict) else None,
                 (res.get("data") or {}).get("items") if isinstance(res, dict) else None,
                 (res.get("data") or {}).get("value") if isinstance(res, dict) else None):
        if isinstance(cand, list):
            items = cand
            break
    if items is None:
        # 生の JSON は stdout に流さず tmp に落としてパスだけ出す（コンテキスト節約）
        os.makedirs(C.JOB, exist_ok=True)
        dump = os.path.join(C.JOB, "onedrive_list_raw.json")
        with open(dump, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print("一覧の形が想定と違います。キー: %s" % (
            list(res)[:10] if isinstance(res, dict) else type(res).__name__))
        print("中身は %s に書きました。file_read で見て s10 を直してください" % dump)
        return []

    pdfs = [x for x in items if str(x.get("name") or x.get("fileName") or "").lower().endswith(".pdf")]
    print("OneDrive の PDF: %d 件" % len(pdfs))

    dl_tool = getattr(at, cfg["download_tool"])
    sem = asyncio.Semaphore(C.CONCURRENCY)

    async def one(it):
        name = _safe(str(it.get("name") or it.get("fileName")))
        out = os.path.join(dst, name)
        if os.path.exists(out):
            return out, None
        async with sem:
            try:
                r = await dl_tool(**{cfg["download_arg"]: it.get(cfg["download_key"])})
            except ToolCallError as e:
                return None, (name, str(e))
        path = None
        if isinstance(r, dict):
            d = r.get("data") if isinstance(r.get("data"), dict) else r
            path = d.get("file") or d.get("path") or d.get("local_path")
        if path and os.path.exists(path):
            shutil.copyfile(path, out)
            return out, None
        return None, (name, "ダウンロード結果にファイルパスが無い")

    results = await asyncio.gather(
        *[one(it) for it in sorted(pdfs, key=lambda x: str(x.get("name") or ""))])
    got = [p for p, _ in results if p]
    ng = [e for _, e in results if e]
    for n, why in ng[:5]:
        print("NG %s: %s" % (n, why))
    if len(ng) > 5:
        print("... 他 %d 件失敗" % (len(ng) - 5))
    return got


async def main(src_dir=None):
    """証憑PDFを tmp/beta/src/ に並べる。

    src_dir を省略すると config.D_NOTE_SRC と uploads/ の両方を見る。
    """
    dst = C.D_SRC
    os.makedirs(dst, exist_ok=True)

    cfg = C.ONEDRIVE
    if cfg.get("list_tool") and cfg.get("download_tool"):
        got = await _from_onedrive(dst)
    else:
        # glob.glob() はNoteのパスを実体化するので、Note配下もそのまま一覧できる
        dirs = [src_dir] if src_dir else [C.D_NOTE_SRC, "uploads"]
        got, found = [], []
        for d in dirs:
            ps = sorted(glob.glob(os.path.join(d, "*.pdf")))
            found.append((d, len(ps)))
            for p in ps:
                out = os.path.join(dst, _safe(os.path.basename(p)))
                if not os.path.exists(out):
                    shutil.copyfile(p, out)
                got.append(out)
        for d, n in found:
            print("  %s から %d 件" % (d, n))
        if not got:
            print("PDFが1件も見つかりません。証憑を %s に置いてください。" % C.D_NOTE_SRC)

    files = sorted(glob.glob(os.path.join(dst, "*.pdf")))
    total = sum(os.path.getsize(f) for f in files)
    print("--- 取り込み済み %d 件 / 合計 %.1f MB → %s ---" % (len(files), total / 1024 / 1024, dst))
    for f in files[:10]:
        print("  %s (%.1f MB)" % (os.path.basename(f), os.path.getsize(f) / 1024 / 1024))
    if len(files) > 10:
        print("  ... 他 %d 件" % (len(files) - 10))
    return files
