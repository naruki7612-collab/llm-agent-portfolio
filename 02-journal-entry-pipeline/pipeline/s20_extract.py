# -*- coding: utf-8 -*-
"""s20 — Step 2: 請求書を読んで、帳票と明細行を取り出す。AI・画像で渡す。

使い方:
    import sys; sys.path.insert(0, "notes/システム")
    import s20_extract as s
    await s.main(limit=3)     # 先頭3本のPDFだけ
    await s.main()            # 全件

やっていること:
  1. PDF を PAGE_DPI の JPEG に変換する
  2. **重なりつきの窓**でページ画像をまとめて渡す（担当8ページ ＋ 前1・後3ページ）
  3. 担当範囲の外で始まった帳票は捨てる → 重なっても二重にならない

PDFではなく画像で渡すのは、手書きの3段コード・赤チェック・欄外の書き込みが
判断の主役で、PDFのままだとテキスト層しか見ない挙動になるため。
窓を重ねるのは、請求書が境目にまたがると前半・後半が別の帳票になるため。

窓ごとに tmp/beta/extract/chunks/ に残す。一部が失敗しても、もう一度実行すれば
失敗した窓だけ処理する。全窓そろったPDFだけを1本にまとめる。
"""
import asyncio, glob, hashlib, json, os

import agent_sdk as at
from agent_sdk import ToolCallError
# Noteのパスは cwd 相対では解決されないので、自分のあるディレクトリを sys.path に通す
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))

import config as C
import prompts as P
import rules as R
import pdfimg


def _cache_name(tag):
    return hashlib.md5(tag.encode("utf-8")).hexdigest() + ".json"


def _img_dir(pdf_path):
    return os.path.join(C.D_IMG, hashlib.md5(os.path.basename(pdf_path).encode("utf-8")).hexdigest()[:10])


def _window_prompt(pages, own_from, own_to):
    """何ページ目の画像を何枚渡したかを、AIに文章で伝える。"""
    return (
        "\n\n## 渡した画像\n"
        + "\n".join("- %d枚目の画像 = このPDFの **%dページ目**" % (i + 1, p) for i, p in enumerate(pages))
        + "\n\n## あなたの担当\n"
        "**%dページ目から%dページ目まで**に**始まる**帳票だけを出してください。\n"
        "それ以外のページは、帳票の切れ目や続きを判断するための文脈として渡しています。\n"
        "- 担当ページより前から始まっている帳票は**出さない**（前の窓が担当します）\n"
        "- 担当ページで始まって次のページに続いている帳票は、**続きのページまで含めて**出す\n"
        "- `page_from` / `page_to` は必ず**このPDFの実際のページ番号**で書く\n" % (own_from, own_to)
    )


async def _one(pdf_path, pages, own_from, own_to, sem, tag):
    """1つの窓を読む。失敗は必ず (tag, None, 理由) で返す（gather 全体を落とさない）。"""
    try:
        img_dir = _img_dir(pdf_path)
        files = [os.path.join(img_dir, "p-%03d.jpg" % p) for p in pages]
        files = [f for f in files if os.path.exists(f)]
        if not files:
            return tag, None, "ページ画像がありません"
        async with sem:
            try:
                r = await at.llm_call(
                    P.EXTRACT_PROMPT + _window_prompt(pages, own_from, own_to),
                    file_paths=files, schema=P.EXTRACT_SCHEMA, model=C.MODEL_EXTRACT)
            except ToolCallError as e:
                return tag, None, "bridge: %s" % e
        if "error" in r:
            return tag, None, r["error"]
        data = r.get("data")
        if not isinstance(data, dict):
            return tag, None, "data が返ってこなかった"
        if "error" in data:
            # llm_call のエラーは r 直下ではなく data の中に入ってくることがある。
            # ここを見ないと documents:[] の顔をして失敗が握り潰される
            return tag, None, data["error"]

        # 担当範囲の外で始まった帳票は捨てる（重なりの二重取りを防ぐ）
        keep, dropped = [], 0
        for d in data.get("documents", []):
            pf = d.get("page_from")
            if pf is None or own_from <= int(pf) <= own_to:
                keep.append(d)
            else:
                dropped += 1
        ids = {d.get("doc_id") for d in keep}
        data["documents"] = keep
        data["line_items"] = [l for l in data.get("line_items", []) if l.get("doc_id") in ids]
        data["_dropped"] = dropped
        return tag, data, None
    except Exception as e:                                # noqa: BLE001
        return tag, None, "%s: %s" % (type(e).__name__, e)


async def main(limit=None, files=None):
    os.makedirs(C.D_EXTRACT, exist_ok=True)
    chunk_dir = os.path.join(C.D_EXTRACT, "chunks")
    os.makedirs(chunk_dir, exist_ok=True)

    # 「再送」を含むPDFは後ろに回す（s30の写し畳みで本編を残すため）
    src = files or R.resend_last(glob.glob(os.path.join(C.D_SRC, "*.pdf")))
    if limit:
        src = src[:limit]
    if not src:
        print("PDF がありません。先に s10_ingest.main() を実行してください")
        return

    # --- 1. 画像に変換する（既にあれば作り直さない）
    print("画像変換の方法:", pdfimg.backend())
    plan, npages, ng_img = {}, 0, []
    for p in src:
        stem = os.path.splitext(os.path.basename(p))[0]
        try:
            imgs = pdfimg.render(p, _img_dir(p), dpi=C.PAGE_DPI)
        except Exception as e:
            ng_img.append((stem, str(e)))
            continue
        n = len(imgs)
        npages += n
        plan[stem] = (p, pdfimg.windows(n, C.PAGE_WINDOW, C.PAGE_BACK, C.PAGE_FWD))
    print("--- PDF %d 件 / %d ページを画像化 ---" % (len(plan), npages))
    for stem, why in ng_img[:3]:
        print("画像化NG %s: %s" % (stem, why))

    # --- 2. 窓ごとに読む
    sem = asyncio.Semaphore(C.CONCURRENCY)
    tasks, cached = [], 0
    for stem, (path, wins) in plan.items():
        for own_from, own_to, show_from, show_to in wins:
            tag = "%s#%d-%d" % (stem, own_from, own_to)
            if os.path.exists(os.path.join(chunk_dir, _cache_name(tag))):
                cached += 1
                continue
            pages = list(range(show_from, show_to + 1))
            tasks.append(_one(path, pages, own_from, own_to, sem, tag))
    print("窓 %d（済 %d / 今回 %d）  1窓 = 担当%dページ + 前%d・後%dページ" % (
        sum(len(w) for _, w in plan.values()), cached, len(tasks),
        C.PAGE_WINDOW, C.PAGE_BACK, C.PAGE_FWD))

    results = await asyncio.gather(*tasks) if tasks else []
    ng, dropped = [], 0
    for tag, data, err in results:
        if err:
            ng.append((tag, err))
            continue
        dropped += data.pop("_dropped", 0)
        with open(os.path.join(chunk_dir, _cache_name(tag)), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)

    # --- 3. 全窓そろったPDFだけをまとめる
    incomplete = []
    for stem, (path, wins) in plan.items():
        tags = ["%s#%d-%d" % (stem, a, b) for a, b, _, _ in wins]
        have = [t for t in tags if os.path.exists(os.path.join(chunk_dir, _cache_name(t)))]
        if len(have) != len(tags):
            incomplete.append((stem, len(have), len(tags)))
            continue
        m = {"documents": [], "line_items": [], "source_pdf": stem, "pdf": path,
             "img_dir": _img_dir(path)}
        # doc_id / line_id は **PDFをまたいでも一意**にする。窓番号だけだと
        # 別のPDFの "w1-D01" と衝突し、下流が別ファイルの帳票を同じ伝票に混ぜる
        fid = hashlib.md5(stem.encode("utf-8")).hexdigest()[:6]
        for i, t in enumerate(tags):
            with open(os.path.join(chunk_dir, _cache_name(t)), encoding="utf-8") as f:
                data = json.load(f)
            pre = "%s-w%d-" % (fid, i + 1)
            for d in data.get("documents", []):
                d["doc_id"] = pre + str(d.get("doc_id"))
                m["documents"].append(d)
            for l in data.get("line_items", []):
                l["line_id"] = pre + str(l.get("line_id"))
                l["doc_id"] = pre + str(l.get("doc_id"))
                m["line_items"].append(l)
        m["documents"].sort(key=lambda d: (d.get("page_from") or 0, d["doc_id"]))
        with open(os.path.join(C.D_EXTRACT, stem + ".json"), "w", encoding="utf-8") as f:
            json.dump(m, f, ensure_ascii=False)

    # --- 4. 要約だけ出す
    all_files = sorted(glob.glob(os.path.join(C.D_EXTRACT, "*.json")))
    ndoc = nline = nlow = nhw = 0
    covered = 0
    for f in all_files:
        with open(f, encoding="utf-8") as fh:
            m = json.load(fh)
        ndoc += len(m["documents"])
        nline += len(m["line_items"])
        nlow += sum(1 for l in m["line_items"] if (l.get("confidence") or 1) < C.CONF_THRESHOLD)
        nhw += sum(1 for l in m["line_items"] if l.get("handwritten_account_code"))
        seen = set()
        for d in m["documents"]:
            seen.update(range(int(d.get("page_from") or 0), int(d.get("page_to") or 0) + 1))
        covered += len(seen - {0})
    print("--- 読めたPDF %d / 帳票 %d / 明細 %d 行 ---" % (len(all_files), ndoc, nline))
    print("手書きコードが読めた明細: %d 行 (%.0f%%)" % (nhw, 100 * nhw / nline if nline else 0))
    print("どの帳票にも入らなかったページ: %d（白紙なら正常）" % max(0, npages - covered))
    print("要確認（confidence<%.1f）: %d 行" % (C.CONF_THRESHOLD, nlow))
    if dropped:
        print("重なりで二重に読んだ帳票を捨てた数: %d" % dropped)
    for stem, h, t in incomplete[:5]:
        print("未完了 %s: %d/%d 窓（もう一度実行すれば残りを処理します）" % (stem, h, t))
    for tag, why in ng[:5]:
        print("NG %s: %s" % (tag, why))
    if len(ng) > 5:
        print("... 他 %d 窓が失敗" % (len(ng) - 5))
