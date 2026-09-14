# -*- coding: utf-8 -*-
"""PDF を1ページずつ画像（JPEG）にする。

PDFをそのまま渡すとテキスト層しか見ない挙動になり、判断の主役である手書きの
3段コード・赤チェック・欄外の書き込みが落ちるため。

使い方:
    import pdfimg
    paths = pdfimg.render("tmp/beta/src/x.pdf", "tmp/beta/img/x", dpi=150)
    # → ["tmp/beta/img/x/p-001.jpg", ...]（1始まりのページ番号）

変換の実装は環境によって使えるものが違うので、上から順に試します:
  1. pypdfium2   … pip の wheel だけで動く。いちばん確実
  2. PyMuPDF     … 速い
  3. pdftoppm    … poppler が入っていれば
  4. pdf2image   … poppler のラッパー
"""
import glob, os, shutil, subprocess, sys

_BACKEND = None


def _ensure(pkg, mod=None):
    mod = mod or pkg
    try:
        __import__(mod)
        return True
    except ImportError:
        pass
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "-q", pkg],
                       check=True, capture_output=True, timeout=300)
        __import__(mod)
        return True
    except Exception:
        return False


def backend():
    """使える変換方法の名前を返す。無ければ None。"""
    global _BACKEND
    if _BACKEND is not None:
        return _BACKEND
    if _ensure("pypdfium2"):
        _BACKEND = "pypdfium2"
    elif _ensure("PyMuPDF", "fitz"):
        _BACKEND = "fitz"
    elif shutil.which("pdftoppm"):
        _BACKEND = "pdftoppm"
    elif _ensure("pdf2image"):
        _BACKEND = "pdf2image"
    else:
        _BACKEND = None
    return _BACKEND


def page_count(pdf_path):
    try:
        from pypdf import PdfReader
    except ImportError:
        from PyPDF2 import PdfReader
    return len(PdfReader(pdf_path).pages)


def render(pdf_path, out_dir, dpi=150, quality=85, max_side=2400):
    """PDF を JPEG にする。既に出来ていれば作り直さない（冪等）。

    戻り: ページ番号順のパス一覧。ページ番号は1始まりで、ファイル名は p-001.jpg。
    """
    os.makedirs(out_dir, exist_ok=True)
    have = sorted(glob.glob(os.path.join(out_dir, "p-*.jpg")))
    try:
        n = page_count(pdf_path)
    except Exception:
        n = None
    if have and (n is None or len(have) == n):
        return have

    b = backend()
    if b is None:
        raise RuntimeError("PDFを画像にする方法がありません（pypdfium2 のインストールに失敗）")

    if b == "pypdfium2":
        import pypdfium2 as pdfium
        doc = pdfium.PdfDocument(pdf_path)
        scale = dpi / 72.0
        for i in range(len(doc)):
            p = os.path.join(out_dir, "p-%03d.jpg" % (i + 1))
            if os.path.exists(p):
                continue
            img = doc[i].render(scale=scale).to_pil().convert("RGB")
            img.thumbnail((max_side, max_side))
            img.save(p, "JPEG", quality=quality, optimize=True)
        doc.close()

    elif b == "fitz":
        import fitz
        doc = fitz.open(pdf_path)
        for i, page in enumerate(doc):
            p = os.path.join(out_dir, "p-%03d.jpg" % (i + 1))
            if os.path.exists(p):
                continue
            pix = page.get_pixmap(dpi=dpi)
            pix.save(p, jpg_quality=quality)
        doc.close()

    elif b == "pdftoppm":
        subprocess.run(["pdftoppm", "-jpeg", "-r", str(dpi), "-jpegopt", "quality=%d" % quality,
                        pdf_path, os.path.join(out_dir, "p")], check=True, timeout=600)
        # pdftoppm は p-1.jpg / p-01.jpg と桁が揃わないので、3桁に直す
        for f in sorted(glob.glob(os.path.join(out_dir, "p-*.jpg"))):
            base = os.path.basename(f)
            num = "".join(ch for ch in base if ch.isdigit())
            want = os.path.join(out_dir, "p-%03d.jpg" % int(num))
            if f != want:
                os.rename(f, want)

    else:
        from pdf2image import convert_from_path
        for i, img in enumerate(convert_from_path(pdf_path, dpi=dpi)):
            p = os.path.join(out_dir, "p-%03d.jpg" % (i + 1))
            if os.path.exists(p):
                continue
            img.convert("RGB").save(p, "JPEG", quality=quality, optimize=True)

    return sorted(glob.glob(os.path.join(out_dir, "p-*.jpg")))


def windows(n_pages, size, back=1, fwd=3):
    """ページを重なりつきの窓に切る。

    重なりを入れる理由:
      1つの請求書が窓の境目にまたがると、前半だけ／後半だけを見て
      別々の帳票として読んでしまいます（実データで2件起きていました）。

      * うしろに重ねる (fwd) … 担当の最後のページから始まる帳票が、
        次のページに続いていても丸ごと見える
      * まえに重ねる (back) … 担当の最初のページが「前の帳票の続き」なのか
        「新しい帳票の1枚目」なのかを判断できる

      そのうえで **担当範囲の外で始まった帳票は捨てる**ので、重なっても二重にはなりません。

    戻り: [(担当開始, 担当終了, 見せる開始, 見せる終了), ...] すべて1始まり・両端を含む
    """
    out = []
    start = 1
    while start <= n_pages:
        end = min(start + size - 1, n_pages)
        out.append((start, end, max(1, start - back), min(n_pages, end + fwd)))
        start = end + 1
    return out
