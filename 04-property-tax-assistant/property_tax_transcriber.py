"""固都税アシスタント 転記エンジン

納税通知書PDF（課税明細書）から物件ごとの
  評価額 / 固定資産税課税標準額 / 都市計画税課税標準額
を読み取り、Excelの自治体シートの該当行（H / I / K 列）へ数値で転記する。

■ 設計方針
  - fail closed: 全件が検証を通ったときだけ書き込む。1件でも落ちたら1セルも書かない
  - 数値の計算は必ずコードで行う（LLMに手計算させない）
  - LLMに任せるのは「画像から数字を読む」ことだけ
  - 見出し・数式・書式・対象外の行/シートには一切触れない

■ 使い方（code_execute から）
    from kotozei_tenki import run
    report = await run(pdf_path, xlsx_path, sheet_name, out_path)
    print(report)

  個別に動かす場合:
    pages  = kt.render_pdf(pdf_path, workdir)        # PDF→画像（/Rotate を適用して正立させる）
    plan   = kt.dry_run(records, covers, xlsx, sheet)  # 検証のみ。ファイルには触れない
    kt.commit(plan, xlsx, out_path)                  # 全部緑のときだけ書く

■ 前提
  - pypdfium2（通常は pdfplumber>=0.11 の必須依存として入っている）。
    無い環境では NEED_PDF_SKILL を投げるので `load_skill("pdf")` してから再実行する
  - openpyxl / pdfplumber は sandbox にプリインストール済み

■ llm_call のスキーマについて
  Gemini の構造化出力は ``type`` に配列（union）を受け付けない。
  ``{"type": ["integer", "null"]}`` は 422 で落ちるので、
  ``{"type": "integer", "nullable": True}`` と書くこと。
"""

from __future__ import annotations

import logging
import math
import re
import sys
import time
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

# ---------------------------------------------------------------------------
# 0. ログ
# ---------------------------------------------------------------------------
#
# code_execute は stdout しか返さないため、stdout とファイルの両方に出す。
# ファイル (``{workdir}/kotozei.log``) は stdout が切り詰められたときの保険。
# 実行のたびに上書きせず追記するので、複数回試したときの比較ができる。

LOG = logging.getLogger("kotozei")
_LOG_READY = False


def setup_logging(workdir: str = "tmp/kotozei", level: int = logging.INFO) -> logging.Logger:
    """ロガーを1度だけ構成する。``code_execute`` はセッション内で状態が残るので冪等にする。"""
    global _LOG_READY
    if _LOG_READY:
        LOG.setLevel(level)
        return LOG
    Path(workdir).mkdir(parents=True, exist_ok=True)
    fmt = logging.Formatter("%(asctime)s %(levelname)-5s %(message)s", datefmt="%H:%M:%S")
    # **stdout に出す**。logging.StreamHandler の既定は stderr で、agent_platform は stderr に
    # 出力があると実行ブロックを「実行エラー」と表示してしまう（正常終了でも赤くなる）。
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    fh = logging.FileHandler(str(Path(workdir) / "kotozei.log"), encoding="utf-8")
    fh.setFormatter(fmt)
    LOG.handlers.clear()
    LOG.addHandler(sh)
    LOG.addHandler(fh)
    LOG.setLevel(level)
    LOG.propagate = False
    _LOG_READY = True
    LOG.info("=" * 60)
    LOG.info("固都税 転記エンジン 開始")
    return LOG

# ---------------------------------------------------------------------------
# 定数
# ---------------------------------------------------------------------------

# 転記先の列。Excelの3行目ヘッダに対応する
COL_KAKAKU = "H"  # 価格（評価額）
COL_KOTEI = "I"  # 固定資産税課税標準額
COL_TOKEI = "K"  # 都市計画税課税標準額

# 転記対象のブロック。B列が「納付書①」「納付書②」の行の直下〜次のブロック手前まで。
# 「納付書③」（償却資産）以下は対象外。
BLOCK_LABELS_TARGET = ("納付書①", "納付書②")
BLOCK_LABEL_EXCLUDED = "納付書③"

# 税率はシート上の I2 / K2 から読む（ハードコードしない）。読めない場合のフォールバック。
FALLBACK_RATE_KOTEI = 0.014
FALLBACK_RATE_TOKEI = 0.003

RENDER_DPI = 300

# 読み取りモデル。行単位の検算と実データ照合が精度の合否判定になるので、
# 「全件通る一番安いモデル」を客観的に選べる。
#
# 実測（福岡市博多区・明細56件、2026-08-17。各3回、判定は数値のみで行った）:
#   モデル                    合格率  平均クレジット
#   gemini-3.1-pro-preview     3/3      108.4   ← 採用（通算でも数値の誤読ゼロ）
#   gemini-3.5-flash           3/3       72.4   ← 通算 5/6。納品後の最適化候補
#   gemini-3.1-flash-lite      1/3       78.0   ← 脱落。単価1/6でも高く、精度も落ちる
#
# flash-lite は出力トークンが増えて単価の安さが打ち消される。
# 崩れ方は桁の誤読ではなく「都市計画税の欄に固定資産税の値を入れる」列の取り違えで、
# 住宅用地特例で都計＝固定×2になる行に集中した。
DEFAULT_READ_MODEL = "gemini/gemini-3.1-pro-preview"

# USD / 1M トークン。コスト実感をログに出すためのもの（正確な課金は backend 側で行われる）
MODEL_RATES: dict[str, tuple[float, float]] = {
    "gemini/gemini-3.1-pro-preview": (2.00, 12.00),
    "gemini/gemini-3-pro-preview": (2.00, 12.00),
    "gemini/gemini-3.5-flash": (1.50, 9.00),
    "gemini/gemini-3-flash-preview": (0.50, 3.00),
    "gemini/gemini-3.1-flash-lite": (0.25, 1.50),
    "anthropic/claude-opus-4-7": (5.00, 25.00),
    "anthropic/claude-sonnet-5": (2.00, 10.00),
    "anthropic/claude-haiku-4-5": (1.00, 5.00),
    "openai/gpt-5.5": (5.00, 30.00),
    "openai/gpt-5.4-mini": (0.75, 4.50),
}


# ---------------------------------------------------------------------------
# 1. PDF → 画像
# ---------------------------------------------------------------------------


def render_pdf(pdf_path: str, workdir: str, dpi: int = RENDER_DPI) -> list[str]:
    """PDFの各ページを画像に描画して保存し、そのパス一覧を返す。

    納税通知書PDFは MRC (Mixed Raster Content) 圧縮されており、文字が
    「1bitマスク層」と「JPEG背景層」に分散して格納されている。
    pypdf / pdfplumber で埋め込み画像を取り出すと **数字の桁が虫食いで欠ける**。
    必ずページ全体をレンダリングして合成した画像を使うこと。

    ページの向きについて:
    課税明細書のページには PDF の ``/Rotate 90`` が付いていることがあるが、
    **pypdfium2 の render() は /Rotate を適用しない**（MediaBox のまま描画する）。
    そのため描画後に自分で回転させる。回さないと表が90度倒れたまま読み取りに渡ることになる。
    ページの縦横比では判定できない（正立の縦長ページと、倒れた縦長ページが混在するため）。

    pypdfium2 は pdfplumber (>=0.11) の必須依存なので、通常はサンドボックスに入っている。
    入っていない環境向けに、明示的な指示付きで落とす。
    """
    try:
        import pypdfium2 as pdfium
    except ImportError as e:  # pragma: no cover - 環境依存
        raise RuntimeError(
            "NEED_PDF_SKILL: pypdfium2 が見つかりません。"
            '`load_skill("pdf")` を実行してから、もう一度このスクリプトを呼び直してください。'
        ) from e

    Path(workdir).mkdir(parents=True, exist_ok=True)
    t0 = time.monotonic()
    doc = pdfium.PdfDocument(pdf_path)
    LOG.info("PDF描画 開始: %s (%d頁, %ddpi)", pdf_path, len(doc), dpi)
    out: list[str] = []
    for i in range(len(doc)):
        page = doc[i]
        img = page.render(scale=dpi / 72).to_pil().convert("L")
        try:
            rot = int(page.get_rotation() or 0) % 360
        except Exception:
            rot = 0
        if rot:
            img = img.rotate(rot, expand=True)
        path = str(Path(workdir) / f"page{i + 1:02d}.png")
        img.save(path)
        out.append(path)
        LOG.debug("  p%02d %dx%d /Rotate=%d -> %s", i + 1, img.width, img.height, rot, path)
    LOG.info("PDF描画 完了: %d頁 (%.1f秒)", len(out), time.monotonic() - t0)
    return out


def page_texts(pdf_path: str) -> list[str]:
    """各ページのテキスト層を返す。OCR由来で文字化けが多いが、頁の仕分けには使える。"""
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        return [(p.extract_text() or "") for p in pdf.pages]


def classify_pages(pdf_path: str) -> list[dict]:
    """【調査用】テキスト層のキーワードでページを仕分ける。**本処理では使わない。**

    テキスト層はOCR由来で文字化けが激しく、この方法では取りこぼす（実測: 宛名ページを
    課税明細書と誤判定、納税通知書ページを取りこぼし）。
    本処理の仕分けは ``run()`` が画像を見て行う（PAGE_SCHEMA の page_kind）。
    この関数はPDFの構成を人が確認するための補助として残している。
    """
    kinds = []
    for i, text in enumerate(page_texts(pdf_path), start=1):
        flat = re.sub(r"\s+", "", text)
        is_shokyaku = "償却資産" in flat
        is_meisai = "課税明細書" in flat
        is_tsuchi = "納税通知書" in flat
        if is_shokyaku:
            kind = "shokyaku"  # 償却資産 → 対象外
        elif is_meisai:
            kind = "meisai"  # 課税明細書 → 抽出対象
        elif is_tsuchi:
            kind = "cover"  # 納税通知書（表紙）→ 合計検算に使う
        else:
            kind = "unknown"
        kinds.append({"page": i, "kind": kind, "chars": len(flat)})
    return kinds


# ---------------------------------------------------------------------------
# 2. 抽出結果のスキーマ（llm_call に渡す JSON Schema）
# ---------------------------------------------------------------------------

PAGE_SCHEMA: dict = {
    "type": "object",
    "properties": {
        "page_kind": {
            "type": "string",
            "enum": ["meisai", "cover", "shokyaku", "other"],
            "description": (
                "meisai=固定資産税・都市計画税の課税明細書（物件の明細が並ぶ表）／"
                "cover=固定資産税・都市計画税の納税通知書（課税標準額の合計が載る）／"
                "shokyaku=償却資産に関する頁（表題に『償却資産』とある）／"
                "other=宛名・案内など上記以外"
            ),
        },
        "cover": {
            "type": "object",
            "description": "page_kind が cover のときだけ埋める",
            "properties": {
                "kotei_tochi": {"type": "integer", "nullable": True, "description": "課税標準額 土地（固定資産税）"},
                "kotei_kaoku": {"type": "integer", "nullable": True, "description": "課税標準額 家屋（固定資産税）"},
                "kotei_goukei": {"type": "integer", "nullable": True, "description": "課税標準額 合計(ア)（固定資産税）"},
                "tokei_tochi": {"type": "integer", "nullable": True, "description": "課税標準額 土地（都市計画税）"},
                "tokei_kaoku": {"type": "integer", "nullable": True, "description": "課税標準額 家屋（都市計画税）"},
                "tokei_goukei": {"type": "integer", "nullable": True, "description": "課税標準額 合計(ア)（都市計画税）"},
            },
        },
        "rows": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "excel_row": {
                        "type": "integer",
                        "nullable": True,
                        "description": "転記先Excelの物件一覧のうち、この明細が該当する行の番号。判断できなければ null",
                    },
                    "shisan": {"type": "string", "description": "土地 または 家屋"},
                    "chomei": {"type": "string", "description": "町名。印字されているとおりに。例: 桜川駅前四丁目"},
                    "banchi": {
                        "type": "string",
                        "description": (
                            "地番。印字は「−11−3−」のようにダッシュで区切られている。"
                            "**区切りを保ったままハイフン1文字に置き換える**（連結しない）。"
                            "例: 「−11−3−」→ 11-3（111-3 ではない）／「−180−1−」→ 180-1／「−178−」→ 178"
                        ),
                    },
                    "taisho_kaoku": {
                        "type": "integer",
                        "nullable": True,
                        "description": "『対象家屋(n)』の注記がある場合の n。無ければ null",
                    },
                    "kotei_hyojun": {"type": "integer", "nullable": True, "description": "固定資産税課税標準額"},
                    "kotei_sotogaku": {"type": "integer", "nullable": True, "description": "固定資産税相当額"},
                    "keigen": {"type": "integer", "nullable": True, "description": "軽減税額"},
                    "hyoka": {"type": "integer", "nullable": True, "description": "評価額（価格）"},
                    "zennendo_tokei": {"type": "integer", "nullable": True, "description": "前年度都市計画税課税標準額"},
                    "tokei_hyojun": {"type": "integer", "nullable": True, "description": "都市計画税課税標準額"},
                    "tokei_sotogaku": {"type": "integer", "nullable": True, "description": "都市計画税相当額"},
                    "goukei_sotogaku": {"type": "integer", "nullable": True, "description": "合計税相当額"},
                    "confidence": {"type": "number", "description": "0.0-1.0"},
                },
                "required": ["shisan", "chomei", "banchi", "confidence"],
            },
            "description": "page_kind が meisai のときだけ埋める。1件も漏らさないこと",
        },
    },
    "required": ["page_kind"],
}

EXTRACT_PROMPT = """\
この画像は日本の市町村が発行した納税通知書の1ページをスキャンしたものです。

まず page_kind を判定してください。
  meisai   … 「固定資産税・都市計画税 課税明細書」。物件ごとの明細が縦に並ぶ表
  cover    … 「固定資産税・都市計画税 納税通知書」。課税標準額の土地/家屋/合計(ア)が表になっている
  shokyaku … 表題に「償却資産」とある頁
  other    … 宛名・案内など上記以外

page_kind が cover のときは cover の各項目を読み取ってください。
page_kind が meisai のときは、記載されている物件を**1件残らず**、印字されているとおりに読み取ってください。
それ以外の page_kind のときは rows と cover を空のままにしてください。

■ 課税明細書のレイアウト（物件1件 = 3行1組）

■ レイアウト（物件1件 = 3行1組）
  1行目: 資産区分(土地/家屋)  町名  −街区・号−本番−枝番−小番−支号−   [評価地目/種類]
  2行目: 地積・床面積(㎡) | 前年度固定資産税課税標準額 | 固定資産税課税標準額 | 固定資産税相当額 | 軽減税額
  3行目: 評価額(円)       | 前年度都市計画税課税標準額 | 都市計画税課税標準額 | 都市計画税相当額 | 合計税相当額

■ 読み取りの注意
  - 数字は等幅で印字され、桁区切りのカンマは無い。空白は桁区切りではなく空欄を意味する
  - 列は右揃え。どの列の値かは、見出しの位置と揃えて判断すること
  - 空欄の項目は null にする。**推測で埋めない**
  - 「− 対象家屋(1)」のような注記がある行は taisho_kaoku にその番号を入れる
    （同一地番が複数明細に分かれているため、この番号が行の識別に必要）
  - banchi は先頭・末尾のダッシュを除き、区切りをハイフン1文字に統一する
    例: 「−180−1−」→ 180-1 ／ 「−42− −2−9」→ 42-2-9 ／ 「−178−」→ 178
  - 判読できない文字が1つでもある項目は、その項目を null にし confidence を下げる

読み取った内容だけを返してください。計算・補完・要約はしないでください。
"""

# 物件の同定を「自由記述」から「台帳からの選択」に変える。
#
# 町名と地番を自由に書かせると誤読が出る（実測）:
#   ・「桜川駅前四丁目」と「桜川駅南四丁目」の取り違え（1回の実行で8〜9件）
#   ・地番の桁誤り「松尾二丁目22-5」→「松尾二丁目222-5」
# 台帳に載っている物件は有限（博多区なら56件）なので、その一覧を渡して
# 「どれに当たるか」を選ばせる。あわせて印字されている文字もそのまま書かせ、
# 選択と読み取りが食い違ったら報告できるようにする。
BUKKEN_HINT_TEMPLATE = """

■ 転記先Excelの物件一覧（重要）
転記先のExcelに登録されている物件は次の {n} 件だけです。行番号つきで示します。

{bukken_list}

各明細について、**この一覧のどれに当たるかを excel_row で答えてください。**
判断は「資産区分（土地/家屋）＋ 町名 ＋ 地番」の3つが揃って一致することを条件とします。

注意点:
  ・「桜川駅前四丁目」と「桜川駅南四丁目」は一文字しか違いません。画像を拡大し、
    1文字ずつ照合してください。近くの行につられて同じ町名にしないでください
  ・地番は桁を間違えやすいので、ハイフンの位置と桁数を必ず数えてください
    （例: 「11-3」と「111-3」は別物です）
  ・一覧のどれにも当てはまらないと判断したら excel_row は null にしてください。
    無理に当てはめないでください
  ・excel_row を答えた場合も、chomei と banchi には**画像に印字されているとおり**を
    書いてください（一覧の表記に合わせて書き換えないでください）
"""


# ---------------------------------------------------------------------------
# 3. 正規化
# ---------------------------------------------------------------------------

_KANJI_DIGITS = {"〇": "0", "一": "1", "二": "2", "三": "3", "四": "4", "五": "5", "六": "6", "七": "7", "八": "8", "九": "9"}
_DASHES = "－−‐‑‒–—―ー─-"


def norm_text(s: str) -> str:
    """全角/半角・空白・ダッシュの揺れを吸収する。"""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    s = re.sub(f"[{re.escape(_DASHES)}]", "-", s)
    s = re.sub(r"\s+", "", s)
    return s


def norm_chomei(s: str) -> str:
    """町名を正規化する。『四丁目』『4丁目』の揺れを漢数字側に寄せる。"""
    s = norm_text(s)
    # 「4丁目」→「四丁目」
    rev = {v: k for k, v in _KANJI_DIGITS.items()}

    def _to_kanji(m: re.Match) -> str:
        return rev.get(m.group(1), m.group(1)) + "丁目"

    s = re.sub(r"(\d)丁目", _to_kanji, s)
    return s


def norm_banchi(s: str) -> str:
    """地番の数値部分を `-` 区切りに正規化する。

    「−180−1−」「180-1」「180‐1」→ いずれも "180-1"
    """
    s = norm_text(s)
    nums = re.findall(r"\d+", s)
    return "-".join(nums)


def split_shozai(s: str) -> tuple[str, str]:
    """Excel D列の「桜川駅前四丁目305-2」を (町名, 地番) に割る。"""
    s = norm_text(s)
    m = re.search(r"[\d\-]+$", s)
    if not m:
        return norm_chomei(s), ""
    return norm_chomei(s[: m.start()]), norm_banchi(m.group(0))


def norm_shisan(s: str) -> str:
    """PDFの「土地/家屋」とExcelの「土地/建物」を揃える。"""
    s = norm_text(s)
    if "家屋" in s or "建物" in s:
        return "家屋"
    if "土地" in s:
        return "土地"
    return s


# ---------------------------------------------------------------------------
# 4. 行単位の検算
# ---------------------------------------------------------------------------


@dataclass
class Issue:
    """検証で見つかった問題。1件でもあれば書き込みは行わない。"""

    kind: str
    where: str
    detail: str


def verify_row(rec: dict, rate_kotei: float, rate_tokei: float) -> list[Issue]:
    """明細1行の内部整合性を検算する。

    課税明細書には冗長性があり、読み取り誤りを行単位で検出できる。
        固定資産税課税標準額 × 税率(固定) を切り捨て = 固定資産税相当額
        都市計画税課税標準額 × 税率(都計) を切り捨て = 都市計画税相当額
        固定相当額 + 都計相当額 − 軽減税額            = 合計税相当額
    """
    where = f"{rec.get('shisan')} {rec.get('chomei')}{rec.get('banchi')}"
    if rec.get("taisho_kaoku"):
        where += f" 対象家屋({rec['taisho_kaoku']})"
    issues: list[Issue] = []

    conf = rec.get("confidence")
    if conf is not None and conf < 0.95:
        issues.append(Issue("low_confidence", where, f"読み取り信頼度 {conf}"))

    kh, ks = rec.get("kotei_hyojun"), rec.get("kotei_sotogaku")
    if kh is not None and ks is not None:
        expect = math.floor(kh * rate_kotei)
        if expect != ks:
            issues.append(
                Issue("kotei_mismatch", where, f"固定: {kh:,} × {rate_kotei} = {expect:,} だが印字は {ks:,}")
            )

    th, ts = rec.get("tokei_hyojun"), rec.get("tokei_sotogaku")
    if th is not None and ts is not None:
        expect = math.floor(th * rate_tokei)
        if expect != ts:
            issues.append(
                Issue("tokei_mismatch", where, f"都計: {th:,} × {rate_tokei} = {expect:,} だが印字は {ts:,}")
            )

    if ks is not None and ts is not None and rec.get("goukei_sotogaku") is not None:
        keigen = rec.get("keigen") or 0
        expect = ks + ts - keigen
        if expect != rec["goukei_sotogaku"]:
            issues.append(
                Issue(
                    "goukei_mismatch",
                    where,
                    f"合計: {ks:,} + {ts:,} - {keigen:,} = {expect:,} だが印字は {rec['goukei_sotogaku']:,}",
                )
            )

    # 評価額（H列）は税率計算にも表紙の合計にも出てこないため、他の2列と違って
    # 独立した検算がかからない。誤読しても素通りしてしまうので、構造から縛る。
    #   家屋 … 評価替えの特例が無いので 評価額 == 課税標準額（実測 22/22 で成立）
    #   土地 … 負担調整・住宅用地特例により 課税標準額 <= 評価額
    #          （実測の比は 0.009〜0.577。評価額を超えるものは無い）
    hyoka, shisan = rec.get("hyoka"), norm_shisan(rec.get("shisan", ""))
    if hyoka is not None and kh is not None:
        if shisan == "家屋" and hyoka != kh:
            issues.append(
                Issue("hyoka_mismatch", where, f"家屋は評価額＝課税標準額のはずですが {hyoka:,} ≠ {kh:,}")
            )
        elif shisan == "土地" and kh > hyoka:
            issues.append(
                Issue("hyoka_mismatch", where, f"課税標準額 {kh:,} が評価額 {hyoka:,} を超えています")
            )
    if hyoka is not None and th is not None and shisan == "土地" and th > hyoka:
        issues.append(
            Issue("hyoka_mismatch", where, f"都市計画税課税標準額 {th:,} が評価額 {hyoka:,} を超えています")
        )

    for key, label in (("hyoka", "評価額"), ("kotei_hyojun", "固定資産税課税標準額"), ("tokei_hyojun", "都市計画税課税標準額")):
        v = rec.get(key)
        if v is None:
            issues.append(Issue("missing_value", where, f"{label} が読み取れていない"))
        elif v == 0:
            # 読み取れなかった項目を null ではなく 0 で埋めてくる場合がある。
            # 評価額・課税標準額が 0 の物件は実在しないので、無条件で疑う。
            issues.append(Issue("zero_value", where, f"{label} が 0。読み取り失敗の可能性が高い"))

    return issues


# ---------------------------------------------------------------------------
# 5. Excel 側のインデックス
# ---------------------------------------------------------------------------


@dataclass
class SheetRow:
    row: int
    block: str  # "納付書①" / "納付書②"
    bukken: str  # B列 物件名
    shisan: str  # 土地 / 家屋
    chomei: str
    banchi: str
    k_is_formula: bool
    k_existing: int | None  # K列にハードコードされた数値（前年度値）。数式なら None


@dataclass
class SheetIndex:
    sheet: str
    rows: list[SheetRow]
    rate_kotei: float
    rate_tokei: float
    col_kakaku: str = COL_KAKAKU  # 価格（評価額）の列
    col_kotei: str = COL_KOTEI  # 固定資産税課税標準額の列
    col_tokei: str = COL_TOKEI  # 都市計画税課税標準額の列
    header_row: int = 3
    subtotal_rows: dict[str, int] = field(default_factory=dict)  # ブロック名 → 小計行


def _norm_header(v: object) -> str:
    """ヘッダ文字列を照合用に潰す（全角空白・改行・記号ゆれを吸収）。"""
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(v or "")))


def detect_layout(ws) -> dict:
    """シートのヘッダ行から、転記先の列と税率セルを見つける。

    列位置はシートごとに違う（実測、①福岡県ファイルの25シート）:
        H/I/K … 博多区・中央区・南区・城南区・大牟田市・飯塚市・嘉麻市・岡垣町・宗像市
        I/J/L … 東区・早良区・小倉北区・小倉南区・八幡西区・門司区ほか
        J/K/M … 西区
        O/P/R … 久留米市
    税率も自治体ごとに違う（0.014/0.003、0.016/0.001、0.014/0 など）。
    どちらも決め打ちにせず、**ヘッダ行の文字列から引く**。

    規則（全25シートで成立を確認済み）:
      ・ヘッダ行 = 「所在地番」を含む行
      ・資産区分・所在地番・価格(評価額)・固定資産税課税標準額・都市計画税課税標準額 を
        その行の文字列で特定する
      ・税率 = 各課税標準額列の **1つ上のセル**
    """
    from openpyxl.utils import get_column_letter

    header_row = None
    for r in range(1, min(ws.max_row, 10) + 1):
        if any("所在地番" in _norm_header(ws.cell(r, c).value) for c in range(1, 30)):
            header_row = r
            break
    if header_row is None:
        raise ValueError(f"シート {ws.title!r} にヘッダ行（「所在地番」を含む行）が見つかりません")

    pos: dict[str, int] = {}
    for c in range(1, 30):
        v = _norm_header(ws.cell(header_row, c).value)
        if not v:
            continue
        if v.startswith("資産"):
            pos.setdefault("shisan", c)
        elif "所在地番" in v:
            pos.setdefault("banchi", c)
        elif "固定資産税課税標準額" in v:
            pos.setdefault("kotei", c)
        elif "都市計画税課税標準額" in v:
            pos.setdefault("tokei", c)
        elif "価格" in v or "評価額" in v:
            pos.setdefault("kakaku", c)

    missing = [k for k in ("shisan", "banchi", "kakaku", "kotei", "tokei") if k not in pos]
    if missing:
        raise ValueError(
            f"シート {ws.title!r} のヘッダ行({header_row}行目)に必要な列が見つかりません: {missing}"
        )

    def rate(col: int, fallback: float) -> float:
        v = ws.cell(header_row - 1, col).value
        return v if isinstance(v, (int, float)) else fallback

    return {
        "header_row": header_row,
        "col_shisan": pos["shisan"],
        "col_banchi": pos["banchi"],
        "col_kakaku": get_column_letter(pos["kakaku"]),
        "col_kotei": get_column_letter(pos["kotei"]),
        "col_tokei": get_column_letter(pos["tokei"]),
        "rate_kotei": rate(pos["kotei"], FALLBACK_RATE_KOTEI),
        "rate_tokei": rate(pos["tokei"], FALLBACK_RATE_TOKEI),
    }


def build_index(xlsx_path: str, sheet_name: str) -> SheetIndex:
    """自治体シートを読み、転記対象行のインデックスを作る。

    列位置と税率は ``detect_layout()`` がヘッダ行から引く（シートごとに違うため）。

    対象行の決め方:
      ・B列に「納付書①②③」のラベルがあるシートは、それでブロックを分け、③以降は見ない
      ・ラベルが無いシート（大牟田市など9枚）は、ヘッダ行より下で
        資産区分と所在地番の両方が埋まっている行を拾う
      いずれも「所在地番が空の行」は拾わないので、償却資産（地番なし）は自然に除外される。
    """
    import openpyxl

    wb = openpyxl.load_workbook(xlsx_path, data_only=False)
    if sheet_name not in wb.sheetnames:
        raise KeyError(f"シート {sheet_name!r} が見つかりません。候補: {wb.sheetnames}")
    ws = wb[sheet_name]
    lay = detect_layout(ws)
    ci_shisan, ci_banchi = lay["col_shisan"], lay["col_banchi"]
    ci_tokei = openpyxl.utils.column_index_from_string(lay["col_tokei"])

    rows: list[SheetRow] = []
    subtotal_rows: dict[str, int] = {}
    block: str | None = None
    has_labels = False
    # ① は NFKC 正規化で "1" に潰れるため、norm_text は使わず素の文字列で照合する
    block_labels = {*BLOCK_LABELS_TARGET, BLOCK_LABEL_EXCLUDED}
    for r in range(lay["header_row"] + 1, ws.max_row + 1):
        label = re.sub(r"\s+", "", str(ws.cell(r, 2).value or ""))
        if label in block_labels:
            has_labels = True
            block = label
            subtotal_rows[label] = r
            if label == BLOCK_LABEL_EXCLUDED:
                break  # ③以降は見ない
            continue
        if has_labels and block not in BLOCK_LABELS_TARGET:
            continue
        shozai = ws.cell(r, ci_banchi).value
        shisan = ws.cell(r, ci_shisan).value
        if not shozai or not shisan:
            continue
        if norm_shisan(str(shisan)) not in ("土地", "家屋"):
            continue  # 償却資産などは対象外
        chomei, banchi = split_shozai(str(shozai))
        kcell = ws.cell(r, ci_tokei).value
        k_is_formula = isinstance(kcell, str) and kcell.startswith("=")
        rows.append(
            SheetRow(
                row=r,
                block=block or "（ブロックなし）",
                bukken=str(ws.cell(r, 2).value or ""),
                shisan=norm_shisan(str(shisan)),
                chomei=chomei,
                banchi=banchi,
                k_is_formula=k_is_formula,
                k_existing=kcell if isinstance(kcell, (int, float)) else None,
            )
        )
    wb.close()
    LOG.info(
        "シート解析: %s / 対象 %d行 / 列 %s%s%s / 税率 固定=%s 都計=%s / 小計行 %s",
        sheet_name, len(rows), lay["col_kakaku"], lay["col_kotei"], lay["col_tokei"],
        lay["rate_kotei"], lay["rate_tokei"], subtotal_rows,
    )
    LOG.debug("  K列が数式の行: %s", [r.row for r in rows if r.k_is_formula])
    return SheetIndex(
        sheet=sheet_name, rows=rows,
        rate_kotei=lay["rate_kotei"], rate_tokei=lay["rate_tokei"],
        col_kakaku=lay["col_kakaku"], col_kotei=lay["col_kotei"], col_tokei=lay["col_tokei"],
        header_row=lay["header_row"], subtotal_rows=subtotal_rows,
    )


# ---------------------------------------------------------------------------
# 6. 突合
# ---------------------------------------------------------------------------


def pick_sheet(
    xlsx_path: str,
    records: list[dict],
    min_rate: float = 0.5,
    min_margin: float = 0.25,
) -> tuple[str | None, list[tuple[float, int, int, str]]]:
    """読み取った明細の地番から、転記先のシートを言い当てる。

    ファイル名やユーザーの指定に頼らず、**資料の中身だけ**で決める。
    自治体ごとに地番は重ならないので、実測では1位と2位の差が
    100% 対 18.8% と大きく開いた（福岡市博多区・明細56件）。
    地番を数件読み違えても1位は動かないため、読み取り時に物件一覧を
    渡せなくても判定できる。

    戻り値は (確定したシート名 or None, 全シートのランキング)。
    1位の一致率が ``min_rate`` に届かない、または2位との差が
    ``min_margin`` 未満のときは **決めずに None を返す**（呼び出し側で停止する）。
    """
    import openpyxl

    wb = openpyxl.load_workbook(xlsx_path, read_only=True)
    names = wb.sheetnames
    wb.close()

    got = {(norm_shisan(r.get("shisan", "")), norm_banchi(r.get("banchi", ""))) for r in records}
    got.discard(("", ""))
    ranking: list[tuple[float, int, int, str]] = []
    for name in names:
        try:
            index = build_index(xlsx_path, name)
        except Exception:  # noqa: BLE001 - 転記対象でないシート（マニュアル等）は素通り
            continue
        if not index.rows:
            continue
        want = {(sr.shisan, sr.banchi) for sr in index.rows}
        hit = len(want & got)
        ranking.append((hit / len(want), hit, len(want), name))
    ranking.sort(reverse=True)

    if not ranking:
        LOG.warning("シート判定: 転記対象になりうるシートが1枚もありません")
        return None, ranking
    top = ranking[0]
    second = ranking[1][0] if len(ranking) > 1 else 0.0
    LOG.info("シート判定: 1位 %s %.1f%% (%d/%d) / 2位 %.1f%%", top[3], top[0] * 100, top[1], top[2], second * 100)
    if top[0] < min_rate or (top[0] - second) < min_margin:
        LOG.warning("シート判定: 決め手に欠けるため確定しません（上位3件 %s）", [(f"{r:.1%}", n) for r, _, _, n in ranking[:3]])
        return None, ranking
    return top[3], ranking


def format_bukken_hint(index: SheetIndex) -> list[str]:
    """台帳の物件一覧を、読み取りプロンプトに渡す形に整える。

    番号は **Excel の行番号そのもの**を使う。通し番号を別に振ると対応表を持つ必要があり、
    ずれたときに気づけない。行番号ならログや報告にそのまま出せて、人が Excel を開いて
    確認するときもそのまま使える。
    """
    return [
        f"  [{sr.row}行] {sr.shisan}  {sr.chomei}{sr.banchi}"
        + (f"  （{sr.bukken.splitlines()[0]}）" if sr.bukken else "")
        for sr in index.rows
    ]


def _rescue_by_banchi(
    leftovers: list[dict], index: SheetIndex, assign: dict[int, dict]
) -> tuple[dict[int, dict], list[Issue]]:
    """町名が食い違った明細を、地番だけで救済する。

    町名は「桜川駅前」と「桜川駅南」のように一文字違いのものがあり、読み取りで
    取り違えが起きる（実測。数値は全件正しいのに町名だけ入れ替わった）。
    地番は台帳内で一意なので、地番＋資産区分で行が確定できるなら救済してよい。
    ただし **黙って通さず、必ず不一致として報告する**。
    """
    rescued: dict[int, dict] = {}
    notes: list[Issue] = []
    free = [sr for sr in index.rows if sr.row not in assign]
    by_banchi: dict[tuple[str, str], list[SheetRow]] = {}
    for sr in free:
        by_banchi.setdefault((sr.shisan, sr.banchi), []).append(sr)

    for rec in leftovers:
        k = (norm_shisan(rec.get("shisan", "")), norm_banchi(rec.get("banchi", "")))
        cands = [sr for sr in by_banchi.get(k, []) if sr.row not in rescued]
        if not cands:
            continue
        if len(cands) > 1:
            # 同一地番が複数行あるケース（住宅用地の按分で1筆が複数明細に分かれる）。
            # 台帳のK列に残っている前年度の都市計画税課税標準額で一意に決める。
            # 町名が誤読されたのがちょうどこのグループだと、この救済が無いと丸ごと落ちる（実測）。
            hit = [
                sr
                for sr in cands
                if sr.k_existing is not None
                and sr.k_existing in (rec.get("zennendo_tokei"), rec.get("tokei_hyojun"))
            ]
            if len(hit) != 1:
                continue
            cands = hit
        sr = cands[0]
        rescued[sr.row] = rec
        notes.append(
            Issue(
                "chomei_mismatch",
                f"行{sr.row} {sr.bukken}",
                f"町名が食い違っています。明細の読み取り「{rec.get('chomei')}{rec.get('banchi')}」"
                f" ↔ 台帳「{sr.chomei}{sr.banchi}」。"
                f"地番が台帳内で一意のため行は確定できましたが、読み取り誤りの可能性があります",
            )
        )
    return rescued, notes


def match(records: list[dict], index: SheetIndex) -> tuple[dict[int, dict], list[Issue]]:
    """明細行 ↔ シート行 を1対1で対応づける。

    突合キーは (資産区分, 町名, 地番)。ただし同一キーが複数行あるケースがある
    （住宅用地の特例を建物ごとに按分するため、1筆が複数明細に分かれる）。
    その場合はシートのK列に残っている前年度の都市計画税課税標準額と、
    明細の「前年度都市計画税課税標準額」を突き合わせて一意に決める。
    それでも決まらなければ **割り当てず、報告して止める**（順序で決め打ちしない）。
    """
    issues: list[Issue] = []
    all_records = list(records)  # ①で unresolved に差し替わるので元の一覧を控える
    by_key: dict[tuple[str, str, str], list[SheetRow]] = {}
    for sr in index.rows:
        by_key.setdefault((sr.shisan, sr.chomei, sr.banchi), []).append(sr)

    rec_by_key: dict[tuple[str, str, str], list[dict]] = {}
    for rec in records:
        key = (norm_shisan(rec.get("shisan", "")), norm_chomei(rec.get("chomei", "")), norm_banchi(rec.get("banchi", "")))
        rec_by_key.setdefault(key, []).append(rec)

    assign: dict[int, dict] = {}
    orphans: list[dict] = []  # 町名が合わず、地番での救済に回すもの

    # ── ① Excelの行番号（excel_row）で確定できるものを先に割り当てる ──────────
    # 町名・地番を自由記述させると誤読が出る（実測: 町名8件、地番の桁誤り1件）。
    # 台帳の一覧から選ばせた番号があれば、それを第一のキーにする。
    # ただし選択を鵜呑みにせず、読み取った地番と台帳の地番が食い違えば報告する。
    id_to_row = {sr.row: sr for sr in index.rows}
    unresolved: list[dict] = []
    for rec in records:
        did = rec.get("excel_row")
        sr = id_to_row.get(did) if isinstance(did, int) else None
        if sr is not None and sr.row in assign:
            # 同じ行を2件が指した。どちらが正しいか決められないので、両方を②③に回す
            LOG.warning(
                "行%d を複数の明細が指しています。読み取り「%s%s」は行番号での確定を見送ります",
                sr.row, rec.get("chomei"), rec.get("banchi"),
            )
        if sr is None or sr.row in assign:
            unresolved.append(rec)
            continue
        assign[sr.row] = rec
        got_b, got_c = norm_banchi(rec.get("banchi", "")), norm_chomei(rec.get("chomei", ""))
        if got_b != sr.banchi or got_c != sr.chomei:
            issues.append(
                Issue(
                    "chomei_mismatch",
                    f"行{sr.row} {sr.bukken}",
                    f"台帳から選ばれた物件と、印字の読み取りが食い違っています。"
                    f"読み取り「{rec.get('chomei')}{rec.get('banchi')}」 ↔ 台帳「{sr.chomei}{sr.banchi}」",
                )
            )
    if assign:
        LOG.info("Excelの行番号で確定: %d件 / 残り %d件", len(assign), len(unresolved))

    # シート選びを間違えたときは、行番号も地番もほとんど当たらない。
    # 56行ぶんの「対応する明細が無い」を並べても原因が分からないので、1件にまとめて報告する。
    sheet_banchi = {(sr.shisan, sr.banchi) for sr in index.rows}
    rec_banchi = {(norm_shisan(r.get("shisan", "")), norm_banchi(r.get("banchi", ""))) for r in records}
    overlap = len(sheet_banchi & rec_banchi)
    hit_rate = max(len(assign), overlap) / max(len(sheet_banchi), 1)
    if sheet_banchi and len(records) + len(assign) >= 5 and hit_rate < 0.2:
        return {}, [
            Issue(
                "wrong_sheet",
                index.sheet,
                f"シート「{index.sheet}」の物件と、読み取った明細がほとんど一致しません"
                f"（{len(sheet_banchi)}行中 行番号 {len(assign)}件 / 地番 {overlap}件）。"
                f"転記先のシートが違う可能性が高いです。"
                f"明細の町名: {sorted({norm_chomei(r.get('chomei', '')) for r in records})}",
            )
        ]
    records = unresolved  # 以降は従来どおり 町名＋地番 で突き合わせる
    rec_by_key = {}
    for rec in records:
        key = (norm_shisan(rec.get("shisan", "")), norm_chomei(rec.get("chomei", "")), norm_banchi(rec.get("banchi", "")))
        rec_by_key.setdefault(key, []).append(rec)

    for key, recs in rec_by_key.items():
        # ① で既に確定した行は候補から外す。外さないと上書きして、
        # 押し出された明細が報告されないまま消える（実測でこの事故が起きた）。
        cands = [sr for sr in by_key.get(key, []) if sr.row not in assign]
        where = f"{key[0]} {key[1]}{key[2]}"
        if not cands:
            if by_key.get(key):
                issues.append(
                    Issue(
                        "duplicate_match",
                        where,
                        f"この物件に対応するシートの行はすべて別の明細に割り当て済みです"
                        f"（該当行: {[sr.row for sr in by_key[key]]}）。"
                        f"同じ物件を2回読み取っている可能性があります",
                    )
                )
            else:
                orphans += recs
            continue
        if len(recs) == 1 and len(cands) == 1:
            assign[cands[0].row] = recs[0]
            continue
        if len(recs) != len(cands):
            issues.append(
                Issue("count_mismatch", where, f"明細 {len(recs)} 件に対しシートは {len(cands)} 行。件数が合わない")
            )
            continue
        # K列に残っている前年度の都市計画税課税標準額で一意に決める。
        # 明細の「前年度」で照合し、決まらなければ「当年度」で照合する
        # （評価替えの無い年は前年度＝当年度になるため）。
        resolved: dict[int, dict] = {}
        used: set[int] = set()
        for probe in ("zennendo_tokei", "tokei_hyojun"):
            for sr in cands:
                if sr.row in resolved or sr.k_existing is None:
                    continue
                hits = [i for i, rec in enumerate(recs) if i not in used and rec.get(probe) == sr.k_existing]
                if len(hits) == 1:
                    resolved[sr.row] = recs[hits[0]]
                    used.add(hits[0])
        if len(resolved) == len(cands):
            assign.update(resolved)
            LOG.info(
                "同一地番の割当: %s → %s",
                where, {row: rec.get("taisho_kaoku") for row, rec in sorted(resolved.items())},
            )
        else:
            issues.append(
                Issue(
                    "ambiguous",
                    where,
                    f"同一地番が {len(cands)} 行あり、前年度の都市計画税課税標準額でも一意に決められない"
                    f"（対象家屋番号: {[r.get('taisho_kaoku') for r in recs]}）",
                )
            )

    # 町名が食い違って行が見つからなかった明細を、地番だけで救済する
    if orphans:
        rescued, notes = _rescue_by_banchi(orphans, index, assign)
        assign.update(rescued)
        issues += notes
        for rec in orphans:
            if rec not in rescued.values():
                issues.append(
                    Issue(
                        "no_sheet_row",
                        f"{norm_shisan(rec.get('shisan', ''))} {rec.get('chomei')}{rec.get('banchi')}",
                        "明細に対応するシートの行が見つからない（地番でも一意に決められない）",
                    )
                )

    # 最後に、双方に1件ずつだけ残ったものを結びつける。
    # 地番を読み違えると（実測: 11-3 → 111-3）、地番が救済のキーそのものなので
    # ①②③のどれでも拾えない。ただし「残った明細1件」と「残った台帳行1件」の
    # 資産区分が一致するなら、組み合わせは1通りしかない。
    # 誤っていれば直後の合計検算が必ず落とすので、報告付きで結びつけてよい。
    left_rows = [sr for sr in index.rows if sr.row not in assign]
    used_recs = set(map(id, assign.values()))
    left_recs = [r for r in all_records if id(r) not in used_recs]
    if len(left_rows) == 1 and len(left_recs) == 1:
        sr, rec = left_rows[0], left_recs[0]
        if sr.shisan == norm_shisan(rec.get("shisan", "")):
            assign[sr.row] = rec
            issues.append(
                Issue(
                    "chomei_mismatch",
                    f"行{sr.row} {sr.bukken}",
                    f"所在地が一致しませんが、双方に1件ずつしか残っていないため対応づけました。"
                    f"読み取り「{rec.get('chomei')}{rec.get('banchi')}」 ↔ 台帳「{sr.chomei}{sr.banchi}」",
                )
            )
            issues = [i for i in issues if not (i.kind == "no_sheet_row" and rec.get("banchi", "") in i.where)]
            left_rows = []

    for sr in left_rows:
        issues.append(
            Issue("no_meisai", f"行{sr.row} {sr.bukken}", f"{sr.shisan} {sr.chomei}{sr.banchi} に対応する明細が無い")
        )

    return assign, issues


# ---------------------------------------------------------------------------
# 7. 合計検算
# ---------------------------------------------------------------------------


def verify_totals(assign: dict[int, dict], index: SheetIndex, covers: list[dict]) -> list[Issue]:
    """転記後の小計が、納税通知書（表紙）の課税標準額の合計と一致するか検算する。

    表紙の合計(ア)は1000円未満切り捨てで表示されているため、
    明細の積み上げ額も同じく1000円未満を切り捨てて比較する。
    """
    issues: list[Issue] = []
    if not covers:
        return issues

    # 納付書ブロックごとに積み上げる。表紙は1通＝1ブロックに対応する
    # （博多区なら 納付書①＝単独名義、納付書②＝共有名義）。
    by_row = {sr.row: sr for sr in index.rows}
    blocks: dict[str, list[int]] = {}
    for row, rec in assign.items():
        b = by_row[row].block
        acc = blocks.setdefault(b, [0, 0])
        acc[0] += rec.get("kotei_hyojun") or 0
        acc[1] += rec.get("tokei_hyojun") or 0

    trunc = lambda v: (v // 1000) * 1000  # 表紙は1000円未満切り捨て表示
    remaining = dict(blocks)
    unmatched: list[dict] = []
    for cov in covers:
        ck, ct = cov.get("kotei_goukei"), cov.get("tokei_goukei")
        hit = None
        for b, (sk, st) in remaining.items():
            if (ck is None or trunc(sk) == ck) and (ct is None or trunc(st) == ct):
                hit = b
                break
        if hit is None:
            unmatched.append(cov)
        else:
            LOG.info(
                "合計検算 %s: 固定 %s / 都計 %s → 表紙と一致",
                hit, f"{trunc(remaining[hit][0]):,}", f"{trunc(remaining[hit][1]):,}",
            )
            del remaining[hit]

    for cov in unmatched:
        ck, ct = cov.get("kotei_goukei") or 0, cov.get("tokei_goukei") or 0
        cand = {b: (trunc(v[0]), trunc(v[1])) for b, v in remaining.items()}
        issues.append(
            Issue(
                "total_mismatch",
                "全体",
                f"表紙の合計（固定 {ck:,} / 都計 {ct:,}）に一致する納付書ブロックがありません。"
                f"明細の積み上げ: {cand}",
            )
        )
    for b, (sk, st) in remaining.items():
        issues.append(
            Issue(
                "total_no_cover",
                b,
                f"積み上げ（固定 {trunc(sk):,} / 都計 {trunc(st):,}）に対応する納税通知書（表紙）が見つかりません",
            )
        )
    return issues


# ---------------------------------------------------------------------------
# 8. 書き込み計画とコミット
# ---------------------------------------------------------------------------


# 書き込みを止めない指摘。行は確定していて数値も検算を通っているが、人が見るべきもの。
WARNING_KINDS = frozenset({"chomei_mismatch"})


@dataclass
class Plan:
    sheet: str
    writes: list[tuple[str, int, int]]  # (列, 行, 値)
    skipped: list[tuple[str, int, str]]  # (列, 行, 理由)
    issues: list[Issue]
    warnings: list[Issue] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.issues


def dry_run(records: list[dict], covers: list[dict], xlsx_path: str, sheet_name: str) -> Plan:
    """検証だけ行い、書き込み計画を返す。ファイルには一切触れない。"""
    index = build_index(xlsx_path, sheet_name)

    issues: list[Issue] = []
    row_ng = 0
    for rec in records:
        found = verify_row(rec, index.rate_kotei, index.rate_tokei)
        if found:
            row_ng += 1
            for x in found:
                LOG.warning("行検算NG [%s] %s: %s", x.kind, x.where, x.detail)
        issues += found
    LOG.info("行単位の検算: OK %d件 / NG %d件", len(records) - row_ng, row_ng)

    assign, match_issues = match(records, index)
    LOG.info("突合: %d行に割当 / 未解決 %d件", len(assign), len(match_issues))
    for x in match_issues:
        LOG.warning("突合NG [%s] %s: %s", x.kind, x.where, x.detail)
    issues += match_issues

    total_issues = verify_totals(assign, index, covers)
    for x in total_issues:
        LOG.warning("合計検算NG [%s] %s", x.kind, x.detail)
    if not total_issues and covers:
        LOG.info("合計の検算: 一致")
    issues += total_issues

    by_row = {sr.row: sr for sr in index.rows}
    writes: list[tuple[str, int, int]] = []
    skipped: list[tuple[str, int, str]] = []
    for row, rec in sorted(assign.items()):
        sr = by_row[row]
        if rec.get("hyoka") is not None:
            writes.append((index.col_kakaku, row, int(rec["hyoka"])))
        if rec.get("kotei_hyojun") is not None:
            writes.append((index.col_kotei, row, int(rec["kotei_hyojun"])))
        # K列: 数式が入っている行は上書きしない（数式は変更禁止）
        tokei = rec.get("tokei_hyojun")
        if sr.k_is_formula:
            if tokei is not None and rec.get("kotei_hyojun") is not None and tokei != rec["kotei_hyojun"]:
                issues.append(
                    Issue(
                        "formula_conflict",
                        f"{sheet_name}!{index.col_tokei}{row}",
                        f"{index.col_tokei}列は数式（固定課税標準を参照）だが、都計課税標準 {tokei:,} は "
                        f"固定課税標準 {rec['kotei_hyojun']:,} と異なる。数式のままでは誤りになる",
                    )
                )
            else:
                skipped.append((index.col_tokei, row, f"{index.col_tokei}列は数式。値は固定資産税課税標準額と一致するため変更不要"))
        elif tokei is not None:
            writes.append((index.col_tokei, row, int(tokei)))

    warnings = [i for i in issues if i.kind in WARNING_KINDS]
    blocking = [i for i in issues if i.kind not in WARNING_KINDS]
    for w in warnings:
        LOG.warning("要確認 [%s] %s: %s", w.kind, w.where, w.detail)
    LOG.info(
        "書き込み計画: %d セル / 意図的にスキップ %d セル / 警告 %d件 / 問題 %d件 → %s",
        len(writes), len(skipped), len(warnings), len(blocking), "書き込む" if not blocking else "中止",
    )
    return Plan(sheet=sheet_name, writes=writes, skipped=skipped, issues=blocking, warnings=warnings)


_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_PKGREL = "http://schemas.openxmlformats.org/package/2006/relationships"


def _col_index(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def _sheet_part(zf, sheet_name: str) -> str:
    """workbook.xml と rels から、シート名に対応する xl/worksheets/sheetN.xml のパスを引く。"""
    import xml.etree.ElementTree as ET

    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rid_to_target = {r.get("Id"): r.get("Target") for r in rels.findall(f"{{{_NS_PKGREL}}}Relationship")}
    for sh in wb.find(f"{{{_NS_MAIN}}}sheets"):
        if sh.get("name") == sheet_name:
            target = rid_to_target[sh.get(f"{{{_NS_REL}}}id")]
            return "xl/" + target.lstrip("/").removeprefix("xl/")
    raise KeyError(f"シート {sheet_name!r} が workbook.xml に見つかりません")


def write_cells_preserving(xlsx_path: str, out_path: str, sheet_name: str, cells: dict[str, int]) -> str:
    """指定セルだけを書き換え、それ以外のファイル内部を一切変えずに保存する。

    openpyxl で開いて保存すると、印刷設定 (printerSettings)・一部の図形・
    calcChain などが失われる（実測: 元110パーツ → 57パーツ）。
    「見出し・数式・書式・他シートを変更しない」という要件を厳密に満たすため、
    xlsx（=ZIP）の対象シートのXMLだけを書き換えて他のパーツはそのままコピーする。

    **XMLは木に読み直さず、対象セルの文字列だけを差し替える。**
    ElementTree で再構築すると名前空間の接頭辞が付け替わり
    （``mc:`` → ``ns1:`` など）、``mc:Ignorable="x14ac"`` のように
    接頭辞を文字列で参照する属性が壊れて、**そのシートだけ Excel で開けなくなる**（実測）。
    元のバイト列を保ったまま最小限だけ触るのが唯一安全な方法。
    """
    import zipfile

    with zipfile.ZipFile(xlsx_path) as zf:
        part = _sheet_part(zf, sheet_name)
        names = zf.namelist()
        payload = {n: zf.read(n) for n in names}

    xml = payload[part].decode("utf-8")

    for ref, value in sorted(cells.items(), key=lambda kv: (int(kv[0][1:]), kv[0][0])):
        m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
        if not m:
            raise ValueError(f"セル参照が不正です: {ref}")
        col, rownum = m.group(1), int(m.group(2))
        newv = f"<v>{int(value)}</v>"

        cell = re.search(rf'<c r="{ref}"(?P<attrs>[^>]*?)(?P<end>/>|>(?P<body>.*?)</c>)', xml, re.S)
        if cell:
            if cell.group("body") and "<f" in cell.group("body"):
                raise RuntimeError(f"{sheet_name}!{ref} は数式です。書き込みを中止しました")
            # t属性（文字列型など）を外し、中身を <v> だけにする
            attrs = re.sub(r'\s+t="[^"]*"', "", cell.group("attrs"))
            xml = xml[: cell.start()] + f'<c r="{ref}"{attrs}>{newv}</c>' + xml[cell.end() :]
            continue

        # 空セルはXML上に存在しないことがある。同じ行の中に列順を保って挿入する
        row = re.search(rf'<row r="{rownum}"[^>]*>.*?</row>', xml, re.S)
        if row is None:
            raise RuntimeError(f"{sheet_name}!{ref}: {rownum}行目が存在しません")
        block = row.group(0)
        pos = len(block) - len("</row>")
        for c in re.finditer(r'<c r="([A-Z]+)\d+"', block):
            if _col_index(c.group(1)) > _col_index(col):
                pos = c.start()
                break
        xml = xml[: row.start()] + block[:pos] + f'<c r="{ref}">{newv}</c>' + block[pos:] + xml[row.end() :]

    payload[part] = xml.encode("utf-8")

    # calcChain は数式の再計算順序のキャッシュ。セルを増やしたので落として Excel に作り直させる。
    # ただし **落とすだけでは壊れる**: [Content_Types].xml の宣言と workbook.xml.rels の参照が
    # 残ったままだと、Excel が「参照先が無い」と判断して修復ダイアログを出す（実測）。3つセットで外す。
    CALC = "xl/calcChain.xml"
    if CALC in payload:
        payload.pop(CALC)
        ct = payload.get("[Content_Types].xml")
        if ct is not None:
            payload["[Content_Types].xml"] = re.sub(
                rb'<Override[^>]*PartName="/xl/calcChain\.xml"[^>]*/>', b"", ct
            )
        rels_name = "xl/_rels/workbook.xml.rels"
        rels = payload.get(rels_name)
        if rels is not None:
            payload[rels_name] = re.sub(
                rb'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', b"", rels
            )

    # 開いた瞬間に再計算させる。calcChain を落としただけだと、数式セルに残った
    # 古いキャッシュ値（未入力時の 0）がそのまま表示されることがある。
    wbx = payload.get("xl/workbook.xml")
    if wbx is not None:
        if b"<calcPr" in wbx:
            wbx = re.sub(rb"<calcPr[^>]*/>", b'<calcPr calcId="0" fullCalcOnLoad="1"/>', wbx)
        else:
            wbx = wbx.replace(b"</workbook>", b'<calcPr calcId="0" fullCalcOnLoad="1"/></workbook>')
        payload["xl/workbook.xml"] = wbx

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zo:
        for name in names:
            if name in payload:
                zo.writestr(name, payload[name])
    return out_path


def commit(plan: Plan, xlsx_path: str, out_path: str) -> str:
    """検証を全部通っているときだけ書き込む。1件でも問題があれば例外で止める。"""
    if not plan.ok:
        raise RuntimeError(
            "検証を通らなかったため書き込みを行いません:\n"
            + "\n".join(f"  [{i.kind}] {i.where}: {i.detail}" for i in plan.issues)
        )
    cells = {f"{col}{row}": int(value) for col, row, value in plan.writes}
    LOG.info("書き込み実行: %s → %s (%d セル)", xlsx_path, out_path, len(cells))
    for ref, v in sorted(cells.items(), key=lambda kv: (int(kv[0][1:]), kv[0][0])):
        LOG.debug("  %s = %s", ref, f"{v:,}")
    path = write_cells_preserving(xlsx_path, out_path, plan.sheet, cells)
    LOG.info("書き込み完了: %s", path)
    return path


async def extract_pdf(
    pdf_path: str,
    workdir: str = "tmp/kotozei",
    bukken_hint: list[str] | None = None,
    model: str = DEFAULT_READ_MODEL,
) -> tuple[list[dict], list[dict], list[tuple]]:
    """PDFを描画して読み取るところまでを行う。Excelには一切触れない。

    戻り値は (明細行のリスト, 表紙のリスト, [(頁番号, 判定), ...])。
    読み取り結果は ``{workdir}/extracted.json`` にも保存する。

    頁の仕分けは読み取りと同じ1回の呼び出しで行う。
    テキスト層はOCR由来で文字化けが激しく、キーワードでの仕分けは取りこぼす（実測）。
    """
    import asyncio
    import json

    from agent_sdk import llm_call  # 実行環境 サンドボックスが注入するツール

    setup_logging(workdir)
    Path(workdir).mkdir(parents=True, exist_ok=True)
    pages = render_pdf(pdf_path, workdir)
    sem = asyncio.Semaphore(6)  # 429 回避
    usage_total = [0, 0]  # [入力トークン, 出力トークン]

    prompt = EXTRACT_PROMPT
    if bukken_hint:
        prompt += BUKKEN_HINT_TEMPLATE.format(n=len(bukken_hint), bukken_list="\n".join(bukken_hint))
        LOG.info("台帳の物件一覧 %d件 をプロンプトに渡します", len(bukken_hint))

    async def _one(page_no: int) -> dict:
        async with sem:
            LOG.info("p%02d 読み取り 開始 (model=%s)", page_no, model)
            t0 = time.monotonic()
            # SDK の llm_call は prompt 以外がキーワード専用。schema は dict、戻り値も dict
            out = await llm_call(
                prompt,
                model=model,
                file_paths=[pages[page_no - 1]],
                schema=PAGE_SCHEMA,
            )
            dt = time.monotonic() - t0
        if out.get("error"):
            LOG.error("p%02d 読み取り 失敗 (%.1f秒): %s", page_no, dt, out["error"])
            raise RuntimeError(f"{page_no}ページ目の読み取りに失敗: {out['error']}")
        data = out.get("data") or {}
        data["_page"] = page_no
        usage = out.get("usage") or {}
        usage_total[0] += usage.get("input_tokens") or 0
        usage_total[1] += usage.get("output_tokens") or 0
        LOG.info(
            "p%02d 読み取り 完了 (%.1f秒) kind=%s rows=%d in=%s out=%s",
            page_no, dt, data.get("page_kind"), len(data.get("rows") or []),
            usage.get("input_tokens", "?"), usage.get("output_tokens", "?"),
        )
        return data

    async def _one_split(page_no: int) -> dict:
        """1回で読めなかったページを、上下に重なりを持たせて2分割して読む。

        ``llm_call`` には120秒の上限がある。明細が23件並ぶページは Pro だと
        80〜115秒かかり、上限に届くことがある（実測でタイムアウト）。
        行数を半分にすれば所要も出力トークンもおおよそ半分になる。

        物件1件は3行1組なので、境目で切ると片方が欠ける。**15%重ねて**切り、
        重複は (資産, 町名, 地番, 対象家屋番号) で除く。
        """
        from PIL import Image

        src = Image.open(pages[page_no - 1])
        w, h = src.size
        parts: list[str] = []
        for i, (top, bottom) in enumerate(((0.0, 0.60), (0.45, 1.0)), start=1):
            path = str(Path(workdir) / f"page{page_no:02d}_part{i}.png")
            src.crop((0, int(h * top), w, int(h * bottom))).save(path)
            parts.append(path)

        async def _part(path: str, label: str) -> dict:
            async with sem:
                LOG.info("p%02d %s 読み取り 開始（分割）", page_no, label)
                t0 = time.monotonic()
                out = await llm_call(prompt, model=model, file_paths=[path], schema=PAGE_SCHEMA)
                dt = time.monotonic() - t0
            if out.get("error"):
                raise RuntimeError(f"{page_no}ページ目({label})の読み取りに失敗: {out['error']}")
            data = out.get("data") or {}
            usage = out.get("usage") or {}
            usage_total[0] += usage.get("input_tokens") or 0
            usage_total[1] += usage.get("output_tokens") or 0
            LOG.info("p%02d %s 読み取り 完了 (%.1f秒) rows=%d", page_no, label, dt, len(data.get("rows") or []))
            return data

        halves = await asyncio.gather(_part(parts[0], "上半分"), _part(parts[1], "下半分"))

        merged: list[dict] = []
        seen: set[tuple] = set()
        for hdata in halves:
            for row in hdata.get("rows") or []:
                k = (
                    norm_shisan(row.get("shisan", "")),
                    norm_chomei(row.get("chomei", "")),
                    norm_banchi(row.get("banchi", "")),
                    row.get("taisho_kaoku"),
                )
                if k in seen:
                    continue
                seen.add(k)
                merged.append(row)
        kind = next((h.get("page_kind") for h in halves if h.get("page_kind") == "meisai"), halves[0].get("page_kind"))
        LOG.info("p%02d 分割して読み直し 完了: %d件（重複除去後）", page_no, len(merged))
        return {"page_kind": kind, "rows": merged, "cover": halves[0].get("cover") or {}, "_page": page_no}

    async def _one_or_split(page_no: int) -> dict:
        try:
            return await _one(page_no)
        except RuntimeError as e:
            if "タイムアウト" not in str(e) and "timeout" not in str(e).lower():
                raise
            LOG.warning("p%02d が時間内に読めませんでした。上下に分割して読み直します: %s", page_no, e)
            return await _one_split(page_no)

    t_all = time.monotonic()
    results = await asyncio.gather(*(_one_or_split(p) for p in range(1, len(pages) + 1)))
    rate_in, rate_out = MODEL_RATES.get(model, (0.0, 0.0))
    cost = usage_total[0] / 1e6 * rate_in + usage_total[1] / 1e6 * rate_out
    LOG.info(
        "全頁の読み取り完了 (%.1f秒) model=%s in=%s out=%s 概算 $%.3f",
        time.monotonic() - t_all, model, f"{usage_total[0]:,}", f"{usage_total[1]:,}", cost,
    )
    if usage_total[1]:
        LOG.info("  ※ コストの大半は出力トークン。出力 %s tok × $%.2f/M", f"{usage_total[1]:,}", rate_out)

    records: list[dict] = []
    covers: list[dict] = []
    for r in results:
        kind = r.get("page_kind")
        if kind == "meisai":
            for row in r.get("rows") or []:
                row["_page"] = r.get("_page")
                records.append(row)
        elif kind == "cover":
            cov = r.get("cover") or {}
            # 償却資産の通知書は課税標準額の体系が違うため合計検算に混ぜない
            if cov.get("kotei_goukei") or cov.get("tokei_goukei"):
                covers.append(cov)

    page_kinds = [(r.get("_page"), r.get("page_kind")) for r in results]
    LOG.info("頁の判定: %s", page_kinds)
    LOG.info("抽出結果: 明細 %d件 / 表紙 %d件", len(records), len(covers))
    for c in covers:
        LOG.info("  表紙: 固定合計=%s 都計合計=%s", c.get("kotei_goukei"), c.get("tokei_goukei"))
    out_json = str(Path(workdir) / "extracted.json")
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump({"pages": page_kinds, "records": records, "covers": covers}, f, ensure_ascii=False, indent=2)
    LOG.info("読み取り結果を保存: %s", out_json)
    return records, covers, page_kinds


def format_extraction(records: list[dict], covers: list[dict], page_kinds: list[tuple],
                      rate_kotei: float = FALLBACK_RATE_KOTEI,
                      rate_tokei: float = FALLBACK_RATE_TOKEI) -> str:
    """抽出結果を人が読める形にまとめる（Excelには触れない、抽出だけの確認用）。"""
    lines = [f"■ 頁の判定: {page_kinds}", f"■ 明細 {len(records)}件 / 表紙 {len(covers)}件", ""]
    ng = 0
    for rec in records:
        issues = verify_row(rec, rate_kotei, rate_tokei)
        mark = "OK " if not issues else "NG "
        if issues:
            ng += 1
        h = rec.get("hyoka")
        i = rec.get("kotei_hyojun")
        k = rec.get("tokei_hyojun")
        f = lambda v: f"{v:,}" if isinstance(v, int) else "—"
        tk = f" 対象家屋({rec['taisho_kaoku']})" if rec.get("taisho_kaoku") else ""
        lines.append(f"{mark}p{rec.get('_page')} {rec.get('shisan')} {rec.get('chomei')}{rec.get('banchi')}{tk}")
        lines.append(f"      H評価額={f(h)}  I固定={f(i)}  K都計={f(k)}")
        for x in issues:
            lines.append(f"      ↑ [{x.kind}] {x.detail}")
    lines.append("")
    lines.append(f"■ 行単位の検算: OK {len(records) - ng}件 / NG {ng}件")
    if covers:
        # ここは Excel を見ないので納付書ブロックが分からず、全体でしか比べられない。
        # 表紙は通知書ごとに1000円未満を切り捨てているため、n通あると
        # 「切り捨て後の和」は「和を切り捨てたもの」より最大 (n-1)*1000 円小さくなる。
        # その範囲内なら読み取りは一致とみなす（厳密なブロック単位の検算は dry_run 側で行う）。
        n = len(covers)
        tol = (n - 1) * 1000
        sk = sum((c.get("kotei_goukei") or 0) for c in covers)
        st = sum((c.get("tokei_goukei") or 0) for c in covers)
        gk = (sum((r.get("kotei_hyojun") or 0) for r in records) // 1000) * 1000
        gt = (sum((r.get("tokei_hyojun") or 0) for r in records) // 1000) * 1000
        judge = lambda got, cov: (
            "一致" if got == cov
            else f"一致（切り捨ての累積誤差 {got - cov:,}円、許容 {tol:,}円以内）" if 0 <= got - cov <= tol
            else f"不一致 差 {got - cov:,}"
        )
        lines.append(f"■ 合計の検算（明細の積み上げ vs 表紙の合計(ア)、表紙 {n}通）")
        lines.append(f"   固定資産税 {gk:,} vs {sk:,}  → {judge(gk, sk)}")
        lines.append(f"   都市計画税 {gt:,} vs {st:,}  → {judge(gt, st)}")
    return "\n".join(lines)


def score_extraction(
    records: list[dict],
    covers: list[dict],
    xlsx_path: str,
    sheet_name: str,
    label: str = "",
) -> dict:
    """抽出結果を台帳と突き合わせて採点する。モデルを比べるときに使う。

    採点の軸は4つ。
      1. 件数        … 台帳の対象行数と一致するか
      2. 行単位の検算 … 課税標準×税率＝相当額 が全行成立するか
      3. 町名        … 地番で引いた台帳の町名と一致するか（数値には影響しないが読み取り品質の指標）
      4. **実データ照合** … 台帳のK列に人が入力済みの値と、抽出した都市計画税課税標準額を比較する。
         検算とは独立した外部の真値なので、これが最も強い精度指標になる。
    """
    index = build_index(xlsx_path, sheet_name)

    ng_calc = sum(1 for r in records if verify_row(r, index.rate_kotei, index.rate_tokei))

    # 地番＋資産区分で台帳を引く（台帳内で一意。同一地番グループだけ複数になる）
    by_banchi: dict[tuple[str, str], list[SheetRow]] = {}
    for sr in index.rows:
        by_banchi.setdefault((sr.shisan, sr.banchi), []).append(sr)

    chomei_ok = chomei_ng = 0
    chomei_bad: list[str] = []
    truth_ok = truth_ng = 0
    truth_bad: list[str] = []
    for rec in records:
        k = (norm_shisan(rec.get("shisan", "")), norm_banchi(rec.get("banchi", "")))
        cands = by_banchi.get(k)
        if not cands:
            continue
        if norm_chomei(rec.get("chomei", "")) == cands[0].chomei:
            chomei_ok += 1
        else:
            chomei_ng += 1
            chomei_bad.append(f"{rec.get('chomei')}{rec.get('banchi')}→{cands[0].chomei}{cands[0].banchi}")
        # 台帳に人が入力済みの値と照合できるのは、K列が数式でなく1行しか候補が無いときだけ
        if len(cands) == 1 and cands[0].k_existing is not None:
            got = rec.get("tokei_hyojun")
            if got == cands[0].k_existing:
                truth_ok += 1
            else:
                truth_ng += 1
                truth_bad.append(f"行{cands[0].row} 台帳{cands[0].k_existing:,} ≠ 抽出{got}")

    total_ok = True
    if covers:
        n = len(covers)
        sk = sum((c.get("kotei_goukei") or 0) for c in covers)
        st = sum((c.get("tokei_goukei") or 0) for c in covers)
        gk = (sum((r.get("kotei_hyojun") or 0) for r in records) // 1000) * 1000
        gt = (sum((r.get("tokei_hyojun") or 0) for r in records) // 1000) * 1000
        total_ok = 0 <= gk - sk <= (n - 1) * 1000 and 0 <= gt - st <= (n - 1) * 1000

    # 合否は **数値だけ** で決める。
    # 町名は excel_row の裏取り用に読ませているだけで、間違っても
    # 行番号と地番で救済できる（実害が無い）。参考値として出すに留める。
    passed = (
        len(records) == len(index.rows)
        and ng_calc == 0
        and truth_ng == 0
        and total_ok
    )
    s = {
        "label": label,
        "判定": "合格" if passed else "不合格",
        "件数": f"{len(records)}/{len(index.rows)}",
        "行検算NG": ng_calc,
        "実データ照合": f"{truth_ok}/{truth_ok + truth_ng}",
        "合計検算": "一致" if total_ok else "不一致",
        "町名(参考)": f"{chomei_ok}/{chomei_ok + chomei_ng}",
        "_町名不一致": chomei_bad,
        "_照合不一致": truth_bad,
    }
    LOG.info(
        "採点[%s] %s — 件数 %s / 行検算NG %d / 実データ照合 %s / 合計 %s（町名 %s は参考値）",
        label, s["判定"], s["件数"], ng_calc, s["実データ照合"], s["合計検算"], s["町名(参考)"],
    )
    return s


async def extract_only(
    pdf_path: str,
    xlsx_path: str,
    sheet_name: str,
    model: str = DEFAULT_READ_MODEL,
    workdir: str = "tmp/kotozei",
) -> str:
    """モデル比較用。**読み取るだけ**して、結果JSONのパスを返す。

    採点も判定もここでは行わない。判定は手元で ``score_extraction()`` に通して
    行う（実行のたびに判定基準がブレないように、採点は1箇所に寄せる）。
    Excelは物件一覧を作るために読むだけで、書き込みはしない。

    戻り値のパスをそのまま ``file_output()`` に渡せばよい。
    """
    setup_logging(workdir)
    index = build_index(xlsx_path, sheet_name)
    hint = format_bukken_hint(index)
    records, covers, kinds = await extract_pdf(pdf_path, workdir, bukken_hint=hint, model=model)
    LOG.info("読み取り完了: 明細 %d件 / 表紙 %d件 / 頁 %s", len(records), len(covers), kinds)
    LOG.info("このJSONを手元に持ち帰って採点してください: %s", Path(workdir) / "extracted.json")
    return str(Path(workdir) / "extracted.json")


async def extract_and_score(
    pdf_path: str,
    xlsx_path: str,
    sheet_name: str,
    model: str = DEFAULT_READ_MODEL,
    label: str = "",
    workdir: str = "tmp/kotozei",
) -> str:
    """モデル比較用。抽出して採点し、結果JSONのパスを返す。Excelには書き込まない。

    呼び出し側が JSON を書き直す余地を無くすためにここで完結させる
    （実測で、アシスタントが親切心から ``json.dump({"records": str(records)})`` と
    書き直してしまい、比較不能なファイルになった）。

    戻り値のパスをそのまま ``file_output()`` に渡せばよい。
    """
    setup_logging(workdir)
    index = build_index(xlsx_path, sheet_name)
    hint = format_bukken_hint(index)
    records, covers, _ = await extract_pdf(pdf_path, workdir, bukken_hint=hint, model=model)
    score = score_extraction(records, covers, xlsx_path, sheet_name, label=label or model)
    print(format_scores([score]))
    return str(Path(workdir) / "extracted.json")


COMPARE_FIELDS = (
    "excel_row", "chomei", "hyoka", "kotei_hyojun", "tokei_hyojun",
    "kotei_sotogaku", "tokei_sotogaku", "goukei_sotogaku", "keigen", "zennendo_tokei",
)


def compare_extractions(paths: list[str], labels: list[str] | None = None) -> str:
    """複数回の抽出結果を1件ずつ突き合わせる。

    採点（``score_extraction``）は「検算を通るか」しか見ないので、
    検算を通ったまま値がブレるケースを取りこぼす。
    実行ごとの ``extracted.json`` を物件単位で比較して、
    **1フィールドでも違えば列挙する**。

    モデルを替えた実行どうしを比べれば、クロスチェックにもなる
    （別モデルが同じ値を出したなら、その値はまず正しい）。
    """
    import json

    # ルームを分けて実行すると sandbox の tmp/ も分かれるため、
    # workdir でも extracted.json のパスでも受けられるようにする。
    labels = labels or [Path(p).stem for p in paths]
    runs = []
    for p in paths:
        f_path = Path(p) if Path(p).suffix == ".json" else Path(p) / "extracted.json"
        with open(str(f_path), encoding="utf-8") as f:
            runs.append(json.load(f))

    lines = []
    for lb, d in zip(labels, runs):
        lines.append(f"{lb}: 明細{len(d['records'])}件 / 表紙{len(d['covers'])}件 / 頁判定 {d['pages']}")

    def key(r: dict) -> tuple[str, str]:
        return (norm_shisan(r.get("shisan", "")), norm_banchi(r.get("banchi", "")))

    maps = []
    for d in runs:
        m: dict[tuple[str, str], list[dict]] = {}
        for r in d["records"]:
            m.setdefault(key(r), []).append(r)
        maps.append(m)

    allkeys = sorted(set().union(*[set(m) for m in maps]), key=str)
    lines.append(f"\n■ 物件キー（資産区分＋地番）: {len(allkeys)}種")

    missing = [k for k in allkeys if any(k not in m for m in maps)]
    for k in missing:
        lines.append(f"   × {k}  出現したのは {[labels[i] for i, m in enumerate(maps) if k in m]} のみ")

    diffs = 0
    for k in allkeys:
        if k in missing:
            continue
        counts = [len(m[k]) for m in maps]
        if len(set(counts)) > 1:
            diffs += 1
            lines.append(f"   件数ちがい {k}: {dict(zip(labels, counts))}")
            continue
        for i in range(counts[0]):  # 同一地番が複数あるケースは出現順に比べる
            for f in COMPARE_FIELDS:
                vals = [m[k][i].get(f) for m in maps]
                if len({str(v) for v in vals}) > 1:
                    diffs += 1
                    lines.append(f"   差分 {k}[{i}] {f}: {dict(zip(labels, vals))}")

    lines.append(f"\n■ 値の食い違い: {diffs}箇所")
    covers = [[(c.get("kotei_goukei"), c.get("tokei_goukei")) for c in d["covers"]] for d in runs]
    lines.append(f"■ 表紙: {dict(zip(labels, covers))}")
    lines.append("■ 判定: " + ("全実行で完全一致 ✓" if diffs == 0 and not missing else "食い違いあり ← 上を確認"))
    return "\n".join(lines)


def format_scores(scores: list[dict]) -> str:
    """複数モデルの採点を並べて表示する。"""
    cols = ["label", "判定", "件数", "行検算NG", "実データ照合", "合計検算", "町名(参考)"]
    w = {c: max(len(c), *(len(str(s.get(c, ""))) for s in scores)) for c in cols}
    lines = ["  ".join(c.ljust(w[c]) for c in cols), "  ".join("-" * w[c] for c in cols)]
    for s in scores:
        lines.append("  ".join(str(s.get(c, "")).ljust(w[c]) for c in cols))
    for s in scores:
        if s["_町名不一致"]:
            lines.append(f"\n[{s['label']}] 町名不一致 {len(s['_町名不一致'])}件: {s['_町名不一致']}")
        if s["_照合不一致"]:
            lines.append(f"[{s['label']}] 実データ照合の不一致: {s['_照合不一致']}")
    return "\n".join(lines)


async def run(
    pdf_path: str,
    xlsx_path: str,
    sheet_name: str | None = None,
    out_path: str | None = None,
    workdir: str = "tmp/kotozei",
    model: str = DEFAULT_READ_MODEL,
) -> str:
    """入口。PDFとExcelを受け取り、検証を全部通ったときだけ転記したExcelを書き出す。

    ``sheet_name`` を省略すると、**読み取った地番から転記先シートを自動で判定する**
    （ファイル名やユーザーの指定に依存しない）。
    指定された場合は、そのシートの物件一覧を読み取りプロンプトに渡せるので
    地番・町名の誤読が減る。どちらでも突合と検算は同じように行う。

    戻り値は人が読むレポート。書き込めなかった場合も例外にせずレポートを返し、
    呼び出し側（アシスタント）がそのままユーザーに提示できるようにする。
    """
    setup_logging(workdir)

    if not sheet_name:
        # 転記先シートはユーザーに指示してもらう前提。
        # 未指定だと台帳の物件一覧を読み取りプロンプトに渡せず、町名・地番の誤読が増える
        # （実測: 一覧ありは地番の誤読 0/505、一覧なしは2実行中1件）。
        # **PDFを読む前に**止めるので、ここで停止してもコストは発生しない。
        import openpyxl

        wb = openpyxl.load_workbook(xlsx_path, read_only=True)
        names = [n for n in wb.sheetnames]
        wb.close()
        LOG.warning("シート名が指定されていないため、読み取りを開始せずに停止します")
        return (
            "■ 結果: 中止 — 読み取りを開始していません（費用は発生していません）\n"
            "  転記先のシート名（自治体名）を指定してください。\n"
            f"  このExcelのシート: {names}\n"
            "  例:「福岡市博多区_2026.pdf を読込み、28期_固都税・償却資産税（①福岡県).xlsx の"
            "福岡市博多区シートを更新して下さい」"
        )

    if sheet_name:
        # 転記先が分かっているので、台帳の物件一覧を渡して読み取り精度を上げる
        index = build_index(xlsx_path, sheet_name)
        hint = format_bukken_hint(index)
        records, cover_data, page_kinds = await extract_pdf(
            pdf_path, workdir, bukken_hint=hint, model=model
        )
    else:
        # 資料の中身だけで転記先を決める。物件一覧は渡せないが、
        # シート判定に必要な精度は低い（1位と2位が大きく開く）ので成立する
        records, cover_data, page_kinds = await extract_pdf(pdf_path, workdir, model=model)
        if records:
            sheet_name, ranking = pick_sheet(xlsx_path, records)
            if sheet_name is None:
                top = [f"{r:.1%} {n}" for r, _, _, n in ranking[:5]]
                return (
                    "■ 結果: 中止 — 1セルも書き込んでいません\n"
                    "  読み取った地番から転記先シートを特定できませんでした。\n"
                    f"  一致率の上位: {top}\n"
                    "  シート名を指定して実行し直してください。"
                )

    if out_path is None:
        out_path = f"output/{Path(xlsx_path).name}"

    if not records:
        return (
            "■ 結果: 中止 — 1セルも書き込んでいません\n"
            "  PDFから課税明細書のページを読み取れませんでした。\n"
            f"  頁の判定結果: {page_kinds}"
        )

    plan = dry_run(records, cover_data, xlsx_path, sheet_name)
    if not plan.ok:
        return format_report(plan)
    commit(plan, xlsx_path, out_path)
    return format_report(plan) + f"\n■ 出力: {out_path}"


def format_report(plan: Plan, index: SheetIndex | None = None) -> str:
    """人が読む用のレポートを組み立てる。"""
    lines = [f"■ 対象シート: {plan.sheet}"]
    if plan.issues:
        lines.append(f"■ 結果: 中止（{len(plan.issues)}件の問題）— 1セルも書き込んでいません")
        lines.append("")
        for i in plan.issues:
            lines.append(f"  [{i.kind}] {i.where}")
            lines.append(f"      {i.detail}")
    else:
        cols: dict[str, int] = {}
        rows = set()
        for col, row, _ in plan.writes:
            cols[col] = cols.get(col, 0) + 1
            rows.add(row)
        lines.append(f"■ 結果: 検証OK — {len(rows)}行 / {len(plan.writes)}セルに転記しました")
        lines.append(f"   内訳: " + " / ".join(f"{c}列 {n}セル" for c, n in sorted(cols.items())))
        if plan.skipped:
            lines.append(f"■ 意図的に書かなかったセル: {len(plan.skipped)}")
            for col, row, why in plan.skipped[:10]:
                lines.append(f"   {col}{row}: {why}")
            if len(plan.skipped) > 10:
                lines.append(f"   ... 他 {len(plan.skipped) - 10} セル")
    if plan.warnings:
        lines.append("")
        lines.append(f"■ 要確認 {len(plan.warnings)}件（書き込みは行いましたが、目視での確認を推奨します）")
        for w in plan.warnings:
            lines.append(f"   [{w.kind}] {w.where}")
            lines.append(f"       {w.detail}")
    return "\n".join(lines)
