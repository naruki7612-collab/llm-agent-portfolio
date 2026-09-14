# -*- coding: utf-8 -*-
"""
請求書突合結果 HTMLレポート生成スクリプト（実行環境想定）

「請求書抽出突合スクリプト.py」の main() が出力する
`抽出結果_raw.json` / `突合結果_raw.json` を読み込み、1つの自己完結型HTMLに
まとめる。デザインは同フォルダの `36協定届一覧_案.html` を参考にしている。

## 画面の考え方
- 上部の一覧テーブル：突合対象（明細グループ）1件＝1行。行をクリックすると
  下に詳細パネルが開く。
- 詳細パネル左：抽出した値の一覧＋リフォーム一覧／修繕売上一覧それぞれの判定結果。
  - **一致**：一致したExcel行のデータをそのまま表示
  - **確度が低い（medium/low）**：候補のExcel行データ＋「〜の理由で確度が低いです。
    ご確認ください」というコメント
  - **該当なし（no_match）**：断定はせず、alternative_candidatesの候補行データを
    「〜ではないかと思われる候補」として複数提示＋「一致させられませんでした。
    ご確認ください」というコメント
- 詳細パネル右：元の請求書PDFのページ画像（複数ページある場合はボタンでページ送り）。

## 使い方
1. 先に `請求書抽出突合スクリプト.py` の main() を実行し、
   `抽出結果_raw.json` / `突合結果_raw.json` を生成しておく。
2. このスクリプトの `generate_html_report()` を呼ぶ（`main()`からも呼べる）。
3. 出力される `突合結果レポート.html` をブラウザで開く。
"""

from __future__ import annotations

import base64
import json
import sys
from pathlib import Path
from typing import Any, Optional

import fitz  # PyMuPDF（請求書PDFをページ画像化するため）

# 抽出突合スクリプト本体から、Excel読み込み関数とパス設定を再利用する
import importlib.util

_SCRIPT_PATH = Path(__file__).resolve().parent / "請求書抽出突合スクリプト.py"
_spec = importlib.util.spec_from_file_location("invoice_agent_core", _SCRIPT_PATH)
_core = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_core)

BASE_DIR = _core.BASE_DIR
UPLOADS_DIR = _core.UPLOADS_DIR
EXTRACTION_RESULTS_JSON_PATH = _core.EXTRACTION_RESULTS_JSON_PATH
MATCH_RESULTS_JSON_PATH = _core.MATCH_RESULTS_JSON_PATH

HTML_REPORT_PATH = BASE_DIR / "突合結果レポート.html"

# Excel側で見せる列（多すぎると読みにくいので主要なものだけに絞る）
RIFORM_DISPLAY_COLS = [
    "row_id", "sheet_name", "vendor_name", "property_name", "room_number",
    "work_type", "billing_content", "vendor_amount", "billing_amount_to_customer",
    "completion_date", "payment_method",
]
SHUZEN_DISPLAY_COLS = [
    "row_id", "construction_no", "vendor_name", "property_name", "room_number",
    "work_type", "billing_content", "vendor_amount", "billing_amount_to_customer",
    "completion_date",
]


# ==============================================================
# 1. PDFをページ画像化してbase64にする
# ==============================================================

def render_pdf_pages_to_base64(pdf_path: Path, zoom: float = 2.0) -> list[str]:
    """PDFの各ページをPNG画像化し、data:image/png;base64,... の配列で返す。"""
    images: list[str] = []
    doc = fitz.open(str(pdf_path))
    mat = fitz.Matrix(zoom, zoom)
    for page in doc:
        pix = page.get_pixmap(matrix=mat)
        png_bytes = pix.tobytes("png")
        b64 = base64.b64encode(png_bytes).decode("ascii")
        images.append(f"data:image/png;base64,{b64}")
    doc.close()
    return images


# ==============================================================
# 2. Excel行の検索（row_idからフルデータを引く）
# ==============================================================

def build_excel_lookup() -> dict[str, dict[str, Any]]:
    """row_id -> 行データ の辞書を作る（リフォーム一覧・修繕売上一覧の両方）。"""
    lookup: dict[str, dict[str, Any]] = {}
    for row in _core.load_riform_rows():
        lookup[row["row_id"]] = row
    for row in _core.load_shuzen_rows():
        lookup[row["row_id"]] = row
    return lookup


def _row_display(row: Optional[dict[str, Any]], cols: list[str]) -> dict[str, Any]:
    if not row:
        return {}
    return {c: row.get(c) for c in cols if c in row}


# ==============================================================
# 3. HTML本体の組み立て
# ==============================================================

_STYLE = """
body{font-family:-apple-system,'Hiragino Sans',sans-serif;font-size:13px;margin:0;background:#fff;color:#111}
h1{font-size:16px;margin:12px 14px}
.legend{margin:0 14px 10px}.legend span{display:inline-block;padding:2px 8px;margin-right:6px;border:1px solid #bbb;border-radius:3px}
#top{margin:0 14px;border-collapse:collapse;width:calc(100% - 28px)}
#top th,#top td{border:1px solid #bbb;padding:5px 7px;vertical-align:top;text-align:left}
#top th{background:#e8e8e8;position:sticky;top:0}
#top tr.r{cursor:pointer}
#top tr.r:hover td{outline:2px solid #4a90e2}
tr.j-match td{background:#E2F0D9}
tr.j-review td{background:#FFF2CC}
tr.j-nomatch td{background:#F8CBAD}
tr.sel td{outline:2px solid #1a5e1a}
#detail{display:none;margin:0 14px 20px;padding:10px 12px;border:2px solid #333;border-top:none}
#detail.on{display:flex;gap:14px}
#vals{flex:0 0 620px;max-height:82vh;overflow:auto}
#vals table{border-collapse:collapse;width:100%;margin-bottom:14px}
#vals th,#vals td{border:1px solid #ccc;padding:4px 7px;vertical-align:top}
#vals th{background:#eee;text-align:left;width:170px}
#viewer{flex:1;min-width:0}
#vtool{margin-bottom:6px;color:#555}
#vtool button{margin-left:6px}
#stage{position:relative;display:inline-block;max-width:100%;border:1px solid #ccc}
#stage img{display:block;max-width:100%;height:auto}
.section-title{font-weight:bold;margin:10px 0 4px;font-size:13px}
.comment{background:#FFF2CC;border:1px solid #E0B040;border-radius:4px;padding:8px 10px;margin-bottom:8px;font-size:12.5px;line-height:1.6}
.comment.nomatch{background:#F8CBAD;border-color:#D07050}
.candlabel{font-size:11px;color:#a35e00;margin:6px 0 2px}
.badge{display:inline-block;padding:1px 7px;border-radius:8px;font-size:11px;margin-left:6px}
.badge.high{background:#E2F0D9;color:#2e6b1f}
.badge.medium,.badge.low{background:#FFF2CC;color:#8a6d00}
.badge.no_match{background:#F8CBAD;color:#a03e1a}
"""

_SCRIPT_TMPL = """
const D = __DATA_JSON__;

function badge(conf){
  const label = {high:'一致', medium:'要確認', low:'要確認', no_match:'該当なし', error:'エラー'}[conf] || conf;
  return `<span class="badge ${conf}">${label}</span>`;
}

function rowClass(t){
  const r = t.match_riform.confidence, s = t.match_shuzen.confidence;
  if (r==='high' && s==='high') return 'j-match';
  if (r==='no_match' && s==='no_match') return 'j-nomatch';
  return 'j-review';
}

function renderTop(){
  const rows = D.targets.map((t,i)=>{
    return `<tr class="r" data-i="${i}"><td>${t.source_file}</td><td>${t.line_index ?? ''}</td>`+
      `<td>${t.vendor_name||''}</td><td>${t.property_name||''}</td><td>${t.room_number||''}</td>`+
      `<td>${t.completion_date||''}</td><td>${t.vendor_amount ?? ''}</td>`+
      `<td>${badge(t.match_riform.confidence)}</td><td>${badge(t.match_shuzen.confidence)}</td></tr>`;
  }).join('');
  document.getElementById('top').innerHTML =
    '<tr><th>請求書</th><th>#</th><th>業者名</th><th>物件名</th><th>号室</th><th>完了日</th><th>金額</th>'+
    '<th>リフォーム一覧</th><th>修繕売上一覧</th></tr>' + rows;
  document.querySelectorAll('#top tr.r').forEach((tr,i)=>{
    tr.classList.add(rowClass(D.targets[i]));
    tr.addEventListener('click', ()=>openDetail(i));
  });
}

function excelRowTable(row, cols){
  if (!row || Object.keys(row).length===0) return '<p style="color:#888">（データなし）</p>';
  const lines = cols.filter(c=>c in row).map(c=>`<tr><th>${c}</th><td>${row[c] ?? ''}</td></tr>`).join('');
  return `<table>${lines}</table>`;
}

function renderExcelSide(label, match){
  let html = `<div class="section-title">${label}</div>`;
  const conf = match.confidence;
  if (conf === 'high'){
    html += excelRowTable(match.matched_row, match.cols);
  } else if (conf === 'medium' || conf === 'low'){
    html += `<div class="comment">確度が低いです。理由: ${match.reasoning || ''}${match.mismatched_fields && match.mismatched_fields.length? '<br>差分: '+match.mismatched_fields.join('; '):''}<br>ご確認をお願いします。</div>`;
    if (match.matched_row) {
      html += `<div class="candlabel">候補行:</div>` + excelRowTable(match.matched_row, match.cols);
    }
  } else if (conf === 'no_match'){
    html += `<div class="comment nomatch">一致させられませんでした。理由: ${match.reasoning || ''}<br>ご確認をお願いします。</div>`;
    if (match.alt_rows && match.alt_rows.length){
      html += `<div class="candlabel">こちらではないかと思われる候補:</div>`;
      match.alt_rows.forEach(r=>{ html += excelRowTable(r, match.cols); });
    } else {
      html += `<p style="color:#888">近い候補も見つかりませんでした。</p>`;
    }
  } else {
    html += `<div class="comment nomatch">判定エラー: ${match.reasoning || ''}</div>`;
  }
  return html;
}

let curFile = null, curPage = 1;

function openDetail(i){
  const t = D.targets[i];
  document.querySelectorAll('#top tr.r').forEach(tr=>tr.classList.remove('sel'));
  document.querySelectorAll('#top tr.r')[i].classList.add('sel');

  let valsHtml = '<div class="section-title">請求書から抽出した値</div><table>';
  valsHtml += `<tr><th>業者名</th><td>${t.vendor_name||''}</td></tr>`;
  valsHtml += `<tr><th>物件名</th><td>${t.property_name||''}</td></tr>`;
  valsHtml += `<tr><th>号室</th><td>${t.room_number||''}</td></tr>`;
  valsHtml += `<tr><th>完了日</th><td>${t.completion_date||''}</td></tr>`;
  valsHtml += `<tr><th>金額</th><td>${t.vendor_amount ?? ''}</td></tr>`;
  valsHtml += `<tr><th>作業内容</th><td>${t.billing_content||''}</td></tr>`;
  valsHtml += '</table>';

  valsHtml += renderExcelSide('リフォーム一覧の判定根拠', t.match_riform);
  valsHtml += renderExcelSide('修繕売上一覧の判定根拠', t.match_shuzen);

  document.getElementById('vals').innerHTML = valsHtml;

  curFile = t.source_file;
  curPage = 1;
  renderViewer();

  document.getElementById('detail').classList.add('on');
  document.getElementById('detail').scrollIntoView({behavior:'smooth'});
}

function changePage(delta){
  const imgs = D.imgs[curFile] || [];
  if (!imgs.length) return;
  curPage = Math.min(imgs.length, Math.max(1, curPage + delta));
  renderViewer();
}

function renderViewer(){
  const imgs = D.imgs[curFile] || [];
  const stage = document.getElementById('stage');
  if (!imgs.length){
    document.getElementById('vtool').textContent = '（この請求書の画像はありません）';
    stage.innerHTML = '';
    return;
  }
  const idx = curPage - 1;
  document.getElementById('vtool').innerHTML =
    `${curFile}　${idx+1}/${imgs.length}ページ ` +
    (imgs.length>1 ? `<button onclick="changePage(-1)">◀前ページ</button> <button onclick="changePage(1)">次ページ▶</button>` : '');
  stage.innerHTML = `<img src="${imgs[idx]}">`;
}

renderTop();
"""


def _match_display(match: dict[str, Any], excel_lookup: dict[str, dict[str, Any]], cols: list[str]) -> dict[str, Any]:
    """突合結果1件（match_riform or match_shuzen）を、HTML表示用に整形する。"""
    matched_row_id = match.get("matched_row_id")
    matched_row = excel_lookup.get(matched_row_id) if matched_row_id else None
    alt_ids = match.get("alternative_candidates") or []
    alt_rows = [excel_lookup[rid] for rid in alt_ids if rid in excel_lookup]
    return {
        "confidence": match.get("confidence", "no_match"),
        "reasoning": match.get("reasoning", ""),
        "mismatched_fields": match.get("mismatched_fields") or [],
        "matched_row": _row_display(matched_row, cols),
        "alt_rows": [_row_display(r, cols) for r in alt_rows],
        "cols": cols,
    }


def generate_html_report() -> Path:
    if not MATCH_RESULTS_JSON_PATH.exists():
        print(f"突合結果が見つかりません: {MATCH_RESULTS_JSON_PATH}", file=sys.stderr)
        sys.exit(1)

    with open(MATCH_RESULTS_JSON_PATH, "r", encoding="utf-8") as f:
        match_results = json.load(f)

    print("[HTML] Excelデータを読み込み中...")
    excel_lookup = build_excel_lookup()

    print("[HTML] 請求書PDFを画像化中...")
    imgs: dict[str, list[str]] = {}
    for source_file in {t["source_file"] for t in match_results}:
        pdf_path = UPLOADS_DIR / source_file
        if not pdf_path.exists():
            print(f"  [警告] PDFが見つかりません: {pdf_path}", file=sys.stderr)
            imgs[source_file] = []
            continue
        imgs[source_file] = render_pdf_pages_to_base64(pdf_path)
        print(f"  {source_file}: {len(imgs[source_file])}ページ")

    targets = []
    for t in match_results:
        t2 = dict(t)
        t2["match_riform"] = _match_display(t.get("match_riform") or {}, excel_lookup, RIFORM_DISPLAY_COLS)
        t2["match_shuzen"] = _match_display(t.get("match_shuzen") or {}, excel_lookup, SHUZEN_DISPLAY_COLS)
        targets.append(t2)

    data = {"targets": targets, "imgs": imgs}
    data_json = json.dumps(data, ensure_ascii=False, default=str)
    # データ中に "</script" という文字列が紛れていた場合に<script>タグが
    # 途中で閉じてしまう事故を防ぐ（reasoning等は請求書内容に応じた自由記述のため）。
    data_json = data_json.replace("</script", "<\\/script")

    html = f"""<!doctype html><html lang="ja"><head><meta charset="utf-8">
<title>請求書突合結果レポート</title>
<style>{_STYLE}</style></head>
<body>
<h1>請求書突合結果レポート（{len(targets)}件）</h1>
<p class="legend">
<span style="background:#E2F0D9">両方一致</span>
<span style="background:#FFF2CC">要確認（確度が低い）</span>
<span style="background:#F8CBAD">両方該当なし</span>
　行をクリックで詳細を展開。右側に元の請求書PDFを表示します。
</p>
<table id="top"></table>
<div id="detail">
  <div id="vals"></div>
  <div id="viewer">
    <div id="vtool"></div>
    <div id="stage"></div>
  </div>
</div>
<script>{_SCRIPT_TMPL.replace("__DATA_JSON__", data_json)}</script>
</body></html>"""

    with open(HTML_REPORT_PATH, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"[HTML] 保存しました: {HTML_REPORT_PATH.name}（{len(html)/1024/1024:.1f}MB）")
    return HTML_REPORT_PATH


if __name__ == "__main__":
    generate_html_report()
