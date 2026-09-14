# -*- coding: utf-8 -*-
"""BETA 請求書自動仕訳 — 設定。全段・s60_feedback・selftest が読む。"""

# --- Note の中の置き場所 ---------------------------------------------------
# **`notes/` は AgentPlatform が Note をサンドボックスに見せるときの固定のマウント先。**
# Note の名前（請求書自動仕訳）はパスに出てこない。中のフォルダが notes/ の直下に並ぶ。
#
#   notes/
#     01_証憑/           … 人がPDFを置く
#     02_仕訳データ/     … 結果CSVが書かれる
#     03_確認済み仕訳/   … 人が直したCSVを置く（辞書に反映される）
#     システム/          … 処理のプログラム一式。触らない
#
# 実際に間違えた: NOTE_ROOT を "請求書自動仕訳" にしてしまい、
# 「ファイルが見つかりません」で止まった。**notes/ を外さないこと。**
NOTE_ROOT  = "notes"
SCRIPT_DIR = NOTE_ROOT + "/システム"

# --- db のテーブル名 ---
T_JOURNAL      = "beta_journal"        # 確定仕訳（学習の元）
T_ITEM_ACCOUNT = "beta_item_account"   # 支払先 × 品名 → 4項目
T_VEHICLE      = "beta_vehicle"        # 車両マスタ
T_CODE_RULE    = "beta_code_rule"      # 手書き3段コードの変換規則
T_GLOBAL_ITEM  = "beta_global_item"    # 全社の品目 → 勘定科目
T_ACCT_MGMT    = "beta_acct_mgmt"      # 管理⇔事業の科目対応
T_PAYEE_ALIAS  = "beta_payee_alias"    # 支払先の別名
T_MASTER       = "beta_master"         # 部門・セグメント・科目・補助の名称

DICT_TABLES = [T_JOURNAL, T_ITEM_ACCOUNT, T_VEHICLE, T_CODE_RULE,
               T_GLOBAL_ITEM, T_ACCT_MGMT, T_PAYEE_ALIAS, T_MASTER]

# --- OneDrive（Composio）。None なら s10 は uploads/ を使う ---
# ツール名と引数は s10_ingest.show_tools() で確認してから埋める。
ONEDRIVE = {
    "list_tool":     None,   # 例: "ONE_DRIVE_LIST_FILES"
    "list_args":     {},     # 例: {"folder_path": "/請求書/未処理"}
    "download_tool": None,   # 例: "ONE_DRIVE_DOWNLOAD_FILE"
    "download_key":  "id",   # list の各要素から download に渡すキー名
    "download_arg":  "file_id",
}

# --- 作業ディレクトリ（すべて tmp/ 配下）---
JOB      = "tmp/beta"
D_SRC    = JOB + "/src"        # 取り込んだ PDF
D_IMG    = JOB + "/img"        # PDFを1ページずつ画像にしたもの

# --- 証憑PDFの置き場所。s10 はここと uploads/ の両方を見る ---
D_NOTE_SRC = NOTE_ROOT + "/01_証憑"

# --- 人が直した確認済みCSVの置き場所。s60 はここと uploads/ の両方を見る ---
D_FEEDBACK = NOTE_ROOT + "/03_確認済み仕訳"

# --- 各段の途中結果 ---
D_EXTRACT= JOB + "/extract"    # s20（窓ごと・PDFごと）
D_VOUCH  = JOB + "/voucher"    # s30/s32
D_SPLIT  = JOB + "/split"      # s35（伝票ごと）
D_CLASS  = JOB + "/class"      # s40（伝票ごと）

# --- Noteに残すのはこの3ファイルだけ（file_create を通すのもここだけ）---
# file_create で書いただけではNoteに保存されない。この会話は自動同期がOFFなので、
# アシスタントが notes_sync を呼ぶまでバックエンドに届かない（このスクリプトからは
# 呼べない）。呼び忘れると3ファイルすべてが失われる。
NOTE_STATE  = NOTE_ROOT + "/02_仕訳データ"
OUT1_NOTE   = NOTE_STATE + "/奉行取込用.csv"   # 結果の本体。セッションをまたいで積み増す
OUT2_NOTE   = NOTE_STATE + "/要確認.csv"       # そのうち要確認フラグが立った行
STATUS_PATH = NOTE_STATE + "/進捗.txt"         # いまどの段か（毎回上書き）

# --- サンドボックス側の控え（ダウンロード用。セッションが終わると消える）---
D_OUTPUT    = "output"

# --- デバッグログ（各段が何を受け取って何を返したか。1行1件のJSON）---
# **お客様が見る 02_仕訳データ/ ではなく システム/ に置く。**
# 要らなくなったら False にすれば、記録も書き出しもしなくなる。
DEBUG_LOG   = True
LOG_PATH    = SCRIPT_DIR + "/ログ.jsonl"

def check_paths():
    """設定したフォルダが本当に在るか調べる。戻り: 見つからなかったものの一覧。

    **glob で確かめる。** Note のファイルはゴーストなので os.path.exists（stat）では
    見えないことがある。glob はディレクトリの一覧を取るので効く。

    NOTE_ROOT を間違えると全段が「ファイルがありません」で止まる。実際に
    "請求書自動仕訳" と書いてしまって止まったので、走り出す前にここで気づけるようにした。
    """
    import glob as _g
    ng = []
    for name, path in (("証憑", D_NOTE_SRC), ("出力先", NOTE_STATE),
                       ("確認済み", D_FEEDBACK), ("スクリプト", SCRIPT_DIR)):
        if not (_g.glob(path) or _g.glob(path + "/*")):
            ng.append("%s: %s" % (name, path))
    return ng


def where_is_note():
    """Note がどこに見えているかを探して、候補を返す（迷子になったとき用）。"""
    import glob as _g
    out = []
    for root in ("notes", ".", "note", "Notes"):
        for d in ("01_証憑", "システム"):
            if _g.glob("%s/%s" % (root, d)):
                out.append("%s/%s" % (root, d))
    return sorted(set(out))


# --- 実行パラメータ ---
CONCURRENCY   = 6        # llm_call の同時実行数（AgentPlatformの並列スロットは全体16）
BATCH_FILES   = 2        # 1バッチで処理するPDFの本数。大きくすると15分に収まらない
                         # （5本だとぎりぎりだったので2本にした）

# --- ページ画像 ---
PAGE_DPI      = 150     # 1241×1756
PAGE_WINDOW   = 8       # 1回の llm_call が「担当」するページ数
PAGE_BACK     = 1       # 担当の前に何ページ重ねて見せるか（続きページの判定用）
PAGE_FWD      = 3       # 担当の後ろに何ページ重ねて見せるか（またがる帳票用）
SPLIT_MAX_IMG = 12      # s35 に渡すページ画像の上限

# --- モデル。完全一致でしか通らない（"" は既定の軽量モデル）---
# Step 5 / Step 6 が「Anthropic returned no structured output」で落ちる場合は
# 既知の問題.md の⑪を見る（prompts.py の2つのスキーマを required 化する手がある）。
MODEL_EXTRACT = "gemini/gemini-3.7-flash"   # 画像から読む段。anthropic/claude-opus-5 も可
MODEL_SPLIT   = "anthropic/claude-opus-5"
MODEL_CLASS   = "anthropic/claude-opus-5"
MODEL_SUMMARY = "anthropic/claude-opus-5"

NEW_RULE_FROM = "202606"  # これ以降の実績を「新」とし、値が割れたらこちらを優先
CONF_THRESHOLD= 0.8       # これ未満は要確認
DB_CHUNK      = 200       # insert_many の1回あたり件数
