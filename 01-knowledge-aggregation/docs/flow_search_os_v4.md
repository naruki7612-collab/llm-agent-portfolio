# 処理フロー — agent_platform_集約_v4（agent_platform DB + Note 版 検索OS）

`notion_集約_v3.py` の agent_platform 版。**検索の仕様は v3 と同じ**（9レシピ・キーワード横断・
人物名寄せ・スレッド復元・深読み・出典必須の合成）。変わったのは問い合わせ先と、
サーバーにできない処理をどこでやるか。

- 実装: `情報集約/情報検索/agent_platform_集約_v4.py`
- 前提: この会話に情報集約用の **Project が attach** されていること（テーブルは Project 配下に自動スコープ）
- 使うもの: `agent_sdk.db`（find / count_documents）、`llm_call`、`notes/` ツリーの直読み

---

## 0. なぜ実装が変わるのか（agent_platform 側の制約）

| v3 が依存していたもの | agent_platform での可否 | v4 での対処 |
|---|---|---|
| `contains`（部分一致） | **無い** | キーワード照合を Python 側へ（`_keyword_match`） |
| `$or` グループ | **無い**（同じ列の複数値は `$in` で代替可） | project は `$in`、人物照合は Python 側へ（`_person_match`） |
| ソート＋ページング併用 | **不可**（sort を付けると `after_id` が使えない） | 取得モードを2本に分岐（下記 §2） |
| 行ID での引き当て `_id` | **backend が 422 で明示拒否** | Relation から行を復元できない。添付は `links`(note_path) の完全一致で引く |
| Notion ページ本文 | 無い | meeting 全文は `content` 列にある。深読みは **file 行の note 実体**を読む |
| DB ID 定数 | 不要 | Project スコープ。未 attach の会話では 422 で落ちる |

`db.search()`（セマンティック検索）は**使っていない**。テーブルに `type=="vector"` の列が
ちょうど1つ必要だが、`create_table_from_csv` は string/number/boolean/date しか作らないため、
現在の raw テーブルには vector 列が無い（呼ぶと 422）。将来 vector 列を付けたら選択肢になる。

---

## 1. 全体フロー

```
answer_query(質問, requester_hint=)
  │
  ├─① plan_query …… プランナー（sonnet-5／失敗時 sonnet-4-6 フォールバック）
  │     9レシピから1つ選択 ＋ project / person_hint / keywords / 期間 /
  │     deep_read / fact_critical を決める
  │
  ├─② _execute_plan …… レシピ実行
  │     ├─ _resolve_plan_person → resolve_person（複数人なら即 clarification を返して終了）
  │     ├─ fetch_records（★常に全ソース。plan の sources は無視＝取りこぼし防止）
  │     ├─ _expand_slack_threads（slack ヒット上位10件を種にスレッド復元）
  │     └─ 深読み Tier 判定 → _deep_read_notes
  │
  ├─③ _needs_replan …… 0件 or 300件超なら **1回だけ**再計画して②をやり直す
  │
  ├─④ （catchup_whole_project のみ）_map_reduce_summarize
  │     月次チャンク × 最大40行で中間要約（flash-lite）
  │
  └─⑤ synthesize_answer …… 材料に [S番号] を振って合成（sonnet-5）
        LLM は source_id だけ引用 → _resolve_sources が本物の出典に復元
        （対応表に無い番号＝幻の出典は破棄。有効出典ゼロの finding も落とす）
```

---

## 2. 取得層（v4 で最も変わったところ）

`fetch_records(sources, project, start_date, end_date, keywords, person)` はソース別に並列実行。

### サーバー側に送る条件（`_build_server_filter`）
構造化条件だけ。ここに**キーワードと人物条件は入らない**。

```python
{"source": "gmail",
 "project": {"$in": ["pjt_アルファリース", "アルファリース"]},   # pjt_ 有無を吸収
 "timestamp": {"$gte": "2026-06-01T00:00:00+09:00",
               "$lte": "2026-06-30T23:59:59+09:00"}}
```
- キーワード指定時は `project` の `$in` に **`未分類` も自動追加**（Zoom の未分類会議の取りこぼし防止・v3 と同じ）

### 取得モードの分岐

| 条件 | 動き | 理由 |
|---|---|---|
| キーワードも人物も無い | `find(sort=[("timestamp",-1)], limit=500)` を**1回** | 新しい順の上位が欲しいだけ。sort が使える |
| キーワード or 人物あり | `after_id` で**全走査**（1000件/ページ・上限20,000件）→ Python 照合 → 降順 → 上位500件 | sort を付けるとページングできない。照合前に打ち切ると古いヒットを落とす |

**打ち切りは必ず WARN を出す**（黙って減らさない）:
- `_top_rows`: 上限に達したら `count_documents` で真の件数を出して「新しい順の上位のみ」と警告
- `_scan_rows`: 20,000件で打ち切ったら「古い行を取りこぼしている可能性」と警告
- 照合後 500件超も件数を明示して警告

### Python 側の照合（v3 のサーバー条件と同じ意味論）
- `_keyword_match`: ソース別の本文系フィールド × キーワードの **OR**。NFKC＋casefold で全角半角・大小文字を吸収（v3 より取りこぼしが減る方向）
- `_person_match`: メール列に email、名前列に表記候補（2文字以上）の **OR**

| source | 本文系フィールド | 人物: email 列 | 人物: 名前列 |
|---|---|---|---|
| slack | content, links | sender | name |
| gmail | content, subject | sender, to, cc_bcc | name |
| drive | content, subject, parents | sender | name |
| meeting | content, subject | sender | name |
| **file** | content, subject, parents | sender | name |

`file` = Gmail 添付を `notes/` に置いた行。drive と同じ扱い（content=要約JSON / subject=ファイル名 / parents=出どころ）。

---

## 3. 解決層

- `resolve_person`: ①敬称除去＋LLM表記変換 → ②**人物辞書を全件読んでローカル照合**
  （contains が無いため。辞書は小さいので全件で足りる）→ 1人確定 / 複数人は candidates を返す
  → ③0件なら raw の slack/drive/file 行から逆引き（1行=1人で name⇄sender が正確）
  → ④それでも不明なら文字列一致にフォールバック
- `resolve_company`: `domain_map` を全件読んで正規化名の相互包含で照合。ドメイン直指定も可
- `company_members`: `find({"company": ドメイン})` の完全一致

---

## 4. 文脈層

- `expand_thread(row)`
  - gmail: `find({"source":"gmail","parents":threadId})` → 時系列昇順
  - slack: 親ts一致の子 ＋ 親本体（`timestamp` ±1分で取ってクライアント側で秒一致）
  - meeting / drive / file: その行のみ
- `expand_attachments(row)`: メール行の `links`（note_path の改行区切り）を1件ずつ
  `find({"note_path": path})` で引く
  - ★`relations_apply` で張ったメール⇄添付の Relation は**読み側では使えない**。
    `relations_query` が返すのは相手の `record_id` だけで、`find({"_id": ...})` は
    backend が 422 で拒否するため行を復元できない

---

## 5. 深読み層

v3 は「meeting 行の Notion ページ本文」だけを深読みしていた。agent_platform では meeting の
議事録全文が `content` 列にそのまま入っているので取り直す必要が無い。代わりに
**DB に載っていない唯一の本文＝添付 note の実体**が深読み対象になる。

```
_collect_deep_read_candidates → file 行の note_path を候補に
  ├─ 候補5件以下 or fact_critical or deep_read → 全部読む
  └─ それ以外 → _triage_deep_read（flash-lite で最大5件に選別）
read_note(note_path) …… notes ルート + note_path を直読み
```

`read_note` が本文を返せないケースは**理由付きで返し、握りつぶさない**:

| reason | 意味 | 合成側の扱い |
|---|---|---|
| `ghost（…）` | セッション開始時の復元でプレースホルダだけが置かれた（バイナリ・大きいファイル） | DB の要約JSONで代替 |
| `binary（…）` | テキストとしてデコードできない（PDF/xlsx 等） | 同上 |
| `notes/ に実体がありません` | 未 sync（書いた当日）or 削除済み | 同上 |

> notes/ は自分たちの資産なので、v3 の「外部ファイルは検索時に再取得しない」原則には反しない
> （外部アプリごとにリーダーを増やす話ではないため）。

---

## 6. 合成層

- `_materials_text`: 各材料に `[S番号]` を振り、`source_map[Sn] = {title, date, url, ref}` を作る
  - `url` = http の一次情報（メール原文 > links 内の Drive リンク）。無ければ空文字
  - `ref` = URL が無い行の「たどれる場所」。file 行は `notes/<note_path>`、それ以外は `raw#<行番号>`
    （agent_platform の行には Notion のようなページURLが無いため v4 で追加した）
- LLM には **source_id だけ**を書かせ、URL は書かせない → `_resolve_sources` が復元し、
  対応表に無い番号は破棄
- fact_critical の finding は `quote` に材料からの逐語引用を必須にする
- 本文の抜粋長は fact_critical=true なら1200字、それ以外400字

---

## 7. 検証状況

`OK 41 / NG 0`（ローカル・ネットワークなし。`adb.table` を偽テーブルに差し替えて実行）

検証済み: フィルタ組み立て（$in / 日付範囲 / 未分類の自動追加）・キーワード照合の
ソース別フィールド・人物照合の列使い分け・Gmail原文URLの復元・出典の代表値選択・
幻の出典の破棄・取得モードの分岐・スレッド復元・添付の引き当て・note 読み出しの4分岐。
偽テーブルには「未対応の演算子（`$or`/`$regex` 等）を送ったら失敗」のガードを入れてある。

**未検証（実データが要る）**:
- `timestamp` 列が `date` 型として作られているか（`agent_platform_db_init.show_schema()` で確認する）。
  `string` になっていた場合も `+09:00` 固定の ISO なので辞書順＝時系列順で動くが、
  `$gte/$lte` の比較意味論が文字列比較になる
- `{"$gte": ..., "$lte": ...}` を1つの列に同時指定できるか（backend 実挙動）
- 20,000件の全走査に実際どれくらい掛かるか（実行環境 の実行窓は約300秒）
- プランナー（sonnet-5）の構造化出力成功率
