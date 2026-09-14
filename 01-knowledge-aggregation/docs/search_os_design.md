# 検索OS v3 設計書（確定版）

> 2026-07-15 確定。raw DB（4ソース統合）＋人物辞書「人」＋domain_map を読む検索エンジン。
> v2（notion_集約_v2.py・旧4DB構成）は温存し、`notion_集約_v3.py` を新規作成する。
> 実装は Phase 1〜3 の段階方式（各Phase末に実データでテスト）。

## 0. 前提データ

- raw DB `<RAW_DB_ID>` — 11プロパティ・source=slack/gmail/drive/meeting
- 人 DB `<PEOPLE_DB_ID>` — name/email/aliases/company（自動構築済み）
- domain_map `<DOMAIN_MAP_DB_ID>` — project(title)⇄domain（手動管理）

## 1. パイプライン

```
answer_query(質問, requester_hint)
  ① プランナー（レシピ選択・project・人物・期間・キーワード・fact_critical）
  ② 解決層（resolve_person 3層 / resolve_company）
  ③ 取得層（raw横断クエリ: source/project/date型timestamp/キーワードcontains）
  ④ 再計画（0件 or 300件超で1回だけ）
  ⑤ 文脈層（parents解釈テーブル経由のスレッド/出どころ復元）
  ⑥ 深読み層（meeting本文＝自分のDBのみ・gmail原文URL復元。Tier判定/fact必読。
     ★Drive/Miro/Canva 等の外部ファイルは再取得せず、書く側の要約JSON＋リンク提示で足りる）
  ⑦ 合成層（findings=出典URL必須 / not_found / factは逐語引用quote）
```

## 2. ソース別セマンティクス表（v3の心臓部。上位層はこの表経由でのみrawを解釈）

| | slack | gmail | drive | meeting |
|---|---|---|---|---|
| content | 本文(2000字上限) | 本文＋[添付]＋[mail:ID] | 要約JSON | 議事録全文 |
| subject | — | 件名 | ファイル名 | 会議タイトル |
| sender/name | 投稿者メール/表示名 | 送信者/送受信者名 | 更新者 | 参加者メール/名前(複数) |
| **parents** | **親メッセージts** | **threadId** | **出どころ**(フォルダ階層/"slack"/"gmail") | **meeting_uuid** |
| links | URL＋要約 | 添付「名前\nURL」 | 同左(matchキー) | — |
| キーワード検索対象 | content, links | content, subject | content, subject, **parents** | content, subject |
| 人物の探し場所 | sender, name | sender, to, cc_bcc, name | sender, name | sender, name |
| スレッド復元 | parents=親ts(＋project絞り。親はts±1分→クライアント側一致) | parents=threadId equals | なし(出どころフィルタ) | なし(1会議=1行) |

**parents解釈の注意**: meetingのuuidは定例の回ごとに異なる（系列で束ねるならsubject一致）。
driveの出どころ判定: `"slack"`/`"gmail"` equals、フォルダ由来は `contains "pjt_"`。

## 3. 解決層

**resolve_person（3層＋安全弁）**
1. クエリ正規化＋LLM表記変換（ましこ→担当A/Mashiko/mashiko）
2. 人DBを候補ごとに name/aliases/email contains → 1人確定=email＋aliases一式でOR検索 /
   **複数人=混ぜずに候補提示** / 0件→3へ
3. raw逆引き（slack/drive行=1行1人ペア、gmailはsender別の共通名）→ 辞書に自己登録
4. 最終フォールバック: 入力文字列のcontains（v2と同じ）

**resolve_company**: domain_mapタイトルcontains → {project, domain} →
project絞り / sender・to contains domain / 人DB company=domain（担当者一覧）

## 4. レシピ（v2の9本を改訂）

決定事項の根拠確認 / 過去資料の参照（parentsフォルダ名検索込み）/ 最新議事録（1クエリ）/
期間キャッチアップ / PJ全体まとめ（月次map-reduce）/ 自分・メンバー横断（1〜2クエリ化）/
顧客温度感（resolve_company起点・軽量）/ 顧客像 / custom（深読みTier判定込み）

**共通ルール**:
- projectで絞る検索でも、キーワード指定時は project=未分類 の行も対象に含める。
- **取得は常に全4ソース**。レシピが違うのは処理（件数の絞り・並べ方・深読みの深さ）だけで、
  ソースは絞らない（統合rawDBの意味／レシピのソース限定で情報を見逃すバグを設計で排除）。
  絞りは project / キーワード / 期間 / 人物で行う。

## 5. 出力仕様

- findings 1件ごとに sources 必須。根拠なしは not_found（推測禁止）
- fact_critical 10分野は sources.quote に原文の逐語引用
- **出典URLは source_id 参照方式**: 材料に [S番号] を振りコード側で {title,date,url} を保持。
  LLMは source_id だけ引用し、合成後にコードが本物URLへ復元（幻IDは破棄）。
  LLMに長いURLを書かせないことで破壊・取り違えを防ぐ
- 各行の url は _best_url が優先順で決定: Gmail原文リンク([mail:ID]復元) >
  Driveファイル直リンク > Notion行URL

## 6. 実装Phase

| Phase | 内容 | 出口 |
|---|---|---|
| 1 | セマンティクス表＋取得層＋search()（自然言語→横断取得） | 期間×PJ×キーワード検索が動く |
| 2 | resolve_person / resolve_company ＋スレッド・出どころ復元 | 人名・会社名クエリが動く |
| 3 | プランナー・レシピ・深読み・合成（v2から移植改訂） | answer_query フル稼働 |

- エントリポイントは v2 互換の answer_query ＋ 部品関数（search/resolve_person/resolve_company）も公開
- 検索エージェント用システムプロンプトは rag_search ファイル名発見＋spec import 方式（Gmail/Zoomと同じ）
