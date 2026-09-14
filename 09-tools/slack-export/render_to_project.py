"""保存済みの生JSONから Markdown を作り直し、01_Projects/ 配下に書き出す。

Slack への再アクセスは行わない（raw_*.json だけを読む）。
差分取得のたびに同じファイルを上書きするため、ファイル名は初回取得日で固定する。
"""
import datetime
import json
import os

import slack_export as SE

CHANNEL_ID = "C0BD8E0NZQW"
CHANNEL_NAME = "pjt_アルファリース"
PROJECT = "アルファリース"
FIRST_FETCH = "2026-08-09"          # ファイル名を固定するための初回取得日

HERE = os.path.dirname(os.path.abspath(__file__))
RAW_PATH = os.path.join(HERE, f"raw_{CHANNEL_ID}.json")
VAULT = "<VAULT_ROOT>"
OUT_DIR = os.path.join(VAULT, "01_Projects", PROJECT)

raw = json.load(open(RAW_PATH))
messages, threads, users = raw["messages"], raw["threads"], raw["users"]

channels = {CHANNEL_ID: CHANNEL_NAME}
body, msgs = SE.render(messages, users, channels, threads)

reply_total = sum(len(v) for v in threads.values())
thread_count = len([v for v in threads.values() if v])
first = datetime.datetime.fromtimestamp(float(msgs[0]["ts"])).strftime("%Y-%m-%d")
last = datetime.datetime.fromtimestamp(float(msgs[-1]["ts"])).strftime("%Y-%m-%d")
today = datetime.date.today().isoformat()

# 発言者を件数順に（親＋返信を合算）
counts = {}
for m in msgs:
    if m.get("user"):
        n = users.get(m["user"], m["user"])
        counts[n] = counts.get(n, 0) + 1
for lst in threads.values():
    for r in lst:
        if r.get("user"):
            n = users.get(r["user"], r["user"])
            counts[n] = counts.get(n, 0) + 1
ranked = sorted(counts.items(), key=lambda kv: -kv[1])

header = f"""---
title: Slackログ_{CHANNEL_NAME}
created: {FIRST_FETCH}
updated: {today}
tags: [slack, アルファリース, ログ, キャッチアップ]
status: reference
source: "Slack #{CHANNEL_NAME} ({CHANNEL_ID})"
---

## 概要

Slack `#{CHANNEL_NAME}` の全履歴。案件のキャッチアップ用。差分取得で更新される。

| 項目 | 値 |
|---|---|
| チャンネル | `#{CHANNEL_NAME}` |
| 期間 | {first} 〜 {last} |
| 親メッセージ | {len(msgs)}件 |
| スレッド返信 | {reply_total}件（{thread_count}スレッド） |
| 初回取得 | {FIRST_FETCH} |
| 最終更新 | {today} |

### 発言者（件数順）

| 発言者 | 件数 |
|---|---|
"""
for name, n in ranked:
    header += f"| {name} | {n} |\n"

header += "\n## ログ\n"

footer = """

## 関連
- [[アルファリース]]
"""

os.makedirs(OUT_DIR, exist_ok=True)
out_path = os.path.join(OUT_DIR, f"{FIRST_FETCH}_slack_{CHANNEL_NAME}.md")
with open(out_path, "w") as f:
    f.write(header + body + footer)

print(f"✅ {out_path}")
print(f"   {len(msgs)}件 + 返信{reply_total}件 / {first}〜{last}")
print(f"   {os.path.getsize(out_path):,} bytes")
