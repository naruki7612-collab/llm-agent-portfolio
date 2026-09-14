"""再開可能版。履歴とスレッドを取得のたびにディスクへ書き出す。

途中でレート制限に負けて落ちても、再実行すれば取得済み分はスキップされる。
"""
import json
import os
import sys
import time

import slack_export as SE

CHANNEL_ID = sys.argv[1] if len(sys.argv) > 1 else "C0BD8E0NZQW"
CHANNEL_NAME = sys.argv[2] if len(sys.argv) > 2 else "pjt_アルファリース"

SCRATCH = os.path.dirname(os.path.abspath(__file__))
HIST_PATH = os.path.join(SCRATCH, f"hist_{CHANNEL_ID}.json")
THREAD_DIR = os.path.join(SCRATCH, f"threads_{CHANNEL_ID}")
USERS_PATH = os.path.join(SCRATCH, "users.json")
RAW_PATH = os.path.join(SCRATCH, f"raw_{CHANNEL_ID}.json")

os.makedirs(THREAD_DIR, exist_ok=True)


def load(path):
    return json.load(open(path)) if os.path.exists(path) else None


def save(path, obj):
    with open(path, "w") as f:
        json.dump(obj, f, ensure_ascii=False)


# ---------- ユーザー辞書（キャッシュ） ----------
users = load(USERS_PATH)
if users:
    SE.log(f"[1/3] ユーザー辞書: キャッシュから {len(users)}人")
else:
    SE.log("[1/3] ユーザー辞書を取得中…")
    users = SE.build_user_map()
    save(USERS_PATH, users)
    SE.log(f"  {len(users)}人")

# ---------- 履歴（キャッシュ） ----------
messages = load(HIST_PATH)
if messages:
    SE.log(f"[2/3] 履歴: キャッシュから {len(messages)}件")
else:
    SE.log("[2/3] チャンネル履歴を取得中…")
    messages = SE.fetch_all_history(CHANNEL_ID)
    save(HIST_PATH, messages)
    SE.log(f"  {len(messages)}件 保存")

# ---------- スレッド（1本ずつ保存） ----------
parents = [m for m in messages if (m.get("reply_count") or 0) > 0]
done = {f[:-5] for f in os.listdir(THREAD_DIR) if f.endswith(".json")}
todo = [p for p in parents if p["ts"] not in done]
SE.log(f"[3/3] スレッド {len(parents)}本中 {len(done)}本取得済み、残り {len(todo)}本")

for i, p in enumerate(todo, 1):
    try:
        replies = SE.fetch_thread(CHANNEL_ID, p["ts"])
    except Exception as e:
        SE.log(f"  [失敗] ts={p['ts']}: {str(e)[:120]} → スキップして継続")
        continue
    save(os.path.join(THREAD_DIR, f"{p['ts']}.json"), replies)
    if i % 5 == 0 or i == len(todo):
        SE.log(f"  {i}/{len(todo)} 本完了（累計 {len(done) + i}/{len(parents)}）")
    time.sleep(1)  # 制限に当たりにくくするための間隔

# ---------- 結合して raw を書き出し ----------
threads = {}
for f in os.listdir(THREAD_DIR):
    if f.endswith(".json"):
        threads[f[:-5]] = json.load(open(os.path.join(THREAD_DIR, f)))

save(RAW_PATH, {"messages": messages, "threads": threads, "users": users})
reply_total = sum(len(v) for v in threads.values())
SE.log(f"\n✅ raw 保存完了: {RAW_PATH}")
SE.log(f"   親 {len(messages)}件 / スレッド {len(threads)}本 / 返信 {reply_total}件")
