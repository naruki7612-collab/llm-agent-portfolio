"""前回取得以降の差分だけを取ってきて raw JSON を更新する。

新しい親メッセージだけでなく「古い親に付いた新しい返信」も拾う必要があるため、
親の一覧（conversations.history）は毎回フルで取り直し、保存済みの reply_count /
latest_reply と突き合わせて、変化のあったスレッドだけ再取得する。
親一覧は10ページ弱で済むので、78スレッドを毎回舐めるより圧倒的に速い。
"""
import datetime
import json
import os
import sys

import slack_export as SE

CHANNEL_ID = sys.argv[1] if len(sys.argv) > 1 else "C0BD8E0NZQW"
CHANNEL_NAME = sys.argv[2] if len(sys.argv) > 2 else "pjt_アルファリース"

HERE = os.path.dirname(os.path.abspath(__file__))
RAW_PATH = os.path.join(HERE, f"raw_{CHANNEL_ID}.json")
STATE_PATH = os.path.join(HERE, f"state_{CHANNEL_ID}.json")


def load_state():
    if os.path.exists(STATE_PATH):
        return json.load(open(STATE_PATH))
    return {"runs": []}


def thread_fingerprint(m):
    """スレッドの更新を検知するための指紋。"""
    return (m.get("reply_count") or 0, str(m.get("latest_reply") or ""))


def main():
    raw = json.load(open(RAW_PATH))
    old_msgs = {m["ts"]: m for m in raw["messages"]}
    threads = raw["threads"]
    users = raw["users"]

    prev_max = max(float(ts) for ts in old_msgs)
    SE.log(f"前回の最新: {datetime.datetime.fromtimestamp(prev_max):%Y-%m-%d %H:%M}")

    # ---- 親メッセージをフル取得（差分検知のため）----
    SE.log("[1/4] 親メッセージ一覧を取得中…")
    fresh = SE.fetch_all_history(CHANNEL_ID)
    SE.log(f"  {len(fresh)}件（保存済み {len(old_msgs)}件）")

    new_parents = [m for m in fresh if m["ts"] not in old_msgs]
    SE.log(f"  新規の親メッセージ: {len(new_parents)}件")

    # ---- 再取得が必要なスレッドを特定 ----
    todo = []
    for m in fresh:
        rc = m.get("reply_count") or 0
        if rc == 0:
            continue
        old = old_msgs.get(m["ts"])
        if old is None:                                   # 新しい親のスレッド
            todo.append(m)
        elif thread_fingerprint(old) != thread_fingerprint(m):  # 返信が増減した
            todo.append(m)
        elif m["ts"] not in threads:                      # 前回取り逃していた
            todo.append(m)
    SE.log(f"[2/4] 再取得が必要なスレッド: {len(todo)}本")

    added_replies = 0
    for i, p in enumerate(todo, 1):
        before = len(threads.get(p["ts"], []))
        try:
            replies = SE.fetch_thread(CHANNEL_ID, p["ts"])
        except Exception as e:
            SE.log(f"  [失敗] ts={p['ts']}: {str(e)[:100]} → スキップ")
            continue
        threads[p["ts"]] = replies
        added_replies += max(0, len(replies) - before)
        SE.log(f"  {i}/{len(todo)} 本完了（+{len(replies) - before}件）")

    # ---- ユーザー辞書を更新（新メンバーがいる場合）----
    unknown = {m.get("user") for m in fresh if m.get("user") and m["user"] not in users}
    for lst in threads.values():
        unknown |= {r.get("user") for r in lst if r.get("user") and r["user"] not in users}
    unknown.discard(None)
    if unknown:
        SE.log(f"[3/4] 未知のユーザー {len(unknown)}人 → 辞書を取り直し")
        users = SE.build_user_map()
    else:
        SE.log("[3/4] ユーザー辞書は最新")

    # ---- 保存 ----
    SE.log("[4/4] 保存中…")
    merged = {m["ts"]: m for m in fresh}
    for ts, m in old_msgs.items():          # 取得窓から外れた古い親も残す
        merged.setdefault(ts, m)
    messages = sorted(merged.values(), key=lambda m: float(m["ts"]))

    with open(RAW_PATH, "w") as f:
        json.dump({"messages": messages, "threads": threads, "users": users}, f, ensure_ascii=False)

    new_max = max(float(m["ts"]) for m in messages)
    state = load_state()
    state["runs"].append({
        "fetched_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "channel": CHANNEL_NAME,
        "latest_parent_ts": str(new_max),
        "latest_parent_at": datetime.datetime.fromtimestamp(new_max).isoformat(timespec="seconds"),
        "total_parents": len(messages),
        "total_threads": len([v for v in threads.values() if v]),
        "total_replies": sum(len(v) for v in threads.values()),
        "new_parents": len(new_parents),
        "new_replies": added_replies,
    })
    with open(STATE_PATH, "w") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

    SE.log(f"\n✅ 差分取得完了")
    SE.log(f"   新規の親 {len(new_parents)}件 / 新規の返信 {added_replies}件")
    SE.log(f"   累計 親{len(messages)}件・返信{sum(len(v) for v in threads.values())}件")
    SE.log(f"   取得ログ: {STATE_PATH}")


if __name__ == "__main__":
    main()
