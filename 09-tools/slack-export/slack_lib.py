"""Composio CLI 経由で Slack を叩く薄いラッパー。"""
import json
import subprocess
import time

ACCOUNT = "slack_pavy-janet"


def call(slug, payload, retries=5):
    """composio execute を叩いて data を返す。ratelimited は指数バックオフ。"""
    for attempt in range(retries):
        proc = subprocess.run(
            ["composio", "execute", slug, "--account", ACCOUNT, "-d", json.dumps(payload)],
            capture_output=True, text=True,
        )
        raw = proc.stdout.strip()
        if not raw:
            raise RuntimeError(f"{slug}: empty output / stderr={proc.stderr[:300]}")
        # Slack のメッセージ本文に生の改行が入るため strict=False
        d = json.loads(raw, strict=False)

        if d.get("successful"):
            return d.get("data") or {}

        err = str(d.get("error") or "")
        if "ratelimit" in err.lower() or "429" in err:
            wait = 2 ** attempt * 5
            print(f"    [rate limited] {wait}s 待機 (attempt {attempt + 1})")
            time.sleep(wait)
            continue
        raise RuntimeError(f"{slug} failed: {err[:300]}")
    raise RuntimeError(f"{slug}: レート制限でリトライ上限")


def find_channel(query):
    data = call("SLACK_FIND_CHANNELS", {"query": query, "limit": 200})
    return data.get("channels") or []


def fetch_history(channel, cursor=None, limit=200, oldest=None):
    payload = {"channel": channel, "limit": limit}
    if cursor:
        payload["cursor"] = cursor
    if oldest:
        payload["oldest"] = oldest
    return call("SLACK_FETCH_CONVERSATION_HISTORY", payload)


def fetch_thread(channel, ts, cursor=None, limit=200):
    payload = {"channel": channel, "ts": str(ts), "limit": limit}
    if cursor:
        payload["cursor"] = cursor
    return call("SLACK_FETCH_MESSAGE_THREAD_FROM_A_CONVERSATION", payload)
