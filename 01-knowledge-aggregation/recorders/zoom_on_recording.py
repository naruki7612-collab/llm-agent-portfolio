"""notion_zoom_録画時.py — Zoom録画1件を raw DB に保存する（トリガー用・全ステップ統合版）

旧 step1_2_get_and_download.py（録画情報取得＋VTTダウンロード）と
step3_summarize_and_save.py（議事録生成＋Notion保存）を1本化したもの。

- 保存先は全ソース統合の raw DB（source=meeting。Gmail/Slack と同じDB）
- 重複判定は parents=meeting_uuid ＋ source=meeting の存在チェック
- 議事録は content（title）に、議事録＋生VTTはページ本文ブロックにも保存
- エントリポイント: run(trigger_payload) を呼ぶだけ
  （trigger_payload はトリガーの "payload" オブジェクト。JSON文字列でもdictでも可）
"""
import asyncio
import json
import os
import random
import re
import urllib.request
from datetime import datetime, timedelta, timezone
from urllib.parse import quote

from agent_sdk import (
    NOTION_QUERY_DATABASE_WITH_FILTER,
    NOTION_UPSERT_ROW_DATABASE,
    ZOOM_GET_MEETING_RECORDINGS,
    ZOOM_GET_PAST_MEETING_PARTICIPANTS,
    llm_call,
)

JST = timezone(timedelta(hours=9))

# raw DB ID（固定。全ソース統合DB「社内情報集約PJ > raw」。Gmail/Slack と同じ保存先）
RAW_DB_ID = "<RAW_DB_ID>"
# プロジェクト判定の候補リスト取得元: ドメイン情報DB（Gmail と同じ。タイトル=pjt_プロジェクト名）
# ※旧プロジェクト一覧DB（36e4a092...）は旧構成のため使わない
DOMAIN_INFO_DB_ID = "<DOMAIN_MAP_DB_ID>"
DOMAIN_PROP = "domain"   # domain_map（旧ドメイン情報DB）のドメイン列名
# 人物辞書DB ID（meta > 人。登場人物の自動登録先。空にすると人物登録をスキップ）
PEOPLE_DB_ID = "<PEOPLE_DB_ID>"

# raw DB のプロパティ名（Notion側の列名が変わったらここだけ直す）
PROP_CONTENT   = "content"    # title: 議事録全文（生VTTはページ本文へ）
PROP_SUBJECT   = "subject"    # rich_text: 会議タイトル
PROP_TIMESTAMP = "timestamp"  # date: 開始日時（JST・ISO 8601）
PROP_SENDER    = "sender"     # rich_text: 参加者のメールアドレス（カンマ結合・ホスト含む）
PROP_NAME      = "name"       # rich_text: 参加者名（カンマ結合）
PROP_PROJECT   = "project"    # rich_text: プロジェクト名
PROP_SOURCE    = "source"     # select: 情報ソース
PROP_PARENTS   = "parents"    # rich_text: meeting_uuid（重複判定キー）
SOURCE_MEETING = "meeting"    # source select の選択肢名

NOTION_RICH_TEXT_LIMIT = 2000
DOWNLOAD_TIMEOUT = 120

# 前処理後の文字起こしがこれ未満なら議事録を作らない（文字起こし未完了/無発話の防御）
_MIN_TRANSCRIPT_CHARS = 200
# LLM に渡す文字起こしの上限（長時間会議でのコンテキスト/コスト爆発防止）
_MAX_TRANSCRIPT_CHARS = 120_000

# デバッグ用チェックポイント（処理自体は tmp ファイルに依存しない）
HANDOFF_PATH = "tmp/zoom_handoff.json"
VTT_PATH     = "tmp/transcript.vtt"
MINUTES_PATH = "tmp/minutes.json"


# ===== 共通ユーティリティ =====

def _page_title(page: dict) -> str:
    """ページの title 型プロパティを名前によらず探して平文で返す。"""
    props = page.get("properties") or {}
    for _, v in props.items():
        if (v or {}).get("type") == "title":
            parts = v.get("title") or []
            return "".join((p.get("plain_text") or "") for p in parts)
    return ""


async def _retry(coro_factory, *, label: str = "api", attempts: int = 3):
    last = None
    for i in range(attempts):
        try:
            return await coro_factory()
        except Exception as e:
            last = e
            # ジッター付き指数バックオフ（同時リトライの衝突＝レート制限の再発を防ぐ）
            wait = 2 ** i + random.uniform(0, 1)
            print(f"{label} retry {i+1}/{attempts} after {wait:.1f}s: {e}")
            await asyncio.sleep(wait)
    raise RuntimeError(f"{label} が {attempts} 回失敗: {last}")


def _split_text(text: str, limit: int = NOTION_RICH_TEXT_LIMIT) -> list[str]:
    chunks = []
    while len(text) > limit:
        cut = text.rfind("\n", 0, limit)
        if cut < limit // 2:
            cut = limit
        chunks.append(text[:cut])
        text = text[cut:].lstrip("\n")
    if text:
        chunks.append(text)
    return chunks


def _rt(value: str) -> list[dict]:
    """rich_text/title 用の配列を作る。Notion の 2000字上限に合わせてチャンク分割する。"""
    s = value or ""
    if not s:
        return []
    return [{"text": {"content": s[i:i + 2000]}} for i in range(0, len(s), 2000)]


def _upsert_failed_detail(result) -> str | None:
    """UPSERT応答の中身まで見て、行の作成/更新が実際に失敗していないか調べる。

    トップレベルの successful はツール呼び出し自体の成否しか示さず、実際の
    行書き込みの成否は data.results[].ok に入る（slack_recorder_v5 の実運用で
    誤検知を確認済み）。失敗があれば詳細文字列、無ければ None を返す。
    """
    if not isinstance(result, dict):
        return "レスポンスがdictでない"
    data = result.get("data") or {}
    results = data.get("results") if isinstance(data, dict) else None
    if results is None:
        return None
    bad = [r for r in results if isinstance(r, dict) and not r.get("ok")]
    if bad:
        return "; ".join(str(r.get("error_message") or r.get("error_code") or r) for r in bad)
    return None


def _upsert_page_info(result) -> dict:
    """UPSERT応答から作成/更新されたページ情報を取り出す。

    ページの id / url はトップレベルの data ではなく data.results[0] に入る
    （2026-07-14 実走で確認。data.get("url") では常に None になる）。
    """
    if not isinstance(result, dict):
        return {}
    data = result.get("data") or {}
    results = data.get("results") if isinstance(data, dict) else None
    if isinstance(results, list) and results and isinstance(results[0], dict):
        return results[0]
    return {}


# ===== 人物辞書DB（people）自動更新 =====
# ※ PEOPLE_DB_ID はファイル冒頭のDB ID定数セクションにある

# 人物辞書DBのプロパティ名
PROP_P_NAME    = "name"     # title: 代表表示名（初回観測時に固定）
PROP_P_EMAIL   = "email"    # rich_text: 主キー（小文字正規化）
PROP_P_ALIASES = "aliases"  # rich_text: 観測した表記ゆれの改行区切り蓄積
PROP_P_COMPANY = "company"  # rich_text: メールの@より後ろの文字列そのまま（判断なしの事実。
                            # 会社名・プロジェクトへの意味づけは手動のドメイン情報DBが担う）

# 人物として登録しない機械アドレス（部分一致）
_PEOPLE_EMAIL_BLOCKLIST = (
    "noreply", "no-reply", "donotreply", "mailer-daemon", "notifications@",
)

# バッチ実行中の人物収集バッファ（書き込みは run() の最後にまとめて行う。
# 並列の保存中に都度書くと同一人物のマージが競合するため）
_people_buffer: list[dict] = []


def _valid_person_email(email: str) -> str:
    """人物辞書に登録可能なメールなら小文字正規化して返す。不可なら空文字。"""
    e = (email or "").strip().lower()
    if not e or "@" not in e:
        return ""
    if any(b in e for b in _PEOPLE_EMAIL_BLOCKLIST):
        return ""
    return e


async def _upsert_people(persons: list[dict]) -> None:
    """登場人物を人物辞書DBに upsert する（name / email / aliases / company の4列）。

    aliases は set マージで増えるだけ＝冪等（再実行しても壊れず、競合しても
    次の登場で自己修復する）。失敗は WARN のみで続行し、本体の保存処理を
    絶対に巻き込まない。PEOPLE_DB_ID 未設定なら何もしない。
    """
    if not PEOPLE_DB_ID:
        return
    # email で集約（同一実行内の重複をまとめ、書き込みを人数分に抑える）
    by_email: dict[str, set] = {}
    for p in persons or []:
        email = _valid_person_email(p.get("email", ""))
        if not email:
            continue
        names = by_email.setdefault(email, set())
        name = (p.get("name") or "").strip()
        if name and name.lower() != email:
            names.add(name)
            # 複合表示名（"担当A / Tomomi Mashiko" 等）は分割した各部分も蓄積する
            # （検索側が contains 検索の起点に使うため、短い単位の方がヒットしやすい）
            for part in re.split(r"[/／|｜]", name):
                part = part.strip()
                if part and part.lower() != email:
                    names.add(part)
    for email, names in by_email.items():
        try:
            resp = await _retry(
                lambda e=email: NOTION_QUERY_DATABASE_WITH_FILTER(
                    database_id=PEOPLE_DB_ID,
                    filter={"property": PROP_P_EMAIL, "rich_text": {"equals": e}},
                    page_size=1,
                ),
                label="people_query",
            )
            data = resp.get("data", resp) if isinstance(resp, dict) else {}
            rows = data.get("results") or []
            if rows:
                props = rows[0].get("properties") or {}
                cur_name = "".join(
                    t.get("plain_text", "")
                    for t in ((props.get(PROP_P_NAME) or {}).get("title") or [])
                )
                cur_aliases = "".join(
                    t.get("plain_text", "")
                    for t in ((props.get(PROP_P_ALIASES) or {}).get("rich_text") or [])
                )
                alias_set = {a.strip() for a in cur_aliases.split("\n") if a.strip()} | names
                display = cur_name or (sorted(names)[0] if names else email)
            else:
                alias_set = set(names)
                display = sorted(names)[0] if names else email
            props_out = {
                PROP_P_NAME:    {"title":     _rt(display)},
                PROP_P_EMAIL:   {"rich_text": _rt(email)},
                PROP_P_ALIASES: {"rich_text": _rt("\n".join(sorted(alias_set)))},
                PROP_P_COMPANY: {"rich_text": _rt(email.split("@", 1)[1])},
            }
            result = await _retry(
                lambda p=props_out, e=email: NOTION_UPSERT_ROW_DATABASE(
                    database_id=PEOPLE_DB_ID,
                    items=[{
                        "match":  {"property": PROP_P_EMAIL, "equals": e},
                        "create": {"properties": p},
                        "update": {"properties": p},
                    }],
                ),
                label="people_upsert",
            )
            err = _upsert_failed_detail(result)
            if err:
                print(f"[WARN] 人物辞書の更新失敗: {email} / {err}")
        except Exception as e:
            print(f"[WARN] 人物辞書の更新失敗: {email} / {str(e)[:120]}")


def _clean_participant_name(name: str) -> str:
    """Zoom表示名から敬称・括弧内の所属・区切り以降のサフィックスを除いた照合用の名前を返す。

    例: "Ayano Ito さん" → "Ayano Ito" / "田中太郎(営業部)" → "田中太郎" /
        "Taro Tanaka - ABC商事" → "Taro Tanaka"
    """
    n = (name or "").strip()
    n = re.sub(r"[（(\[【].*?[）)\]】]", "", n)
    n = re.split(r"\s+[-–—|｜/／]\s*", n)[0]
    n = re.sub(r"\s*(さん|様|さま|君|くん|ちゃん|先生|氏)$", "", n)
    return n.strip()


def _name_tokens(cleaned: str) -> list[str]:
    """照合用トークン（姓・名など）に分割する。

    空白（全半角）区切りに加え、"AyanoIto" のような camelCase 境界でも分割する。
    1文字トークンは誤ヒットの元なので除外する。
    """
    s = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", cleaned)
    return [t for t in re.split(r"[\s\u3000]+", s) if len(t) >= 2]


async def _resolve_emails_from_people(names: list[str]) -> list[dict]:
    """メールが取れなかった参加者名を人物辞書DBで逆引きする。

    実際のZoom表示名は「Ayano Ito さん」「田中太郎(営業部)」「Ito Ayano」のような
    揺れが多いため、2段階で照合する:
      段階1: 敬称・括弧・サフィックスを除去した名前で contains 照合
      段階2: 0件ならトークン分割（空白/camelCase境界）の AND 照合
             （姓名逆順・スペース無し表記に対応）
    どの段階でも一致が**ちょうど1件**のときだけ採用する
    （複数ヒット＝同名別人の可能性があるため推測しない）。
    戻り値: [{"name": 元の名前, "email": 解決したメール}]
    """
    if not PEOPLE_DB_ID or not names:
        return []

    async def _query_people(filter_obj: dict) -> list[dict]:
        resp = await _retry(
            lambda f=filter_obj: NOTION_QUERY_DATABASE_WITH_FILTER(
                database_id=PEOPLE_DB_ID, filter=f, page_size=2,
            ),
            label="people_reverse",
        )
        data = resp.get("data", resp) if isinstance(resp, dict) else {}
        return data.get("results") or []

    def _contains_group(q: str) -> dict:
        return {"or": [
            {"property": PROP_P_NAME, "title": {"contains": q}},
            {"property": PROP_P_ALIASES, "rich_text": {"contains": q}},
        ]}

    def _email_of(row: dict) -> str:
        props = row.get("properties") or {}
        return "".join(
            t.get("plain_text", "")
            for t in ((props.get(PROP_P_EMAIL) or {}).get("rich_text") or [])
        ).strip()

    resolved: list[dict] = []
    for name in names:
        n = (name or "").strip()
        cleaned = _clean_participant_name(n)
        if not cleaned:
            continue
        try:
            # 段階1: クリーニング済みの名前そのままで照合
            rows = await _query_people(_contains_group(cleaned))
            if len(rows) > 1:
                print(f"  ⚠ 人物辞書で複数ヒットのためメール解決を見送り: {n}")
                continue
            if not rows:
                # 段階2: トークンAND照合（姓名逆順・スペース無し・camelCase対応）
                tokens = _name_tokens(cleaned)
                if len(tokens) >= 2:
                    rows = await _query_people({"and": [_contains_group(t) for t in tokens]})
                if len(rows) > 1:
                    print(f"  ⚠ 人物辞書で複数ヒットのためメール解決を見送り: {n}")
                    continue
            if not rows:
                # 未登録＝辞書がまだ育っていないだけ。後日この人がメール付きで
                # 観測されれば辞書に入り、以降の会議では自動解決される
                print(f"  人物辞書に未登録のためメール無しで記録: {n}")
                continue
            email = _email_of(rows[0])
            if email:
                resolved.append({"name": n, "email": email})
                print(f"  人物辞書からメール解決: {n} → {email}")
        except Exception as e:
            print(f"  ⚠ 人物辞書の逆引き失敗: {n} / {str(e)[:80]}")
    return resolved


def _to_jst(utc_str: str) -> str:
    """JST の ISO 8601 形式（例: 2026-07-09T13:53:01+09:00）で返す。

    raw DB の timestamp（date型）にそのまま入れられる形式。
    Gmail/Slack スクリプトと表記を統一する。
    """
    if not utc_str:
        return ""
    try:
        dt = datetime.fromisoformat(utc_str.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(JST).isoformat(timespec="seconds")
    except Exception as e:
        print(f"⚠ 時刻変換失敗 ({utc_str}): {e}")
        return utc_str


def _normalize(s: str) -> str:
    """大文字小文字・空白・括弧内テキストを正規化する（プロジェクト名の比較用）"""
    s = (s or "").strip()
    s = re.sub(r"[（(\[].*?[）)\]]", "", s)
    s = re.sub(r"\s+", "", s)
    return s.lower()


def _encode_uuid_double(uuid: str) -> str:
    """Zoom UUID は '/' や '==' を含む場合ダブル URL エンコードが必要（録画取得API用）"""
    if uuid.startswith("/") or "//" in uuid or "+" in uuid or "=" in uuid:
        return quote(quote(uuid, safe=""), safe="")
    return uuid


def _encode_uuid_single(uuid: str) -> str:
    """参加者取得APIはシングル URL エンコード"""
    if uuid.startswith("/") or "//" in uuid or "+" in uuid or "=" in uuid:
        return quote(uuid, safe="")
    return uuid


# ===== Step1: 録画メタ情報の取得 =====

async def _fetch_recording_meta(trigger_payload: str | dict) -> dict:
    """トリガー payload から録画メタ情報（VTTのURL・token等）を取得する。

    UUID で取得 → 失敗なら数値 ID にフォールバック
    （実測 2026-07-12: UUID 検索が 404 になるケースがあり、数値 ID フォールバックが必須）。
    """
    payload = json.loads(trigger_payload) if isinstance(trigger_payload, str) else trigger_payload
    meeting = payload["meeting"]
    raw_uuid = meeting["uuid"]
    meeting_id_num = str(meeting.get("id", ""))

    response = await ZOOM_GET_MEETING_RECORDINGS(
        meetingId=_encode_uuid_double(raw_uuid),
        include_fields="download_access_token",
    )
    raw = json.loads(response) if isinstance(response, str) else response
    http_status = raw.get("mercury_last_http_status_code")
    data = raw.get("data") or {}

    if (http_status and http_status >= 400) or not data.get("recording_files"):
        print(f"UUID 検索で取得失敗 (http={http_status})、数値 Meeting ID で再試行...")
        response = await ZOOM_GET_MEETING_RECORDINGS(
            meetingId=meeting_id_num,
            include_fields="download_access_token",
        )
        raw = json.loads(response) if isinstance(response, str) else response
        http_status = raw.get("mercury_last_http_status_code")
        data = raw.get("data") or {}

    if http_status and http_status >= 400:
        raise RuntimeError(
            f"Zoom API HTTP {http_status} (uuid={raw_uuid}, id={meeting_id_num})"
        )
    if not raw.get("successful"):
        raise RuntimeError(f"Zoom API 失敗: {raw.get('error')}")

    access_token = data.get("download_access_token")
    recording_files = data.get("recording_files", [])

    transcript_file = next(
        (f for f in recording_files
         if f.get("file_type") == "TRANSCRIPT" or f.get("recording_type") == "audio_transcript"),
        None,
    )
    if not transcript_file:
        statuses = [(f.get("file_type"), f.get("status")) for f in recording_files]
        raise RuntimeError(f"TRANSCRIPT がまだ生成されていません: {statuses}")
    if not access_token:
        raise RuntimeError("download_access_token が取得できませんでした")

    handoff = {
        "download_url": transcript_file["download_url"],
        "access_token": access_token,
        "meeting_topic": data.get("topic", ""),
        "meeting_start_time": data.get("start_time", ""),
        "meeting_uuid": data.get("uuid", "") or raw_uuid,
        "meeting_id": meeting_id_num,
        "host_email": data.get("host_email", "") or "",
    }
    with open(HANDOFF_PATH, "w", encoding="utf-8") as f:
        json.dump(handoff, f, ensure_ascii=False, indent=2)
    print(f"Step1 OK: {handoff['meeting_topic']} / host={handoff['host_email']}")
    return handoff


# ===== Step2: VTTダウンロード（実行環境内で直接） =====

def _download_sync(url: str, token: str) -> bytes:
    """Bearer 認証付きで URL の内容を取得する（同期。to_thread 経由で呼ぶ）"""
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    with urllib.request.urlopen(req, timeout=DOWNLOAD_TIMEOUT) as resp:
        return resp.read()


async def _download_vtt(url: str, token: str, *, attempts: int = 3) -> bytes:
    """VTT をダウンロードして内容を検証する。失敗時は指数バックオフでリトライ"""
    last = None
    for i in range(attempts):
        try:
            content = await asyncio.to_thread(_download_sync, url, token)
            head = content[:200].decode("utf-8", errors="replace")
            if len(content) < 100 or "WEBVTT" not in head:
                raise RuntimeError(f"VTT 内容が不正: size={len(content)} 先頭={head[:80]!r}")
            return content
        except Exception as e:
            last = e
            wait = 2 ** i
            print(f"  ダウンロード retry {i+1}/{attempts} after {wait}s: {e}")
            await asyncio.sleep(wait)
    raise RuntimeError(f"VTT ダウンロードが {attempts} 回失敗: {last}")


# ===== 参加者収集（Zoom API + VTT発言者、名前のみ） =====

async def _fetch_zoom_participants(meeting_uuid: str) -> list[dict]:
    """Zoom API から参加者の {name, email} 一覧を取得する。

    メールアドレスは取れる場合（同一アカウント/サインイン参加者等）のみ入る。
    """
    if not meeting_uuid:
        return []
    participants: list[dict] = []
    token = ""
    while True:
        kwargs = {"meetingId": _encode_uuid_single(meeting_uuid), "page_size": 300}
        if token:
            kwargs["next_page_token"] = token
        try:
            resp = await ZOOM_GET_PAST_MEETING_PARTICIPANTS(**kwargs)
        except Exception as e:
            print(f"⚠ Zoom 参加者取得失敗: {e}")
            break
        if isinstance(resp, dict) and resp.get("successful") is False:
            print(f"⚠ Zoom 参加者取得エラー: {resp.get('error')}")
            break
        data = resp.get("data", resp) if isinstance(resp, dict) else {}
        for user in data.get("participants", []):
            name = (user.get("name") or "").strip()
            email = (user.get("user_email") or user.get("email") or "").strip()
            if name or email:
                participants.append({"name": name, "email": email})
        token = data.get("next_page_token") or ""
        if not token:
            break
    return participants


def _extract_speakers_from_vtt(vtt: str) -> list[str]:
    """
    VTT テキストの各行から "発言者名: テキスト" 形式の発言者名を抽出する。
    タイムスタンプ・ヘッダ・シーケンス番号行はスキップする。
    """
    speakers: set[str] = set()
    for line in vtt.splitlines():
        s = line.strip()
        if not s or s == "WEBVTT":
            continue
        if re.match(r'^\d+$', s) or re.match(r'^\d{2}:\d{2}', s):
            continue
        m = re.match(r'^(.{1,40})[:：]\s+\S', s)
        if m:
            name = m.group(1).strip()
            if not re.match(r'^\d{1,2}:\d{2}', name):
                speakers.add(name)
    return list(speakers)


# Zoom API が返す実在しないダミー/テスト用の参加者名（正規化キーで比較して除外。
# 2026-07-14 実走で "aeiou" の混入を確認）
_PARTICIPANT_BLOCKLIST = {"aeiou"}


def _merge_participants(zoom_participants: list[dict],
                        vtt_names: list[str]) -> tuple[list[str], list[str], list[str]]:
    """(参加者名リスト, メールアドレスリスト, メール未解決の名前リスト) を返す。

    - 名前: Zoom API 参加者 + VTT 発言者を正規化キー（大文字小文字・空白無視）で
      重複排除（name プロパティ用）
    - メール: 取得できた参加者分を重複排除して集める（sender プロパティ用）
    - メール未解決: 名前はあるがメールが取れなかった参加者
      （呼び出し元が人物辞書DBで逆引き補完する）
    Zoom API 側の取得は uuid だけで呼べるため、呼び出し元で他の処理と
    並行実行してからここでマージする。
    """
    entries = [((p.get("name") or "").strip(), (p.get("email") or "").strip())
               for p in zoom_participants]
    entries += [(n.strip(), "") for n in vtt_names]

    # メールが取れている名前の正規化キー（未解決判定に使う）
    keys_with_email = {_normalize(n) for n, e in entries if e and n}

    seen_norm: set[str] = set()
    seen_email: set[str] = set()
    names: list[str] = []
    emails: list[str] = []
    for name, email in entries:
        key = _normalize(name or email)
        if not key or key in _PARTICIPANT_BLOCKLIST:
            continue
        if email and email.lower() not in seen_email:
            seen_email.add(email.lower())
            emails.append(email)
        if key not in seen_norm:
            seen_norm.add(key)
            names.append(name or email)
    unresolved = [n for n in names if "@" not in n and _normalize(n) not in keys_with_email]
    return names, emails, unresolved


# ===== 議事録生成 =====

def _preprocess_vtt(vtt: str) -> str:
    """タイムスタンプ・ヘッダ・シーケンス番号を除去し、重複行を削除する"""
    lines = []
    for line in vtt.splitlines():
        s = line.strip()
        if not s or s == "WEBVTT":
            continue
        if re.match(r'^\d+$', s):
            continue
        if re.match(r'^\d{2}:\d{2}:\d{2}', s):
            continue
        lines.append(s)

    deduped = []
    for i, line in enumerate(lines):
        text = re.sub(r'^[^:：]+[:：]\s*', '', line)
        next_line = lines[i + 1] if i + 1 < len(lines) else ""
        next_text = re.sub(r'^[^:：]+[:：]\s*', '', next_line)
        if text and next_text.startswith(text):
            continue
        deduped.append(line)

    result = "\n".join(deduped)
    if len(result) > _MAX_TRANSCRIPT_CHARS:
        # 3時間級の会議などで LLM のコンテキスト/コストが爆発しないよう上限を設ける
        print(f"⚠ 文字起こしが {len(result)} 字と長いため {_MAX_TRANSCRIPT_CHARS} 字に切り詰め")
        result = result[:_MAX_TRANSCRIPT_CHARS] + "\n（以降省略）"
    return result


def _build_prompt(topic: str, meeting_at: str, transcript: str) -> str:
    return f"""以下のZoom会議の文字起こしを元に、詳細な議事録を作成してください。

会議タイトル: {topic}
開始日時: {meeting_at}

【出力フォーマット（この順序・見出しで必ず出力）】

## 会議概要（テーマ）
（会議の目的・背景・主な議題を2〜3文で簡潔に）

## 主な議論内容
- 議論されたトピックを箇条書きで

## 決定事項
- 会議で合意・決定された内容を箇条書きで

## 課題・懸念点
- 未解決の問題、懸念、リスクを箇条書きで

## 次のアクション
- 担当者: タスク内容（期限があれば記載）

## 参加者の主要発言
- 発言者名: 印象的・重要な発言を要約

【重要】
- 前置き・挨拶・断り書きは一切書かず、必ず「## 会議概要（テーマ）」の見出しから直接出力すること
- 文字起こしから読み取れない内容は創作せず、不明な箇所は「不明」と明記すること

【文字起こし（前処理済み）】
{transcript}
"""


async def _summarize_with_retry(prompt: str, *, attempts: int = 3) -> str:
    last_err = None
    for attempt in range(attempts):
        try:
            result = await llm_call(prompt=prompt)
            text = result["text"] or ""
            # プロンプトで禁止していても前置きが付くことがあるため、
            # 最初の見出しより前を捨てる（2026-07-14 実走で混入を確認）
            head = text.find("## ")
            if head > 0:
                text = text[head:]
            if text and len(text) > 50:
                return text
            raise RuntimeError(f"LLM 応答が短すぎる: {len(text)} chars")
        except Exception as e:
            last_err = e
            wait = 2 ** attempt
            print(f"LLM retry {attempt+1}/{attempts} after {wait}s: {e}")
            await asyncio.sleep(wait)
    raise RuntimeError(f"LLM 呼び出しに {attempts} 回失敗: {last_err}")


# ===== Notion ユーティリティ =====

async def _query_all(db_id: str) -> list[dict]:
    rows, cursor = [], None
    while True:
        kwargs: dict = {"database_id": db_id, "page_size": 100}
        if cursor:
            kwargs["start_cursor"] = cursor
        # フィルタなし全件取得も WITH_FILTER で行う（filter は省略可・Gmail側で実績あり）
        r = await _retry(
            lambda k=kwargs: NOTION_QUERY_DATABASE_WITH_FILTER(**k),
            label=f"query({db_id[:8]})",
        )
        d = r.get("data", r) if isinstance(r, dict) else {}
        rows.extend(d.get("results", []))
        if not d.get("has_more"):
            break
        cursor = d.get("next_cursor")
        if not cursor:
            break
    return rows


# ===== プロジェクト判定 =====

async def _load_all_projects() -> list[dict]:
    """ドメイン情報DBから [{name, domain}] 一覧を返す。

    Gmail スクリプトと同じ取得元（タイトル=pjt_ prefix付き）にすることで、
    raw DB の project 列の値が全ソースで一致する。domain は参加者メールの
    ドメイン一致判定（静的ルール）に使う。
    """
    rows = await _query_all(DOMAIN_INFO_DB_ID)
    projects = []
    for row in rows:
        name = _page_title(row).strip()
        props = row.get("properties") or {}
        d = props.get(DOMAIN_PROP) or {}
        domain = "".join(x.get("plain_text", "") for x in (d.get("rich_text") or []))
        domain = domain.strip().lstrip("@").lower()
        if name:
            projects.append({"name": name, "domain": domain})
    return projects


def _match_by_topic(topic: str, projects: list[dict]) -> tuple[str, str] | None:
    """
    meeting topic にプロジェクト名が含まれていれば (project_name, strategy) を返す。
    完全一致 → 部分一致の順で最初に見つかったものを採用する。
    """
    norm_topic = _normalize(topic)
    if not norm_topic:
        return None

    for proj in projects:
        if not proj["name"]:
            continue
        if _normalize(proj["name"]) == norm_topic:
            return proj["name"], "topic完全一致"

    for proj in projects:
        if not proj["name"]:
            continue
        norm_proj = _normalize(proj["name"])
        if norm_proj and norm_proj in norm_topic:
            return proj["name"], f"topic部分一致 ({proj['name']})"

    return None


def _match_by_participant_domain(participant_emails: list[str],
                                 projects: list[dict]) -> tuple[str, str] | None:
    """参加者メールのドメインをドメイン情報DBと突き合わせる静的ルール。

    社外参加者のドメインが登録済みならそのプロジェクトに確定する。
    LLM推定より確実なため topic 一致の次・LLM の前に評価する
    （自社ドメインはDBに登録されていないため誤マッチしない）。
    """
    seen: list[str] = []
    for e in participant_emails:
        if "@" in e:
            d = e.split("@", 1)[1].lower()
            if d not in seen:
                seen.append(d)
    for d in seen:
        for proj in projects:
            if proj.get("domain") and proj["domain"] == d:
                return proj["name"], f"参加者ドメイン一致 ({d})"
    return None


async def _match_by_content(
    topic: str, transcript: str, projects: list[dict]
) -> tuple[str, str] | None:
    """
    meeting topic + VTT テキスト（抜粋）を LLM に渡してプロジェクトを推定する。
    信頼度が medium 以上の場合のみ結果を採用する。
    """
    valid_projects = [p for p in projects if p["name"]]
    if not valid_projects:
        return None

    project_list_text = "\n".join(f"- {p['name']}" for p in valid_projects)

    prompt = f"""以下の会議情報から、最も関連するプロジェクトを選んでください。

会議タイトル: {topic}

プロジェクト一覧:
{project_list_text}

【文字起こし（抜粋）】
{transcript[:3000]}

一致するプロジェクト名をそのままの表記で返してください。
判断が難しい場合や一致するものがない場合は project_name を空文字列にしてください。
"""
    schema = {
        "type": "object",
        "properties": {
            "project_name": {
                "type": "string",
                "description": "最も一致するプロジェクト名。なければ空文字列。",
            },
            "confidence": {
                "type": "string",
                "enum": ["high", "medium", "low"],
                "description": "判断の信頼度。不明・複数候補あり・判断根拠薄の場合は low。",
            },
            "reason": {"type": "string", "description": "判断根拠（短く）"},
        },
        "required": ["project_name", "confidence"],
    }

    try:
        result = await llm_call(prompt=prompt, schema=schema)
        data = result["data"]
        matched_name = (data.get("project_name") or "").strip()
        confidence = data.get("confidence", "low")
        reason = data.get("reason", "")
        print(f"内容推定: '{matched_name}' 信頼度={confidence} 理由={reason}")

        if not matched_name or confidence == "low":
            print("⚠ 内容推定を棄却 (信頼度が low またはプロジェクト名なし)")
            return None

        for proj in valid_projects:
            if proj["name"] == matched_name:
                return proj["name"], f"内容推定 (信頼度:{confidence})"

        norm_matched = _normalize(matched_name)
        for proj in valid_projects:
            if _normalize(proj["name"]) == norm_matched:
                return proj["name"], f"内容推定 (信頼度:{confidence})"

    except Exception as e:
        print(f"⚠ LLM プロジェクト推定失敗: {e}")

    return None


async def _find_project(
    topic: str, transcript: str, projects: list[dict],
    participant_emails: list[str],
) -> tuple[str, str]:
    """
    プロジェクトを次の優先順で判定し (project_name, strategy) を返す:
      1. topic 一致
      2. 参加者メールのドメイン一致（静的ルール。LLMより確実）
      3. 内容 (LLM) 推定
    どれも一致しなければ未分類で処理を続行する。
    """
    result = _match_by_topic(topic, projects)
    if result:
        return result

    result = _match_by_participant_domain(participant_emails, projects)
    if result:
        return result

    result = await _match_by_content(topic, transcript, projects)
    if result:
        return result

    print(f"⚠ プロジェクト未特定: topic='{topic}' → '未分類' で保存")
    return "未分類", "フォールバック"


# ===== raw DB 保存 =====

async def _already_exists(dedup_key: str) -> bool:
    """parents=dedup_key かつ source=meeting の行が raw DB に存在するか確認する。

    dedup_key（通常は meeting_uuid）は会議1回ごとに一意なので、Gmail の
    [mail:ID] と同じく確実な重複判定キーになる。LLM要約より前に呼び、
    保存済み会議の再要約コストを避ける。
    """
    resp = await _retry(
        lambda: NOTION_QUERY_DATABASE_WITH_FILTER(
            database_id=RAW_DB_ID,
            filter={
                "and": [
                    {"property": PROP_PARENTS, "rich_text": {"equals": dedup_key}},
                    {"property": PROP_SOURCE, "select": {"equals": SOURCE_MEETING}},
                ]
            },
            page_size=1,
        ),
        label="dup_check",
    )
    data = resp.get("data", resp) if isinstance(resp, dict) else {}
    return len(data.get("results", [])) > 0


def _heading_block(level: str, text: str) -> dict:
    return {"object": "block", "type": level,
            level: {"rich_text": [{"text": {"content": text}}]}}


def _build_child_blocks(minutes_text: str, vtt_raw: str) -> list[dict]:
    """ページ本文のブロック（議事録＋生VTT）を Notion 生API形式で構築する。

    content プロパティにも議事録を入れるが、本文にも常に両方を残す
    （VTT の行き先はページ本文＝2026-07-14 ユーザー確定）。
    """
    blocks = [_heading_block("heading_2", "議事録")]
    for chunk in _split_text(minutes_text):
        blocks.append({"object": "block", "type": "paragraph",
                       "paragraph": {"rich_text": [{"text": {"content": chunk}}]}})
    blocks.append({"object": "block", "type": "divider", "divider": {}})
    blocks.append(_heading_block("heading_3", "生の文字起こし (VTT)"))
    for chunk in _split_text(vtt_raw):
        blocks.append({"object": "block", "type": "paragraph",
                       "paragraph": {"rich_text": [{"text": {"content": chunk}}]}})
    return blocks


# ===== エントリポイント =====

async def run(trigger_payload: str | dict) -> dict:
    """録画情報取得 → VTTダウンロード → 議事録生成 → raw DB 保存 を1本で実行する。

    Returns:
        {"status": "inserted" | "skipped_duplicate",
         "project_name": str,
         "page_url": str | None}
    """
    os.makedirs("tmp", exist_ok=True)

    payload = json.loads(trigger_payload) if isinstance(trigger_payload, str) else trigger_payload
    raw_uuid = (payload.get("meeting") or {}).get("uuid", "")

    # ---- Step1: 録画メタ情報の取得 ----
    # 参加者取得は payload の uuid だけで呼べるため、メタ取得と並行して開始する
    participants_task = asyncio.create_task(_fetch_zoom_participants(raw_uuid))
    handoff = await _fetch_recording_meta(payload)

    # ---- Step2: VTTダウンロード（token が新鮮なうちに実行） ----
    content = await _download_vtt(handoff["download_url"], handoff["access_token"])
    with open(VTT_PATH, "wb") as f:
        f.write(content)
    vtt = content.decode("utf-8", errors="replace")
    print(f"Step2 OK: VTT {len(content)} bytes")

    # ---- 重複チェック（LLM要約より前に行い、保存済み会議の再処理コストを避ける） ----
    dedup_key = handoff.get("meeting_uuid", "") or (
        f"zoom:{handoff.get('meeting_start_time', '')}:{handoff.get('meeting_topic', '')}"
    )
    if await _already_exists(dedup_key):
        participants_task.cancel()  # 以降使わないため
        print(f"⚠ 重複スキップ（保存済み）: {handoff.get('meeting_topic', '')}")
        return {"status": "skipped_duplicate", "project_name": None, "page_url": None}

    # ---- Step3: 議事録生成 → raw DB 保存 ----
    projects = await _load_all_projects()
    print(f"プロジェクト一覧: {len(projects)} 件")

    topic = handoff.get("meeting_topic") or "（タイトル未設定）"
    meeting_at = _to_jst(handoff.get("meeting_start_time", ""))

    transcript = _preprocess_vtt(vtt)
    if len(transcript) < _MIN_TRANSCRIPT_CHARS:
        # VTTはあるが中身がほぼ無い（Zoom側の文字起こし未完了 or 無発話）。
        # 議事録を創作させないためスキップ。後で再実行すれば拾える
        participants_task.cancel()
        print(f"⚠ 文字起こしが短すぎるためスキップ（{len(transcript)}字）: {topic}")
        return {"status": "skipped_short_vtt", "project_name": None, "page_url": None}
    minutes_text = await _summarize_with_retry(_build_prompt(topic, meeting_at, transcript))
    print(f"議事録生成: {len(minutes_text)} chars")

    # Step1 と並行で取得しておいた Zoom 参加者に VTT 発言者をマージ
    zoom_participants = await participants_task
    participants, participant_emails, unresolved_names = _merge_participants(
        zoom_participants, _extract_speakers_from_vtt(vtt)
    )
    # メールが取れなかった参加者は人物辞書で逆引きして補完する（一意ヒット時のみ）
    resolved_people = await _resolve_emails_from_people(unresolved_names)
    for r in resolved_people:
        if r["email"].lower() not in {e.lower() for e in participant_emails}:
            participant_emails.append(r["email"])
    # ホストのメールは参加者APIから取れない場合に備えて必ず含める
    host_email = (handoff.get("host_email") or "").strip()
    if host_email and host_email.lower() not in {e.lower() for e in participant_emails}:
        participant_emails.insert(0, host_email)
    print(f"参加者: {len(participants)} 名: {participants} / メール: {participant_emails}")
    # name=参加者名 / sender=参加者メール（どちらもカンマ結合）
    attendees = ", ".join(participants)
    sender_emails = ", ".join(participant_emails)

    # チェックポイント保存（デバッグ用）
    with open(MINUTES_PATH, "w", encoding="utf-8") as f:
        json.dump({
            "meeting_title": topic,
            "meeting_at": meeting_at,
            "meeting_uuid": handoff.get("meeting_uuid", ""),
            "host_email": handoff.get("host_email", ""),
            "participants": participants,
            "participant_emails": participant_emails,
            "minutes_text": minutes_text,
        }, f, ensure_ascii=False, indent=2)

    project_name, strategy = await _find_project(
        topic, transcript, projects, participant_emails
    )
    print(f"判定: {project_name} ({strategy}) / 保存先: raw DB = {RAW_DB_ID}")

    child_blocks = _build_child_blocks(minutes_text, vtt)
    print(f"child_blocks: {len(child_blocks)} 個")

    # raw DB の統一プロパティ（Notion 生API形式・2000字チャンク分割で全文を切らず保存）
    props = {
        PROP_CONTENT:   {"title":     _rt(minutes_text)},
        PROP_SUBJECT:   {"rich_text": _rt(topic)},
        PROP_SENDER:    {"rich_text": _rt(sender_emails)},
        PROP_NAME:      {"rich_text": _rt(attendees)},
        PROP_PROJECT:   {"rich_text": _rt(project_name)},
        PROP_SOURCE:    {"select":    {"name": SOURCE_MEETING}},
        PROP_PARENTS:   {"rich_text": _rt(dedup_key)},
    }
    if meeting_at:
        # 空文字を date に入れるとエラーになるため、日時がある場合のみ付与する
        props[PROP_TIMESTAMP] = {"date": {"start": meeting_at}}

    # 書き込みは Gmail/Slack と同じ NOTION_UPSERT_ROW_DATABASE に統一。
    # match=parents（dedup_key）なので万一の再実行も既存行の上書きで済む。
    # ページ本文（議事録＋生VTT）は新規作成時のみ children で渡す
    # （update にも渡すと再実行のたびにブロックが積み増しされるため）
    result = await _retry(
        lambda: NOTION_UPSERT_ROW_DATABASE(
            database_id=RAW_DB_ID,
            items=[{
                "match":  {"property": PROP_PARENTS, "equals": dedup_key},
                "create": {"properties": props, "children": child_blocks},
                "update": {"properties": props},
            }],
        ),
        label="upsert",
    )
    err = _upsert_failed_detail(result)
    if err:
        raise RuntimeError(f"Notion書き込み失敗: {err}")

    # 人物辞書の自動更新（参加者＋ホスト。失敗しても保存結果には影響しない）
    await _upsert_people(
        list(zoom_participants) + resolved_people
        + [{"name": "", "email": handoff.get("host_email", "")}]
    )

    info = _upsert_page_info(result)
    # url が応答に無い場合はページIDから組み立てる（NotionのURLはIDだけで開ける）
    url = info.get("url") or (
        f"https://www.notion.so/{str(info['id']).replace('-', '')}" if info.get("id") else None
    )
    print(f"\nupsert OK → {url}")
    return {"status": "inserted", "project_name": project_name, "page_url": url}
