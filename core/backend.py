# -*- coding: utf-8 -*-
"""저장소 뒤쪽 — 로컬 파일 또는 GitHub 저장소.

클라우드(Streamlit Cloud)에서는 디스크가 재시작마다 비워진다. 그래서 이월명부·
마감이력·수당표·거래처를 **코드와 별개인 비공개 저장소**에 파일로 담는다.
접속 정보가 없으면 예전처럼 ``data/`` 폴더를 쓴다 — 사내 PC와 클라우드를 같은
코드로 돌리기 위해서다.

데이터를 코드 저장소에 같이 두면 저장할 때마다 Streamlit Cloud 가 앱을
다시 배포해 버린다. 그래서 반드시 따로 둔다.

저장은 곧 커밋이라 이력이 영구히 남는다. "9월 마감 직후 명부가 어땠는지"를
GitHub 에서 그대로 펼쳐볼 수 있다. 따로 백업을 돌릴 필요가 없다.

접속 정보는 둘 중 아무 데나 있으면 된다.

    환경변수          GH_DATA_REPO / GH_DATA_TOKEN / GH_DATA_EXPIRES
    secrets.toml     [github] repo / token / expires

위쪽(storage·rates·stores)은 여기 있는 read_text/write_text 만 쓴다. 어디에
저장되는지는 몰라도 된다.
"""
from __future__ import annotations

import base64
import os
import shutil
import time
from datetime import date, datetime
from functools import lru_cache
from pathlib import Path
from urllib.parse import quote

ROOT = Path(__file__).resolve().parent.parent / "data"
SNAP_DIR = ROOT / "마감본"
SNAP_PREFIX = "마감본"          # 저장소 안에서 엑셀 사본이 들어갈 폴더
백업보관 = 20                   # 로컬에서만. GitHub 는 커밋이 곧 이력이다.

API = "https://api.github.com"
UA = "department-store-incentive-app"   # 사람 정보를 넣지 않는다
TIMEOUT = 20
캐시수명 = 30                   # 초. 같은 값을 매 화면 갱신마다 다시 받지 않는다

_cache: dict[str, tuple[float, bytes]] = {}


# ─────────────────────────────────────────────────────────────
# 접속
# ─────────────────────────────────────────────────────────────
def _secret(이름: str, 키: str) -> str:
    """환경변수를 먼저 보고, 없으면 Streamlit 금고를 본다."""
    v = os.environ.get(이름, "")
    if v:
        return v
    try:
        import streamlit as st
        return str(st.secrets.get("github", {}).get(키, ""))
    except Exception:                      # noqa: BLE001  금고가 없는 환경
        return ""


@lru_cache(maxsize=1)
def _conn() -> tuple[str, str] | None:
    repo = _secret("GH_DATA_REPO", "repo").strip()
    token = _secret("GH_DATA_TOKEN", "token").strip()
    return (repo, token) if (repo and token) else None


def 접속정보_문제() -> str | None:
    """secrets 의 repo·token 이 쓸 수 없는 모양이면 무엇이 문제인지 알려 준다.

    값 자체는 화면에 내지 않는다 — 몇 번째 글자가 문제인지만 말한다.
    """
    c = _conn()
    if not c:
        return None
    repo, token = c
    for 이름, 값 in (("repo", repo), ("token", token)):
        나쁜 = [i + 1 for i, ch in enumerate(값) if not (33 <= ord(ch) <= 126)]
        if 나쁜:
            return (f"Secrets 의 [github] {이름} 값에 영문·숫자가 아닌 글자가 있습니다 "
                    f"(전체 {len(값)}글자 중 {나쁜[0]}번째부터 {len(나쁜)}글자). "
                    "예시 문구나 한글·공백이 섞였는지 확인하고 실제 값으로 바꿔 주세요.")
    if "/" not in repo:
        return "Secrets 의 [github] repo 는 '계정이름/저장소이름' 모양이어야 합니다."
    return None


def cloud() -> bool:
    return _conn() is not None


def 어디() -> str:
    c = _conn()
    return f"GitHub · {c[0]}" if c else str(ROOT)


def token_만료일() -> date | None:
    """secrets 에 적어둔 토큰 만료일. 안 적었으면 None."""
    t = _secret("GH_DATA_EXPIRES", "expires").strip()
    try:
        return date.fromisoformat(t) if t else None
    except ValueError:
        return None


def token_남은일수() -> int | None:
    d = token_만료일()
    return (d - date.today()).days if d else None


# ─────────────────────────────────────────────────────────────
# GitHub 호출
# ─────────────────────────────────────────────────────────────
def _url(path: str) -> str:
    repo, _ = _conn()                       # type: ignore[misc]
    return f"{API}/repos/{repo}/contents/{quote(path)}"


def _head(raw: bool = False) -> dict:
    _, token = _conn()                      # type: ignore[misc]
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.raw" if raw
                  else "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": UA,
    }


def _sha(path: str) -> str | None:
    """있으면 현재 파일의 sha, 없으면 None. 덮어쓸 때 필요하다."""
    import requests
    r = requests.get(_url(path), headers=_head(), timeout=TIMEOUT)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.json().get("sha")


def _download(path: str) -> bytes | None:
    import requests
    now = time.time()
    hit = _cache.get(path)
    if hit and now - hit[0] < 캐시수명:
        return hit[1]
    r = requests.get(_url(path), headers=_head(raw=True), timeout=TIMEOUT)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    _cache[path] = (now, r.content)
    return r.content


def _upload(path: str, data: bytes, 메시지: str) -> None:
    import requests
    body = {"message": 메시지,
            "content": base64.b64encode(data).decode("ascii")}
    sha = _sha(path)
    if sha:
        body["sha"] = sha                   # 없으면 '이미 있다'며 거부당한다
    r = requests.put(_url(path), headers=_head(), json=body, timeout=TIMEOUT)
    r.raise_for_status()
    _cache.pop(path, None)                  # 방금 바꿨으니 캐시를 버린다


# ─────────────────────────────────────────────────────────────
# 문서 읽고 쓰기
# ─────────────────────────────────────────────────────────────
def read_text(key: str) -> str | None:
    """없으면 None. CSV 의 BOM 은 떼고 돌려준다(그대로 두면 첫 열 이름이 깨진다)."""
    if not cloud():
        p = ROOT / key
        return p.read_text(encoding="utf-8-sig") if p.exists() else None
    b = _download(key)
    return b.decode("utf-8-sig") if b is not None else None


def write_text(key: str, body: str) -> None:
    if not cloud():
        ROOT.mkdir(parents=True, exist_ok=True)
        p = ROOT / key
        _backup_local(p)
        # CSV 는 엑셀에서 바로 열 수 있도록 BOM 을 붙인다.
        p.write_text(body, encoding="utf-8-sig" if p.suffix == ".csv" else "utf-8")
        return
    enc = "utf-8-sig" if key.endswith(".csv") else "utf-8"
    _upload(key, body.encode(enc), f"{key} 저장 {datetime.now():%Y-%m-%d %H:%M}")


def _backup_local(p: Path) -> None:
    if not p.exists():
        return
    bak = p.with_name(f"{p.stem}.{datetime.now():%Y%m%d_%H%M%S}{p.suffix}")
    shutil.copy2(p, bak)
    for o in sorted(p.parent.glob(f"{p.stem}.*{p.suffix}"))[:-백업보관]:
        o.unlink(missing_ok=True)


# ─────────────────────────────────────────────────────────────
# 올린 파일 보관 — 다시 들어와도 같은 파일을 또 올리지 않게
# ─────────────────────────────────────────────────────────────
UP_PREFIX = "올린파일"


def _up_key(마감월: str, 종류: str) -> str:
    return f"{UP_PREFIX}/{마감월}/{종류}.xlsx"


def read_upload(마감월: str, 종류: str) -> bytes | None:
    """그 달에 마지막으로 올린 파일. 없으면 None."""
    key = _up_key(마감월, 종류)
    if not cloud():
        p = ROOT / key
        return p.read_bytes() if p.exists() else None
    return _download(key)


def write_upload(마감월: str, 종류: str, data: bytes) -> bool:
    """올린 파일을 보관한다. 이미 같은 내용이면 아무것도 하지 않고 False."""
    if read_upload(마감월, 종류) == data:
        return False
    key = _up_key(마감월, 종류)
    if not cloud():
        p = ROOT / key
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(data)
        return True
    _upload(key, data, f"{마감월} {종류} 올림 {datetime.now():%Y-%m-%d %H:%M}")
    _cache[key] = (time.time(), data)       # 방금 올린 것을 다시 받지 않는다
    return True


# ─────────────────────────────────────────────────────────────
# 확정 마감본 (엑셀 사본)
# ─────────────────────────────────────────────────────────────
def write_snapshot(이름: str, data: bytes) -> str:
    if not cloud():
        SNAP_DIR.mkdir(parents=True, exist_ok=True)
        p = SNAP_DIR / 이름
        p.write_bytes(data)
        return str(p)
    path = f"{SNAP_PREFIX}/{이름}"
    _upload(path, data, f"마감본 {이름} {datetime.now():%Y-%m-%d %H:%M}")
    return path


def list_snapshots() -> list[str]:
    if not cloud():
        return sorted(p.name for p in SNAP_DIR.glob("*.xlsx")) if SNAP_DIR.exists() else []
    import requests
    r = requests.get(_url(SNAP_PREFIX), headers=_head(), timeout=TIMEOUT)
    if r.status_code == 404:
        return []
    r.raise_for_status()
    return sorted(o["name"] for o in r.json()
                  if str(o.get("name", "")).endswith(".xlsx"))


def read_snapshot(이름: str) -> bytes:
    if not cloud():
        return (SNAP_DIR / 이름).read_bytes()
    b = _download(f"{SNAP_PREFIX}/{이름}")
    if b is None:
        raise FileNotFoundError(이름)
    return b
