# -*- coding: utf-8 -*-
"""서버 저장소 — 이월명부와 마감이력.

이월명부는 '월마다 새 파일'이 아니라 하나의 살아 있는 명부다. 마감을 돌릴 때마다
배송 상태를 갱신하고, 새로 생긴 이월 건을 덧붙인다.

실제로 어디에 담기는지(로컬 data/ 폴더냐 Supabase 냐)는 backend 가 정한다.
여기서는 문서 이름만 안다.
"""
from __future__ import annotations

import io
import json
from datetime import datetime

import pandas as pd

from . import backend, engine

CARRY = "이월명부.csv"
HIST = "마감이력.json"


def load_carryover() -> pd.DataFrame:
    t = backend.read_text(CARRY)
    if not t:
        return engine.empty_carryover()
    df = pd.read_csv(io.StringIO(t), dtype=str, keep_default_na=False)
    for c in engine.CARRY_COLS:
        if c not in df.columns:
            df[c] = ""
    return df[engine.CARRY_COLS]


def save_carryover(df: pd.DataFrame) -> None:
    backend.write_text(CARRY, df.to_csv(index=False))


def load_history() -> list[dict]:
    t = backend.read_text(HIST)
    return json.loads(t) if t else []


def save_history(rows: list[dict]) -> None:
    backend.write_text(HIST, json.dumps(rows, ensure_ascii=False, indent=2))


def record_close(마감월: str, share: pd.DataFrame, 메모: str = "",
                 적용수당표: str = "", 입력요약: dict | None = None) -> list[dict]:
    """마감 확정 1건을 이력에 남긴다. 같은 달이 있으면 덮어쓰되 회차를 올린다."""
    rows = load_history()
    회차 = 1 + sum(1 for r in rows if r["마감월"] == 마감월)
    rows = [r for r in rows if r["마감월"] != 마감월]
    rows.append(dict(
        마감월=마감월, 회차=회차,
        확정시각=datetime.now().strftime("%Y-%m-%d %H:%M"),
        정상가합계=float(share["정상가합계"].sum()),
        건수=int(share["건수"].sum()),
        인센티브총액=float(share["인센티브총액"].sum()),
        당월지급=float(share["당월지급"].sum()),
        이월건수=int(share["이월건수"].sum()),
        이월금액=float(share["이월금액"].sum()),
        적용수당표=적용수당표, 입력요약=입력요약 or {}, 메모=메모,
        지점별=[{k: (float(v) if isinstance(v, (int, float)) else str(v))
               for k, v in r.items()} for r in share.to_dict("records")],
    ))
    rows.sort(key=lambda r: r["마감월"])
    save_history(rows)
    return rows


def save_snapshot(마감월: str, data: bytes, 이름: str) -> str:
    return backend.write_snapshot(f"{마감월}_{이름}", data)


def snapshots() -> list[str]:
    return backend.list_snapshots()


def snapshot_bytes(이름: str) -> bytes:
    return backend.read_snapshot(이름)
