# -*- coding: utf-8 -*-
"""지점(거래처) 마스터 — 별칭 통합과 지점별 목표금액."""
from __future__ import annotations

import json

from . import backend

KEY = "거래처.json"


def load() -> dict:
    t = backend.read_text(KEY)
    if not t:
        raise FileNotFoundError(f"{KEY} 가 저장소에 없습니다.")
    return json.loads(t)


def save(d: dict) -> None:
    backend.write_text(KEY, json.dumps(d, ensure_ascii=False, indent=2))


def active(d: dict | None = None) -> list[str]:
    d = d or load()
    return [s["이름"] for s in d["지점"] if s.get("사용", True)]


def alias_map(d: dict | None = None) -> dict[str, str]:
    """{별칭 또는 본명: 본명}. 엔진이 거래처명을 대표 이름으로 접을 때 쓴다."""
    d = d or load()
    m = {}
    for s in d["지점"]:
        m[s["이름"]] = s["이름"]
        for a in s.get("별칭") or []:
            m[str(a).strip()] = s["이름"]
    return m


def targets(d: dict | None = None) -> dict[str, tuple[float, float, float]]:
    """{지점: (1차, 2차, 3차)}. 지점 목표가 없으면 기본목표를 쓴다."""
    d = d or load()
    base = d["기본목표"]
    out = {}
    for s in d["지점"]:
        t = s.get("목표") or base
        out[s["이름"]] = (float(t["1차"]), float(t["2차"]), float(t["3차"]))
    return out


def 기본목표(d: dict | None = None) -> tuple[float, float, float]:
    b = (d or load())["기본목표"]
    return float(b["1차"]), float(b["2차"]), float(b["3차"])


def 키워드(d: dict | None = None) -> list[str]:
    """거래처명이 이 조각을 품고 있으면 지점으로 본다.

    실제 유통사 이름이라 코드가 아니라 거래처.json 에 둔다(저장소에 올라가지 않는다).
    """
    d = d or load()
    return [str(k).strip() for k in (d.get("지점키워드") or []) if str(k).strip()]


def 수수료율(d: dict | None = None) -> float:
    """판매원장에 수수료율이 안 적힌 행에 쓰는 기본값.

    회사 숫자라 코드가 아니라 거래처.json 에 둔다(저장소에 올라가지 않는다).
    없으면 조용히 0 으로 계산하지 않고 멈춘다.
    """
    d = d or load()
    v = d.get("기본수수료율")
    if not isinstance(v, (int, float)) or isinstance(v, bool) or not (0 <= v < 1):
        raise ValueError(f"{KEY} 에 기본수수료율(0~1 사이 숫자)이 없습니다.")
    return float(v)


def is_custom(s: dict) -> bool:
    return bool(s.get("목표"))
