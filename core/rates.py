# -*- coding: utf-8 -*-
"""적용월별 판매수당표 저장소.

수당표는 매달 바뀐다(7·8월 동일 → 9월 대폭 개정). 그래서 전역 설정 하나가 아니라
'적용월'이 붙은 이력으로 보관하고, 해당 월 표가 없으면 그 이전 가장 가까운 달을 이어받는다.
여기에 없는 품목코드는 수당 대상이 아니다(0원).
"""
from __future__ import annotations

import json

from . import backend

KEY = "수당표.json"
차수키 = {1: "1차", 2: "2차", 3: "3차"}


def load_all() -> dict:
    t = backend.read_text(KEY)
    if not t:
        raise FileNotFoundError(f"{KEY} 가 저장소에 없습니다.")
    return json.loads(t)


def months() -> list[str]:
    return sorted(load_all()["표"])


def for_month(마감월_text: str) -> tuple[dict, str]:
    """(품목표, 실제로 적용된 월). 해당 월이 없으면 직전 월을 상속한다."""
    표 = load_all()["표"]
    후보 = [m for m in sorted(표) if m <= 마감월_text]
    if not 후보:
        raise KeyError(f"{마감월_text} 이전에 등록된 수당표가 없습니다. "
                       f"등록된 월: {sorted(표) or '없음'}")
    적용월 = 후보[-1]
    return 표[적용월]["품목"], 적용월


def save_month(마감월_text: str, 품목: dict, 출처: str = "", 비고: str = "") -> None:
    d = load_all()
    entry = {"품목": 품목}
    if 출처:
        entry["_출처"] = 출처
    if 비고:
        entry["_비고"] = 비고
    d["표"][마감월_text] = entry
    backend.write_text(KEY, json.dumps(d, ensure_ascii=False, indent=2))


def copy_from_previous(새월: str) -> dict:
    """새 달을 시작할 때 직전 달 표를 그대로 복사해온다(저장은 하지 않음)."""
    품목, _ = for_month(새월)
    return json.loads(json.dumps(품목, ensure_ascii=False))


def apply_to(settings) -> str:
    """Settings 의 단가표·온라인전용을 해당 마감월 수당표로 교체. 적용월을 돌려준다."""
    품목, 적용월 = for_month(settings.마감월_text)
    settings.단가표 = {
        코드: {n: float(v.get(k) or 0) for n, k in 차수키.items()}
        for 코드, v in 품목.items()
    }
    settings.온라인전용 = {
        코드 for 코드, v in 품목.items()
        if "온라인만" in str(v.get("판매채널") or "")
    }
    return 적용월


def diff(이전월: str, 현재월: str) -> list[dict]:
    """두 달 수당표의 변경점. 설정 화면에서 '이번 달 뭐가 바뀌었나'를 보여주는 용도."""
    a, _ = for_month(이전월)
    b, _ = for_month(현재월)
    out = []
    for 코드 in sorted(set(a) | set(b)):
        x, y = a.get(코드), b.get(코드)
        if x == y:
            continue
        out.append(dict(
            품목코드=코드,
            품목명=(y or x).get("품목명", ""),
            구분="신규" if x is None else "삭제" if y is None else "변경",
            이전={k: x.get(k) for k in ("1차", "2차", "3차", "판매채널")} if x else None,
            현재={k: y.get(k) for k in ("1차", "2차", "3차", "판매채널")} if y else None,
        ))
    return out
