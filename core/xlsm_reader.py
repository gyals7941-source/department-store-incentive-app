# -*- coding: utf-8 -*-
"""기존 인센티브 .xlsm 에서 입력·설정·기대값을 읽어온다 (검증 및 이관용)."""
from __future__ import annotations

from dataclasses import dataclass, field

import openpyxl
import pandas as pd


@dataclass
class Settings:
    마감월: object = None
    # 수수료율·목표금액의 실제 값은 코드에 두지 않는다 — 거래처.json(데이터 저장소)
    # 또는 기존 엑셀의 설정 시트에서 채운다.
    기본수수료율: float = 0.0
    기준1: float = 0.0
    기준2: float = 0.0
    기준3: float = 0.0
    # {품목코드: {1: 1차단가, 2: 2차단가, 3: 3차단가}}
    단가표: dict = field(default_factory=dict)
    # 판매채널이 '온라인만'인 품목코드 — 오프라인 판매는 수당 제외
    온라인전용: set = field(default_factory=set)
    # {이카운트 거래처명: 대표 지점명} — 같은 매장의 다른 표기를 묶는다
    별칭: dict = field(default_factory=dict)
    # {지점: (1차, 2차, 3차)} — 비어 있으면 기준1/2/3(공통 목표)을 쓴다
    목표: dict = field(default_factory=dict)
    # 거래처명 중 지점으로 볼 조각. 거래처.json 에서 온다.
    지점키워드: list = field(default_factory=list)

    @property
    def 마감월_text(self) -> str:
        d = self.마감월
        return f"{d.year:04d}-{d.month:02d}" if d is not None else ""


RAW_MAP = ["거래처명", "주문서일자", "판매일자", "회계전표일자", "비고", "주문자",
           "품목명", "수량", "단가", "합계", "정상가", "수수료율", "적요",
           "품목코드", "창고명", "프로젝트명", "거래처코드"]

DELIV_COLS = {1: "구분", 2: "날짜", 6: "설치예정일", 7: "설치확정일", 13: "주문코드"}

CARRY_COLS = {1: "판매월", 2: "지점", 3: "건키", 4: "정규화키", 5: "품목코드",
              6: "품목명", 7: "확정차수", 8: "확정단가", 9: "지급상태",
              10: "실제배송완료", 11: "비고", 13: "지급월"}


def _rows(ws, first_row, last_row, colmap, keycol):
    out = []
    for r in range(first_row, last_row + 1):
        if ws.cell(r, keycol).value in (None, ""):
            continue
        out.append({name: ws.cell(r, c).value for c, name in colmap.items()})
    return pd.DataFrame(out)


def _키워드() -> list:
    """거래처.json 이 있으면 지점 키워드를 가져온다. 없으면 빈 목록."""
    try:
        from . import stores
        return stores.키워드()
    except Exception:                      # noqa: BLE001
        return []


def load(path) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True, keep_vba=True)

    # ── 설정 ──
    s = wb["설정"]
    st = Settings(
        마감월=s.cell(3, 2).value,
        기본수수료율=float(s.cell(4, 2).value or 0),
        기준1=float(s.cell(7, 2).value or 0),
        기준2=float(s.cell(8, 2).value or 0),
        기준3=float(s.cell(9, 2).value or 0),
        지점키워드=_키워드(),
    )
    # 단가표: 13행부터 빈 행을 만날 때까지 (최종본은 '판매채널' G열이 추가됨)
    for r in range(13, 40):
        code = s.cell(r, 1).value
        if not code:
            if s.cell(r + 1, 1).value:      # 한 줄 비고 이어지는 경우 허용
                continue
            break
        code = str(code).strip()
        if not code.upper().startswith("ZP"):
            break
        st.단가표[code] = {
            1: float(s.cell(r, 4).value or 0),
            2: float(s.cell(r, 5).value or 0),
            3: float(s.cell(r, 6).value or 0),
        }
        if "온라인만" in str(s.cell(r, 7).value or ""):
            st.온라인전용.add(code)

    # ── 원본붙여넣기 ──
    ws = wb["원본붙여넣기"]
    raw = []
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 7).value in (None, ""):      # G 품목명
            continue
        raw.append({n: ws.cell(r, i + 1).value for i, n in enumerate(RAW_MAP)})
    raw = pd.DataFrame(raw)

    # ── 배송완료원본 ──
    d = wb["배송완료원본"]
    deliv = _rows(d, 2, min(d.max_row, 2001), DELIV_COLS, 13)

    # ── 지점 목록 (지점차수 A5:A30) ──
    t = wb["지점차수"]
    지점목록 = [str(t.cell(r, 1).value).strip() for r in range(5, 31)
             if t.cell(r, 1).value not in (None, "")]

    # ── 기대값: 지점차수 ──
    exp_tier = pd.DataFrame([
        dict(지점=str(t.cell(r, 1).value).strip(),
             정상가합계=float(t.cell(r, 2).value or 0),
             차수=int(t.cell(r, 3).value or 0))
        for r in range(5, 5 + len(지점목록))
    ])

    # ── 기대값: 영업팀공유 (5행~, '합계' 전까지) ──
    sh = wb["영업팀공유"]
    exp_share = []
    for r in range(5, 31):
        name = sh.cell(r, 1).value
        if name in (None, "") or str(name).strip() == "합계":
            break
        exp_share.append(dict(
            지점=str(name).strip(),
            정상가합계=float(sh.cell(r, 2).value or 0),
            차수=str(sh.cell(r, 3).value or ""),
            건수=int(sh.cell(r, 4).value or 0),
            인센티브총액=float(sh.cell(r, 5).value or 0),
            당월지급=float(sh.cell(r, 6).value or 0),
            이월건수=int(sh.cell(r, 7).value or 0),
            이월금액=float(sh.cell(r, 8).value or 0),
        ))
    exp_share = pd.DataFrame(exp_share)

    # ── 이월명부 (5행~) ──
    c = wb["이월명부"]
    carry = _rows(c, 5, min(c.max_row, 304), CARRY_COLS, 3)
    if not carry.empty:
        for col in CARRY_COLS.values():
            if col not in carry.columns:
                carry[col] = ""
        carry = carry.fillna("")

    # ── 기대값: 대상추출 (행 단위 대조용) ──
    e = wb["대상추출"]
    exp_ex = []
    for r in range(2, min(e.max_row, 601) + 1):
        if e.cell(r, 1).value in (None, ""):
            continue
        exp_ex.append(dict(
            지점=e.cell(r, 1).value, 건키=e.cell(r, 2).value,
            정규화키=e.cell(r, 4).value, 품목코드=e.cell(r, 6).value,
            순매출=e.cell(r, 11).value, 판정=e.cell(r, 12).value,
            배송=e.cell(r, 13).value, 차수=e.cell(r, 14).value,
            인센단가=e.cell(r, 15).value, 지급구분=e.cell(r, 16).value,
            정상가=e.cell(r, 17).value, 상계합=e.cell(r, 19).value,
        ))
    exp_ex = pd.DataFrame(exp_ex)

    # ── 템플릿 세대 판별 ──
    #   v1: _배송인덱스 없음(배송완료원본 AE/AF 보조열), 취소상계·익월배송 없음  (~2026-07)
    #   v2: _배송인덱스 + 취소상계 + 익월배송                                  (2026-08~)
    #   v3: v2 + 설정 '판매채널'(온라인만) + 이월명부 '지급월'                    (2026-09 최종본~)
    if "_배송인덱스" not in wb.sheetnames:
        ver = "v1"
    elif str(s.cell(12, 7).value or "").strip() == "판매채널":
        ver = "v3"
    else:
        ver = "v2"

    wb.close()
    return dict(settings=st, raw=raw, delivery=deliv, 지점목록=지점목록,
                carryover=carry, exp_tier=exp_tier, exp_share=exp_share,
                exp_extract=exp_ex, 버전=ver)
