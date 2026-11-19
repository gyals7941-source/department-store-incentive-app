# -*- coding: utf-8 -*-
"""결과를 기존 엑셀 포맷 그대로 내보낸다 (영업팀공유 / 대상추출 / 이월명부)."""
from __future__ import annotations

import io

import pandas as pd

from . import engine

SHARE_HDR = ["거래처명", "마감 합계금액(정상가)", "달성차수", "대상 건수",
             "판매월귀속 인센티브", "당월 실지급 인센티브", "익월이월 건수",
             "익월이월(참고)"]

EX_HDR = ["거래처명", "고유키", "주문번호", "매칭키(적요)", "주문/수취인", "품목코드",
          "품목명", "단가", "합계", "수수료율", "수수료제외 발행액", "상태판정",
          "배송완료(자동)", "달성차수", "인센티브단가", "지급구분", "정상가(차수기준)",
          "판매채널", "_수량", "_순수량(상계)"]

EX_SRC = ["지점", "건키", "주문서일자", "정규화키", "주문자", "품목코드", "품목명",
          "단가", "합계", "수수료율", "순매출", "판정", "배송", "차수", "인센단가",
          "지급구분", "정상가", "채널", "수량", "상계합"]

CARRY_HDR = ["판매월", "거래처명", "고유키", "매칭키(적요)", "품목코드", "품목명",
             "달성차수", "확정 인센티브단가", "지급상태", "실제배송완료(자동)",
             "비고", "_원장취소(자동)", "지급월"]


def build(res: dict, 마감월_text: str, carryover: pd.DataFrame | None = None,
          제목: str = "백화점 오프라인 마감 — 거래처별 인센티브") -> bytes:
    share, ex = res["영업팀공유"], res["대상추출"]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
        bk = xw.book
        money = bk.add_format({"num_format": "#,##0"})
        bold = bk.add_format({"bold": True, "bg_color": "#E8EDE4", "border": 1,
                              "align": "center", "valign": "vcenter", "text_wrap": True})
        title = bk.add_format({"bold": True, "font_size": 14})
        tot = bk.add_format({"bold": True, "num_format": "#,##0", "top": 1})

        # ── 영업팀공유 ──
        ws = bk.add_worksheet("영업팀공유")
        xw.sheets["영업팀공유"] = ws
        y, m = 마감월_text.split("-")
        ws.write(0, 0, 제목, title)
        ws.write(1, 0, f"마감월: {int(y)}년 {int(m)}월")
        for c, h in enumerate(SHARE_HDR):
            ws.write(3, c, h, bold)
        r = 4
        for _, x in share.iterrows():
            ws.write(r, 0, x["지점"])
            ws.write_number(r, 1, float(x["정상가합계"]), money)
            ws.write(r, 2, f"{x['차수']}차" if str(x["차수"]).isdigit() else str(x["차수"]))
            ws.write_number(r, 3, int(x["건수"]))
            for i, k in enumerate(["인센티브총액", "당월지급"]):
                ws.write_number(r, 4 + i, float(x[k]), money)
            ws.write_number(r, 6, int(x["이월건수"]))
            ws.write_number(r, 7, float(x["이월금액"]), money)
            r += 1
        ws.write(r, 0, "합계", bk.add_format({"bold": True, "top": 1}))
        for c, k in [(1, "정상가합계"), (3, "건수"), (4, "인센티브총액"),
                     (5, "당월지급"), (6, "이월건수"), (7, "이월금액")]:
            ws.write_number(r, c, float(share[k].sum()), tot)
        ws.set_column(0, 0, 24)
        ws.set_column(1, 7, 17)
        ws.freeze_panes(4, 1)

        # ── 대상추출 ──
        d = pd.DataFrame({h: (ex[s] if s in ex.columns else "")
                          for h, s in zip(EX_HDR, EX_SRC)})
        d.to_excel(xw, sheet_name="대상추출", index=False)
        w = xw.sheets["대상추출"]
        for c, h in enumerate(EX_HDR):
            w.write(0, c, h, bold)
        w.set_column(0, 0, 22); w.set_column(1, 3, 20); w.set_column(6, 6, 30)
        w.freeze_panes(1, 1)
        w.autofilter(0, 0, max(len(d), 1), len(EX_HDR) - 1)

        # ── 이월명부 ──
        co = carryover if carryover is not None else engine.empty_carryover()
        cd = pd.DataFrame({h: (co[s] if s in co.columns else "")
                           for h, s in zip(CARRY_HDR, engine.CARRY_COLS)})
        cd.to_excel(xw, sheet_name="이월명부", index=False)
        w = xw.sheets["이월명부"]
        for c, h in enumerate(CARRY_HDR):
            w.write(0, c, h, bold)
        w.set_column(0, 0, 10); w.set_column(1, 3, 22); w.set_column(5, 5, 28)
        w.freeze_panes(1, 0)
        w.autofilter(0, 0, max(len(cd), 1), len(CARRY_HDR) - 1)
    return buf.getvalue()
