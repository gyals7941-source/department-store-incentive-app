# -*- coding: utf-8 -*-
"""업로드 파일 읽기 — 이카운트확인(판매원장)과 배송관리 대장."""
from __future__ import annotations

import warnings

import openpyxl
import pandas as pd

warnings.filterwarnings("ignore")

# 이카운트확인 시트의 A~Q 열 (18번째 '회계반영일자'는 쓰지 않는다)
RAW_MAP = ["거래처명", "주문서일자", "판매일자", "회계전표일자", "비고", "주문자",
           "품목명", "수량", "단가", "합계", "정상가", "수수료율", "적요",
           "품목코드", "창고명", "프로젝트명", "거래처코드"]

# 배송관리 대장에서 쓰는 다섯 칸. 열 번호는 내보내기 양식에 따라 달라지므로
# 머리글 이름으로 찾는다. 아래는 머리글을 못 찾았을 때만 쓰는 옛 .xlsm 기준 자리.
DELIV_COLS = {1: "구분", 2: "날짜", 6: "설치예정일", 7: "설치확정일", 13: "주문코드"}

# 같은 칸이 양식마다 다르게 적힌다. 공백을 뗀 뒤 맞춰 본다.
DELIV_ALIAS = {
    "구분": ("구분",),
    "날짜": ("날짜", "주문일", "주문일자"),
    "설치예정일": ("설치예정일", "설치예정", "예정일"),
    "설치확정일": ("설치확정일", "설치확정", "확정일", "설치일"),
    "주문코드": ("주문코드", "주문번호", "오더코드"),
}

_HEAD0 = "거래처명"          # 머리글 판별용


def _squash(v) -> str:
    return "".join(str(v or "").split())


def deliv_map(header) -> dict[str, int]:
    """머리글 한 줄에서 {칸이름: 0부터 센 열번호} 를 만든다. 못 찾으면 빈 dict."""
    pos: dict[str, int] = {}
    cells = [_squash(v) for v in header]
    for name, names in DELIV_ALIAS.items():
        for i, c in enumerate(cells):
            if c in names:
                pos[name] = i
                break
    # 구분·주문코드 둘 다 있어야 머리글로 인정한다.
    return pos if {"구분", "주문코드"} <= set(pos) else {}


def ecount_map(header) -> tuple[dict[str, int | None], list[str]]:
    """머리글 한 줄에서 {칸이름: 열번호} 를 만든다. 둘째 값은 이 표에 없는 칸.

    기본은 자리 번호(RAW_MAP 순서)다 — 정상가·수수료율처럼 머리글이 빈 칸이 있어서
    이름만으로는 찾을 수 없다. 다만 그 자리 머리글에 **다른 이름이 적혀 있으면**
    자리를 믿지 않는다. 같은 시트 아래쪽에 열 구성이 다른 표가 붙어 있을 때
    (예: K열이 정상가가 아니라 적요) 주문번호를 금액으로 읽는 일을 막는다.
    """
    cells = [_squash(v) for v in header]
    pos: dict[str, int | None] = {}
    없음: list[str] = []
    for i, name in enumerate(RAW_MAP):
        c = cells[i] if i < len(cells) else ""
        if not c or c.startswith(name):
            pos[name] = i                     # 머리글이 비었거나 맞다 — 자리 그대로
            continue
        다른곳 = [j for j, x in enumerate(cells) if x.startswith(name)]
        pos[name] = 다른곳[0] if 다른곳 else None
        if not 다른곳:
            없음.append(name)
    return pos, 없음


def read_ecount(src) -> tuple[pd.DataFrame, dict]:
    """이카운트확인_N월.xlsx — 지점별 시트를 모두 이어붙인다.

    반환: (원본 DataFrame, {시트명: 행수})
    열 구성이 다른 표를 만나 일부 칸을 비워 읽었으면 DataFrame.attrs["알림"] 에 적는다.
    """
    wb = openpyxl.load_workbook(src, data_only=True, read_only=True)
    rows, per, 알림 = [], {}, []
    기본 = {k: i for i, k in enumerate(RAW_MAP)}
    for sn in wb.sheetnames:
        ws, n = wb[sn], 0
        pos, 없음, 빈행수 = 기본, [], 0

        def 닫기():
            if 없음 and 빈행수:
                알림.append(f"{sn} 시트의 열 구성이 다른 표 {빈행수}행 — "
                            f"{'·'.join(없음)} 칸이 없어 비워 읽음")

        for r in ws.iter_rows(min_row=1, values_only=True):
            if not r or len(r) < 14:
                continue
            if str(r[0] or "").strip().startswith(_HEAD0):
                닫기()                        # 머리글 행 — 여기서부터 이 머리글을 따른다
                pos, 없음 = ecount_map(r)
                빈행수 = 0
                continue
            if r[6] in (None, ""):            # G 품목명이 비면 무효 행
                continue
            행 = {k: (r[i] if i is not None and i < len(r) else None)
                 for k, i in pos.items()}
            # 주문번호가 숫자 칸으로 저장돼 있으면 소수로 읽힌다(2.0261E+14) — 정수로 되돌린다.
            if isinstance(행["적요"], float) and 행["적요"].is_integer():
                행["적요"] = int(행["적요"])
            rows.append(행)
            n += 1
            빈행수 += 1
        닫기()
        if n:
            per[sn] = n
    wb.close()
    df = pd.DataFrame(rows, columns=RAW_MAP)
    df.attrs["알림"] = 알림
    return df, per


def read_delivery(src) -> pd.DataFrame:
    """배송관리 대장 — 구분/날짜/설치예정일/설치확정일/주문코드만 추린다.

    열 순서는 내보내기 양식에 따라 바뀐다(2026년 중 소비자명이 4열에서 13열로
    옮겨가며 그 사이가 한 칸씩 당겨졌다). 그래서 자리 번호를 믿지 않고
    머리글 이름으로 칸을 찾는다. 머리글이 없으면 옛 자리로 되돌아간다.
    """
    wb = openpyxl.load_workbook(src, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    pos, out = {}, []
    for r in ws.iter_rows(min_row=1, values_only=True):
        if not r:
            continue
        if not pos:
            pos = deliv_map(r)
            if pos:
                continue                      # 이 줄은 머리글
            if len(r) < 13:
                continue
            pos = {n: c - 1 for c, n in DELIV_COLS.items()}   # 머리글 없는 파일
        키 = pos["주문코드"]
        if len(r) <= 키 or r[키] in (None, ""):
            continue                          # 주문코드 없으면 무효
        out.append({n: (r[i] if len(r) > i else None) for n, i in pos.items()})
    wb.close()
    return pd.DataFrame(out, columns=list(DELIV_COLS.values()))


def read_from_xlsm(src) -> tuple[pd.DataFrame, pd.DataFrame]:
    """기존 인센티브 .xlsm 에서 원본붙여넣기·배송완료원본을 그대로 꺼낸다(이관용)."""
    wb = openpyxl.load_workbook(src, data_only=True, keep_vba=True)
    ws = wb["원본붙여넣기"]
    raw = [{n: ws.cell(r, i + 1).value for i, n in enumerate(RAW_MAP)}
           for r in range(2, ws.max_row + 1) if ws.cell(r, 7).value not in (None, "")]
    d = wb["배송완료원본"]
    머리 = [d.cell(1, c).value for c in range(1, d.max_column + 1)]
    pos = deliv_map(머리) or {n: c - 1 for c, n in DELIV_COLS.items()}
    키 = pos["주문코드"]
    dl = [{n: d.cell(r, i + 1).value for n, i in pos.items()}
          for r in range(2, min(d.max_row, 2001) + 1)
          if d.cell(r, 키 + 1).value not in (None, "")]
    wb.close()
    return (pd.DataFrame(raw, columns=RAW_MAP),
            pd.DataFrame(dl, columns=list(DELIV_COLS.values())))
