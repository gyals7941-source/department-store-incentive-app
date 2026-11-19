# -*- coding: utf-8 -*-
"""백화점 인센티브 마감 계산 엔진.

기존 26.09백화점_인센티브.xlsm (VBA v6) 의 수식/매크로 로직을 1:1 재현한다.
자세한 근거는 ../SPEC.md 참고.
"""
from __future__ import annotations

import datetime as _dt
import math
import re
import unicodedata

import pandas as pd

# ─────────────────────────────────────────────────────────────
# 엑셀 호환 유틸
# ─────────────────────────────────────────────────────────────
_EXCEL_EPOCH = _dt.datetime(1899, 12, 30)

대형안마의자 = "✅대형안마의자"
취소상계 = "❌취소상계"
검토필요 = "⚠검토필요"
대상아님 = "대상아님"

당월지급 = "당월지급"
취소제외 = "취소제외"
익월이월 = "익월이월"

배송완료표기 = "O(배송완료 확인됨)"


def to_serial(v) -> float:
    """엑셀 N() 과 같은 의미의 숫자값. 날짜는 시리얼, 숫자는 그대로, 그 외 0."""
    if v is None:
        return 0.0
    if isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        if isinstance(v, float) and math.isnan(v):
            return 0.0
        return float(v)
    if isinstance(v, _dt.datetime):
        return (v - _EXCEL_EPOCH).total_seconds() / 86400.0
    if isinstance(v, _dt.date):
        return float((_dt.datetime(v.year, v.month, v.day) - _EXCEL_EPOCH).days)
    return 0.0


def serial_to_date(s: float):
    if not s or s <= 0:
        return None
    return _EXCEL_EPOCH + _dt.timedelta(days=float(s))


def is_number(v) -> bool:
    """엑셀 ISNUMBER(). 숫자형 셀만 True (숫자처럼 보이는 문자열은 False)."""
    if isinstance(v, bool) or v is None:
        return False
    if isinstance(v, (int, float)):
        return not (isinstance(v, float) and math.isnan(v))
    if isinstance(v, (_dt.datetime, _dt.date)):
        return True
    return False


def loose_num(v, default: float = 0.0) -> float:
    """숫자로 '적혀 있기만' 해도 숫자로 읽는다.

    num() 은 엑셀 ISNUMBER() 를 그대로 흉내내므로 '250000' 같은 문자열을
    0 으로 본다. 원본 시트에서는 그게 맞지만, 이월명부를 CSV 로 저장했다가
    다시 읽으면 전 열이 문자열이 되므로 금액이 통째로 0 이 되어 버린다.
    저장소에서 돌아온 금액에는 이 함수를 쓴다.
    """
    if is_number(v):
        return to_serial(v)
    t = text(v).replace(",", "").replace("원", "").strip()
    if not t:
        return default
    try:
        return float(t)
    except ValueError:
        return default


def sum_num(s) -> float:
    """합계를 항상 float 로 돌려준다.

    pandas 의 문자열 dtype 에서는 빈 시리즈의 sum() 이 0 이 아니라 '' 를 내놓아
    float 과 더할 때 터진다. 이월명부를 CSV 에서 읽으면 전 열이 문자열이므로
    금액 합계는 반드시 이 함수를 거친다.
    """
    if s is None or len(s) == 0:
        return 0.0
    return float(sum(loose_num(v) for v in s))


def num(v, default=0.0) -> float:
    return to_serial(v) if is_number(v) else default


def text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and math.isnan(v):
        return ""
    return str(v)


def excel_clean(s: str) -> str:
    """엑셀 CLEAN(): 인쇄 불가 문자(코드 0~31) 제거."""
    return "".join(ch for ch in s if ord(ch) >= 32 or ord(ch) > 127)


def excel_trim(s: str) -> str:
    """엑셀 TRIM(): 앞뒤 공백 제거 + 내부 연속 공백을 하나로."""
    return re.sub(r" +", " ", s.strip())


def search(needle: str, haystack) -> bool:
    """엑셀 ISNUMBER(SEARCH(..)) — 대소문자 무시 부분일치."""
    h = text(haystack)
    if not h:
        return False
    return needle.upper() in h.upper()


def excel_round(x: float, digits: int = 0) -> float:
    """엑셀 ROUND(): banker's rounding 이 아니라 half-away-from-zero."""
    if x is None:
        return 0.0
    f = 10 ** digits
    return math.floor(abs(x) * f + 0.5) / f * (1 if x >= 0 else -1)


def eomonth(d) -> float:
    """EOMONTH(d, 0) → 그 달 말일의 시리얼."""
    if isinstance(d, (int, float)):
        d = serial_to_date(d)
    if d is None:
        return to_serial(_dt.datetime(2099, 12, 31))
    y, m = d.year, d.month
    last = (_dt.datetime(y + (m // 12), (m % 12) + 1, 1) - _dt.timedelta(days=1))
    return to_serial(last)


def norm_key(v) -> str:
    """_배송인덱스 A열: SUBSTITUTE(TRIM(CLEAN(주문코드)), " ", "")"""
    s = text(v)
    if not s:
        return ""
    return excel_trim(excel_clean(s)).replace(" ", "")


# 온라인몰 주문번호: 날짜 8자리(20YYMMDD) 뒤에 영숫자가 붙고, 하이픈이 없다.
# 꼬리는 숫자일 수도(202610036231684) 영문이 섞일 수도(20261001D1B9CD) 있다.
_온라인주문 = re.compile(r"^20\d{6}[0-9A-Za-z]{5,8}$")


def is_online(적요) -> bool:
    """온라인몰 주문건 판별 — 적요의 주문번호 모양으로 가른다.

    예) '202610036231684'          → 온라인
        '20261001D1B9CD'           → 온라인 (꼬리에 영문이 섞인 것도 온라인몰)
        '20261008-2차오프-004'      → 오프라인 (하이픈이 들어간 특판 번호)
        '현장판매' · '매장지급'       → 오프라인
        '(품목명)'                  → 오프라인 (사은품 줄의 메모)
        ''                          → 오프라인 (적요 없음)
    """
    return bool(_온라인주문.match(excel_trim(text(적요))))


def head_key(v) -> str:
    """대상추출 D열: 적요에서 '/' 또는 ' ' 앞까지 잘라 TRIM."""
    s = text(v)
    if not s:
        return ""
    i = s.find("/")
    if i < 0:
        i = s.find(" ")
    cut = s if i < 0 else s[:i]
    return excel_trim(cut)


# ─────────────────────────────────────────────────────────────
# 1) _배송인덱스
# ─────────────────────────────────────────────────────────────
def build_delivery_index(delivery: pd.DataFrame, 마감월) -> pd.DataFrame:
    """배송완료원본 → 보조 인덱스.

    delivery 는 열 이름이 ['구분','날짜','설치예정일','설치확정일','주문코드'] 를
    포함해야 한다 (원본 A,B,F,G,M).
    """
    기준일 = eomonth(마감월)
    rows = []
    for _, r in delivery.iterrows():
        key = norm_key(r.get("주문코드"))
        if not key:
            rows.append(dict(정규화키="", 취소반품=0, 배송완료=0, 배송완료일=0.0,
                             마감월내배송완료=0))
            continue
        구분 = r.get("구분")
        취소 = 1 if (search("취소", 구분) or search("반품", 구분)) else 0
        완료 = 1 if excel_trim(text(구분)) == "배송완료" else 0
        g, f, b = num(r.get("설치확정일")), num(r.get("설치예정일")), num(r.get("날짜"))
        완료일 = g if g > 0 else (f if f > 0 else b)
        당월 = 1 if (완료 == 1 and 완료일 > 0 and 완료일 <= 기준일) else 0
        rows.append(dict(정규화키=key, 취소반품=취소, 배송완료=완료,
                         배송완료일=완료일, 마감월내배송완료=당월))
    idx = pd.DataFrame(rows)
    idx.attrs["마감기준일"] = 기준일
    return idx


def _idx_agg(idx: pd.DataFrame):
    """정규화키별 집계 — SUMIFS/MAX 를 한 번에."""
    live = idx[idx["정규화키"] != ""]
    g = live.groupby("정규화키", sort=False)
    return {
        "당월": g["마감월내배송완료"].sum().to_dict(),
        "취소": g["취소반품"].sum().to_dict(),
        "완료": g["배송완료"].sum().to_dict(),
        "완료일": (live.assign(_v=live["배송완료"] * live["배송완료일"])
                   .groupby("정규화키", sort=False)["_v"].max().to_dict()),
    }


# ─────────────────────────────────────────────────────────────
# 2) 대상추출
# ─────────────────────────────────────────────────────────────
_RAW_COLS = ["거래처명", "주문서일자", "판매일자", "회계전표일자", "비고", "주문자",
             "품목명", "수량", "단가", "합계", "정상가", "수수료율", "적요",
             "품목코드", "창고명", "프로젝트명", "거래처코드"]


def build_extract(raw: pd.DataFrame, idx: pd.DataFrame, settings) -> pd.DataFrame:
    """원본붙여넣기 + _배송인덱스 → 대상추출."""
    agg = _idx_agg(idx)
    rows = []
    for _, r in raw.iterrows():
        품목명 = r.get("품목명")
        if text(품목명).strip() == "":
            continue                      # 원본 수식: G 가 비면 전 열 ""

        거래처명 = text(r.get("거래처명"))
        # 어떤 거래처명을 지점으로 볼지는 거래처.json 이 정한다(유통사 이름이라 코드에 두지 않는다).
        키워드 = getattr(settings, "지점키워드", None) or []
        지점 = 거래처명 if any(search(k, 거래처명) for k in 키워드) else "기타"
        # 같은 매장이 이카운트에 다른 이름으로 찍히는 경우 대표 이름으로 접는다.
        지점 = getattr(settings, "별칭", {}).get(지점.strip(), 지점)
        D = head_key(r.get("적요"))
        품목코드 = text(r.get("품목코드"))
        수수료율 = r.get("수수료율")
        J = num(수수료율) if is_number(수수료율) else float(settings.기본수수료율)
        합계 = r.get("합계")
        K = excel_round(num(합계) * (1 - J), 0) if is_number(합계) else None

        rows.append(dict(
            지점=지점,
            건키=f"{거래처명}_{text(r.get('주문서일자'))}_{text(품목명)[:20]}",
            주문서일자=r.get("주문서일자"),
            정규화키=D,
            주문자=r.get("주문자"),
            품목코드=품목코드,
            품목명=품목명,
            단가=r.get("단가"),
            합계=합계,
            수수료율=J,
            순매출=K,
            정상가=num(r.get("정상가")),
            수량=num(r.get("수량")) if is_number(r.get("수량")) else 0.0,
            채널="온라인" if is_online(r.get("적요")) else "오프라인",
            _적요=r.get("적요"),
            _비고=r.get("비고"),
        ))
    ex = pd.DataFrame(rows)
    if ex.empty:
        return ex

    # S 상계합 : 같은 (정규화키, 품목코드) 의 수량 합
    ex["상계합"] = ex.groupby(["정규화키", "품목코드"], sort=False)["수량"].transform("sum")
    # T/U/V : 배송인덱스 조회
    ex["취소반품"] = ex["정규화키"].map(lambda k: agg["취소"].get(k, 0) if k else 0)
    ex["배송완료수"] = ex["정규화키"].map(lambda k: agg["완료"].get(k, 0) if k else 0)
    ex["배송완료일"] = ex["정규화키"].map(lambda k: agg["완료일"].get(k, 0) if k else 0)
    ex["_당월배송"] = ex["정규화키"].map(lambda k: agg["당월"].get(k, 0) if k else 0)

    # L 판정
    def judge(r):
        zpc = search("ZPC", r["품목코드"]) or search("ZPC", r["품목명"])
        if zpc and r["상계합"] <= 0 and r["수량"] != 0:
            return 취소상계
        if (search("교환", r["_적요"]) or search("교환", r["_비고"])
                or search("취소", r["_적요"]) or search("취소", r["_비고"])
                or (is_number(r["수량"]) and r["수량"] < 0)):
            return 검토필요
        if zpc and is_number(r["단가"]) and num(r["단가"]) > 0:
            return 대형안마의자
        return 대상아님

    ex["판정"] = ex.apply(judge, axis=1)

    # M 배송
    def delivery_state(r):
        if r["판정"] != 대형안마의자:
            return ""
        if r["_당월배송"] > 0:
            return "O"
        if r["배송완료수"] > 0:
            d = serial_to_date(r["배송완료일"])
            return f"익월배송({d:%m/%d})" if d else "익월배송()"
        return "미배송"

    ex["배송"] = ex.apply(delivery_state, axis=1)
    return ex


def apply_tier(ex: pd.DataFrame, tier: dict, settings) -> pd.DataFrame:
    """지점차수가 정해진 뒤 N(차수)·O(단가)·P(지급구분) 을 채운다."""
    if ex.empty:
        return ex

    def 차수(r):
        return tier.get(r["지점"], 0) if r["판정"] == 대형안마의자 else ""

    ex["차수"] = ex.apply(차수, axis=1)

    def 단가(r):
        if r["판정"] != 대형안마의자:
            return ""
        n = r["차수"] or 0
        if n == 0:
            return 0
        # 판매채널이 '온라인만'인 품목은 오프라인 판매에 수당을 주지 않는다.
        if r["품목코드"] in getattr(settings, "온라인전용", set()) and r["채널"] != "온라인":
            return 0
        return settings.단가표.get(r["품목코드"], {}).get(int(n), 0)

    ex["인센단가"] = ex.apply(단가, axis=1)

    def 지급구분(r):
        if r["판정"] != 대형안마의자:
            return ""
        if r["배송"] == "O":
            return 당월지급
        if r["취소반품"] > 0:
            return 취소제외
        return 익월이월

    ex["지급구분"] = ex.apply(지급구분, axis=1)
    return ex


# ─────────────────────────────────────────────────────────────
# 3) 지점차수
# ─────────────────────────────────────────────────────────────
def build_tier(ex: pd.DataFrame, 지점목록, settings) -> pd.DataFrame:
    """지점별 정상가 합계 → 차수.  ※ 판정 필터 없음(원본 동작 그대로)."""
    합 = (ex.groupby("지점", sort=False)["정상가"].sum().to_dict()
          if not ex.empty else {})
    목표 = getattr(settings, "목표", None) or {}
    기본 = (settings.기준1, settings.기준2, settings.기준3)
    rows = []
    for s in 지점목록:
        v = float(합.get(s, 0))
        g1, g2, g3 = 목표.get(s, 기본)       # 지점별 목표가 없으면 공통 목표
        if v >= g3:
            t = 3
        elif v >= g2:
            t = 2
        elif v >= g1:
            t = 1
        else:
            t = 0
        rows.append(dict(지점=s, 정상가합계=v, 차수=t,
                         목표1=g1, 목표2=g2, 목표3=g3))
    return pd.DataFrame(rows)


# ─────────────────────────────────────────────────────────────
# 4) 영업팀공유
# ─────────────────────────────────────────────────────────────
def 이월지급건(carryover: pd.DataFrame | None, 마감월_text: str = "") -> pd.DataFrame:
    """이월명부 중 이번 달에 지급되는 건 — 당월 실지급의 ②·③.

    ② 배송완료·대기이면서 판매월이 이번 달이 아닌 건
    ③ '지급완료' 이고 지급월이 이번 달인 건
    """
    if carryover is None or carryover.empty:
        return empty_carryover()
    c = carryover
    상태 = c["지급상태"].astype(str).str.strip()
    대기 = ((c["실제배송완료"] == 배송완료표기) & (상태 == "대기")
          & (c["판매월"].astype(str).str.strip() != 마감월_text))
    if "지급월" in c.columns:
        완료 = (상태 == "지급완료") & (c["지급월"].astype(str).str.strip() == 마감월_text)
    else:
        완료 = pd.Series(False, index=c.index)
    return pd.concat([c[대기], c[완료]])


def build_share(ex: pd.DataFrame, tier: pd.DataFrame,
                carryover: pd.DataFrame | None = None,
                마감월_text: str = "") -> pd.DataFrame:
    """영업팀공유.

    당월 실지급 = ① 이번 달 '당월지급' 건
                + ② 이월명부 중 배송완료·대기이면서 판매월이 이번 달이 아닌 건
                + ③ 이월명부 중 '지급완료' 이고 지급월이 이번 달인 건
    """
    rows = []
    이월지급 = 이월지급건(carryover, 마감월_text)
    for _, t in tier.iterrows():
        s = t["지점"]
        sub = ex[ex["지점"] == s] if not ex.empty else ex
        big = sub[(sub["판정"] == 대형안마의자) & (sub["지급구분"] != 취소제외)] \
            if not ex.empty else sub

        당월 = sum_num(sub[sub["지급구분"] == 당월지급]["인센단가"]) if not ex.empty else 0.0
        이월대기 = 0
        if not 이월지급.empty:
            이월대기 = sum_num(이월지급[이월지급["지점"] == s]["확정단가"])

        rows.append(dict(
            지점=s,
            정상가합계=float(t["정상가합계"]),
            차수=f'{int(t["차수"])}차',
            건수=int(len(big)),
            인센티브총액=sum_num(big["인센단가"]) if not ex.empty else 0.0,
            당월지급=float(당월 + 이월대기),
            이월건수=int((sub["지급구분"] == 익월이월).sum()) if not ex.empty else 0,
            이월금액=sum_num(sub[sub["지급구분"] == 익월이월]["인센단가"])
            if not ex.empty else 0.0,
        ))
    df = pd.DataFrame(rows)
    return df


# ─────────────────────────────────────────────────────────────
# 5) 이월명부
# ─────────────────────────────────────────────────────────────
CARRY_COLS = ["판매월", "지점", "건키", "정규화키", "품목코드", "품목명",
              "확정차수", "확정단가", "지급상태", "실제배송완료", "비고", "지급월"]


def empty_carryover() -> pd.DataFrame:
    return pd.DataFrame(columns=CARRY_COLS)


def register_carryover(ex: pd.DataFrame, carryover: pd.DataFrame,
                       마감월_text: str) -> tuple[pd.DataFrame, int, int]:
    """대상추출의 '익월이월' 건을 이월명부에 누적. 건키 기준 중복 제거.

    수당이 0원인 건(0차 지점, 단가 0원 품목)은 올리지 않는다 — 나중에 배송돼도
    줄 돈이 없어 추적할 것이 없다.
    """
    if carryover is None or carryover.empty:
        carryover = empty_carryover()
    기존 = set(carryover["건키"].astype(str))
    added, skipped, new_rows = 0, 0, []
    if not ex.empty:
        for _, r in ex[ex["지급구분"] == 익월이월].iterrows():
            if loose_num(r["인센단가"]) <= 0:
                continue
            k = text(r["건키"])
            if k and k in 기존:
                skipped += 1
                continue
            기존.add(k)
            new_rows.append(dict(
                판매월=마감월_text, 지점=r["지점"], 건키=k, 정규화키=r["정규화키"],
                품목코드=r["품목코드"], 품목명=r["품목명"],
                확정차수=r["차수"], 확정단가=r["인센단가"],
                지급상태="대기", 실제배송완료="", 비고="",
            ))
            added += 1
    if new_rows:
        carryover = pd.concat([carryover, pd.DataFrame(new_rows)], ignore_index=True)
    return carryover, added, skipped


def 지급처리(carryover: pd.DataFrame, 마감월_text: str) -> tuple[pd.DataFrame, int]:
    """마감 확정 때 — 이번 달에 지급되는 '대기' 이월 건을 지급완료로 넘긴다.

    대상은 당월 실지급의 ②(배송완료·대기·판매월이 이번 달이 아님)와 같다. 지급월을
    이번 달로 적어 두므로 같은 달을 다시 계산해도 ③으로 똑같이 잡히고, 다음 달에는
    다시 지급되지 않는다.
    """
    if carryover is None or carryover.empty:
        return empty_carryover(), 0
    c = carryover.copy()
    if "지급월" not in c.columns:
        c["지급월"] = ""
    대상 = ((c["실제배송완료"] == 배송완료표기)
          & (c["지급상태"].astype(str).str.strip() == "대기")
          & (c["판매월"].astype(str).str.strip() != 마감월_text))
    c.loc[대상, "지급상태"] = "지급완료"
    c.loc[대상, "지급월"] = 마감월_text
    return c, int(대상.sum())


def refresh_carryover(carryover: pd.DataFrame, idx: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """J(실제배송완료)·L(원장취소) 재계산 후 취소 자동판정."""
    if carryover is None or carryover.empty:
        return empty_carryover(), 0
    agg = _idx_agg(idx)
    carryover = carryover.copy()
    carryover["실제배송완료"] = carryover["정규화키"].map(
        lambda k: 배송완료표기 if (k and agg["당월"].get(k, 0) > 0) else "")
    원장취소 = carryover["정규화키"].map(
        lambda k: "취소" if (k and agg["취소"].get(k, 0) > 0) else "")

    changed = 0
    for i in carryover.index:
        if (text(carryover.at[i, "지급상태"]).strip() == "대기"
                and 원장취소.at[i] == "취소"
                and text(carryover.at[i, "실제배송완료"]).strip() == ""):
            carryover.at[i, "지급상태"] = "취소종결"
            carryover.at[i, "비고"] = "자동판정: 배송완료원본 구분=취소/반품 (무산)"
            changed += 1
    return carryover, changed


# ─────────────────────────────────────────────────────────────
# 6) 전체 파이프라인
# ─────────────────────────────────────────────────────────────
def run(raw: pd.DataFrame, delivery: pd.DataFrame, settings,
        carryover: pd.DataFrame | None = None, 지점목록=None) -> dict:
    idx = build_delivery_index(delivery, settings.마감월)
    ex = build_extract(raw, idx, settings)

    if 지점목록 is None:
        지점목록 = sorted(ex["지점"].unique()) if not ex.empty else []
    tier = build_tier(ex, 지점목록, settings)
    ex = apply_tier(ex, dict(zip(tier["지점"], tier["차수"])), settings)

    carryover, 취소변경 = refresh_carryover(carryover, idx)
    share = build_share(ex, tier, carryover, settings.마감월_text)

    return dict(배송인덱스=idx, 대상추출=ex, 지점차수=tier,
                영업팀공유=share, 이월명부=carryover, 취소자동변경=취소변경)
