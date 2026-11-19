# -*- coding: utf-8 -*-
"""백화점 오프라인 인센티브 마감 — 사내 웹앱.

실행:  streamlit run app.py --server.address 0.0.0.0 --server.port 8502
"""
from __future__ import annotations

import hashlib
import io
import os
import sys
import warnings
from datetime import date, datetime
from pathlib import Path

warnings.filterwarnings("ignore")
sys.path.insert(0, str(Path(__file__).resolve().parent))

import pandas as pd
import streamlit as st

import ui
from core import backend, engine, excel_out, gate, io_in, rates, storage, stores
from core.xlsm_reader import Settings

st.set_page_config(page_title="백화점 인센티브 마감", page_icon="🏬", layout="wide")
ui.apply()

# 핀번호가 설정돼 있으면 여기서 막는다. 통과 전에는 데이터를 읽지도 않는다.
if not gate.열렸나():
    st.stop()

# 접속 정보가 잘못 적혀 있으면 긴 오류 대신 무엇을 고칠지 알려 준다.
_문제 = backend.접속정보_문제()
if _문제:
    st.error(_문제)
    st.stop()


def won(v) -> str:
    return f"{float(v or 0):,.0f}"


def 기본마감월() -> str:
    t = date.today()
    y, m = (t.year, t.month - 1) if t.month > 1 else (t.year - 1, 12)
    return f"{y:04d}-{m:02d}"


def 월목록(n: int = 18) -> list[str]:
    t = date.today()
    out, y, m = [], t.year, t.month
    for _ in range(n):
        out.append(f"{y:04d}-{m:02d}")
        m -= 1
        if m == 0:
            y, m = y - 1, 12
    return out


def make_settings(마감월_text: str):
    y, m = (int(x) for x in 마감월_text.split("-"))
    sd = stores.load()
    s = Settings(마감월=date(y, m, 1))
    b = sd["기본목표"]
    s.기준1, s.기준2, s.기준3 = float(b["1차"]), float(b["2차"]), float(b["3차"])
    적용 = rates.apply_to(s)
    s.별칭 = stores.alias_map(sd)
    s.목표 = stores.targets(sd)
    s.지점키워드 = stores.키워드(sd)
    try:
        s.기본수수료율 = stores.수수료율(sd)
    except ValueError as e:
        st.error(f"{e} 데이터 저장소의 거래처.json 에 적어 주세요.")
        st.stop()
    return s, 적용, sd


def 숫자(s, default: float = 0.0) -> float:
    """'1,000,000원' 처럼 적어도 숫자로 읽는다."""
    t = "".join(ch for ch in str(s) if ch.isdigit() or ch == ".")
    try:
        return float(t) if t else default
    except ValueError:
        return default


def 차수표기(x) -> str:
    s = str(x)
    return s if s.endswith("차") else f"{s}차"


# ══════════════════════════════════════════════════════════════
목록 = 월목록()
기본 = 기본마감월()
if "마감월" not in st.session_state:
    st.session_state.마감월 = 기본 if 기본 in 목록 else 목록[0]

ui.topbar(st.session_state.마감월)

# 데이터 저장소 토큰은 유효기간이 있다. 만료되면 계산은 되는데 저장만 안 되므로
# 마감 당일에야 알아차리기 쉽다. 한 달 전부터 모든 화면 위에 띄워 둔다.
남은 = backend.token_남은일수()
if 남은 is not None and 남은 <= 30:
    끝 = backend.token_만료일()
    if 남은 < 0:
        st.error(f"**저장소 토큰이 {-남은}일 전에 만료됐습니다 ({끝}).** "
                 "계산은 되지만 저장이 안 됩니다. 새 토큰을 발급해 "
                 "Streamlit 설정의 Secrets 에 바꿔 넣으세요 — README 참고.")
    else:
        st.warning(f"**저장소 토큰이 {남은}일 뒤 만료됩니다 ({끝}).** "
                   "만료되면 저장이 안 됩니다. 미리 갱신해 두세요 — README 참고.")

tab실행, tab현황, tab이월, tab이력, tab설정 = st.tabs(
    ["마감 실행", "지점 현황", "이월명부", "마감이력", "설정"])


# ────────────────────────────── 마감 실행
with tab실행:
    ui.section("마감 실행",
               "파일 두 개를 올리면 차수 판정·인센티브 산정·배송 대조·이월 등록이 "
               "한 번에 끝납니다. 엑셀 매크로를 열거나 지난달 파일을 복제할 필요가 "
               "없습니다.")

    c = st.columns([1, 3])
    with c[0]:
        st.session_state.마감월 = st.selectbox(
            "마감월", 목록, index=목록.index(st.session_state.마감월))
    마감월 = st.session_state.마감월
    설정, 적용수당표, sd = make_settings(마감월)

    u1, u2 = st.columns(2)
    with u1:
        f판매 = st.file_uploader(
            "이카운트 지점별 마감", type=["xlsx", "xlsm"],
            help="지점별 시트를 그대로 올리세요. 머리글과 빈 행은 자동으로 걸러집니다.")
    with u2:
        f배송 = st.file_uploader(
            "배송완료 원장", type=["xlsx", "xlsm"],
            help="구분·설치확정일·주문코드 기준으로 대조합니다.")

    # 올린 파일은 보관해 두고, 다시 들어왔을 때 올린 것이 없으면 그 달의 보관본을 쓴다.
    def 보관하거나_불러오기(f, 종류: str):
        try:
            if f is not None:
                data = f.getvalue()
                지문 = hashlib.sha256(data).hexdigest()
                칸 = f"_보관_{마감월}_{종류}"
                if st.session_state.get(칸) != 지문:     # 같은 파일을 화면 갱신마다 올리지 않는다
                    backend.write_upload(마감월, 종류, data)
                    st.session_state[칸] = 지문
                return f, False
            b = backend.read_upload(마감월, 종류)
        except Exception as e:                       # noqa: BLE001
            st.warning(f"{종류} 보관본을 다루지 못했습니다 — {type(e).__name__}", icon="⚠️")
            return f, False
        if b is None:
            return None, False
        보관본 = io.BytesIO(b)
        보관본.name = f"{마감월} {종류} (지난번에 올린 파일)"
        return 보관본, True

    f판매, 판매_보관본 = 보관하거나_불러오기(f판매, "판매원장")
    f배송, 배송_보관본 = 보관하거나_불러오기(f배송, "배송원장")

    raw = deliv = None
    데모판매, 데모배송 = os.environ.get("DEMO_SALES"), os.environ.get("DEMO_DELIV")
    if f판매 is None and 데모판매 and Path(데모판매).exists():
        f판매 = 데모판매
    if f배송 is None and 데모배송 and Path(데모배송).exists():
        f배송 = 데모배송

    시트별, 읽기알림 = {}, []
    if f판매 is not None:
        try:
            raw, 시트별 = io_in.read_ecount(f판매)
            읽기알림 = raw.attrs.get("알림", [])
        except Exception as e:                       # noqa: BLE001
            st.error(f"판매원장을 읽지 못했습니다 — {e}")
    if f배송 is not None:
        try:
            deliv = io_in.read_delivery(f배송)
        except Exception as e:                       # noqa: BLE001
            st.error(f"배송원장을 읽지 못했습니다 — {e}")

    def 파일명(f) -> str:
        return Path(f).name if isinstance(f, str) else getattr(f, "name", "")

    i1, i2 = st.columns(2)
    i1.markdown(
        f"<div class='filenote'>{파일명(f판매)} · {len(raw):,}행</div>"
        if raw is not None else "<div class='filenote dim'>파일을 올려주세요</div>",
        unsafe_allow_html=True)
    i2.markdown(
        f"<div class='filenote'>{파일명(f배송)} · {len(deliv):,}행</div>"
        if deliv is not None else "<div class='filenote dim'>파일을 올려주세요</div>",
        unsafe_allow_html=True)
    if 시트별:
        i1.markdown("<div class='filenote dim'>"
                    + " · ".join(f"{k} {v}" for k, v in 시트별.items())
                    + "</div>", unsafe_allow_html=True)

    개별 = {k for k, v in 설정.목표.items()
          if v != (설정.기준1, 설정.기준2, 설정.기준3)}
    ui.cards([
        ("적용 수당표",
         f"{적용수당표} · {len(설정.단가표)}품목"
         + ("  ← 상속" if 적용수당표 != 마감월 else ""),
         적용수당표 != 마감월),
        ("적용 목표금액",
         f"거래처별 설정 ({len(개별)}곳)" if 개별 else f"공통 {설정.기준1:,.0f}원",
         False),
        ("마감기준일 — 배송완료 인정 한도",
         f"{engine.serial_to_date(engine.eomonth(설정.마감월)):%Y-%m-%d}", False),
    ])
    if 적용수당표 != 마감월:
        st.caption(f"{마감월} 수당표가 없어 {적용수당표} 표를 이어받았습니다. "
                   f"수당안이 바뀌었다면 설정 탭에서 {마감월}을 등록하세요.")

    if raw is not None and deliv is not None:
        carry0 = storage.load_carryover()
        res = engine.run(raw, deliv, 설정, carryover=carry0,
                         지점목록=stores.active(sd))
        share, ex, idx = res["영업팀공유"], res["대상추출"], res["배송인덱스"]
        이월후, 신규, 중복 = engine.register_carryover(ex, res["이월명부"], 마감월)

        대상 = ex[ex["판정"] == "✅대형안마의자"]
        당월건 = int((ex["지급구분"] == "당월지급").sum())
        이월건 = int((ex["지급구분"] == "익월이월").sum())
        기타 = ex[ex["지점"] == "기타"]
        미등재 = sorted({str(c) for c in 대상["품목코드"] if str(c) not in 설정.단가표})

        줄 = [
            f"<span class='hd'>실행 완료</span> · "
            f"{datetime.now():%Y-%m-%d %H:%M}",
            f"판매원장 <b>{len(raw):,}행</b>  "
            + " / ".join(f"{k} {v}" for k, v in 시트별.items()),
            f"배송원장 <b>{len(deliv):,}행</b> / 기준일 내 배송완료 "
            f"<b>{int(idx['마감월내배송완료'].sum())}건</b> "
            f"(전체 {int(idx['배송완료'].sum())}건) / 취소·반품 "
            f"<b>{int(idx['취소반품'].sum())}건</b>",
            f"대상 건수 <b>{len(대상)}건</b> · 당월지급 <b>{당월건}건</b> · "
            f"익월이월 <b>{이월건}건</b>",
            f"이월명부 신규 <b>{신규}건</b> 추가, 중복 {중복}건 건너뜀 · "
            f"취소 자동판정 <b>{res['취소자동변경']}건</b> 취소종결",
            (f"<span class='ok'>✓ 거래처 표에 없는 거래처 없음</span>" if 기타.empty
             else f"<span class='no'>⚠ 거래처 표에 없는 행 {len(기타)}건</span> — "
                  f"설정 탭에서 ‘다른 표기’를 등록하세요"),
        ]
        for 말 in 읽기알림:
            줄.append(f"<span class='no'>⚠ {말}</span>")
        if 미등재:
            줄.append(f"<span class='no'>⚠ 수당표에 없는 품목 {len(미등재)}종</span> "
                     f"→ {', '.join(미등재)} (수당 0원으로 계산)")
        ui.runbox(줄)

        st.write("")
        m = st.columns(5)
        m[0].metric("정상가 합계", won(share["정상가합계"].sum()))
        m[1].metric("대상 건수", f"{int(share['건수'].sum())}건")
        m[2].metric("판매월귀속 인센티브", won(share["인센티브총액"].sum()))
        m[3].metric("당월 실지급", won(share["당월지급"].sum()))
        이월지급 = engine.이월지급건(res["이월명부"], 마감월)
        이월지급 = 이월지급[이월지급["지점"].isin(share["지점"])
                        & (이월지급["확정단가"].map(engine.loose_num) > 0)]
        이월지급액 = engine.sum_num(이월지급["확정단가"]) if not 이월지급.empty else 0.0
        이번달분 = float(share["당월지급"].sum()) - 이월지급액
        if 이월지급액:
            m[3].caption(f"이번 달 {won(이번달분)} + 이월 {won(이월지급액)}")
        m[4].metric("익월이월", won(share["이월금액"].sum()))

        ui.section("거래처별 결과")
        v = share.copy()
        v.insert(2, "1차목표", [won(설정.목표.get(s, (0, 0, 0))[0]) for s in v["지점"]])
        v["차수"] = v["차수"].map(차수표기)
        for col in ["정상가합계", "인센티브총액", "당월지급", "이월금액"]:
            v[col] = v[col].map(won)
        st.dataframe(v, hide_index=True, use_container_width=True)

        ui.section("대상 건 상세")
        보기 = st.radio("보기", ["수당 발생 건만", "대형안마의자 전체", "전체"],
                       horizontal=True, label_visibility="collapsed")
        d = (ex[ex["인센단가"].map(engine.num) > 0] if 보기 == "수당 발생 건만"
             else 대상 if 보기 == "대형안마의자 전체" else ex)
        열 = [c for c in ["지점", "주문서일자", "품목코드", "품목명", "채널", "정상가",
                        "판정", "배송", "차수", "인센단가", "지급구분"] if c in d.columns]
        st.dataframe(d[열], hide_index=True, use_container_width=True, height=330)

        if not 이월지급.empty:
            ui.section("이번 달 지급되는 이월 건")
            이월열 = [c for c in ["판매월", "지점", "품목코드", "품목명", "확정차수",
                              "확정단가", "지급상태"] if c in 이월지급.columns]
            st.dataframe(이월지급[이월열].sort_values(["판매월", "지점"]),
                         hide_index=True, use_container_width=True)
            st.caption(f"{won(이번달분)} + {won(이월지급액)} = "
                       f"{won(share['당월지급'].sum())} 당월 실지급")

        st.divider()
        xlsx = excel_out.build(res, 마감월, 이월후)
        b = st.columns([1, 1, 2])
        b[0].download_button(
            "엑셀 내려받기", xlsx, file_name=f"백화점_인센티브_{마감월}.xlsx",
            use_container_width=True,
            mime="application/vnd.openxmlformats-officedocument."
                 "spreadsheetml.sheet")
        확정 = b[1].button("마감 확정", type="primary", use_container_width=True)
        메모 = b[2].text_input("마감 메모", label_visibility="collapsed",
                             placeholder="마감 메모 (선택)")
        확정명부, 지급건 = engine.지급처리(이월후, 마감월)
        st.caption(f"‘마감 확정’을 누르면 이월명부에 신규 {신규}건이 등록되고, "
                   f"이번 달 지급되는 이월 {지급건}건이 지급완료(지급월 {마감월})로 "
                   f"바뀌며, 마감이력에 기록됩니다. 엑셀만 받아볼 때는 누르지 않아도 됩니다.")
        if 확정:
            storage.save_carryover(확정명부)
            storage.record_close(마감월, share, 메모, 적용수당표,
                                 {"판매원장": len(raw), "배송원장": len(deliv)})
            storage.save_snapshot(마감월, xlsx, "마감본.xlsx")
            st.success(f"{마감월} 마감을 확정했습니다 — 이월명부 {len(이월후)}건 "
                       f"(신규 {신규}건, 지급완료 처리 {지급건}건).")
    else:
        st.caption("두 파일을 모두 올리면 자동으로 계산합니다.")


# ────────────────────────────── 지점 현황
with tab현황:
    h = storage.load_history()
    if not h:
        ui.section("지점 현황", "확정된 마감이 아직 없습니다. "
                             "마감 실행 탭에서 첫 마감을 확정하세요.")
    else:
        ui.section("지점 현황")
        월 = st.selectbox("마감월", [r["마감월"] for r in reversed(h)], key="현황월")
        rec = next(r for r in h if r["마감월"] == 월)
        tg = stores.targets()
        st.markdown(f"<span class='muted'>확정 {rec['확정시각']} · {rec['회차']}회차 "
                    f"· 수당표 {rec.get('적용수당표', '-')}</span>",
                    unsafe_allow_html=True)
        st.write("")
        for r in rec["지점별"]:
            g1, g2, g3 = tg.get(r["지점"], stores.기본목표())
            v = float(r["정상가합계"])
            try:
                차 = int(float(str(r["차수"]).replace("차", "")))
            except (TypeError, ValueError):
                차 = 0
            pct = min(v / g3 * 100, 100) if g3 else 0
            색 = {3: "#15532F", 2: "#1C6B45", 1: "#5FA47B"}.get(차, "#C7D2CB")
            다음 = {0: ("1차", g1), 1: ("2차", g2), 2: ("3차", g3)}.get(차)
            남 = (f"{다음[0]}까지 {won(다음[1] - v)}원" if 다음 and 다음[1] > v
                 else "최고 차수 달성")
            c = st.columns([3, 1, 1, 1, 1])
            c[0].markdown(
                f"**{r['지점']}** <span class='tag t{차}'>{차}차</span><br>"
                f"<span class='muted'>{won(v)}원 · {남}</span>"
                f"<div class='bar'><div style='width:{pct:.1f}%;background:{색}'>"
                f"</div></div>", unsafe_allow_html=True)
            c[1].metric("건수", f"{int(float(r['건수']))}")
            c[2].metric("인센티브", won(r["인센티브총액"]))
            c[3].metric("당월지급", won(r["당월지급"]))
            c[4].metric("이월", won(r["이월금액"]))
            st.write("")


# ────────────────────────────── 이월명부
with tab이월:
    co = storage.load_carryover()
    if co.empty:
        ui.section("이월명부", "아직 비어 있습니다. 마감을 확정하면 미배송 건이 "
                            "여기에 쌓이고, 다음 달 배송이 끝나면 자동으로 "
                            "‘지급 가능’으로 올라옵니다.")
    else:
        ui.section("이월명부", "미배송이라 이번 달에 못 준 수당을 추적합니다. "
                            "배송이 끝나면 자동으로 표시되고, 원장에서 취소된 건은 "
                            "자동으로 취소종결 처리됩니다.")
        # 예전에 올라간 0원 건은 지우지 않되, 대기 목록과 숫자에서는 뺀다.
        대기 = co[(co["지급상태"].astype(str).str.strip() == "대기")
                & (co["확정단가"].map(engine.loose_num) > 0)]
        완 = 대기[대기["실제배송완료"].astype(str).str.strip() != ""]
        c = st.columns(4)
        c[0].metric("전체", f"{len(co)}건")
        c[1].metric("대기", f"{len(대기)}건")
        c[2].metric("대기 금액", won(대기["확정단가"].map(engine.loose_num).sum()))
        c[3].metric("배송완료·지급대기", f"{len(완)}건")
        if len(완):
            st.success(f"배송이 끝나 지급 가능한 건이 {len(완)}건 "
                       f"({won(완['확정단가'].map(engine.loose_num).sum())}원) 있습니다. "
                       f"그 달 마감을 확정하면 자동으로 ‘지급완료’가 됩니다.",
                       icon="💰")
        f = st.radio("필터", ["대기만", "전체"], horizontal=True,
                     label_visibility="collapsed")
        d = 대기 if f == "대기만" else co
        ed = st.data_editor(
            d, hide_index=True, use_container_width=True, height=430,
            num_rows="fixed", key="co_ed",
            column_config={
                "지급상태": st.column_config.SelectboxColumn(
                    options=["대기", "지급완료", "취소종결"]),
                "지급월": st.column_config.TextColumn(
                    help="지급완료로 바꿀 때 실제 지급한 달 (예 2026-10)"),
            })
        if st.button("이월명부 저장", type="primary"):
            merged = co.copy()
            merged.loc[ed.index, ed.columns] = ed
            storage.save_carryover(merged)
            st.success("저장했습니다.")
            st.rerun()


# ────────────────────────────── 마감이력
with tab이력:
    h = storage.load_history()
    if not h:
        ui.section("마감이력", "확정된 마감이 아직 없습니다.")
    else:
        ui.section("마감이력")
        t = pd.DataFrame([{
            "마감월": r["마감월"], "회차": r["회차"], "확정시각": r["확정시각"],
            "수당표": r.get("적용수당표", ""),
            "정상가합계": won(r["정상가합계"]), "건수": r["건수"],
            "인센티브": won(r["인센티브총액"]), "당월지급": won(r["당월지급"]),
            "이월건수": r["이월건수"], "이월금액": won(r["이월금액"]),
            "메모": r.get("메모", ""),
        } for r in reversed(h)])
        st.dataframe(t, hide_index=True, use_container_width=True)
        files = storage.snapshots()
        if files:
            ui.section("확정 마감본")
            for name in reversed(files):
                st.download_button(name, storage.snapshot_bytes(name),
                                   file_name=name, key=f"dl_{name}")


# ────────────────────────────── 설정
with tab설정:
    sd = stores.load()
    st.markdown(
        f"<div class='muted'>저장 위치 — {backend.어디()}"
        + (f" · 토큰 만료 {backend.token_만료일()}"
           f" (남은 {backend.token_남은일수()}일)"
           if backend.token_만료일() else "")
        + "</div>", unsafe_allow_html=True)
    if gate.핀():
        if st.button("나가기", help="이 브라우저에서 핀번호를 다시 묻게 합니다"):
            gate.나가기()
            st.rerun()
    s1, s2 = st.tabs(["거래처 · 목표", "판매수당표"])

    with s1:
        ui.section("공통 목표금액", "거래처별로 따로 정하지 않으면 이 금액을 씁니다. "
                                "쉼표를 넣어 적어도 됩니다.")
        b = sd["기본목표"]
        g = st.columns(4)
        t1 = g[0].text_input("1차", value=f"{int(b['1차']):,}")
        t2 = g[1].text_input("2차", value=f"{int(b['2차']):,}")
        t3 = g[2].text_input("3차", value=f"{int(b['3차']):,}")
        n1, n2, n3 = 숫자(t1), 숫자(t2), 숫자(t3)
        g[3].markdown(f"<div style='height:28px'></div><span class='muted'>"
                      f"{n1:,.0f} → {n2:,.0f} → {n3:,.0f}</span>",
                      unsafe_allow_html=True)
        if not (0 < n1 <= n2 <= n3):
            st.warning("1차 ≤ 2차 ≤ 3차 순서로 입력해 주세요.", icon="⚠️")

        ui.section("기본 수수료율", "판매원장에 수수료율이 비어 있는 행에 씁니다.")
        tr = st.columns(4)[0].text_input(
            "수수료율(%)", value=f"{float(sd.get('기본수수료율', 0)) * 100:g}")
        수수료 = round(숫자(tr, -1.0) / 100, 6)
        수수료_맞음 = 0 <= 수수료 < 1
        if not 수수료_맞음:
            st.warning("0 이상 100 미만의 숫자로 입력해 주세요.", icon="⚠️")

        ui.section("거래처",
                   "‘다른 표기’는 이카운트에 다른 이름으로 찍히지만 같은 매장인 "
                   "경우입니다 (쉼표로 여러 개). ‘공통 목표’를 켜두면 위 금액을 "
                   "따라가고, 끄면 그 거래처만 다른 금액을 줄 수 있습니다.")
        rows = [{
            "사용": s.get("사용", True), "거래처명": s["이름"],
            "다른 표기": ", ".join(s.get("별칭") or []),
            "공통 목표": not s.get("목표"),
            "1차": float((s.get("목표") or {}).get("1차", n1)),
            "2차": float((s.get("목표") or {}).get("2차", n2)),
            "3차": float((s.get("목표") or {}).get("3차", n3)),
        } for s in sd["지점"]]
        ed = st.data_editor(
            pd.DataFrame(rows), hide_index=True, use_container_width=True,
            num_rows="dynamic", key="store_ed",
            column_config={
                "사용": st.column_config.CheckboxColumn(width="small"),
                "거래처명": st.column_config.TextColumn(width="medium"),
                "다른 표기": st.column_config.TextColumn(
                    width="medium",
                    help="이카운트에 다른 이름으로 찍히는 경우 여기에 적습니다"),
                "공통 목표": st.column_config.CheckboxColumn(
                    width="small", help="켜면 위 공통 목표금액을 따라갑니다"),
                "1차": st.column_config.NumberColumn(format="localized", step=1_000_000),
                "2차": st.column_config.NumberColumn(format="localized", step=1_000_000),
                "3차": st.column_config.NumberColumn(format="localized", step=1_000_000),
            })
        st.caption("‘공통 목표’가 켜진 줄은 1·2·3차를 고쳐도 저장할 때 공통 금액으로 "
                   "되돌아갑니다. 따로 주려면 체크를 먼저 해제하세요.")
        if st.button("거래처 저장", type="primary", disabled=not 수수료_맞음):
            out = []
            for _, r in ed.iterrows():
                nm = str(r["거래처명"] or "").strip()
                if not nm:
                    continue
                tgt = None
                if not bool(r.get("공통 목표", True)):
                    tgt = {"1차": 숫자(r["1차"], n1), "2차": 숫자(r["2차"], n2),
                           "3차": 숫자(r["3차"], n3)}
                out.append({"이름": nm,
                            "별칭": [x.strip() for x in
                                   str(r["다른 표기"] or "").split(",") if x.strip()],
                            "목표": tgt, "사용": bool(r["사용"])})
            sd["지점"] = out
            sd["기본목표"] = {"1차": float(n1), "2차": float(n2), "3차": float(n3)}
            sd["기본수수료율"] = 수수료
            stores.save(sd)
            st.success("저장했습니다.")
            st.rerun()

    with s2:
        all_r = rates.load_all()
        등록 = sorted(all_r["표"])
        ui.section("판매수당표",
                   "수당안은 달마다 바뀝니다. 월별로 보관하므로 지난달을 다시 열어도 "
                   "그때 기준 그대로 재현됩니다. 표에 없는 품목코드는 수당 대상이 "
                   "아닙니다(0원).")
        c = st.columns([2, 2, 3])
        고른달 = c[0].selectbox("적용월", 등록 + ["＋ 새 달 추가"], index=len(등록) - 1)
        if 고른달 == "＋ 새 달 추가":
            후보 = [m for m in 월목록() if m not in 등록]
            월 = c[1].selectbox("만들 달", 후보) if 후보 else 등록[-1]
            c[2].markdown(f"<div style='height:28px'></div><span class='muted'>"
                          f"{rates.for_month(월)[1]} 표를 복사해 시작합니다</span>",
                          unsafe_allow_html=True)
            품목, 기존출처 = rates.copy_from_previous(월), ""
        else:
            월 = 고른달
            품목 = all_r["표"][월]["품목"]
            기존출처 = all_r["표"][월].get("_출처", "")

        df = pd.DataFrame([{
            "품목코드": k, "품목명": v.get("품목명", ""),
            "정상가": int(v.get("정상가") or 0),
            "1차": int(v.get("1차") or 0), "2차": int(v.get("2차") or 0),
            "3차": int(v.get("3차") or 0),
            "온라인만": "온라인만" in str(v.get("판매채널") or ""),
        } for k, v in 품목.items()])
        ed2 = st.data_editor(
            df, hide_index=True, use_container_width=True, num_rows="dynamic",
            key=f"rate_ed_{월}",
            column_config={
                "품목코드": st.column_config.TextColumn(width="small"),
                "정상가": st.column_config.NumberColumn(format="localized"),
                "1차": st.column_config.NumberColumn(format="localized"),
                "2차": st.column_config.NumberColumn(format="localized"),
                "3차": st.column_config.NumberColumn(format="localized"),
                "온라인만": st.column_config.CheckboxColumn(
                    width="small", help="오프라인 판매는 수당 제외"),
            })
        st.caption("‘온라인만’은 적요의 주문번호가 숫자로만 된 온라인몰 건에만 수당을 "
                   "줍니다. 수당 대상이 아닌 품목은 단가 0으로 남겨두면 매달 경고가 "
                   "뜨지 않습니다.")
        출처 = st.text_input("출처 메모", value=기존출처,
                           placeholder="예) 26년 10월 백화점 판매 수당안")
        if st.button(f"{월} 수당표 저장", type="primary"):
            새품 = {}
            for _, r in ed2.iterrows():
                k = str(r["품목코드"] or "").strip().upper()
                if not k:
                    continue
                새품[k] = {"품목명": str(r["품목명"] or ""),
                          "정상가": float(r["정상가"] or 0),
                          "1차": float(r["1차"] or 0), "2차": float(r["2차"] or 0),
                          "3차": float(r["3차"] or 0),
                          "판매채널": "온라인만" if bool(r["온라인만"]) else ""}
            rates.save_month(월, 새품, 출처)
            st.success(f"{월} 수당표를 저장했습니다 ({len(새품)}품목).")
            st.rerun()

        전 = [m for m in 등록 if m < 월]
        if 전:
            ui.section(f"{전[-1]} → {월} 변경점")
            dd = rates.diff(전[-1], 월)
            if not dd:
                st.markdown("<span class='muted'>변경 없음</span>",
                            unsafe_allow_html=True)
            for x in dd:
                a, b2 = x["이전"] or {}, x["현재"] or {}
                st.markdown(
                    f"- **{x['구분']}** {x['품목명']} (`{x['품목코드']}`) — "
                    f"{won(a.get('1차'))} / {won(a.get('2차'))} / {won(a.get('3차'))}"
                    f"  →  {won(b2.get('1차'))} / {won(b2.get('2차'))} / "
                    f"{won(b2.get('3차'))}"
                    + ("　`온라인만`" if "온라인만" in str(b2.get("판매채널") or "")
                       else ""))
