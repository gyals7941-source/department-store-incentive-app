# -*- coding: utf-8 -*-
"""핀번호 출입문.

앱을 '주소를 아는 누구나' 로 열어두고, 대신 앞단에 핀번호를 세운다. 구글 로그인
없이 들어갈 수 있는 대신, **주소와 핀이 함께 새면 누구나 열 수 있다.** 그래서
무차별 대입만이라도 막아 둔다.

핀은 secrets 에 둔다. 코드에는 들어가지 않는다.

    [app]
    pin = "..."            # 숫자만 말고 길게 쓸수록 좋다

핀이 설정돼 있지 않으면 문을 세우지 않는다(사내 PC·로컬 실행용).

한 번 통과하면 그 브라우저에 8시간짜리 통행증(쿠키)을 남긴다. 새로고침하거나 탭을
닫았다 열어도 그동안은 다시 묻지 않는다. 통행증은 만료 시각에 핀으로 서명한 값이라
핀을 바꾸면 전부 무효가 된다. 핀 자체는 브라우저에 남지 않는다.
"""
from __future__ import annotations

import hashlib
import hmac
import os
import time
import unicodedata

import streamlit as st

틀린횟수 = "_출입_틀린횟수"
잠김해제 = "_출입_잠김해제"
통과 = "_출입_통과"
심을것 = "_출입_심을통행증"      # 다음 화면에서 브라우저에 남길 값("" 이면 지운다)
나감 = "_출입_나감"              # 이 접속에서는 통행증을 믿지 않는다

봐주는횟수 = 5          # 이 횟수까지는 그냥 다시 물어본다
잠그는초 = 60           # 그 뒤로는 틀릴 때마다 이만큼씩 잠근다(최대 10분)
최대잠금 = 600
유지초 = 8 * 60 * 60     # 한 번 통과하면 이 시간 동안 다시 묻지 않는다
쿠키이름 = "incentive_pass"


def 핀() -> str:
    v = os.environ.get("APP_PIN", "")
    if v:
        return v
    try:
        return str(st.secrets.get("app", {}).get("pin", ""))
    except Exception:                      # noqa: BLE001  금고가 없는 환경
        return ""


def _다듬기(값: str) -> str:
    """앞뒤 공백을 떼고 전각 글자(１２３ 등)를 보통 글자로 바꾼다 — 눈에는 같은데
    글자가 달라 안 맞는 일을 막는다."""
    return unicodedata.normalize("NFKC", str(값)).strip()


def _맞나(입력: str) -> bool:
    # 글자 수로 정답을 짐작하지 못하게 상수 시간 비교를 쓴다.
    # 한글이 섞인 핀도 되도록 바이트로 바꿔 비교한다(문자열 비교는 영문·숫자만 받는다).
    return hmac.compare_digest(_다듬기(입력).encode("utf-8"),
                               _다듬기(핀()).encode("utf-8"))


def _서명(만료: int) -> str:
    열쇠 = hashlib.sha256(("통행증:" + _다듬기(핀())).encode("utf-8")).digest()
    return hmac.new(열쇠, str(만료).encode("ascii"), hashlib.sha256).hexdigest()


def _통행증() -> str:
    만료 = int(time.time()) + 유지초
    return f"{만료}.{_서명(만료)}"


def _통행증_맞나(값: str) -> bool:
    try:
        만료글, 서명 = str(값).split(".", 1)
        만료 = int(만료글)
    except ValueError:
        return False
    return 만료 > time.time() and hmac.compare_digest(서명, _서명(만료))


def _브라우저_통행증() -> str:
    try:
        return str(st.context.cookies.get(쿠키이름, ""))
    except Exception:                      # noqa: BLE001  쿠키를 못 읽는 환경
        return ""


def _심기() -> None:
    """통과 직후(또는 나간 직후) 한 번, 브라우저에 통행증을 남기거나 지운다."""
    if 심을것 not in st.session_state:
        return
    값 = st.session_state.pop(심을것)
    수명 = 유지초 if 값 else 0
    import streamlit.components.v1 as components
    components.html(
        "<script>try{window.parent.document.cookie="
        f"'{쿠키이름}={값}; max-age={수명}; path=/; SameSite=Strict; Secure';"
        "}catch(e){}</script>", height=0)


def 열렸나() -> bool:
    """핀을 통과했으면 True. 아니면 입력 화면을 그리고 False."""
    if not 핀():
        return True                        # 핀 미설정 — 문을 세우지 않는다
    if st.session_state.get(통과):
        _심기()
        return True
    if not st.session_state.get(나감) and _통행증_맞나(_브라우저_통행증()):
        st.session_state[통과] = True   # 8시간 안에 다시 온 브라우저
        return True
    _심기()                              # 나간 직후라면 브라우저의 통행증을 지운다

    남은 = int(st.session_state.get(잠김해제, 0) - time.time())
    st.markdown("<div style='height:8vh'></div>", unsafe_allow_html=True)
    _, 가운데, _ = st.columns([1, 1.1, 1])
    with 가운데:
        st.markdown(
            "<div style='text-align:center;margin-bottom:18px'>"
            "<div style='width:44px;height:44px;border-radius:12px;background:#1C6B45;"
            "color:#fff;font-weight:800;font-size:21px;display:flex;align-items:center;"
            "justify-content:center;margin:0 auto 12px'>제</div>"
            "<div style='font-size:1.18rem;font-weight:800;color:#1F2937'>"
            "백화점 인센티브 마감</div></div>",
            unsafe_allow_html=True)

        if 남은 > 0:
            st.error(f"여러 번 틀려서 **{남은}초** 동안 잠겼습니다. "
                     "잠시 뒤 다시 시도하세요.")
            st.stop()

        with st.form("출입", clear_on_submit=True):
            값 = st.text_input("핀번호", type="password",
                              placeholder="핀번호를 입력하세요")
            보냄 = st.form_submit_button("들어가기", type="primary",
                                       use_container_width=True)

        if 보냄:
            if _맞나(값):
                st.session_state[통과] = True
                st.session_state[틀린횟수] = 0
                st.session_state[심을것] = _통행증()
                st.session_state.pop(나감, None)
                st.rerun()
            n = st.session_state.get(틀린횟수, 0) + 1
            st.session_state[틀린횟수] = n
            if n > 봐주는횟수:
                벌 = min(잠그는초 * (n - 봐주는횟수), 최대잠금)
                st.session_state[잠김해제] = time.time() + 벌
                st.error(f"핀번호가 맞지 않습니다. {벌}초 동안 잠깁니다.")
            else:
                # 한 번 틀릴 때마다 조금씩 느리게 — 자동 대입을 쓸모없게 만든다.
                time.sleep(min(0.4 * n, 2.0))
                st.error(f"핀번호가 맞지 않습니다. ({n}/{봐주는횟수})")

        st.markdown(
            "<div style='text-align:center;color:#6B7280;font-size:.82rem;"
            "margin-top:14px;line-height:1.7'>사내 인센티브 마감 자료입니다.<br>"
            "핀번호를 외부에 공유하지 마세요.</div>",
            unsafe_allow_html=True)
    return False


def 나가기() -> None:
    for k in (통과, 틀린횟수, 잠김해제):
        st.session_state.pop(k, None)
    st.session_state[나감] = True
    st.session_state[심을것] = ""       # 브라우저의 통행증도 지운다
