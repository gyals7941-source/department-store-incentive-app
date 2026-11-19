# -*- coding: utf-8 -*-
"""화면 껍데기 — 색·카드·헤더. 계산 로직은 core/ 에만 있다."""
from __future__ import annotations

import streamlit as st

# 한글·숫자를 한 벌로 가진 글꼴. 고정폭을 쓰지 않고 '표 숫자'로 자릿수를 맞춘다.
# (Consolas 같은 고정폭에는 한글이 없어 한 줄 안에서 글꼴이 갈라진다.)
FONT_CSS = ("https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9"
            "/dist/web/variable/pretendardvariable-dynamic-subset.css")

CSS = f"""
<style>
  @import url("{FONT_CSS}");
  :root{{
    --g1:#1C6B45; --g2:#2E8B5F; --g-bg:#E9F4EC; --g-bd:#CBE3D4;
    --ink:#1F2937; --mute:#6B7280; --line:#E5E7EB; --box:#F4F5F6;
    --ui:"Pretendard Variable", Pretendard, -apple-system, "Segoe UI",
         "Malgun Gothic", sans-serif;
  }}

  /* 전체 글꼴 한 벌로. 숫자는 폭이 고른 '표 숫자'. */
  html, body, .stApp, [data-testid="stAppViewContainer"],
  button, input, select, textarea, [data-baseweb]{{
      font-family:var(--ui) !important}}
  .num, .pill b, .card .v, .filenote, .runbox, .sec,
  [data-testid="stMetricValue"], [data-testid="stDataFrame"]{{
      font-variant-numeric:tabular-nums; font-feature-settings:"tnum" 1}}
  #MainMenu, footer, [data-testid="stToolbar"], [data-testid="stDecoration"],
  [data-testid="stStatusWidget"]{{display:none !important}}
  header[data-testid="stHeader"]{{height:0;background:transparent}}
  .block-container{{padding:1.1rem 2.2rem 4rem;max-width:1320px}}

  /* ── 상단 바 ── */
  .topbar{{display:flex;align-items:center;gap:12px;padding:4px 0 14px}}
  .logo{{width:30px;height:30px;border-radius:8px;background:var(--g1);color:#fff;
        font-weight:800;font-size:15px;display:flex;align-items:center;
        justify-content:center;flex:none}}
  .topttl{{font-size:1.12rem;font-weight:800;color:var(--ink);letter-spacing:-.01em}}
  .spacer{{flex:1}}
  .pill{{background:var(--box);border:1px solid var(--line);border-radius:999px;
        padding:6px 15px;font-size:.84rem;color:var(--mute);white-space:nowrap}}
  .pill b{{color:var(--ink);font-weight:700;margin-left:6px}}

  /* ── 탭 ── */
  [data-testid="stTabs"] [data-baseweb="tab-list"]{{
      gap:26px;border-bottom:1px solid var(--line);padding:0}}
  [data-testid="stTabs"] [data-baseweb="tab"]{{
      padding:9px 0;font-size:.95rem;font-weight:600;color:var(--mute)}}
  [data-testid="stTabs"] [aria-selected="true"]{{color:var(--g1)}}
  [data-testid="stTabs"] [data-baseweb="tab-highlight"]{{background:var(--g1);height:2px}}
  [data-testid="stTabs"] [data-baseweb="tab-border"]{{display:none}}

  /* ── 섹션 ── */
  .sec{{font-size:1.32rem;font-weight:800;color:var(--ink);margin:22px 0 6px}}
  .lead{{color:var(--mute);font-size:.9rem;line-height:1.65;margin-bottom:16px;
        max-width:640px}}
  .muted{{color:var(--mute);font-size:.84rem}}

  /* ── 업로드 카드 ── */
  [data-testid="stFileUploader"]{{
      background:var(--g-bg);border:1px solid var(--g-bd);border-radius:14px;
      padding:16px 16px 12px;text-align:center}}
  [data-testid="stFileUploader"] > label{{
      width:100%;display:block;text-align:center;margin-bottom:2px}}
  [data-testid="stFileUploader"] > label p{{
      font-size:.98rem !important;font-weight:700;color:var(--ink);margin:0}}
  [data-testid="stFileUploaderDropzone"]{{
      background:transparent;border:none;min-height:0;padding:4px 0 0;
      justify-content:center}}
  [data-testid="stFileUploaderDropzoneInstructions"]{{
      flex-direction:column;gap:2px;align-items:center;color:var(--mute)}}
  [data-testid="stFileUploaderDropzoneInstructions"] span{{display:none}}
  [data-testid="stFileUploaderDropzoneInstructions"] small{{
      font-size:.78rem;color:var(--mute)}}
  [data-testid="stFileUploader"] button{{
      background:#fff;border:1px solid var(--g-bd);color:var(--g1);
      font-weight:600;border-radius:8px}}
  [data-testid="stFileUploaderFile"]{{background:#fff;border-radius:8px;
      padding:6px 10px;margin-top:8px}}

  .filenote{{text-align:center;font-size:.83rem;font-weight:700;color:var(--g1);
            margin:-6px 0 2px}}
  .filenote.dim{{color:var(--mute);font-weight:400}}

  /* ── 정보 카드 ── */
  .cards{{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin:16px 0 4px}}
  .card{{border:1px solid var(--line);border-radius:12px;padding:13px 16px 15px}}
  .card .k{{color:var(--mute);font-size:.8rem;margin-bottom:9px}}
  .card .v{{font-size:1.04rem;font-weight:700;color:var(--ink);
           letter-spacing:-.005em}}
  .card.warn{{border-color:#F0D9A8;background:#FEFBF3}}

  /* ── 실행 요약 ── */
  .runbox{{background:var(--box);border-radius:12px;padding:15px 18px;
          font-size:.86rem;line-height:2.0;color:#374151;
          letter-spacing:-.003em}}
  .runbox b{{color:var(--ink)}}
  .runbox .ok{{color:var(--g1);font-weight:700}}
  .runbox .no{{color:#B45309;font-weight:700}}
  .runbox .hd{{color:var(--ink);font-weight:700}}

  /* ── 지표 ── */
  [data-testid="stMetric"]{{border:1px solid var(--line);border-radius:12px;
      padding:11px 15px 13px}}
  [data-testid="stMetricLabel"] p{{font-size:.79rem !important;color:var(--mute)}}
  [data-testid="stMetricValue"]{{font-size:1.32rem;font-weight:800;
      letter-spacing:-.015em;color:var(--ink)}}

  /* ── 버튼 ── */
  .stButton>button, .stDownloadButton>button{{border-radius:9px;font-weight:700;
      padding:.5rem 1.15rem;border:1px solid var(--line)}}
  .stButton>button[kind="primary"]{{background:var(--g1);border-color:var(--g1)}}
  .stButton>button[kind="primary"]:hover{{background:#17583a;border-color:#17583a}}

  /* ── 차수 배지·막대 ── */
  .tag{{display:inline-block;padding:2px 10px;border-radius:999px;font-size:.76rem;
       font-weight:700;line-height:1.7;vertical-align:2px}}
  .t3{{background:#15532F;color:#fff}}.t2{{background:var(--g1);color:#fff}}
  .t1{{background:#5FA47B;color:#fff}}.t0{{background:#EEF1EF;color:#8B9A92}}
  .bar{{height:8px;border-radius:5px;background:#EEF1EF;overflow:hidden;margin-top:7px;
       max-width:420px}}
  .bar>div{{height:100%;border-radius:5px}}

  /* ── 표 ── */
  [data-testid="stDataFrame"]{{border-radius:10px}}
  hr{{margin:1.1rem 0;border-color:var(--line)}}
</style>
"""


def apply() -> None:
    st.markdown(CSS, unsafe_allow_html=True)


def topbar(마감월: str, 부제: str = "") -> None:
    sub = f"<span class='pill'>{부제}</span>" if 부제 else ""
    st.markdown(
        f"<div class='topbar'><div class='logo'>제</div>"
        f"<div class='topttl'>백화점 인센티브 마감</div><div class='spacer'></div>"
        f"<span class='pill'>마감월<b>{마감월}</b></span>{sub}</div>",
        unsafe_allow_html=True)


def section(title: str, lead: str = "") -> None:
    st.markdown(f"<div class='sec'>{title}</div>"
                + (f"<div class='lead'>{lead}</div>" if lead else ""),
                unsafe_allow_html=True)


def cards(items: list[tuple[str, str, bool]]) -> None:
    """[(제목, 값, 경고여부)] 를 3열 카드로."""
    html = "".join(
        f"<div class='card{' warn' if w else ''}'><div class='k'>{k}</div>"
        f"<div class='v'>{v}</div></div>" for k, v, w in items)
    st.markdown(f"<div class='cards'>{html}</div>", unsafe_allow_html=True)


def runbox(lines: list[str]) -> None:
    st.markdown("<div class='runbox'>" + "<br>".join(lines) + "</div>",
                unsafe_allow_html=True)
