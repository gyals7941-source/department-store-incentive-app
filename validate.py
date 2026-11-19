# -*- coding: utf-8 -*-
"""기존 엑셀 결과와 1원 단위로 대조하는 회귀검증.

    python validate.py <마감파일.xlsm|xlsx> [...]
"""
from __future__ import annotations

import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
sys.path.insert(0, str(Path(__file__).resolve().parent))

import pandas as pd  # noqa: E402

from core import engine, xlsm_reader  # noqa: E402

OK, NG = "OK", "불일치"


def cmp_num(a, b, tol=0.5):
    return abs(float(a or 0) - float(b or 0)) <= tol


def validate(path) -> bool:
    d = xlsm_reader.load(path)
    st = d["settings"]
    res = engine.run(d["raw"], d["delivery"], st,
                     carryover=d["carryover"], 지점목록=d["지점목록"])
    got, exp = res["영업팀공유"], d["exp_share"]

    ver = d.get("버전", "?")
    print("=" * 78)
    print(f"■ {Path(path).name}   마감월 {st.마감월_text}   템플릿 {ver}")
    if ver == "v1":
        print("  ⚠ 구버전 템플릿(_배송인덱스·취소상계·익월배송 없음) — 최신 스펙과 계산식이 다릅니다.")
        print("    대조 대상에서 제외합니다.")
        print()
        return None
    print(f"  단가표 {len(st.단가표)}품목 / 온라인전용 {sorted(st.온라인전용) or '없음'}")
    print(f"  원본 {len(d['raw'])}행 · 배송원장 {len(d['delivery'])}행 · 이월명부 {len(d['carryover'])}건")
    print("-" * 78)

    cols = ["정상가합계", "건수", "인센티브총액", "당월지급", "이월건수", "이월금액"]
    hdr = f"{'지점':<22}{'차수':>6}" + "".join(f"{c:>14}" for c in cols)
    print(hdr)
    bad = 0
    for _, e in exp.iterrows():
        g = got[got["지점"] == e["지점"]]
        if g.empty:
            print(f"{e['지점']:<22}  ← 엔진 결과에 없음"); bad += 1; continue
        g = g.iloc[0]
        line = f"{e['지점']:<22}{str(g['차수']):>6}"
        row_bad = str(g["차수"]) != str(e["차수"])
        for c in cols:
            ok = cmp_num(g[c], e[c])
            row_bad |= not ok
            line += f"{g[c]:>14,.0f}" if ok else f"{'✗' + format(g[c], ',.0f'):>14}"
        print(line)
        if row_bad:
            bad += 1
            for c in cols:
                if not cmp_num(g[c], e[c]):
                    print(f"{'':<22}   · {c}: 엔진 {g[c]:,.0f}  ≠  엑셀 {e[c]:,.0f}"
                          f"   (차이 {g[c]-e[c]:+,.0f})")
            if str(g["차수"]) != str(e["차수"]):
                print(f"{'':<22}   · 차수: 엔진 {g['차수']}  ≠  엑셀 {e['차수']}")

    tot_g = {c: got[c].sum() for c in cols}
    tot_e = {c: exp[c].sum() for c in cols}
    print("-" * 78)
    print(f"{'합계':<22}{'':>6}" + "".join(f"{tot_g[c]:>14,.0f}" for c in cols))
    print(f"{'(엑셀)':<22}{'':>6}" + "".join(f"{tot_e[c]:>14,.0f}" for c in cols))
    tot_ok = all(cmp_num(tot_g[c], tot_e[c]) for c in cols)
    print()
    print(f"  결과: 지점 {len(exp)-bad}/{len(exp)} 일치 · 합계 {OK if tot_ok else NG}")
    return bad == 0 and tot_ok


if __name__ == "__main__":
    files = sys.argv[1:]
    if not files:
        print(__doc__); sys.exit(2)
    results = [(f, validate(f)) for f in files]
    print("=" * 78)
    for f, ok in results:
        mark = "⏭ 건너뜀(구버전)" if ok is None else ("✅ 일치" if ok else "❌ 불일치")
        print(f"  {mark}  {Path(f).name}")
    checked = [ok for _, ok in results if ok is not None]
    print(f"\n  대조 {len(checked)}건 중 {sum(1 for x in checked if x)}건 일치")
    sys.exit(0 if all(checked) else 1)
