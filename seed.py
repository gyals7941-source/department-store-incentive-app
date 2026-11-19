# -*- coding: utf-8 -*-
"""최초 1회 — 기존 엑셀에서 이월명부와 지난 마감이력을 서버로 옮긴다.

    python seed.py <이월명부를_가져올_최종파일.xlsx> [지난마감파일 ...]

이월명부는 첫 번째 파일에서 가져오고, 마감이력은 각 파일의 '영업팀공유' 시트에
적힌 값을 그대로(재계산하지 않고) 기록한다. 실제로 지급된 숫자가 이력이어야 하기 때문이다.
"""
from __future__ import annotations

import sys
import warnings
from datetime import datetime
from pathlib import Path

warnings.filterwarnings("ignore")
sys.path.insert(0, str(Path(__file__).resolve().parent))

from core import storage, xlsm_reader  # noqa: E402


def main(argv: list[str]) -> int:
    if not argv:
        print(__doc__)
        return 2

    주파일 = argv[0]
    d = xlsm_reader.load(주파일)
    carry = d["carryover"]
    storage.save_carryover(carry)
    대기 = (carry["지급상태"].astype(str).str.strip() == "대기").sum() if len(carry) else 0
    print(f"이월명부 {len(carry)}건 저장 (대기 {대기}건)  ← {Path(주파일).name}")

    rows = []
    for p in argv:
        try:
            x = xlsm_reader.load(p)
        except Exception as e:                       # noqa: BLE001
            print(f"  ⚠ {Path(p).name} 건너뜀: {e}")
            continue
        s, 월 = x["exp_share"], x["settings"].마감월_text
        if s.empty:
            print(f"  ⚠ {Path(p).name} 영업팀공유가 비어 건너뜀")
            continue
        rows = [r for r in rows if r["마감월"] != 월]
        rows.append(dict(
            마감월=월, 회차=1,
            확정시각=datetime.fromtimestamp(Path(p).stat().st_mtime)
                      .strftime("%Y-%m-%d %H:%M"),
            정상가합계=float(s["정상가합계"].sum()), 건수=int(s["건수"].sum()),
            인센티브총액=float(s["인센티브총액"].sum()),
            당월지급=float(s["당월지급"].sum()),
            이월건수=int(s["이월건수"].sum()), 이월금액=float(s["이월금액"].sum()),
            적용수당표=월, 입력요약={"원본": len(x["raw"]), "배송대장": len(x["delivery"])},
            메모=f"기존 엑셀에서 이관 ({Path(p).name}, 템플릿 {x['버전']})",
            지점별=[{"지점": r["지점"], "정상가합계": float(r["정상가합계"]),
                   "차수": str(r["차수"]), "건수": int(r["건수"]),
                   "인센티브총액": float(r["인센티브총액"]),
                   "당월지급": float(r["당월지급"]),
                   "이월건수": int(r["이월건수"]), "이월금액": float(r["이월금액"])}
                  for _, r in s.iterrows()],
        ))
        print(f"마감이력 {월}  인센티브 {s['인센티브총액'].sum():,.0f} · "
              f"당월지급 {s['당월지급'].sum():,.0f}  ← {Path(p).name}")

    rows.sort(key=lambda r: r["마감월"])
    storage.save_history(rows)
    print(f"\n완료 — 이월명부 {len(carry)}건 · 마감이력 {len(rows)}개월")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
