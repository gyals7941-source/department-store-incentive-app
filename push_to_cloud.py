# -*- coding: utf-8 -*-
"""한 번만 — 로컬 data/ 폴더의 내용을 데이터 저장소로 올린다.

    python push_to_cloud.py          # 무엇이 올라갈지 보기만 한다
    python push_to_cloud.py --실행    # 실제로 올린다

접속 정보는 .streamlit/secrets.toml 의 [github] 또는 환경변수에서 읽는다.
이미 같은 파일이 저장소에 있으면 덮어쓰며, 그 내용은 커밋 이력에 남는다.
"""
from __future__ import annotations

import os
import sys
import tomllib
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

문서 = ["이월명부.csv", "마감이력.json", "수당표.json", "거래처.json"]


def 금고읽기() -> None:
    """secrets.toml 이 있으면 환경변수로 옮겨 둔다(여긴 Streamlit 이 아니다)."""
    p = ROOT / ".streamlit" / "secrets.toml"
    if not p.exists():
        return
    with open(p, "rb") as f:
        d = tomllib.load(f)
    gh = d.get("github", {})
    for 환경, 키 in (("GH_DATA_REPO", "repo"), ("GH_DATA_TOKEN", "token"),
                    ("GH_DATA_EXPIRES", "expires")):
        if gh.get(키) and not os.environ.get(환경):
            os.environ[환경] = str(gh[키])


def main(argv: list[str]) -> int:
    금고읽기()
    from core import backend

    if not backend.cloud():
        print("데이터 저장소 접속 정보가 없습니다.\n"
              "  .streamlit/secrets.toml 에 [github] repo / token 을 넣거나\n"
              "  GH_DATA_REPO / GH_DATA_TOKEN 환경변수를 설정하세요.")
        return 2

    print(f"대상: {backend.어디()}\n")
    실행 = "--실행" in argv or "--run" in argv
    로컬 = ROOT / "data"

    for name in 문서:
        p = 로컬 / name
        if not p.exists():
            print(f"  · {name:<14} 로컬에 없음 — 건너뜀")
            continue
        body = p.read_text(encoding="utf-8-sig")
        표시 = "덮어씀" if backend.read_text(name) is not None else "새로 올림"
        print(f"  · {name:<14} {len(body):>8,}자  {표시}")
        if 실행:
            backend.write_text(name, body)

    snaps = sorted((로컬 / "마감본").glob("*.xlsx")) if (로컬 / "마감본").exists() else []
    for p in snaps:
        print(f"  · 마감본 {p.name}  {p.stat().st_size:,}바이트")
        if 실행:
            backend.write_snapshot(p.name, p.read_bytes())

    if not 실행:
        print("\n보기만 했습니다. 실제로 올리려면 --실행 을 붙이세요.")
    else:
        print("\n완료 — 이제 앱이 데이터 저장소를 바라봅니다.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
