#!/usr/bin/env python -I
# -*- coding: utf-8 -*-
"""사회보험 통합 조회 — 건강보험 + 국민연금 + 고용산재 순차 실행.

사용법:
  1) 크롬을 디버그 포트로 띄운다:
     start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" ^
       --remote-debugging-port=9333 --user-data-dir=C:\hometax-axtax\chrome-profile ^
       --disable-features=LocalNetworkAccessChecks --new-window https://hometax.go.kr/
  2) 세 EDI 사이트에 각각 로그인한다 (사람이 직접):
     - edi.nhis.or.kr    (건강보험)
     - edi.nps.or.kr     (국민연금)
     - total.comwel.or.kr (고용산재)
  3) 이 스크립트를 실행한다:
     python -I social_insurance_axtax.py
"""
import subprocess, sys, os, time

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))

SCRIPTS = [
    ("건강보험 (NHIS)",  "nhis_edi_axtax.py"),
    ("국민연금 (NPS)",   "nps_edi_axtax.py"),
    ("고용산재 (Comwel)", "comwel_gy_axtax.py"),
]

def run_script(label, filename):
    path = os.path.join(SCRIPTS_DIR, filename)
    if not os.path.exists(path):
        print(f"  파일 없음: {path}")
        return False

    print(f"\n{'='*50}")
    print(f"  {label}  시작")
    print(f"{'='*50}\n")

    result = subprocess.run(
        [sys.executable, "-I", path],
        cwd=SCRIPTS_DIR,
    )

    if result.returncode != 0:
        print(f"\n  !! {label} 비정상 종료 (코드 {result.returncode})")
        return False
    return True

if __name__ == "__main__":
    print("사회보험 통합 조회")
    print(f"저장 위치: C:\\Users\\admin\\Desktop\\사회보험\\{{회사명}}\\\n")

    success = []
    fail = []

    for label, script in SCRIPTS:
        ok = run_script(label, script)
        if ok:
            success.append(label)
        else:
            fail.append(label)

    print(f"\n{'='*50}")
    print(f"전체 완료")
    if success:
        print(f"  성공: {', '.join(success)}")
    if fail:
        print(f"  실패: {', '.join(fail)}")
    print(f"저장 위치: C:\\Users\\admin\\Desktop\\사회보험\\")
