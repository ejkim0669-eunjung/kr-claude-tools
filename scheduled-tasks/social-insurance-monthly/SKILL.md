---
name: social-insurance-monthly
description: 매월 21일 건강보험·국민연금·고용산재 파이썬 스크립트를 실행해 사업장별 PDF/엑셀 저장
---

사회보험료(건강보험·국민연금·고용산재) 월별 조회를 파이썬 스크립트로 실행해줘.

(이 PC(DESKTOP-9VHJIM0)에서 매월 21~24일 오전 10시(한국 시간)에 실행되는 예약 작업. 21일에 처음 실행되고, 휴무일이거나 실패하면 다음날 다시 실행된다. 사용자는 자리에 없을 수 있으니 질문하지 말고 합리적으로 판단해 진행하고, 마지막에 결과와 확인 필요 사항을 짧게 정리할 것. 비밀번호·공동인증서 입력과 로그인은 하지 말 것. 개인정보는 답변에 적지 말 것. 이메일 등 메시지는 보내지 말 것.)

[0. 실행 여부 판단 — 가장 먼저]
- 완료 기록 파일 C:\Users\admin\.claude\scheduled-tasks\social-insurance-monthly\done_YYYY-MM.txt (YYYY-MM = 이번 달)가 있으면 "이번 달 사회보험료는 이미 저장됨"이라고 알리고 종료.
- 오늘(한국 시간)이 토·일요일 또는 대한민국 공휴일(대체공휴일 포함)이면 작업하지 말고 "휴무일이라 내일 10시에 다시 실행합니다"라고 알리고 종료.

[1. 사전 조건 확인]
- 크롬이 --remote-debugging-port=9333 으로 떠 있는지 확인한다. 안 떠 있으면 다음 명령으로 띄운다:
  start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9333 --user-data-dir=C:\hometax-axtax\chrome-profile --disable-features=LocalNetworkAccessChecks --new-window https://hometax.go.kr/
- 3개 EDI 사이트에 로그인이 필요하다. 로그인 화면이 보이면 로그인하지 말고 "로그인 필요"로 기록한 뒤 다른 사이트를 계속 진행한다.
  - edi.nhis.or.kr (건강보험)
  - edi.nps.or.kr (국민연금)
  - total.comwel.or.kr (고용산재)

[2. 스크립트 순차 실행]
작업 디렉토리 C:\hometax-axtax 에서 아래 3개 파이썬 스크립트를 순서대로 실행한다.
통합 실행 스크립트도 있지만(social_insurance_axtax.py), 개별 실행해서 각 결과를 확인하는 것이 좋다.

PowerShell로 실행:
  Set-Location C:\hometax-axtax; python -I nhis_edi_axtax.py
  Set-Location C:\hometax-axtax; python -I nps_edi_axtax.py
  Set-Location C:\hometax-axtax; python -I comwel_gy_axtax.py

각 스크립트 설명:
1) nhis_edi_axtax.py — 건강보험 EDI(edi.nhis.or.kr) 가입자고지내역서 PDF 다운로드 (24개 사업장)
2) nps_edi_axtax.py — 국민연금 EDI(edi.nps.or.kr) 결정내역 PDF 다운로드 (24개 사업장)
3) comwel_gy_axtax.py — 고용산재 EDI(total.comwel.or.kr) 고용보험료·산재보험료·지원금 엑셀 다운로드 (15개 사업장)

★ 각 스크립트는 Playwright로 127.0.0.1:9333 에 연결해 작동한다. localhost를 쓰면 안 된다.
★ 각 스크립트 사이에 10초 이상 간격을 둔다.

[3. 저장 위치]
모든 파일은 C:\Users\admin\Desktop\사회보험\{사업장명}\ 폴더에 저장된다 (스크립트 내부에서 자동 생성).

[4. 예외 처리]
- 스크립트가 비정상 종료(exit code ≠ 0)하면 오류 메시지를 기록하고 다음 스크립트를 계속 진행한다.
- 로그인 만료로 실패한 경우 "로그인 필요"로 기록한다.
- 조회·저장만 할 것. 신고·신청·수정 등 다른 동작은 하지 말 것.

[5. 마무리]
- 3개 스크립트가 모두 성공했으면 done_YYYY-MM.txt를 만들고 결과를 적는다. 하나라도 실패하면 만들지 않는다(다음날 다시 실행, 24일이 마지막).
- 보고: 각 스크립트 실행 결과(성공/실패), 저장된 파일 수, 로그인 필요·실패 사이트와 이유.

[제외 사업장 — 스크립트 내부에 이미 반영됨]
도화구조, 피티솔루션, 도화기술, 도화건축사사무소 4곳은 제외 (스크립트 FIRMS 목록에 없음).