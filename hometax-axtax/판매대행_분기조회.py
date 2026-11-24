"""
홈택스 신용카드/판매(결제)대행 매출자료 조회 - 거래처별 분기 조회 후 바탕화면 저장

흐름: 9333 크롬에 붙기 -> 로그인 확인 -> 판매대행 조회 화면 확인 ->
      거래처별로 사업자번호 입력 -> 결제년도/분기 선택 -> 조회 -> 결과 표 읽기 -> CSV 저장

규칙 (C:\\hometax-axtax\\CLAUDE.md 기준):
- 로그인은 사람이 한다. 이 코드는 아이디/비밀번호/인증서를 다루지 않는다.
- 로그인 화면이 보이면 즉시 멈춘다.
- 조회만 한다. 신고/제출/발급 화면은 열지 않는다.
- 거래처 사이 2~5초 쉰다.
- 확인되지 않은 것은 값을 내놓지 않고 멈춘다.
- 금액은 콘솔에 출력하지 않고 파일로만 저장한다.

사용 예:
    python 판매대행_분기조회.py --csv 판매대행매출_조회대상.csv --year 2026 --quarter 3분기

※ 아직 확인되지 않은 칸은 TODO로 표시했다. 실행 전에 화면지도(판매대행)를 채워야 한다.
"""

import argparse
import csv
import os
import random
import re
import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

CDP_URL = "http://127.0.0.1:9333"
DESKTOP = Path(os.path.expanduser("~")) / "Desktop"

# 화면 제목에 이 문구가 있어야 올바른 화면이다 (WebSquare 제목 기준, 확인됨)
TARGET_TITLE_KEYWORD = "판매(결제)대행"

# 확인됨: 접근성 트리에 'textbox "사업자등록번호 10자리"' 로 보였다.
BIZ_INPUT_NAME = "사업자등록번호 10자리"

# TODO: 아래 항목은 아직 화면에서 확인하지 못했다. 실행 전에 채워야 한다.
QUERY_BUTTON_NAME = "조회"          # TODO: 실제 버튼 이름 확인
YEAR_SELECT_NAME = "결제년도"        # TODO: 실제 선택칸 확인
RESULT_TABLE_TEXT = "판매(결제)대행 매출자료"  # TODO: 결과 표 위치 확인

# 짐작 (미확인): 2기 예정 = 3분기 등 기간 대응. 화면에서 확인되면 수정할 것.
PERIOD_MAP = {"1분기": "1분기", "2분기": "2분기", "3분기": "3분기", "4분기": "4분기"}


class StopForSafety(Exception):
    """확인이 안 되거나 위험하면 멈춘다. 이 예외가 나오면 값을 저장하지 않는다."""


def log(msg):
    print(msg, flush=True)


def find_hometax_page(browser):
    for ctx in browser.contexts:
        for pg in ctx.pages:
            if "hometax" in pg.url:
                return pg
    return None


def check_login_and_screen(page):
    # 로그인 화면이면 멈춘다 (비밀번호 칸이 보이면 로그인 화면으로 본다)
    has_pw = page.evaluate(
        """() => Array.from(document.querySelectorAll('input[type="password"]'))
            .some(e => e.offsetWidth > 0 && e.offsetHeight > 0)"""
    )
    if has_pw:
        raise StopForSafety("로그인 화면이 보입니다. 사람이 로그인한 뒤 다시 실행하세요.")

    # '로그아웃' 글자 하나로만 판단하지 않는다 (CLAUDE.md). 한 번 더 잠깐 뒤 확인한다.
    logged_in = False
    for _ in range(2):
        logged_in = page.get_by_role("link", name="로그아웃").count() > 0
        if logged_in:
            break
        time.sleep(1.5)
    if not logged_in:
        raise StopForSafety("로그인 여부를 확신할 수 없어 멈춥니다.")

    title = page.title()
    if TARGET_TITLE_KEYWORD not in title:
        raise StopForSafety(f"판매(결제)대행 조회 화면이 아닙니다 (제목: {title!r}). 화면을 맞춘 뒤 다시 실행하세요.")


def query_one(page, biz_no, year, quarter):
    digits = re.sub(r"[^0-9]", "", biz_no)
    if len(digits) != 10:
        raise StopForSafety(f"사업자등록번호 형식이 이상합니다: {biz_no!r}")

    box = page.get_by_role("textbox", name=BIZ_INPUT_NAME)
    if box.count() != 1:
        raise StopForSafety("사업자등록번호 입력칸을 정확히 하나 찾지 못했습니다. 멈춥니다.")
    box.fill("")
    box.fill(digits)

    # TODO: 결제년도/분기 선택 방법 확인 후 채우기. 지금은 선택하지 않고 멈춘다.
    raise StopForSafety(
        "결제년도/분기 선택칸을 아직 코드에 맞게 확인하지 못했습니다 (TODO). "
        "화면지도를 채운 뒤 다시 실행하세요."
    )


def read_result_table(page):
    # TODO: 결과 표 구조 확인 후 실제 행/열 읽기로 바꾸기
    raise StopForSafety("결과 표 읽기는 아직 확인되지 않았습니다 (TODO).")


def save_csv(path, header, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def main():
    ap = argparse.ArgumentParser(description="판매(결제)대행 매출자료 거래처별 분기 조회")
    ap.add_argument("--csv", required=True, help="거래처 목록 CSV (사업자등록번호 열 포함)")
    ap.add_argument("--year", required=True, help="예: 2026")
    ap.add_argument("--quarter", required=True, help="예: 3분기")
    args = ap.parse_args()

    if args.quarter not in PERIOD_MAP:
        log(f"[중단] 분기 값을 모릅니다: {args.quarter}")
        sys.exit(1)

    with open(args.csv, encoding="utf-8-sig") as f:
        clients = list(csv.DictReader(f))
    out_dir = DESKTOP / f"판매대행매출_{args.year}_{args.quarter}"

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(CDP_URL)
        page = find_hometax_page(browser)
        if page is None:
            log("[중단] 홈택스 탭을 찾지 못했습니다. 9333 크롬에서 홈택스를 여세요.")
            sys.exit(1)

        try:
            check_login_and_screen(page)
            log("확인됨: 로그인돼 있고 판매대행 조회 화면이 맞습니다.")

            for i, c in enumerate(clients, 1):
                name = c.get("상호(성명)", "").strip() or f"거래처{i}"
                biz = c.get("사업자등록번호", "")
                log(f"[{i}/{len(clients)}] 조회 중...")
                query_one(page, biz, args.year, args.quarter)
                rows = read_result_table(page)
                safe_name = re.sub(r'[\\/:*?"<>|]', "_", name)
                save_csv(out_dir / f"{safe_name}.csv", ["승인년월", "건수", "매출금액", "판매대행업체"], rows)
                log(f"      저장 완료 (파일: {out_dir}). 금액은 파일에만 있습니다.")

                if i < len(clients):
                    time.sleep(random.uniform(2.0, 5.0))

        except StopForSafety as e:
            log(f"[중단] {e}")
            sys.exit(1)
        finally:
            browser.close()


if __name__ == "__main__":
    main()
