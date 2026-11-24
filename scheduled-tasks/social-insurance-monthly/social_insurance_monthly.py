#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
social_insurance_monthly.py
매월 사회보험료(건강보험·국민연금) 개인별 내역을 사업장별로 조회해 PDF로 저장.

사용법:
  python social_insurance_monthly.py                        # 기존 Chrome(9333)에 붙어 실행
  python social_insurance_monthly.py --force                # 휴무일/done 체크 건너뜀
  python social_insurance_monthly.py --site nhis            # 건강보험만
  python social_insurance_monthly.py --site nps             # 국민연금만
  python social_insurance_monthly.py --screenshots          # 각 단계 스크린샷
  python social_insurance_monthly.py --wait-login 180       # 로그인 대기 180초

사전 조건:
  - Chrome이 --remote-debugging-port=9333 으로 띄워져 있어야 한다
  - 각 사이트(edi.nhis.or.kr, edi.nps.or.kr)에 사람이 로그인한 상태여야 한다
  - 외부 패키지: playwright (브라우저 내려받기 없이 기존 Chrome에 붙는다)
"""

import argparse
import base64
import datetime
import json
import logging
import os
import random
import re
import sys
import time
from pathlib import Path
from typing import Optional

from playwright.sync_api import (
    sync_playwright,
    Page,
    Browser,
    BrowserContext,
    CDPSession,
    TimeoutError as PwTimeout,
)

# ═══════════════════════════════════════════════════════════════
# 경로·상수
# ═══════════════════════════════════════════════════════════════
SCRIPT_DIR = Path(__file__).resolve().parent
SAVE_BASE = Path(r"C:\Users\admin\Desktop\건강보험개인별내역")
DONE_DIR = SCRIPT_DIR
LOG_FILE = SCRIPT_DIR / "run.log"
SCREENSHOT_DIR = SCRIPT_DIR / "screenshots"
CDP_URL = "http://127.0.0.1:9333"

HOLIDAYS: dict[int, set[tuple[int, int]]] = {
    2026: {
        (1,1),(1,26),(1,27),(1,28),(1,29),
        (3,1),(3,2),
        (5,5),(5,24),(5,25),
        (6,6),
        (8,15),(8,17),
        (9,24),(9,25),(9,26),
        (10,3),(10,5),(10,9),
        (12,25),
    },
    2027: {
        (1,1),
        (2,15),(2,16),(2,17),
        (3,1),
        (5,5),(5,13),
        (6,6),
        (8,15),(8,16),
        (10,3),(10,4),(10,5),(10,6),(10,9),
        (12,25),
    },
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(
            open(sys.stdout.fileno(), mode="w", encoding="utf-8",
                 closefd=False, buffering=1)
        ),
    ],
)
log = logging.getLogger("social_insurance")


# ═══════════════════════════════════════════════════════════════
# 유틸리티
# ═══════════════════════════════════════════════════════════════

def now_kst() -> datetime.datetime:
    try:
        import zoneinfo
        return datetime.datetime.now(zoneinfo.ZoneInfo("Asia/Seoul"))
    except Exception:
        return datetime.datetime.now() + datetime.timedelta(hours=9)


def is_holiday(dt: datetime.date) -> bool:
    if dt.weekday() >= 5:
        return True
    return (dt.month, dt.day) in HOLIDAYS.get(dt.year, set())


def done_path(dt: datetime.date) -> Path:
    return DONE_DIR / f"done_{dt.year}-{dt.month:02d}.txt"


def query_delay():
    """조회 사이에 2~5초 쉰다 (사이트 보호)."""
    time.sleep(random.uniform(2.0, 5.0))


def safe_filename(name: str) -> str:
    """파일명에 쓸 수 없는 문자 제거."""
    return re.sub(r'[\\/:*?"<>|]', '', name).strip()


def cdp_print_pdf(cdp: CDPSession, path: Path) -> bool:
    """CDP Page.printToPDF — headed 모드에서도 동작."""
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        log.info(f"  이미 존재 → 건너뜀: {path.name}")
        return True
    try:
        result = cdp.send("Page.printToPDF", {
            "landscape": False,
            "printBackground": True,
            "preferCSSPageSize": True,
        })
        pdf_bytes = base64.b64decode(result["data"])
        if len(pdf_bytes) < 5000:
            log.warning(f"  PDF가 너무 작음 ({len(pdf_bytes)} bytes) — 빈 페이지일 수 있음")
        path.write_bytes(pdf_bytes)
        log.info(f"  저장 완료: {path.name} ({len(pdf_bytes):,} bytes)")
        return True
    except Exception as e:
        log.error(f"  PDF 저장 실패: {e}")
        return False


def screenshot(page: Page, name: str, do_save: bool) -> None:
    if not do_save:
        return
    SCREENSHOT_DIR.mkdir(exist_ok=True)
    fp = SCREENSHOT_DIR / f"{name}.png"
    try:
        page.screenshot(path=str(fp), full_page=False)
        log.info(f"  [스크린샷] {fp.name}")
    except Exception as e:
        log.warning(f"  스크린샷 실패: {e}")


# ═══════════════════════════════════════════════════════════════
# 로그인 판별
# ═══════════════════════════════════════════════════════════════

def is_logged_in(page: Page) -> bool:
    """'로그아웃' 링크/버튼이 보이면 로그인 상태.
    단, 화면이 새로 그려지는 중일 수 있으므로 작업 요소도 함께 본다."""
    for sel in [
        "a:has-text('로그아웃')",
        "button:has-text('로그아웃')",
        "[class*='logout']",
        "a[href*='logout']",
    ]:
        try:
            if page.locator(sel).first.is_visible(timeout=1000):
                return True
        except Exception:
            continue
    for frame in page.frames:
        if frame == page.main_frame:
            continue
        for sel in ["a:has-text('로그아웃')", "button:has-text('로그아웃')"]:
            try:
                if frame.locator(sel).first.is_visible(timeout=500):
                    return True
            except Exception:
                continue
    return False


def check_logged_in(page: Page) -> bool:
    """한 번 읽고 단정하지 말고 잠깐 뒤 다시 확인한다."""
    if is_logged_in(page):
        return True
    page.wait_for_timeout(2000)
    return is_logged_in(page)


def wait_for_login(page: Page, site_name: str, timeout: int = 180) -> bool:
    log.info(f"  {site_name} 로그인 대기 중 (최대 {timeout}초)...")
    deadline = time.time() + timeout
    while time.time() < deadline:
        if check_logged_in(page):
            log.info(f"  {site_name} 로그인 확인됨")
            return True
        time.sleep(3)
    log.warning(f"  {site_name} 로그인 대기 시간 초과")
    return False


# ═══════════════════════════════════════════════════════════════
# 사이트 결과
# ═══════════════════════════════════════════════════════════════

class Result:
    def __init__(self, name: str):
        self.name = name
        self.ok = False
        self.login_required = False
        self.error: Optional[str] = None
        self.files: list[str] = []
        self.skipped: list[str] = []

    def summary(self) -> str:
        if self.login_required:
            return f"[{self.name}] 로그인 필요"
        if self.error:
            return f"[{self.name}] 오류: {self.error}"
        parts = [f"[{self.name}] 저장 {len(self.files)}건"]
        if self.skipped:
            parts.append(f"건너뜀 {len(self.skipped)}건")
        return ", ".join(parts)


# ═══════════════════════════════════════════════════════════════
# 건강보험 EDI — 개인별보험료고지내역
# ═══════════════════════════════════════════════════════════════

def _nhis_get_workplaces(page: Page) -> list[dict]:
    """메인 페이지에서 수임 사업장 목록을 수집한다.
    반환: [{"name": 사업장명, "index": 순번}, ...]
    """
    workplaces = []
    # EDI 메인의 사업장 드롭다운 또는 사업장 목록 테이블에서 추출
    # 사이트 구조에 따라 셀렉터 조정 필요
    selectors = [
        "select[name*='workplace'] option",
        "select[name*='biz'] option",
        "select#bizList option",
        "#workplaceList option",
    ]
    for sel in selectors:
        try:
            options = page.locator(sel)
            count = options.count()
            if count > 1:
                for i in range(count):
                    text = options.nth(i).inner_text(timeout=500).strip()
                    value = options.nth(i).get_attribute("value") or ""
                    if text and value and text != "선택":
                        workplaces.append({"name": text, "value": value, "index": i})
                if workplaces:
                    return workplaces
        except Exception:
            continue

    # iframe 안에서도 탐색
    for frame in page.frames:
        if frame == page.main_frame:
            continue
        for sel in selectors:
            try:
                options = frame.locator(sel)
                count = options.count()
                if count > 1:
                    for i in range(count):
                        text = options.nth(i).inner_text(timeout=500).strip()
                        value = options.nth(i).get_attribute("value") or ""
                        if text and value and text != "선택":
                            workplaces.append({"name": text, "value": value, "index": i})
                    if workplaces:
                        return workplaces
            except Exception:
                continue

    return workplaces


def _nhis_close_overlay(page: Page) -> None:
    """open_nav 등 떠 있는 오버레이를 숨겨서 클릭 차단 방지."""
    try:
        page.evaluate("""() => {
            const el = document.getElementById('open_nav');
            if (el) el.style.display = 'none';
        }""")
    except Exception:
        pass


def _nhis_find_billing_docs(page: Page) -> list[dict]:
    """받은문서 테이블에서 고지내역서 행 정보를 수집한다."""
    docs = []
    rows = page.locator("table tr")
    count = rows.count()
    doc_idx = 0
    for i in range(count):
        row = rows.nth(i)
        try:
            text = row.inner_text(timeout=500)
        except Exception:
            continue
        if "고지내역서" not in text:
            continue
        cells = row.locator("td")
        for j in range(cells.count()):
            try:
                ct = cells.nth(j).inner_text(timeout=200)
            except Exception:
                continue
            if "고지내역서" in ct:
                docs.append({
                    "index": doc_idx,
                    "row": i,
                    "cell": j,
                    "text": text.strip()[:100],
                    "cell_text": ct.strip(),
                })
                doc_idx += 1
                break
    return docs


def _nhis_click_doc(page: Page, doc_index: int) -> bool:
    """받은문서 테이블에서 n번째 고지내역서 행을 클릭."""
    _nhis_close_overlay(page)
    rows = page.locator("table tr")
    count = rows.count()
    found = 0
    for i in range(count):
        row = rows.nth(i)
        try:
            text = row.inner_text(timeout=500)
        except Exception:
            continue
        if "고지내역서" not in text:
            continue
        if found == doc_index:
            cells = row.locator("td")
            for j in range(cells.count()):
                try:
                    ct = cells.nth(j).inner_text(timeout=200)
                except Exception:
                    continue
                if "고지내역서" in ct:
                    link = cells.nth(j).locator("a").first
                    try:
                        if link.is_visible(timeout=500):
                            link.click(force=True)
                            return True
                    except Exception:
                        pass
                    cells.nth(j).click(force=True)
                    return True
            return False
        found += 1
    return False


def run_nhis(page: Page, cdp: CDPSession, year: int, month: int,
             save_base: Path, do_ss: bool, wait_login: int = 0) -> Result:
    """건강보험 EDI — 개인별보험료고지내역을 사업장별로 PDF 저장."""
    r = Result("건강보험EDI")
    ctx = page.context
    ym = f"{year}년{month}월"

    try:
        log.info("── 건강보험 EDI ──")
        page.goto("https://edi.nhis.or.kr", timeout=30_000,
                  wait_until="domcontentloaded")
        page.wait_for_timeout(3000)
        screenshot(page, "nhis_01_main", do_ss)

        if not check_logged_in(page):
            if wait_login > 0:
                if not wait_for_login(page, "건강보험 EDI", wait_login):
                    r.login_required = True
                    return r
            else:
                r.login_required = True
                log.warning("  건강보험 EDI: 로그인 필요")
                return r

        log.info("  로그인 확인됨")

        # ── 받은문서에서 고지내역서 수집 ──
        log.info("  받은문서에서 고지내역서 탐색...")
        docs = _nhis_find_billing_docs(page)
        if not docs:
            log.warning("  받은문서에 고지내역서 없음 — 사업장 전환으로 시도")

        # 사업장 목록 파악
        workplaces = _nhis_get_workplaces(page)
        if workplaces:
            log.info(f"  사업장 {len(workplaces)}곳 발견")
            for wp in workplaces:
                wp_name = safe_filename(wp["name"])
                save_dir = save_base / wp_name
                fp = save_dir / f"건강보험개인별내역_{wp_name}_{ym}.pdf"
                if fp.exists():
                    log.info(f"  [{wp_name}] 이미 존재 → 건너뜀")
                    r.files.append(str(fp))
                    continue

                log.info(f"  [{wp_name}] 사업장 전환...")
                # 사업장 전환 후 개인별고지내역 조회
                # (실제 셀렉터는 사이트 탐색 후 확인 필요)
                query_delay()
        else:
            log.info("  사업장 드롭다운 없음 — 받은문서 방식으로 진행")
            # 받은문서에서 하나씩 열어 PDF 저장
            if docs:
                for doc in docs:
                    idx = doc["index"]
                    label = doc["cell_text"]
                    log.info(f"  [{idx+1}/{len(docs)}] '{label}' 열기...")

                    if idx > 0:
                        page.goto("https://edi.nhis.or.kr", timeout=30_000,
                                  wait_until="domcontentloaded")
                        page.wait_for_timeout(3000)
                        if not check_logged_in(page):
                            log.warning("  세션 만료 — 중단")
                            break

                    pages_before = set(id(p) for p in ctx.pages)
                    if not _nhis_click_doc(page, idx):
                        log.warning(f"  '{label}' 클릭 실패")
                        r.skipped.append(label)
                        continue

                    page.wait_for_timeout(4000)
                    screenshot(page, f"nhis_doc_{idx}", do_ss)

                    new_pages = [p for p in ctx.pages
                                 if id(p) not in pages_before
                                 and p.url and p.url != "about:blank"]

                    target = page
                    popup_cdp = None
                    if new_pages:
                        target = new_pages[0]
                        try:
                            target.wait_for_load_state("domcontentloaded",
                                                       timeout=15_000)
                        except Exception:
                            pass
                        popup_cdp = ctx.new_cdp_session(target)

                    active_cdp = popup_cdp or cdp
                    safe_label = safe_filename(label)[:40]
                    save_dir = save_base / "기본사업장"
                    fp = save_dir / f"건강보험_{safe_label}_{ym}.pdf"
                    try:
                        if cdp_print_pdf(active_cdp, fp):
                            r.files.append(str(fp))
                    finally:
                        if popup_cdp:
                            popup_cdp.detach()
                            target.close()
                    query_delay()

        r.ok = len(r.files) > 0
    except Exception as e:
        r.error = str(e)
        log.error(f"  건강보험 오류: {e}")
    return r


# ═══════════════════════════════════════════════════════════════
# 국민연금 EDI — 2차최종결정내역
# ═══════════════════════════════════════════════════════════════

def _nps_get_workplaces(page: Page) -> list[dict]:
    """국민연금 EDI에서 사업장 목록 수집."""
    workplaces = []
    selectors = [
        "select[name*='wrkpl'] option",
        "select[name*='biz'] option",
        "#wrkplList option",
        "select option",
    ]
    for sel in selectors:
        try:
            options = page.locator(sel)
            count = options.count()
            if count > 1:
                for i in range(count):
                    text = options.nth(i).inner_text(timeout=500).strip()
                    value = options.nth(i).get_attribute("value") or ""
                    if text and value and text not in ("선택", "전체", ""):
                        workplaces.append({"name": text, "value": value, "index": i})
                if workplaces:
                    return workplaces
        except Exception:
            continue
    return workplaces


def _nps_navigate_to_decision(page: Page) -> bool:
    """결정내역 조회 화면으로 이동."""
    menu_paths = [
        ["보험료", "결정내역조회"],
        ["보험료", "결정내역"],
        ["보험료", "2차최종결정내역"],
        ["보험료관리", "결정내역조회"],
    ]
    for path in menu_paths:
        try:
            for item in path:
                for loc_fn in [
                    lambda: page.get_by_role("link", name=item, exact=False).first,
                    lambda: page.get_by_role("button", name=item, exact=False).first,
                    lambda: page.get_by_text(item, exact=False).first,
                ]:
                    try:
                        el = loc_fn()
                        el.wait_for(state="visible", timeout=2000)
                        el.click()
                        page.wait_for_timeout(1500)
                        break
                    except Exception:
                        continue
                else:
                    raise Exception(f"메뉴 '{item}' 찾기 실패")
            return True
        except Exception:
            continue
    return False


def run_nps(page: Page, cdp: CDPSession, year: int, month: int,
            save_base: Path, do_ss: bool, wait_login: int = 0) -> Result:
    """국민연금 EDI — 2차최종결정내역·국고지원내역을 사업장별로 PDF 저장."""
    r = Result("국민연금EDI")
    ym = f"{year}년{month}월"

    try:
        log.info("── 국민연금 EDI ──")
        page.goto("https://edi.nps.or.kr", timeout=30_000,
                  wait_until="domcontentloaded")
        page.wait_for_timeout(3000)
        screenshot(page, "nps_01_main", do_ss)

        if not check_logged_in(page):
            if wait_login > 0:
                if not wait_for_login(page, "국민연금 EDI", wait_login):
                    r.login_required = True
                    return r
            else:
                r.login_required = True
                log.warning("  국민연금 EDI: 로그인 필요")
                return r

        log.info("  로그인 확인됨")
        screenshot(page, "nps_02_loggedin", do_ss)

        # 사업장 목록
        workplaces = _nps_get_workplaces(page)
        if workplaces:
            log.info(f"  사업장 {len(workplaces)}곳 발견")
        else:
            log.info("  사업장 드롭다운 없음 — 단일 사업장으로 진행")
            workplaces = [{"name": "기본사업장", "value": "", "index": 0}]

        for wp in workplaces:
            wp_name = safe_filename(wp["name"])
            save_dir = save_base / wp_name

            # 2차최종결정내역
            fp_decision = save_dir / f"국민연금_2차최종결정내역_{wp_name}_{ym}.pdf"
            if fp_decision.exists():
                log.info(f"  [{wp_name}] 2차최종결정내역 이미 존재 → 건너뜀")
                r.files.append(str(fp_decision))
            else:
                log.info(f"  [{wp_name}] 2차최종결정내역 조회...")
                # 사업장 전환 (값이 있을 때만)
                if wp["value"]:
                    # 실제 셀렉터는 사이트 탐색 후 확인 필요
                    pass

                if _nps_navigate_to_decision(page):
                    page.wait_for_timeout(2000)
                    screenshot(page, f"nps_decision_{wp_name}", do_ss)
                    if cdp_print_pdf(cdp, fp_decision):
                        r.files.append(str(fp_decision))
                else:
                    log.warning(f"  [{wp_name}] 결정내역 메뉴 찾기 실패")
                    r.skipped.append(f"{wp_name}_결정내역")

                query_delay()

            # 국고지원내역 (대상 사업장만 — 메뉴가 있으면 저장)
            for subsidy_name, subsidy_menu in [
                ("국고지원예정대상자", ["국고지원예정대상자", "지원예정"]),
                ("국고지원대상자", ["국고지원대상자", "지원대상"]),
            ]:
                fp_sub = save_dir / f"국민연금_{subsidy_name}_{wp_name}_{ym}.pdf"
                if fp_sub.exists():
                    r.files.append(str(fp_sub))
                    continue
                for menu_text in subsidy_menu:
                    try:
                        el = page.get_by_text(menu_text, exact=False).first
                        el.wait_for(state="visible", timeout=2000)
                        el.click()
                        page.wait_for_timeout(2000)
                        if cdp_print_pdf(cdp, fp_sub):
                            r.files.append(str(fp_sub))
                        break
                    except Exception:
                        continue

            query_delay()

        r.ok = len(r.files) > 0
    except Exception as e:
        r.error = str(e)
        log.error(f"  국민연금 오류: {e}")
    return r


# ═══════════════════════════════════════════════════════════════
# 메인
# ═══════════════════════════════════════════════════════════════

def main() -> int:
    parser = argparse.ArgumentParser(description="사회보험료 월간 조회 자동화")
    parser.add_argument("--force", action="store_true",
                        help="휴무일/완료 체크 건너뜀")
    parser.add_argument("--site", choices=["nhis", "nps"],
                        help="특정 사이트만")
    parser.add_argument("--month", type=str,
                        help="대상 월 (YYYY-MM)")
    parser.add_argument("--screenshots", action="store_true",
                        help="각 단계 스크린샷 저장")
    parser.add_argument("--wait-login", type=int, default=0, metavar="SEC",
                        help="로그인 대기 시간 (초). 0이면 미로그인 사이트 건너뜀")
    parser.add_argument("--cdp-url", type=str, default=CDP_URL,
                        help=f"Chrome CDP 주소 (기본: {CDP_URL})")
    args = parser.parse_args()

    dt = now_kst()
    if args.month:
        y, m = map(int, args.month.split("-"))
        target = datetime.date(y, m, 1)
    else:
        target = dt.date()

    year, month = target.year, target.month
    log.info(f"═══ 사회보험료 조회 시작: {year}년 {month:02d}월 ═══")

    if not args.force:
        if done_path(target).exists():
            log.info("이번 달 이미 완료됨")
            return 0
        if is_holiday(dt.date()):
            log.info("휴무일 — 다음 영업일에 재실행")
            return 0

    results: list[Result] = []

    with sync_playwright() as pw:
        log.info(f"기존 Chrome에 연결: {args.cdp_url}")
        try:
            browser: Browser = pw.chromium.connect_over_cdp(args.cdp_url)
        except Exception as e:
            log.error(f"Chrome 연결 실패: {e}")
            log.error("Chrome이 --remote-debugging-port=9333 으로 실행 중인지 확인하세요.")
            return 1

        ctx = browser.contexts[0] if browser.contexts else browser.new_context()
        page = ctx.new_page()
        cdp = ctx.new_cdp_session(page)

        try:
            wl = args.wait_login
            save_base = SAVE_BASE

            if args.site == "nhis" or not args.site:
                results.append(
                    run_nhis(page, cdp, year, month, save_base,
                             args.screenshots, wl))

            if args.site == "nps" or not args.site:
                results.append(
                    run_nps(page, cdp, year, month, save_base,
                            args.screenshots, wl))
        finally:
            cdp.detach()
            page.close()
            browser.close()

    # 결과 정리
    log.info("═══ 결과 요약 ═══")
    all_files: list[str] = []
    for r in results:
        log.info(f"  {r.summary()}")
        all_files.extend(r.files)

    all_ok = all(r.ok for r in results) and not any(
        r.login_required for r in results)

    if not args.site and all_ok:
        dp = done_path(target)
        dp.write_text(
            f"완료: {dt.isoformat()}\n\n" + "\n".join(all_files),
            encoding="utf-8",
        )
        log.info(f"done 파일 생성: {dp}")
    elif not all_ok:
        log.info("일부 미완료 → done 파일 미생성")

    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
