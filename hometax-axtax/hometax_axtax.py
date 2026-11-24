"""
홈택스 조회 도구 (조회만 한다)

명령 8가지:
  check-list    거래처 명단(xlsx)을 읽어 개수·형식만 점검한다 (상호·번호는 출력하지 않는다)
  tonghap       부가가치세 신고자료 통합조회 -> 거래처별 폴더에 저장
  scheduled     예약 실행: 통합조회 -> 신용카드/판매대행을 차례로 (크롬 띄우기·메뉴 이동은 코드가, 로그인은 사람이)
  vat-sum       부가가치세신고용 합계표조회 (화면을 아직 조사하지 못해 지금은 멈춘다. 1월 15일 이후 조사)
  cash-sum      현금영수증 매출총액(매출내역누계) 조회 -> 거래처별 폴더에 저장 (캡처 + 엑셀)
  card-sum      신용카드/판매(결제)대행 매출자료 조회 -> 거래처별 폴더에 저장 (캡처 + 엑셀)
  export-sum    수출실적명세서 조회 -> 수출이 있는 거래처만 거래처별 폴더에 저장
  invoice-sum   전자(세금)계산서 합계표 조회 - 총합계 읽기 (세무대리 공통 조회 메뉴의 화면)

흐름(tonghap): 크롬(9333)에 붙기 -> 로그인/화면 확인 -> 거래처마다
      [조건 넣기 -> 결과 칸 비우기 -> 조회 -> 다시 채워졌는지 확인 -> 표 읽기 -> 화면 캡처 -> 저장]

이 파일은 조회만 한다. 로그인은 절대 다루지 않는다 (로그인 화면이 보이면 멈춘다).
확인이 안 되는 것이 있으면 값을 내놓지 않고 멈춘다 (화면지도.md 5번 항목 참고).

사용 예 (번호는 예시용 가짜 번호):
    python hometax_axtax.py check-list --list "C:\\경로\\거래처명단.xlsx"
    python hometax_axtax.py tonghap --list "C:\\경로\\거래처명단.xlsx" --year 2026 --ki 2기 --rtn 예정+확정
    python hometax_axtax.py tonghap --biz-no 1234567891 --year 2026 --ki 2기 --rtn 예정+확정
    python hometax_axtax.py invoice-sum --biz-no 1234567891 --gubun 둘다 --period-mode 분기별 --year 2026 --quarter "1기 확정"

개인정보 원칙:
  - 거래처 명단에서는 '상호(성명)'과 '사업자등록번호' 두 열만 읽는다. 주민등록번호 열은 읽지도 저장하지도 않는다.
  - 결과 폴더는 바탕화면에 만든다. 바탕화면이 OneDrive 등 클라우드 동기화 폴더이면 멈춘다.
  - 화면에 찍는 글에는 상호·금액을 넣지 않는다 (거래처는 순번과 번호 끝 4자리로만 표시).
"""

import argparse
import base64
import ctypes
import html
import json
import os
import random
import re
import subprocess
import sys
import time
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeout
from playwright.sync_api import sync_playwright

CDP_URL = "http://127.0.0.1:9333"

# 화면지도.md 1번: 이 값은 조건을 바꿔도 안 변하는 것으로 확인된 화면 식별자.
# URL의 tmIdx 계열 번호는 최초 진입 시점에 고정되고 이후 안 바뀌므로 화면 식별에 쓰지 않는다.
WORK_AREA_ID = "mf_txppWframe"
WORK_AREA_CLASS = "w2pageFrame PP ET"

CLSF_RADIO_NAME = "mf_txppWframe_radioEtxivClsfCd_input"
CLSF_OPTIONS = {"전자세금계산서": 0, "전자계산서": 1}

GUBUN_RADIO_NAME = "mf_txppWframe_radio7_input"
GUBUN_OPTIONS = {"매출": 0, "매입": 1}

PERIOD_RADIO_NAME = "mf_txppWframe_file"
PERIOD_MODE_OPTIONS = {"작성일자": 0, "월별": 1, "분기별": 2}

YEAR_SELECT_ID = "mf_txppWframe_selectbox6"
MONTH_SELECT_ID = "mf_txppWframe_selectbox7"
QUARTER_SELECT_ID = "mf_txppWframe_selectbox8"

QUERY_BUTTON_ID = "mf_txppWframe_trigger23"

CLIENT_SWITCH_BUTTON_ID = "mf_txppWframe_trigger501"
CLIENT_POPUP_ID = "mf_txppWframe_UTEETZZA21"
CLIENT_POPUP_PREFIX = "mf_txppWframe_UTEETZZA21_wframe_"


class StopForSafety(Exception):
    """확인이 안 되거나 위험할 수 있어서 멈출 때 쓴다. 이 예외가 나오면 값을 내놓지 않는다."""


_LOG_FILE = None   # 예약 실행(scheduled)에서만 정한다. 화면에 찍는 글은 번호 끝 4자리만 있고 상호·금액은 없다.


def log(msg):
    print(msg, flush=True)
    if _LOG_FILE:
        try:
            with open(_LOG_FILE, "a", encoding="utf-8") as f:
                f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} {msg}\n")
        except OSError:
            pass


# ---------------------------------------------------------------------------
# 크롬에 붙기
# ---------------------------------------------------------------------------

def connect_chrome(p):
    try:
        return p.chromium.connect_over_cdp(CDP_URL)
    except Exception:
        raise StopForSafety(
            "크롬(9333)에 연결하지 못했습니다. 홈택스 전용 크롬 창이 닫혔을 수 있습니다. "
            "CLAUDE.md의 '홈택스 크롬 띄우는 명령'으로 다시 띄운 뒤, 사람이 로그인하고 "
            "해당 화면까지 옮겨 주세요."
        ) from None


def find_hometax_page(browser):
    # 팝업 창(popup.html)은 본 화면이 아니므로 건너뛴다.
    fallback = (None, None)
    for ctx in browser.contexts:
        for pg in ctx.pages:
            if "hometax" not in pg.url:
                continue
            if "popup.html" in pg.url:
                fallback = (ctx, pg) if fallback == (None, None) else fallback
                continue
            return ctx, pg
    return fallback


# ---------------------------------------------------------------------------
# 로그인 판정
#
# 처음에는 머리말의 "로그아웃" 글자 + 이 화면 전용 class(WORK_AREA_CLASS)로 판단했는데,
# 실제로 다른 조회 화면(예: 신용카드 매출자료 조회)으로 옮겨가 있을 때 이 class가
# 안 맞아서 "로그인을 확신 못함"으로 잘못 멈추는 게 확인됐다 (머리말은 멀쩡히
# "로그아웃"을 보여주고 있었는데도). 화면마다 달라지는 class를 로그인 판정에 쓴 게 문제였다.
#
# 그래서 "로그인됐는지"와 "지금 합계표 조회 화면인지"를 완전히 분리한다.
# 로그인 판정은 어떤 조회 화면에서든 공통적으로 있는 것 — 조회 조건 칸(입력/선택/라디오)과
# "조회"라고 적힌 버튼이 실제로 보이는지로 한다. 화면이 바뀌어도 이 두 가지는
# 거의 모든 홈택스 조회 화면에 있으므로, 특정 화면 하나에 종속되지 않는다.
#
# (1) 비밀번호 칸이 실제로 보이면 그건 확실한 로그인 화면 신호로 즉시 멈추고,
# (2) 그게 아니면 조회 조건 칸 + 조회 버튼이 같이 보이는지 보고,
# (3) 한 번에 애매하면 화면이 막 다시 그려지는 중일 수 있으니 잠깐 뒤 다시 본 뒤에만 최종 판단한다.
# ---------------------------------------------------------------------------

def _is_login_screen_visible(frame):
    return frame.evaluate(
        """
        () => {
          const pwFields = document.querySelectorAll('input[type="password"]');
          for (const pw of pwFields) {
            const rect = pw.getBoundingClientRect();
            const style = window.getComputedStyle(pw);
            if (rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none') {
              return true;
            }
          }
          return false;
        }
        """
    )


def _has_query_screen_elements(frame):
    """조회 조건 칸(입력/선택/라디오)과 '조회' 버튼이 실제로 보이는지. 화면이 무엇이든
    (합계표든 다른 조회 화면이든) 공통으로 있는 요소로 로그인 여부를 판단하기 위한 것 —
    특정 화면 하나의 class 값에 기대지 않는다."""
    return frame.evaluate(
        """
        () => {
          const isVisible = (el) => !!(el.offsetWidth || el.offsetHeight);
          const conditionFields = document.querySelectorAll('input[type="text"], input[type="radio"], select');
          const hasConditionField = Array.from(conditionFields).some(isVisible);

          const buttons = document.querySelectorAll('input[type="button"], button');
          const hasQueryButton = Array.from(buttons).some(
            (b) => isVisible(b) && (b.value || b.innerText || '').trim() === '조회'
          );

          return hasConditionField && hasQueryButton;
        }
        """
    )


def _work_area_class(frame):
    return frame.evaluate(
        f"""
        () => {{
          const el = document.getElementById('{WORK_AREA_ID}');
          return el ? el.className : null;
        }}
        """
    )


def ensure_logged_in(frame):
    """로그인 여부만 본다. 어떤 조회 화면에 있는지는 안 본다 (그건 ensure_on_target_screen).
    머리말 글자 하나로 판단하지 않는다."""
    if _is_login_screen_visible(frame):
        raise StopForSafety("로그인 화면이 보입니다. 사람이 로그인해야 합니다. (이 코드는 로그인을 다루지 않습니다)")

    if _has_query_screen_elements(frame):
        return

    # 애매하면 화면이 막 다시 그려지는 중일 수 있으니 잠깐 뒤 다시 본다.
    time.sleep(1.5)

    if _is_login_screen_visible(frame):
        raise StopForSafety("로그인 화면이 보입니다. 사람이 로그인해야 합니다.")

    if _has_query_screen_elements(frame):
        return

    raise StopForSafety(
        "로그인 여부를 확신할 수 없어 멈춥니다. "
        "(조회 조건 칸과 '조회' 버튼이 화면에 안 보입니다. 로그인 화면도 아닙니다 — "
        "예상치 못한 화면일 수 있습니다.)"
    )


def _class_has_all(actual, expected):
    """작업영역 class에 필요한 낱말이 모두 들어 있는지. 같은 창에서 메뉴로 화면을 옮기면 이 class가
    갱신되지 않고 'PP HT PP PP'처럼 낱말이 쌓이므로(2026-10-28 확인), 글자 그대로 비교하면 안 된다.
    또 이전 화면의 낱말이 남아 있을 수도 있어서, class만으로 화면을 확신해서는 안 된다
    (반드시 화면 제목도 함께 확인한다)."""
    return set(expected.split()) <= set((actual or "").split())


def ensure_on_target_screen(frame):
    """지금 화면이 이 코드가 다루는 화면(합계표 조회)이 맞는지. 로그인 여부와는 별개다."""
    work_class = _work_area_class(frame)
    has_title = frame.evaluate(
        """
        () => {
          const root = document.getElementById('mf_txppWframe');
          return !!root && Array.from(root.querySelectorAll('div')).some(
            d => (d.offsetWidth || d.offsetHeight) && (d.innerText || '').trim().startsWith('전자세금계산서 합계표 조회'));
        }
        """
    )
    if not _class_has_all(work_class, WORK_AREA_CLASS) or not has_title:
        raise StopForSafety(
            f"로그인은 돼 있지만 지금 화면이 합계표 조회 화면이 아닌 것 같습니다 "
            f"(화면class={work_class!r}, 기대에 포함돼야 함={WORK_AREA_CLASS!r}, 제목확인={has_title}). "
            "메뉴에서 화면을 맞게 옮겨주세요."
        )


def ensure_logged_in_and_on_screen(frame):
    ensure_logged_in(frame)
    ensure_on_target_screen(frame)


# ---------------------------------------------------------------------------
# 거래처 바꾸기
#
# 화면지도.md 5번 함정3: 결과 목록의 "선택" 라디오는 화면 밖(음수 좌표)에 있어서
# 라디오 자체를 클릭하면 안 되고, 그 줄의 보이는 셀을 클릭해야 한다.
# ---------------------------------------------------------------------------

def normalize_biz_no(raw):
    return re.sub(r"[^0-9]", "", raw or "")


def switch_client(frame, biz_no_raw):
    digits = normalize_biz_no(biz_no_raw)

    if digits == "0000000000":
        # 사용자가 정해둔 규칙: 이 값이 그대로면 사업자번호를 바꾸는 걸 잊은 것이다.
        raise StopForSafety("사업자등록번호를 바꿔서 다시 붙여넣어 주세요")

    if len(digits) != 10:
        raise StopForSafety(f"사업자등록번호 형식이 이상합니다 (10자리 숫자가 아님): {biz_no_raw!r}")

    frame.click(f"#{CLIENT_SWITCH_BUTTON_ID}")
    frame.wait_for_selector(f"#{CLIENT_POPUP_ID}", state="visible", timeout=10000)

    frame.select_option(f"#{CLIENT_POPUP_PREFIX}selectbox5", label="사업자등록번호")
    frame.fill(f"#{CLIENT_POPUP_PREFIX}txprDscmNoA", digits)
    frame.click(f"#{CLIENT_POPUP_PREFIX}trigger85")
    time.sleep(1.5)

    # 화면에 실제로 보이는 줄만 본다. 최대 30줄까지 확인하고, 각 줄(0~7번 칸)의
    # 글자에 찾는 번호가 들어있는지 본다. 정확히 하나가 아니면 멈춘다.
    matched_rows = []
    for r in range(30):
        first_cell_id = f"{CLIENT_POPUP_PREFIX}grdResult_cell_{r}_0"
        exists = frame.evaluate(f"() => !!document.getElementById('{first_cell_id}')")
        if not exists:
            break
        row_text = frame.evaluate(
            f"""
            () => {{
              let parts = [];
              for (let c = 0; c < 8; c++) {{
                const el = document.getElementById('{CLIENT_POPUP_PREFIX}grdResult_cell_{r}_' + c);
                if (el) parts.push(el.innerText || '');
              }}
              return parts.join(' ');
            }}
            """
        )
        if digits in row_text.replace("-", ""):
            matched_rows.append(r)

    if len(matched_rows) != 1:
        try:
            frame.click(f"#{CLIENT_POPUP_PREFIX}btnClose")
        except Exception:
            pass
        raise StopForSafety(
            f"사업자등록번호 {biz_no_raw} 로 화면에 보이는 줄 중 정확히 하나가 나오지 않았습니다 "
            f"(일치한 줄 수: {len(matched_rows)}). 멈춥니다."
        )

    row = matched_rows[0]
    # 화면지도.md 함정3: 라디오가 아니라 보이는 셀을 클릭해야 실제로 선택된다.
    frame.click(f"#{CLIENT_POPUP_PREFIX}grdResult_cell_{row}_0")
    time.sleep(0.3)
    frame.click(f"#{CLIENT_POPUP_PREFIX}btnProcess")
    time.sleep(1.5)

    popup_gone = frame.evaluate(
        f"""
        () => {{
          const el = document.getElementById('{CLIENT_POPUP_ID}');
          return !el || !(el.offsetWidth || el.offsetHeight);
        }}
        """
    )
    if not popup_gone:
        raise StopForSafety("거래처 전환 창이 닫히지 않았습니다. 멈춥니다.")

    applied = frame.evaluate(
        """
        () => {
          const w = window['mf_txppWframe_scwin'];
          return w ? (w.txprDscmNoS || null) : null;
        }
        """
    )
    if applied != digits:
        raise StopForSafety(
            f"거래처 전환이 실제로 반영됐는지 확인할 수 없습니다 (화면에서 읽은 값: {applied!r}). 값을 내놓지 않고 멈춥니다."
        )


# ---------------------------------------------------------------------------
# 조회 조건 넣기
# ---------------------------------------------------------------------------

def _set_radio(frame, name, index):
    ids = frame.evaluate(f"() => Array.from(document.querySelectorAll('input[name=\"{name}\"]')).map(r => r.id)")
    if index >= len(ids):
        raise StopForSafety(f"{name} 라디오에 {index}번째 옵션이 없습니다 (현재 {len(ids)}개뿐).")
    target_id = ids[index]
    frame.click(f"label[for='{target_id}']")
    time.sleep(0.3)
    checked = frame.evaluate(f"() => document.getElementById('{target_id}').checked")
    if not checked:
        raise StopForSafety(f"{name} 조건을 바꾸려 했지만 실제로 반영되지 않았습니다. 멈춥니다.")


def _select_and_verify(frame, select_id, label):
    visible = frame.evaluate(
        f"""
        () => {{
          const s = document.getElementById('{select_id}');
          return s ? !!(s.offsetWidth || s.offsetHeight) : null;
        }}
        """
    )
    if visible is None:
        raise StopForSafety(f"목록 요소를 찾지 못했습니다: {select_id}")
    if not visible:
        raise StopForSafety(
            f"{select_id} 목록이 화면에 보이지 않습니다 (기간구분을 먼저 맞는 모드로 바꿨는지 확인하세요). 멈춥니다."
        )
    frame.select_option(f"#{select_id}", label=label)
    time.sleep(0.3)
    actual = frame.evaluate(
        f"""
        () => {{
          const s = document.getElementById('{select_id}');
          return s.options[s.selectedIndex] ? s.options[s.selectedIndex].text : null;
        }}
        """
    )
    if actual != label:
        raise StopForSafety(f"{select_id} 값을 '{label}'로 바꾸려 했지만 실제로는 '{actual}'입니다. 멈춥니다.")


def set_conditions(frame, clsf, gubun, period_mode, year=None, quarter=None, month=None):
    if clsf not in CLSF_OPTIONS:
        raise StopForSafety(f"분류 값을 모릅니다: {clsf!r} (가능: {list(CLSF_OPTIONS)})")
    _set_radio(frame, CLSF_RADIO_NAME, CLSF_OPTIONS[clsf])

    if gubun not in GUBUN_OPTIONS:
        raise StopForSafety(f"구분 값을 모릅니다: {gubun!r} (가능: {list(GUBUN_OPTIONS)})")
    _set_radio(frame, GUBUN_RADIO_NAME, GUBUN_OPTIONS[gubun])

    if period_mode not in PERIOD_MODE_OPTIONS:
        raise StopForSafety(f"기간구분 값을 모릅니다: {period_mode!r} (가능: {list(PERIOD_MODE_OPTIONS)})")

    if period_mode == "작성일자":
        # 화면지도.md TODO: 거래기간 달력 두 칸의 정확한 요소 이름을 아직 확인하지 못했다.
        # 확인 안 된 것으로 조작하지 않고 멈춘다.
        raise StopForSafety(
            "작성일자(날짜범위) 모드는 달력 칸의 정확한 이름을 아직 확인하지 못해 지원하지 않습니다. "
            "화면지도.md의 TODO를 먼저 채운 뒤에 이 기능을 추가하세요."
        )

    _set_radio(frame, PERIOD_RADIO_NAME, PERIOD_MODE_OPTIONS[period_mode])

    if period_mode == "월별":
        if not year or not month:
            raise StopForSafety("월별 모드에는 year와 month가 모두 필요합니다.")
        _select_and_verify(frame, YEAR_SELECT_ID, year)
        _select_and_verify(frame, MONTH_SELECT_ID, month)
    elif period_mode == "분기별":
        if not year or not quarter:
            raise StopForSafety("분기별 모드에는 year와 quarter가 모두 필요합니다.")
        _select_and_verify(frame, YEAR_SELECT_ID, year)
        _select_and_verify(frame, QUARTER_SELECT_ID, quarter)


# ---------------------------------------------------------------------------
# 조회 + 총합계 읽기
#
# 화면지도.md 5번 함정1: "조회"와 "명세서 조회"는 다른 버튼이다. 이 함수는
# 위쪽 총합계 요약표만 읽는다 (거래처별 상세 명세는 다루지 않는다 - 필요하면
# mf_txppWframe_trigger25를 별도로 눌러야 한다).
#
# 함정2: 조건을 바꿨는데 화면이 실제로 갱신됐는지 숫자만으로는 알 수 없다.
# 그래서 화면 제목에 지금 조건(구분)이 실제로 반영됐는지 먼저 확인한 뒤에만
# 숫자를 믿는다.
# ---------------------------------------------------------------------------

def run_query_and_read_summary(frame, gubun):
    frame.click(f"#{QUERY_BUTTON_ID}")
    time.sleep(2)

    unexpected = frame.evaluate(
        """
        () => Array.from(document.querySelectorAll('.w2window, .w2alert, .w2confirm'))
          .filter(e => e.offsetWidth > 0 && e.offsetHeight > 0)
          .map(e => (e.innerText || '').trim().slice(0, 200))
        """
    )
    if unexpected:
        raise StopForSafety(f"조회 후 예상치 못한 창이 떠 있습니다: {unexpected}. 값을 내놓지 않고 멈춥니다.")

    title = frame.evaluate(
        """
        () => {
          const h = Array.from(document.querySelectorAll('div'))
            .find(d => d.innerText && d.innerText.trim().startsWith('전자세금계산서 합계표 조회'));
          return h ? h.innerText.trim().split('\\n')[0] : null;
        }
        """
    )
    expect_suffix = f"({gubun})"
    if not title or expect_suffix not in title:
        raise StopForSafety(
            f"조회 후 화면 제목이 요청한 구분과 맞지 않습니다 (기대: '...{expect_suffix}', 실제: {title!r}). "
            "화면이 아직 안 바뀐 것일 수 있어 값을 내놓지 않고 멈춥니다."
        )

    summary_text = frame.evaluate(
        """
        () => {
          const anchor = Array.from(document.querySelectorAll('*'))
            .find(e => e.children.length === 0 && (e.innerText || '').includes('총합계'));
          if (!anchor) return null;
          let box = anchor;
          for (let i = 0; i < 6 && box; i++) {
            if ((box.innerText || '').includes('공급가액')) break;
            box = box.parentElement;
          }
          return box ? box.innerText : null;
        }
        """
    )
    if not summary_text:
        raise StopForSafety("총합계 표를 화면에서 찾지 못했습니다. 값을 내놓지 않고 멈춥니다.")

    lines = [l for l in summary_text.split("\n") if l.strip()]
    header_line = None
    total_line = None
    for l in lines:
        if l.startswith("구분") and "매수" in l:
            header_line = l
        if l.startswith("합계"):
            total_line = l

    if not header_line or not total_line:
        raise StopForSafety(f"총합계 표에서 '합계' 줄을 못 찾았습니다. 읽은 내용 일부: {summary_text[:200]!r}")

    cols = header_line.split("\t")
    vals = total_line.split("\t")
    if len(cols) != len(vals):
        raise StopForSafety("총합계 표의 열 개수와 값 개수가 서로 안 맞습니다. 값을 내놓지 않고 멈춥니다.")

    return dict(zip(cols, vals)), title


# ---------------------------------------------------------------------------
# 매출/매입 이어서 조회하기
#
# CLAUDE.md 규칙: 연속 조회 사이에 2~5초 쉰다 (국세청 사이트 보호). 거래처는
# 한 번만 바꾸고, 구분(매출/매입)만 바꿔가며 같은 조건으로 이어서 조회한다.
# ---------------------------------------------------------------------------

REST_BETWEEN_QUERIES_SECONDS = (2.0, 5.0)


def run_multi_gubun(frame, clsf, period_mode, gubun_list, year=None, quarter=None, month=None):
    results = []
    for i, gubun in enumerate(gubun_list):
        set_conditions(frame, clsf, gubun, period_mode, year=year, quarter=quarter, month=month)
        result, title = run_query_and_read_summary(frame, gubun)
        results.append((gubun, result, title))

        if i < len(gubun_list) - 1:
            wait_s = random.uniform(*REST_BETWEEN_QUERIES_SECONDS)
            log(f"      (다음 조회까지 {wait_s:.1f}초 쉬는 중 — 연속 조회 사이 쉬는 규칙)")
            time.sleep(wait_s)

    return results


def format_combined_table(results):
    count_key = {"매출": "매출처수", "매입": "매입처수"}
    header = ["구분", "처수", "매수", "공급가액", "세액", "합계금액"]
    rows = []
    for gubun, result, _title in results:
        key = count_key.get(gubun, "처수")
        rows.append([
            gubun,
            result.get(key, "?"),
            result.get("매수", "?"),
            result.get("공급가액", "?"),
            result.get("세액", "?"),
            result.get("합계금액", "?"),
        ])

    widths = [max(len(str(r[i])) for r in ([header] + rows)) for i in range(len(header))]
    lines = []
    lines.append(" | ".join(str(header[i]).ljust(widths[i]) for i in range(len(header))))
    lines.append("-+-".join("-" * widths[i] for i in range(len(header))))
    for r in rows:
        lines.append(" | ".join(str(r[i]).ljust(widths[i]) for i in range(len(header))))
    return "\n".join(lines)


# ===========================================================================
# 여기부터: 부가가치세 신고자료 통합조회 / 거래처 명단 / 거래처별 폴더 저장
# ===========================================================================

class NoDataForClient(Exception):
    """서버와 화면이 모두 '자료 없음'이라고 확인된 거래처. 실패가 아니며, 폴더·파일을 만들지 않고 기록만 남긴다."""


class ClientFailed(Exception):
    """거래처 한 곳만 확인이 안 될 때. 그 거래처는 값을 내놓지 않고 실패로 기록한 뒤 다음으로 넘어간다.
    (로그인이 풀리거나 예상 못한 화면처럼 전체가 위험할 때는 StopForSafety로 전부 멈춘다.)"""


# ---------------------------------------------------------------------------
# 누르기 / 기다리기 도구
#
# 화면지도.md 함정4: 머리말 메뉴(세무대리·납세관리 등)가 펼쳐져 있으면 화면 전체를 덮는 어두운
# 덮개(head_dim on)가 켜져서 본문 버튼 클릭을 가로챈다. 또 일반 클릭은 버튼을 화면 맨 위로
# 스크롤하는데, 그러면 상단 고정 도구상자에 가려진다. 그래서 버튼을 화면 가운데로 올리고,
# 그 자리 맨 위에 실제로 그 버튼이 있는지 확인한 뒤에만 그 좌표를 누른다.
# 다른 것이 덮고 있으면 누르지 않고 멈춘다.
# ---------------------------------------------------------------------------

_PREP_CLICK_JS = r"""
(sel) => {
  const e = document.querySelector(sel);
  if (!e) return {found: false};
  e.scrollIntoView({block: 'center', behavior: 'instant'});
  const b = e.getBoundingClientRect();
  const x = b.x + b.width / 2, y = b.y + b.height / 2;
  const top = document.elementFromPoint(x, y);
  const ok = !!top && (top === e || e.contains(top) || top.contains(e));
  return {found: true, ok, x, y, cover: top ? (top.id || top.tagName) : 'null', coverCls: top ? String(top.className) : ''};
}
"""

_MENU_OVERLAY_JS = r"""
() => Array.from(document.querySelectorAll('.head_dim.on')).some(e => e.offsetWidth || e.offsetHeight)
"""

_VISIBLE_WINDOWS_JS = r"""
() => Array.from(document.querySelectorAll('.w2window, .w2alert, .w2confirm'))
  .filter(e => e.offsetWidth || e.offsetHeight)
  .map(e => ({
    id: e.id,
    text: (e.innerText || '').replace(/\s+/g, ' ').trim().slice(0, 150),
    buttons: Array.from(e.querySelectorAll('input[type=button],button'))
      .filter(b => b.offsetWidth || b.offsetHeight)
      .map(b => ({id: b.id, t: (b.value || b.innerText || '').trim()})).filter(b => b.t)
  }))
"""


def safe_click(page, frame, selector):
    r = frame.evaluate(_PREP_CLICK_JS, selector)
    if not r["found"]:
        raise StopForSafety(f"화면에서 요소를 찾지 못했습니다: {selector}")
    if not r["ok"]:
        if "head_dim" in r["coverCls"]:
            raise StopForSafety(
                "머리말 메뉴가 펼쳐져 있어 화면이 어둡게 덮여 있습니다. 상단 메뉴 제목을 눌러 접은 뒤 다시 실행하세요."
            )
        raise StopForSafety(f"{selector} 위를 다른 요소가 덮고 있어 누르지 않았습니다 (덮은 요소: {r['cover']}).")
    page.mouse.click(r["x"], r["y"])


def ensure_no_menu_overlay(frame):
    if frame.evaluate(_MENU_OVERLAY_JS):
        raise StopForSafety(
            "머리말 메뉴가 펼쳐져 있어 화면이 어둡게 덮여 있습니다. 상단 메뉴 제목을 눌러 접은 뒤 다시 실행하세요."
        )


def _wait_until(fn, timeout, interval=0.5):
    end = time.time() + timeout
    while True:
        v = fn()
        if v or time.time() >= end:
            return v
        time.sleep(interval)


def _mask_digits(text):
    return re.sub(r"\d{6,}", "******", text or "")


def _dismiss_windows(page, frame, wins):
    for w in wins:
        btn = next((b for b in w["buttons"] if b["t"] in ("확인", "닫기") and b["id"]), None)
        if not btn:
            raise StopForSafety("예상치 못한 창이 떴는데 닫는 버튼을 찾지 못했습니다. 멈춥니다.")
        safe_click(page, frame, "#" + btn["id"])
        time.sleep(0.5)
    if frame.evaluate(_VISIBLE_WINDOWS_JS):
        raise StopForSafety("예상치 못한 창이 닫히지 않았습니다. 멈춥니다.")


# ---------------------------------------------------------------------------
# 저장 위치 (바탕화면)
# ---------------------------------------------------------------------------

def desktop_dir():
    try:
        buf = ctypes.create_unicode_buffer(260)
        # 0x0010 = CSIDL_DESKTOPDIRECTORY (실제 바탕화면 폴더)
        if ctypes.windll.shell32.SHGetFolderPathW(None, 0x0010, None, 0, buf) == 0 and buf.value:
            return Path(buf.value)
    except Exception:
        pass
    return Path.home() / "Desktop"


def resolve_out_root(out_arg, allow_cloud):
    root = Path(out_arg) if out_arg else desktop_dir() / "홈택스_조회결과"
    low = str(root).lower()
    cloud_words = ("onedrive", "dropbox", "google drive", "googledrive", "mybox")
    if not allow_cloud and any(w in low for w in cloud_words):
        raise StopForSafety(
            "저장 위치가 클라우드 동기화 폴더로 보입니다 (거래처 실자료가 PC 밖으로 나갈 수 있음). "
            "다른 폴더를 --out으로 정하거나, 정말 괜찮다면 --allow-cloud-folder를 붙이세요."
        )
    return root


_BAD_NAME_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def client_folder_name(name, bizno, mode):
    """mode: 'name' = 상호만 / 'both' = 상호_사업자번호 / 'bizno' = 사업자번호만."""
    if mode == "bizno" or not name:
        return bizno
    clean = re.sub(r"\s+", " ", _BAD_NAME_CHARS.sub("", name)).strip().rstrip(". ")[:40]
    if not clean:
        return bizno
    return clean if mode == "name" else f"{clean}_{bizno}"


def assign_folder_names(all_clients, mode):
    """{사업자번호: 폴더 이름}. 'name' 방식에서 상호가 겹치는(윈도우는 대소문자 구분 안 함) 거래처들만
    뒤에 사업자번호를 붙여 구분한다. --only/--limit 와 상관없이 명단 전체로 정해서 이름이 늘 같게 한다."""
    names = {c["bizno"]: client_folder_name(c["name"], c["bizno"], mode) for c in all_clients}
    if mode == "name":
        count = {}
        for n in names.values():
            count[n.lower()] = count.get(n.lower(), 0) + 1
        for c in all_clients:
            if count[names[c["bizno"]].lower()] > 1:
                names[c["bizno"]] = client_folder_name(c["name"], c["bizno"], "both")
    return names


def _write_json(path, obj):
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(obj, ensure_ascii=False, indent=1), encoding="utf-8")
    os.replace(tmp, path)


# ---------------------------------------------------------------------------
# 거래처 명단(xlsx) 읽기 — 외부 패키지 없이 표준 라이브러리(zipfile + XML)만 쓴다.
#
# 명단에는 주민등록번호 열이 같이 있다. 여기서는 '상호(성명)'과 '사업자등록번호' 두 열의 값만
# 꺼내고, 나머지 열(주민등록번호 포함)은 어디에도 저장·출력하지 않는다.
# ---------------------------------------------------------------------------

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PKG_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"

HEADER_BIZNO = "사업자등록번호"
HEADER_NAME = "상호(성명)"


def _col_index(cell_ref):
    letters = re.match(r"[A-Za-z]+", cell_ref).group(0).upper()
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _col_letter(idx):
    s, idx = "", idx + 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def _first_visible_sheet_path(z):
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    targets = {r.get("Id"): r.get("Target") for r in rels.iter(_PKG_REL_NS + "Relationship")}
    for sh in wb.find(_NS + "sheets"):
        if sh.get("state") in ("hidden", "veryHidden"):
            continue
        target = targets.get(sh.get(_REL_NS + "id"))
        if target:
            target = target.lstrip("/")
            return target if target.startswith("xl/") else "xl/" + target
    raise StopForSafety("명단 파일에서 볼 수 있는 시트를 찾지 못했습니다.")


def iter_xlsx_rows(path):
    """첫 번째 보이는 시트를 (행번호, {열번호: 글자}) 로 한 줄씩 내보낸다."""
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(_NS + "si"):
                shared.append("".join(t.text or "" for t in si.iter(_NS + "t")))
        sheet = ET.fromstring(z.read(_first_visible_sheet_path(z)))
        for row in sheet.iter(_NS + "row"):
            cells = {}
            for c in row.findall(_NS + "c"):
                t, v = c.get("t"), c.find(_NS + "v")
                if t == "s" and v is not None and v.text is not None:
                    val = shared[int(v.text)]
                elif t == "inlineStr":
                    val = "".join(x.text or "" for x in c.iter(_NS + "t"))
                elif v is not None and v.text is not None:
                    val = v.text
                else:
                    continue
                if re.fullmatch(r"-?\d+(\.\d+)?[eE][+-]?\d+", val.strip()):
                    val = str(int(float(val)))  # 큰 숫자가 지수 표기로 저장된 경우
                cells[_col_index(c.get("r"))] = val
            yield int(row.get("r")), cells


def biz_no_checksum_ok(d):
    w = [1, 3, 7, 1, 3, 7, 1, 3, 5]
    s = sum(int(d[i]) * w[i] for i in range(9)) + (int(d[8]) * 5) // 10
    return (10 - s % 10) % 10 == int(d[9])


def check_biz_no(raw):
    """(숫자 10자리, None) 또는 (None, 못 쓰는 이유)."""
    d = re.sub(r"\D", "", raw or "")
    if not d:
        return None, "사업자등록번호 칸이 비어 있음"
    if d == "0" * 10:
        return None, "0000000000 (바꾸는 것을 잊은 자리표시 값)"
    if len(d) != 10:
        return None, f"자릿수가 10이 아님({len(d)}자리)"
    if not biz_no_checksum_ok(d):
        return None, "검증번호가 맞지 않음(오타 가능)"
    return d, None


def mask_biz_no(d):
    return "******" + d[-4:]


DEFAULT_CLIENT_LIST = Path.home() / "Desktop" / "2. 거래처 인수인계" / "김은정 부가세신고 거래처.xlsx"


def read_client_list(path):
    bizno_col = name_col = header_row = None
    clients, skipped, dup_rows, seen, blank = [], [], [], set(), 0
    try:
        for rownum, cells in iter_xlsx_rows(path):
            if bizno_col is None:
                if rownum > 20:
                    break
                for ci, v in cells.items():
                    h = re.sub(r"\s+", "", v or "")
                    if h == HEADER_BIZNO:
                        bizno_col = ci
                    elif h == re.sub(r"\s+", "", HEADER_NAME):
                        name_col = ci
                if bizno_col is not None:
                    header_row = rownum
                continue
            raw = cells.get(bizno_col, "").strip()
            name = cells.get(name_col, "").strip() if name_col is not None else ""
            if not raw and not name:
                blank += 1
                continue
            d, why = check_biz_no(raw)
            if why:
                skipped.append((rownum, why))
                continue
            if d in seen:
                dup_rows.append(rownum)
                continue
            seen.add(d)
            clients.append({"row": rownum, "bizno": d, "name": name})
    except (zipfile.BadZipFile, KeyError, ET.ParseError, FileNotFoundError, PermissionError) as e:
        raise StopForSafety(f"명단 파일을 읽지 못했습니다 ({type(e).__name__}). 엑셀에서 열려 있으면 닫고 다시 해보세요.")
    if bizno_col is None:
        raise StopForSafety(f"명단 앞 20행 안에서 '{HEADER_BIZNO}' 머리글을 찾지 못했습니다.")
    return {
        "clients": clients, "skipped": skipped, "duplicates": dup_rows, "blank": blank,
        "header_row": header_row, "bizno_col": _col_letter(bizno_col),
        "name_col": _col_letter(name_col) if name_col is not None else None,
    }


def run_check_list(args):
    info = read_client_list(args.list)
    log("명단 점검 결과 (상호·번호·주민등록번호는 출력하지 않습니다)")
    log(f"  머리글 행: {info['header_row']}행 / 사업자등록번호 열: {info['bizno_col']} / 상호(성명) 열: {info['name_col'] or '없음'}")
    log(f"  조회에 쓸 거래처 수(중복 제거): {len(info['clients'])}")
    log(f"  건너뛴 행: {len(info['skipped'])}개 / 같은 번호가 또 나온 행: {len(info['duplicates'])}개 / 빈 행: {info['blank']}개")
    reasons = {}
    for _, why in info["skipped"]:
        key = re.sub(r"\(\d+자리\)", "", why)
        reasons[key] = reasons.get(key, 0) + 1
    for why, n in reasons.items():
        log(f"    - {why}: {n}개")
    for rownum, why in info["skipped"][:40]:
        log(f"    · {rownum}행: {why}")
    if info["duplicates"]:
        log(f"    · 중복 행 번호: {info['duplicates'][:30]}")


# ---------------------------------------------------------------------------
# 부가가치세 신고자료 통합조회 (세무대리인) 화면
#
# 2026-10-28 화면을 직접 조사해서 확인된 것만 적었다 (화면지도.md 참고).
#  - 작업영역 class는 'w2pageFrame PP HT' (합계표 화면은 'PP ET' — 화면마다 다르다)
#  - 거래처는 사업자등록번호 입력칸에 직접 넣는다 ('수임사업자' 버튼은 거래처 전환이 아니라
#    '납세자등록상황' 창을 새로 여는 버튼이라 절대 누르지 않는다)
#  - 첫 번째 '조회' 버튼이 위쪽 조건에 대한 조회 버튼이다
#  - 결과는 표 3개(사업자 기본사항 / 매출·매입 항목 / 공제·부속서류)에 나온다
# ---------------------------------------------------------------------------

TONGHAP_KEY = "통합조회"
TONGHAP_TITLE = "부가가치세 신고자료 통합조회 서비스"
TONGHAP_WORK_CLASS = "w2pageFrame PP HT"
TONGHAP_BSNO_INPUT = "mf_txppWframe_inputBsno"
TONGHAP_YEAR_INPUT = "mf_txppWframe_edtTxnrmY"
TONGHAP_KI_SELECT = "mf_txppWframe_selectHt"
TONGHAP_RTN_RADIO_NAME = "mf_txppWframe_radioRtnClCd_input"
TONGHAP_QUERY_BUTTON = "mf_txppWframe_trigger113"

KI_OPTIONS = ("1기", "2기")
RTN_OPTIONS = {"예정": 0, "확정": 1, "예정+확정": 2}

_RESULT_STATE_JS = r"""
() => {
  const vis = e => !!(e.offsetWidth || e.offsetHeight);
  const root = document.getElementById('mf_txppWframe');
  const f = Array.from(root.querySelectorAll('table input[type=text]')).filter(vis);
  return {n: f.length, filled: f.filter(i => i.value !== '').length};
}
"""

# 결과 칸을 화면에서 비운다. 조회 뒤 다시 채워지면 "방금 조회로 채워진 값"이라는 증거가 된다.
_RESULT_BLANK_JS = r"""
() => {
  const vis = e => !!(e.offsetWidth || e.offsetHeight);
  const root = document.getElementById('mf_txppWframe');
  Array.from(root.querySelectorAll('table input[type=text]')).filter(vis).forEach(i => { i.value = ''; });
}
"""

_TONGHAP_READBACK_JS = r"""
() => {
  const ki = document.getElementById('mf_txppWframe_selectHt');
  const rtn = Array.from(document.querySelectorAll('input[name="mf_txppWframe_radioRtnClCd_input"]')).find(r => r.checked);
  const lab = rtn ? document.querySelector('label[for="' + rtn.id + '"]') : null;
  return {
    year: document.getElementById('mf_txppWframe_edtTxnrmY').value,
    ki: ki.options[ki.selectedIndex] ? ki.options[ki.selectedIndex].text : null,
    rtn: lab ? lab.innerText.trim() : null,
    bizno: document.getElementById('mf_txppWframe_inputBsno').value.replace(/\D/g, '')
  };
}
"""

# 화면의 표를 칸 합치기(colspan/rowspan)까지 그대로 읽는다. 입력칸이 들어 있는 칸은 입력값을 읽는다.
_TABLES_JS = r"""
() => {
  const vis = e => !!(e.offsetWidth || e.offsetHeight);
  const root = document.getElementById('mf_txppWframe');
  const cellText = c => {
    const inp = c.querySelector('input[type=text]');
    const t = inp ? inp.value : (c.innerText || '');
    return t.replace(/\s+/g, ' ').trim();
  };
  const heads = Array.from(root.querySelectorAll('h1,h2,h3,h4,.tit_h1,.tit_h2,.tit_h3')).filter(vis);
  const capOf = t => {
    let cap = null;
    for (const hd of heads) {
      if (hd.compareDocumentPosition(t) & Node.DOCUMENT_POSITION_FOLLOWING) {
        cap = (hd.innerText || '').replace(/\s+/g, ' ').trim();
      }
    }
    return cap;
  };
  return Array.from(root.querySelectorAll('table')).filter(vis).map(t => ({
    id: t.id,
    caption: capOf(t),
    rows: Array.from(t.rows).filter(r => r.offsetHeight > 0).map(r =>
      Array.from(r.cells).filter(c => c.offsetWidth > 0 && c.offsetHeight > 0).map(c => ({
        tag: c.tagName.toLowerCase(), text: cellText(c), colspan: c.colSpan, rowspan: c.rowSpan
      })))
  }));
}
"""


def ensure_on_tonghap_screen(frame):
    cls = _work_area_class(frame)
    has_title = frame.evaluate(
        """
        (title) => Array.from(document.querySelectorAll('h1,h2,h3,h4,.tit_h1,.tit_h2,.tit_h3'))
          .some(e => (e.offsetWidth || e.offsetHeight) && (e.innerText || '').includes(title))
        """,
        TONGHAP_TITLE,
    )
    if not _class_has_all(cls, TONGHAP_WORK_CLASS) or not has_title:
        raise StopForSafety(
            f"로그인은 돼 있지만 지금 화면이 '{TONGHAP_TITLE}'이(가) 아닌 것 같습니다 "
            f"(화면class={cls!r}, 제목확인={has_title}). 메뉴에서 그 화면으로 옮겨 주세요."
        )


def _type_into(page, frame, selector, text):
    safe_click(page, frame, selector)
    page.keyboard.press("Control+A")
    page.keyboard.type(text, delay=50)
    time.sleep(0.3)


def _set_radio_safe(page, frame, name, index):
    ids = frame.evaluate(f"() => Array.from(document.querySelectorAll('input[name=\"{name}\"]')).map(r => r.id)")
    if index >= len(ids):
        raise StopForSafety(f"{name} 라디오에 {index}번째 옵션이 없습니다 (현재 {len(ids)}개뿐).")
    target_id = ids[index]
    safe_click(page, frame, f"label[for='{target_id}']")
    time.sleep(0.3)
    if not frame.evaluate(f"() => document.getElementById('{target_id}').checked"):
        raise StopForSafety(f"{name} 조건을 바꾸려 했지만 실제로 반영되지 않았습니다. 멈춥니다.")


def set_tonghap_conditions(page, frame, year, ki, rtn):
    if not re.fullmatch(r"20\d\d", str(year)):
        raise StopForSafety(f"연도 형식이 이상합니다: {year!r} (예: 2026)")
    if ki not in KI_OPTIONS:
        raise StopForSafety(f"기 값을 모릅니다: {ki!r} (가능: {list(KI_OPTIONS)})")
    if rtn not in RTN_OPTIONS:
        raise StopForSafety(f"신고구분 값을 모릅니다: {rtn!r} (가능: {list(RTN_OPTIONS)})")

    _type_into(page, frame, f"#{TONGHAP_YEAR_INPUT}", str(year))
    _select_and_verify(frame, TONGHAP_KI_SELECT, ki)
    _set_radio_safe(page, frame, TONGHAP_RTN_RADIO_NAME, RTN_OPTIONS[rtn])

    rb = frame.evaluate(_TONGHAP_READBACK_JS)
    if rb["year"] != str(year) or rb["ki"] != ki or rb["rtn"] != rtn:
        raise StopForSafety(
            f"과세기간 조건이 요청과 다르게 들어갔습니다 (요청: {year}년 {ki} {rtn} / "
            f"화면: {rb['year']}년 {rb['ki']} {rb['rtn']}). 값을 내놓지 않고 멈춥니다."
        )
    return rb


def query_tonghap(page, frame, bizno, dialogs):
    """번호 입력 -> 결과 칸 비우기 -> 조회 -> 다시 채워질 때까지 기다리기 -> 표 읽기."""
    _type_into(page, frame, f"#{TONGHAP_BSNO_INPUT}", bizno)
    if frame.evaluate(_TONGHAP_READBACK_JS)["bizno"] != bizno:
        raise StopForSafety("사업자등록번호가 입력칸에 요청한 대로 들어가지 않았습니다. 멈춥니다.")

    frame.evaluate(_RESULT_BLANK_JS)
    blank = frame.evaluate(_RESULT_STATE_JS)
    if blank["n"] == 0 or blank["filled"] != 0:
        raise StopForSafety(f"결과 칸을 비우지 못했습니다 (칸 {blank['n']}개 중 {blank['filled']}개가 안 비워짐). 멈춥니다.")

    dialogs.clear()
    safe_click(page, frame, f"#{TONGHAP_QUERY_BUTTON}")

    def poll():
        st = frame.evaluate(_RESULT_STATE_JS)
        wins = frame.evaluate(_VISIBLE_WINDOWS_JS)
        if (st["n"] > 0 and st["filled"] == st["n"]) or wins or dialogs:
            return {"st": st, "wins": wins}
        return None

    res = _wait_until(poll, 15)
    if res and (res["wins"] or dialogs):
        msg = _mask_digits(" / ".join([w["text"] for w in res["wins"]] + list(dialogs)))[:100]
        if res["wins"]:
            _dismiss_windows(page, frame, res["wins"])
        raise ClientFailed(f"화면 메시지가 떴습니다: {msg}")
    if not res:
        raise ClientFailed("조회 후 15초 안에 결과 칸이 다시 채워지지 않았습니다 (값을 내놓지 않음)")

    rb = frame.evaluate(_TONGHAP_READBACK_JS)
    if rb["bizno"] != bizno:
        raise StopForSafety("조회 뒤 입력칸의 사업자등록번호가 바뀌어 있습니다. 멈춥니다.")
    tables = frame.evaluate(_TABLES_JS)
    if not tables:
        raise ClientFailed("결과 표를 화면에서 찾지 못했습니다 (값을 내놓지 않음)")
    return tables, rb


def capture_work_area(page, frame):
    ensure_no_menu_overlay(frame)
    frame.evaluate("window.scrollTo(0, 0)")
    png = page.locator("#" + WORK_AREA_ID).screenshot(timeout=20000)
    frame.evaluate("window.scrollTo(0, 0)")
    return png


# ---------------------------------------------------------------------------
# 거래처별 폴더 저장
#
# 폴더 안에: <화면>.png (화면 캡처) + 내려받거나 만든 엑셀 — 열어 볼 수 있는 파일만 둔다 (.json 은 저장하지 않음).
# 한페이지.html 은 더 만들지 않는다 (사용자 요청 2026-11-06).
# ---------------------------------------------------------------------------

CASH_SUM_KEY = "현금영수증매출총액"
CARD_SUM_KEY = "신용카드판매대행매출"
EXPORT_SUM_KEY = "수출실적명세서"

SCREENS = [
    (TONGHAP_KEY, "부가가치세 신고자료 통합조회 서비스 (세무대리인)"),
    ("부가세합계표", "부가가치세신고용 합계표조회"),
    (CASH_SUM_KEY, "현금영수증 매출총액(매출내역누계) 조회"),
    (CARD_SUM_KEY, "신용카드/판매(결제)대행 매출자료 조회"),
    (EXPORT_SUM_KEY, "수출실적명세서 조회"),
]

# 화면을 조사해서 구현까지 마친 화면만 True. False인 화면은 예약 실행 기록에 "조회하지 않는 화면"으로 남긴다.
SCREEN_READY = {
    TONGHAP_KEY: True, "부가세합계표": False, CASH_SUM_KEY: True,
    CARD_SUM_KEY: True, EXPORT_SUM_KEY: True,
}


def save_screen_result(client_dir, key, png):
    client_dir.mkdir(parents=True, exist_ok=True)
    (client_dir / f"{key}.png").write_bytes(png)


def save_print_pdf(page, client_dir, key):
    """CDP Page.printToPDF로 화면을 PDF로 저장한다 (인쇄하기 버튼 대용)."""
    cdp = page.context.new_cdp_session(page)
    try:
        result = cdp.send("Page.printToPDF", {
            "printBackground": True,
            "paperWidth": 11.69,
            "paperHeight": 8.27,
            "marginTop": 0.4,
            "marginBottom": 0.4,
            "marginLeft": 0.4,
            "marginRight": 0.4,
            "landscape": True,
        })
        pdf_data = base64.b64decode(result["data"])
        client_dir.mkdir(parents=True, exist_ok=True)
        pdf_path = client_dir / f"{key}_인쇄.pdf"
        pdf_path.write_bytes(pdf_data)
        return pdf_path.name
    finally:
        cdp.detach()


# ---------------------------------------------------------------------------
# 진행 상황 기록 (중간에 멈춰도 이어서 하려고)
# ---------------------------------------------------------------------------

def load_status(out_root):
    # 메모장으로 열리는 .txt 로 둔다 (내용은 같은 형식). 옛 이름(.json)만 있으면 그걸 읽는다.
    for name in ("_진행상황.txt", "_진행상황.json"):
        p = out_root / name
        if p.exists():
            return json.loads(p.read_text(encoding="utf-8"))
    return {}


def save_status(out_root, status):
    _write_json(out_root / "_진행상황.txt", status)


# ---------------------------------------------------------------------------
# 화면별 부품 + 공통 실행부
#
# 화면마다 다른 것(조건 넣기 / 조회 / 읽기)은 Screen 클래스에 두고, 거래처를 돌며 저장하는
# 공통 흐름은 run_batch 에 둔다. 새 화면은 Screen 클래스 하나를 더 만들면 된다.
# ---------------------------------------------------------------------------

class ResponseWatcher:
    """조회 버튼을 누른 뒤 서버 응답(wqAction.do POST)이 실제로 왔는지 센다.
    결과가 입력칸이 아니라 표인 화면에서 '방금 조회가 끝났다'는 증거로 쓴다
    (조회 한 번에 표마다 요청이 한 번씩 나가고 모두 200이었다 — 신용카드/판매대행 화면)."""

    def __init__(self, page):
        self.statuses = []
        self.responses = []
        page.on("response", self._on_response)

    def _on_response(self, r):
        try:
            if r.request.method == "POST" and "wqAction.do" in r.url:
                self.statuses.append(r.status)
                self.responses.append(r)
        except Exception:
            pass

    def reset(self):
        self.statuses.clear()
        self.responses.clear()

    def json_bodies(self):
        """조회 응답 본문(JSON)들. 화면 표와 별개로 서버가 알려준 줄 수 같은 독립 증거를 읽을 때 쓴다."""
        out = []
        for r in list(self.responses):
            try:
                out.append(json.loads(r.text()))
            except Exception:
                pass
        return out

    def wait(self, page, timeout=20.0, quiet=1.2):
        end = time.time() + timeout
        last_n, last_change = -1, time.time()
        while time.time() < end:
            page.wait_for_timeout(300)   # time.sleep 은 Playwright 이벤트(응답·대화상자)를 처리하지 못한다
            n = len(self.statuses)
            if n != last_n:
                last_n, last_change = n, time.time()
            elif n > 0 and time.time() - last_change >= quiet:
                break
        return list(self.statuses)


class RunContext:
    def __init__(self, page):
        self.dialogs = []
        page.on("dialog", lambda d: (self.dialogs.append(d.message[:80]), d.dismiss()))
        self.watcher = ResponseWatcher(page)


_GRID_OVERFLOW_JS = r"""
() => {
  const vis = e => !!(e.offsetWidth || e.offsetHeight);
  const root = document.getElementById('mf_txppWframe');
  return Array.from(root.querySelectorAll('table[id$="_body_table"]')).filter(vis).some(t => {
    let n = t.parentElement, k = 0;
    while (n && n !== root && k < 8) {
      const cs = getComputedStyle(n);
      if ((cs.overflowY === 'auto' || cs.overflowY === 'scroll') && n.scrollHeight > n.clientHeight + 2) return true;
      n = n.parentElement; k++;
    }
    return false;
  });
}
"""


def work_heading_visible(frame, text):
    """작업영역 안의 보이는 소제목 중에 이 글자가 들어 있는 것이 있는지 (띄어쓰기 무시)."""
    return frame.evaluate(
        """
        (t) => {
          const root = document.getElementById('mf_txppWframe');
          return !!root && Array.from(root.querySelectorAll('h1,h2,h3,h4,.tit_h1,.tit_h2,.tit_h3'))
            .some(e => (e.offsetWidth || e.offsetHeight) && (e.innerText || '').replace(/\\s+/g, '').includes(t.replace(/\\s+/g, '')));
        }
        """,
        text,
    )


def read_result_tables(frame):
    tables = frame.evaluate(_TABLES_JS)
    return [t for t in tables if any(c["text"] for r in t["rows"] for c in r)]


def raise_if_screen_message(page, frame, ctx):
    """조회 뒤 알림창/대화상자가 떠 있으면 닫고 이 거래처만 실패로 처리한다."""
    wins = frame.evaluate(_VISIBLE_WINDOWS_JS)
    if wins or ctx.dialogs:
        msg = _mask_digits(" / ".join([w["text"] for w in wins] + list(ctx.dialogs)))[:100]
        if wins:
            _dismiss_windows(page, frame, wins)
        raise ClientFailed(f"화면 메시지가 떴습니다: {msg}")


class TonghapScreen:
    key = TONGHAP_KEY
    title = SCREENS[0][1]
    menu_item_id = "menuAtag_0602190000"   # 세무대리/납세관리 > 부가가치세 신고자료 통합조회 서비스 (2026-11 메뉴 개편)
    arrival_heading = TONGHAP_TITLE

    @staticmethod
    def add_arguments(p):
        p.add_argument("--year", required=True, help="과세기간 연도, 예: 2026 (기본값 없음: 꼭 적는다)")
        p.add_argument("--ki", required=True, choices=list(KI_OPTIONS), help="1기 / 2기")
        p.add_argument("--rtn", required=True, choices=list(RTN_OPTIONS), help="예정 / 확정 / 예정+확정")

    def conditions(self, args):
        return {"year": args.year, "ki": args.ki, "rtn": args.rtn}

    def describe(self, cond):
        return f"{cond['year']}년 {cond['ki']} {cond['rtn']}"

    def ensure_on(self, frame):
        ensure_on_tonghap_screen(frame)

    def apply(self, page, frame, cond):
        set_tonghap_conditions(page, frame, cond["year"], cond["ki"], cond["rtn"])

    def query(self, page, frame, bizno, cond, ctx):
        tables, rb = query_tonghap(page, frame, bizno, ctx.dialogs)
        return tables, {"과세기간": f"{rb['year']}년 {rb['ki']}", "신고구분": rb["rtn"]}

    def after_save(self, page, frame, client_dir, tables, cond, ctx):
        name = save_print_pdf(page, client_dir, self.key)
        return [name] if name else []


def _is_done(rec, desc, redo_before):
    """같은 조건으로 이미 끝난 거래처인가. redo_before(ISO 시각)가 있으면 그보다 이전에 끝낸 기록은 '끝난 것'으로 안 친다."""
    if not rec or rec.get("결과") not in ("성공", "자료없음") or rec.get("조건") != desc:
        return False
    return not redo_before or rec.get("시각", "") >= redo_before


def _is_conflict(rec, desc, redo_before):
    """다른 기간(또는 기간 기록 없음)의 성공 결과가 있어서 덮어쓰면 안 되는 거래처인가. redo_before 이전 기록은 덮어써도 된다."""
    if not rec or rec.get("결과") != "성공" or rec.get("조건") == desc:   # 자료없음은 파일이 없어 덮어쓸 것도 없다
        return False
    return not redo_before or rec.get("시각", "") >= redo_before


def run_batch(args, screen):
    redo_before = getattr(args, "redo_before", None)
    if redo_before:
        try:
            datetime.fromisoformat(redo_before)
        except ValueError:
            raise StopForSafety(f"--redo-before 형식이 이상합니다: {redo_before!r} (예: 2027-01-17T00:00:00)")
    if bool(args.list) == bool(args.biz_no):
        raise StopForSafety("--list(명단 파일)와 --biz-no(번호 하나) 중 정확히 하나만 주세요.")
    cond = screen.conditions(args)   # 기간 조건이 이상하면 여기서 바로 멈춘다 (브라우저를 건드리기 전)

    if args.list:
        info = read_client_list(args.list)
        clients = info["clients"]
        log(f"명단에서 조회할 거래처 {len(clients)}곳 (건너뛴 행 {len(info['skipped'])}개, 중복 {len(info['duplicates'])}개)")
    else:
        d, why = check_biz_no(args.biz_no)
        if why:
            if d is None and why.startswith("0000000000"):
                raise StopForSafety("사업자등록번호를 바꿔서 다시 붙여넣어 주세요")
            raise StopForSafety(f"사업자등록번호를 쓸 수 없습니다: {why}")
        clients = [{"row": 0, "bizno": d, "name": args.name or ""}]
        if not args.name:
            # 번호만 줘도 폴더 이름은 회사명(상호)으로 만든다: 명단에서 상호를 찾는다.
            # 명단 전체로 폴더 이름을 정해야 --list 로 돌릴 때와 같은 폴더가 된다 (겹치는 상호는 번호가 붙는다).
            listed = read_client_list(DEFAULT_CLIENT_LIST)["clients"] if DEFAULT_CLIENT_LIST.exists() else []
            if not any(c["bizno"] == d and c["name"] for c in listed):
                raise StopForSafety(
                    "이 사업자번호의 상호(회사명)를 명단에서 찾지 못해 폴더 이름을 정할 수 없습니다. "
                    "--name 으로 상호를 함께 주세요. (사업자번호로 된 폴더는 만들지 않습니다)")
            clients = listed
            args.only = [d]

    folder_of = assign_folder_names(clients, args.folder_name)   # 명단 전체로 정해서 --only/--limit 와 무관하게 늘 같다
    if args.only:
        wanted = {re.sub(r"\D", "", x) for x in args.only}
        clients = [c for c in clients if c["bizno"] in wanted]
    if args.limit:
        clients = clients[: args.limit]
    if not clients:
        raise StopForSafety("조회할 거래처가 없습니다.")

    out_root = resolve_out_root(args.out, args.allow_cloud_folder)
    desc = screen.describe(cond)

    if args.dry_run:
        names = [folder_of[c["bizno"]] for c in clients]
        prev_all = (load_status(out_root) if out_root.exists() else {}).get(screen.key, {})
        same = sum(1 for c in clients if _is_done(prev_all.get(c["bizno"]), desc, redo_before))
        longest = max(len(str(out_root / n / f"{screen.key}.png")) for n in names)
        log("점검만 합니다 (--dry-run): 크롬을 건드리지 않고, 폴더도 만들지 않습니다.")
        log(f"  화면: {screen.title} / 조회 조건: {desc}")
        log(f"  조회 대상 거래처: {len(clients)}곳 / 저장 폴더 이름 중복: {len(names) - len(set(names))}개")
        log(f"  같은 조건으로 이미 끝난 곳: {same}곳 -> 실제 조회할 곳: {len(clients) - same}곳 (--redo 이면 전부)")
        log(f"  가장 긴 저장 경로: {longest}자 (윈도우 한도 260자)")
        log(f"  저장 위치: {out_root}")
        return

    out_root.mkdir(parents=True, exist_ok=True)
    status = load_status(out_root)
    status.setdefault(screen.key, {})

    # 같은 폴더에 다른 기간(또는 기간 기록이 없는 옛 결과)이 이미 있으면, 조용히 덮어쓰지 않고 멈춘다.
    if not args.redo:
        conflict = [c for c in clients if _is_conflict(status[screen.key].get(c["bizno"]), desc, redo_before)]
        if conflict:
            raise StopForSafety(
                f"저장 폴더에 이번 조건({desc})과 다른 기간의 결과가 이미 {len(conflict)}곳 있습니다. "
                "덮어쓰지 않도록 멈춥니다. 다른 --out 폴더를 쓰거나, 덮어써도 되면 --redo 를 붙이세요."
            )

    total = len(clients)
    done = failed = skipped = nodata = 0
    stopped = False
    streak_reason, streak_n = None, 0   # 같은 사유로 연속 실패하면 조건·화면 문제일 가능성이 커서 멈춘다

    with sync_playwright() as p:
        browser = connect_chrome(p)
        try:
            ctx, page = find_hometax_page(browser)
            if page is None:
                raise StopForSafety("홈택스 탭을 찾지 못했습니다.")
            frame = page.main_frame
            run_ctx = RunContext(page)

            for i, c in enumerate(clients, 1):
                label = f"[{i}/{total}] 사업자번호 {mask_biz_no(c['bizno'])}"
                client_dir = out_root / folder_of[c["bizno"]]
                prev = status[screen.key].get(c["bizno"])
                if (_is_done(prev, desc, redo_before) and not args.redo
                        and (prev.get("결과") == "자료없음" or (client_dir / f"{screen.key}.png").exists())):
                    log(f"{label} → 이미 끝난 거래처라 건너뜀 (다시 하려면 --redo)")
                    skipped += 1
                    continue

                try:
                    ensure_logged_in(frame)
                    screen.ensure_on(frame)
                    ensure_no_menu_overlay(frame)
                    screen.apply(page, frame, cond)
                    tables, _basis = screen.query(page, frame, c["bizno"], cond, run_ctx)
                    png = capture_work_area(page, frame)
                    save_screen_result(client_dir, screen.key, png)
                    files = []
                    after = getattr(screen, "after_save", None)
                    if after:   # 화면이 따로 저장할 파일(예: 엑셀 내려받기)이 있으면 받는다
                        files = after(page, frame, client_dir, tables, cond, run_ctx)
                    status[screen.key][c["bizno"]] = {
                        "결과": "성공", "조건": desc, "시각": datetime.now().isoformat(timespec="seconds")}
                    if files:
                        status[screen.key][c["bizno"]]["파일"] = files
                    done += 1
                    streak_reason, streak_n = None, 0
                    log(f"{label} → 성공 ({desc})" + (f" + 엑셀 {len(files)}개" if files else ""))
                except NoDataForClient as e:
                    status[screen.key][c["bizno"]] = {
                        "결과": "자료없음", "조건": desc, "근거": str(e), "시각": datetime.now().isoformat(timespec="seconds")}
                    nodata += 1
                    streak_reason, streak_n = None, 0
                    log(f"{label} → 자료 없음 ({desc}) — 저장하지 않음")
                except ClientFailed as e:
                    status[screen.key][c["bizno"]] = {
                        "결과": "실패", "사유": str(e), "시각": datetime.now().isoformat(timespec="seconds")}
                    failed += 1
                    log(f"{label} → 실패: {e}")
                    streak_n = streak_n + 1 if str(e) == streak_reason else 1
                    streak_reason = str(e)
                    if streak_n >= 3:
                        save_status(out_root, status)
                        raise StopForSafety(
                            "같은 사유로 3곳 연속 실패해서 멈춥니다 (조건이나 화면 문제일 수 있어 계속하지 않습니다)."
                        )
                save_status(out_root, status)

                if i < total:
                    wait_s = random.uniform(*REST_BETWEEN_QUERIES_SECONDS)
                    time.sleep(wait_s)
        except StopForSafety as e:
            log(f"[중단] {e}")
            stopped = True
        except Exception as e:
            first_line = (str(e).splitlines() or [""])[0][:150]
            log(f"[중단] 예상하지 못한 오류로 멈춥니다 ({type(e).__name__}): {first_line}")
            stopped = True
        finally:
            save_status(out_root, status)
            browser.close()

    log("")
    log(f"끝: 성공 {done}곳 / 자료 없음 {nodata}곳 / 실패 {failed}곳 / 건너뜀 {skipped}곳 / 전체 {total}곳"
        + (" (중간에 멈춤 — 다시 실행하면 끝난 곳은 건너뛰고 이어서 합니다)" if stopped else ""))
    log(f"저장 위치: {out_root}")
    if stopped:
        sys.exit(1)
    if failed:
        sys.exit(2)


def run_tonghap(args):
    run_batch(args, TonghapScreen())


# ---------------------------------------------------------------------------
# 신용카드/판매(결제)대행 매출자료 조회 (세무대리 공통 조회 메뉴, 코드 4801140000)
#
# 2026-10-28 화면을 직접 조사해서 확인된 것:
#  - 조건: 결제년도(select) + 시작 분기 + 끝 분기(select 2개) + 사업자등록번호 입력칸(최대 10자리)
#  - 조회 버튼은 하나(trigger163). '엑셀 내려받기' 버튼 3개는 파일을 내려받으므로 절대 누르지 않는다.
#  - 결과는 표 3개: 매출자료구분별 합계 / 신용카드·제로(온누리)페이 매출 / 판매(결제)대행 매출
#  - 조회 한 번에 서버 요청(wqAction.do POST)이 표마다 한 번씩 나가고, 응답이 다 와야 조회가 끝난 것이다.
#    결과가 입력칸이 아니라 표라서 '칸 비우고 다시 채워지는지'는 못 쓰고, 이 응답 확인을 쓴다.
#  - 이 화면의 작업영역 class는 이전 화면의 낱말이 남아 있어 화면 식별에 쓸 수 없다. 소제목으로 확인한다.
# 기간은 분기 단위다 (통합조회는 1기/2기, 이 화면은 결제년도 + 분기 범위) — 결과에 본 기간을 꼭 함께 저장한다.
# ---------------------------------------------------------------------------

CARD_TITLE_HINT = "신용카드, 제로(온누리)페이 매출자료 조회"
CARD_BSNO_INPUT = "mf_txppWframe_edtTxprDscmNo1"
CARD_YEAR_SELECT = "mf_txppWframe_selectStlYr"
CARD_QFROM_SELECT = "mf_txppWframe_selectQrtFrom"
CARD_QTO_SELECT = "mf_txppWframe_selectQrtTo"
CARD_QUERY_BUTTON = "mf_txppWframe_trigger163"
QUARTERS = ("1분기", "2분기", "3분기", "4분기")

_CARD_READBACK_JS = r"""
() => {
  const t = id => { const s = document.getElementById(id); return s && s.options[s.selectedIndex] ? s.options[s.selectedIndex].text : null; };
  return {
    year: t('mf_txppWframe_selectStlYr'), qfrom: t('mf_txppWframe_selectQrtFrom'), qto: t('mf_txppWframe_selectQrtTo'),
    bizno: document.getElementById('mf_txppWframe_edtTxprDscmNo1').value.replace(/\D/g, '')
  };
}
"""


class CardScreen:
    key = CARD_SUM_KEY
    title = dict(SCREENS)[CARD_SUM_KEY]
    menu_item_id = "menuAtag_0602060000"   # 세무대리/납세관리 > 신용카드/판매(결제)대행매출자료조회 (2026-11 메뉴 개편)
    arrival_heading = CARD_TITLE_HINT

    @staticmethod
    def add_arguments(p):
        p.add_argument("--year", required=True, help="결제년도, 예: 2026 (기본값 없음: 꼭 적는다)")
        p.add_argument("--from-q", required=True, choices=list(QUARTERS), help="시작 분기")
        p.add_argument("--to-q", required=True, choices=list(QUARTERS), help="끝 분기")
        p.add_argument("--no-excel", action="store_true",
                       help="표마다 있는 '엑셀 내려받기' 파일을 받지 않는다 (기본은 자료가 있는 표의 엑셀을 받아 저장)")

    def conditions(self, args):
        if not re.fullmatch(r"20\d\d", str(args.year)):
            raise StopForSafety(f"연도 형식이 이상합니다: {args.year!r} (예: 2026)")
        fi, ti = QUARTERS.index(args.from_q), QUARTERS.index(args.to_q)
        # 화면이 직접 알려준 규칙(2026-10-28): "분기별 또는 상반기(1분기~2분기), 하반기(3분기~4분기)로
        # 선택하면 조회 가능합니다." 1분기~4분기처럼 통째로는 조회되지 않는다.
        if not (fi == ti or (fi, ti) in ((0, 1), (2, 3))):
            raise StopForSafety(
                f"{args.from_q}~{args.to_q}는 이 화면에서 조회되지 않는 기간입니다. "
                "한 분기씩(예: 3분기~3분기), 상반기(1분기~2분기), 하반기(3분기~4분기)만 가능합니다."
            )
        return {"year": str(args.year), "from_q": args.from_q, "to_q": args.to_q,
                "excel": not getattr(args, "no_excel", False)}

    def describe(self, cond):
        return f"{cond['year']}년 {cond['from_q']}~{cond['to_q']}"

    def ensure_on(self, frame):
        if not work_heading_visible(frame, CARD_TITLE_HINT):
            raise StopForSafety(
                f"로그인은 돼 있지만 지금 화면이 '{self.title}'이(가) 아닌 것 같습니다 "
                f"(소제목 '{CARD_TITLE_HINT}'이 안 보임). 메뉴에서 그 화면으로 옮겨 주세요."
            )

    def apply(self, page, frame, cond):
        _select_and_verify(frame, CARD_YEAR_SELECT, cond["year"])
        # 시작 분기가 끝 분기보다 뒤인 순간이 생기지 않게: 끝을 4분기로 올리고 -> 시작 -> 끝 순서로 맞춘다.
        _select_and_verify(frame, CARD_QTO_SELECT, "4분기")
        _select_and_verify(frame, CARD_QFROM_SELECT, cond["from_q"])
        _select_and_verify(frame, CARD_QTO_SELECT, cond["to_q"])

    def query(self, page, frame, bizno, cond, ctx):
        # 번호가 입력칸에 안 들어가는 일이 간헐적으로 있어서, 읽어 확인하고 몇 번 다시 시도한다 (틀린 번호로는 조회하지 않는다).
        for _ in range(3):
            _type_into(page, frame, f"#{CARD_BSNO_INPUT}", bizno)
            if frame.evaluate(_CARD_READBACK_JS)["bizno"] == bizno:
                break
            page.wait_for_timeout(700)
        else:
            raise StopForSafety("사업자등록번호가 입력칸에 3번 시도해도 요청한 대로 들어가지 않았습니다. 멈춥니다.")

        ctx.dialogs.clear()
        ctx.watcher.reset()
        safe_click(page, frame, f"#{CARD_QUERY_BUTTON}")
        statuses = ctx.watcher.wait(page)

        raise_if_screen_message(page, frame, ctx)
        if not statuses:
            raise ClientFailed("조회 후 20초 안에 서버 응답이 오지 않았습니다 (값을 내놓지 않음)")
        bad = [s for s in statuses if s >= 400]
        if bad:
            raise ClientFailed(f"서버 응답이 오류입니다 (상태 {bad}). 값을 내놓지 않음")

        rb = frame.evaluate(_CARD_READBACK_JS)
        if (rb["year"], rb["qfrom"], rb["qto"], rb["bizno"]) != (cond["year"], cond["from_q"], cond["to_q"], bizno):
            raise StopForSafety(
                f"조회 뒤 화면의 조건이 요청과 다릅니다 (화면: {rb['year']}년 {rb['qfrom']}~{rb['qto']}). 멈춥니다."
            )
        if frame.evaluate(_GRID_OVERFLOW_JS):
            raise ClientFailed("결과 표가 스크롤될 만큼 길어서 일부만 읽혔을 수 있습니다 (값을 내놓지 않음)")
        tables = read_result_tables(frame)
        if not tables:
            raise ClientFailed("결과 표를 화면에서 찾지 못했습니다 (값을 내놓지 않음)")
        return tables, {"결제년도": rb["year"], "분기": f"{rb['qfrom']} ~ {rb['qto']}"}

    def after_save(self, page, frame, client_dir, tables, cond, ctx):
        """자료가 있는 표의 '엑셀 내려받기' 파일을 거래처 폴더에 저장한다. 저장한 파일 이름 목록을 돌려준다."""
        if not cond.get("excel"):
            return []
        saved = []
        for prefix, button_id, stem in CARD_EXCEL_BUTTONS:
            if not _section_has_data(tables, prefix):
                continue   # 자료 없는 표는 누르지 않는다 (홈택스가 '조회된 데이터가 없습니다.'라고 알린다)
            name = download_excel(page, frame, ctx, button_id, stem, client_dir)
            if name:
                saved.append(name)
            page.wait_for_timeout(2500)   # 내려받기 사이에도 쉰다
        return saved


# 신용카드/판매(결제)대행 화면의 '엑셀 내려받기' 버튼 3개는 표마다 하나다 (표 소제목과 위치로 짝을 확인함, 2026-10-28).
# 사용자 요청(2026-10-28)으로 파일 내려받기를 허용했다. 자료가 있는 표의 버튼만 누른다.
CARD_EXCEL_BUTTONS = (
    # (표 소제목 앞부분, 버튼 id, 저장 파일 이름)
    ("매출자료구분별 합계조회", "mf_txppWframe_trigger166", "신용카드판매대행_매출자료구분별합계"),
    ("신용카드, 제로(온누리)페이 매출자료 조회", "mf_txppWframe_trigger165", "신용카드판매대행_신용카드제로페이매출"),
    ("판매(결제)대행 매출자료 조회", "mf_txppWframe_trigger164", "신용카드판매대행_판매결제대행매출"),
)
_EXCEL_SIGNATURES = (b"\xd0\xcf\x11\xe0", b"PK\x03\x04")   # 옛 엑셀(.xls) / 새 엑셀(.xlsx) 파일 표지


def _section_has_data(tables, caption_prefix):
    want = caption_prefix.replace(" ", "")
    for t in tables:
        if (t.get("caption") or "").replace(" ", "").startswith(want):
            for r in t["rows"]:
                for c in r:
                    if re.fullmatch(r"-?[\d,]+", c["text"]) and c["text"].replace(",", "").strip("0-"):
                        return True
    return False


def download_excel(page, frame, ctx, button_id, stem, client_dir):
    """버튼을 눌러 내려받은 엑셀을 client_dir/<stem>.xls(x) 로 저장하고 파일 이름을 돌려준다.
    '조회된 데이터가 없습니다.' 알림이 뜨면 None. 파일이 비었거나 엑셀 형식이 아니면 저장하지 않고 이 거래처를 실패 처리한다."""
    # 2026-10-28 시험: 내려받기가 간헐적으로 안 오는 때가 있었다 (같은 버튼이 어떤 때는 0.5초에 되고 어떤 때는 안 됨,
    # 원인 모름 — 시간 규칙도 아니었다). 실패하면 12초 쉬고 다시 누르는 것이 통했다 (30초 쉬는 시험 6번 모두 2번째에 성공,
    # 12초 쉬는 시험 6번 중 5번은 첫 시도, 1번은 2번째 시도에 성공). 그래서 최대 4번까지 다시 시도한다.
    dl = None
    for attempt in range(1, 5):
        ctx.dialogs.clear()
        try:
            with page.expect_download(timeout=10000) as info:
                safe_click(page, frame, f"#{button_id}")
            dl = info.value
            break
        except PlaywrightTimeout:
            if any("조회된 데이터가 없습니다" in m for m in ctx.dialogs):
                return None
            if attempt < 4:
                page.wait_for_timeout(12000)
    if dl is None:
        raise ClientFailed(
            f"'엑셀 내려받기'를 4번 눌러도 파일이 오지 않았습니다 ({stem}) — 홈택스나 크롬이 연속 내려받기를 막고 있을 수 있습니다 "
            "(크롬 창에 허용 안내가 있으면 '허용'을 눌러 주세요)")
    if dl.failure():
        raise ClientFailed(f"엑셀 내려받기에 실패했습니다 ({stem}): {dl.failure()}")
    ext = Path(dl.suggested_filename).suffix.lower()
    if ext not in (".xls", ".xlsx"):
        raise ClientFailed(f"내려받은 파일 형식이 엑셀이 아닙니다 ({stem}, 확장자 {ext!r})")
    dest = client_dir / f"{stem}{ext}"
    dl.save_as(str(dest))
    if dest.stat().st_size == 0 or not any(dest.read_bytes()[:4].startswith(s) for s in _EXCEL_SIGNATURES):
        dest.unlink(missing_ok=True)
        raise ClientFailed(f"내려받은 파일이 비었거나 엑셀 형식 표지가 맞지 않습니다 ({stem}) — 저장하지 않음")
    return dest.name


def run_card_sum(args):
    run_batch(args, CardScreen())


# ---------------------------------------------------------------------------
# vat-sum: 부가가치세신고용 합계표조회 (아직 화면을 조사하지 못함)
#
# 이 화면은 1월 15일 이후에 열린다. 열린 뒤에 할 일 (화면지도.md에도 기록):
#  1) 읽기 전용으로 조사: 작업영역 class, 제목, 조회 조건 칸, 거래처 넣는 곳, 조회 버튼, 결과 표
#  2) tonghap와 같은 순서로 구현: 조건 넣기 -> 결과 칸 비우기 -> 조회 -> 다시 채워졌는지 확인
#     -> _TABLES_JS 방식으로 표 읽기 -> capture_work_area -> save_screen_result(client_dir, "부가세합계표", png) + 엑셀
#  3) SCREEN_READY["부가세합계표"] = True 로 바꾸고 이 함수에 구현을 넣는다.
# 안 본 화면의 요소 이름을 짐작해서 쓰지 않는다 — 그래서 지금은 멈춘다.
# ---------------------------------------------------------------------------

def run_vat_sum(args):
    raise StopForSafety(
        "부가가치세신고용 합계표조회 화면은 아직 조사하지 못했습니다 (1월 15일 이후 화면이 열리면 조사). "
        "요소 이름을 짐작해서 조회하지 않고 멈춥니다."
    )


# ---------------------------------------------------------------------------
# 현금영수증 매출총액 조회 (메뉴 이름) = 화면 소제목은 '현금영수증 매출내역누계 조회'
# (세무대리 공통 조회 메뉴, 코드 4801120000)
#
# 2026-10-28 화면을 직접 조사해서 확인된 것:
#  - 조건: 조회년도(select, 값은 '2026년' 꼴) + 분기(select: -전체- / 1~4분기)
#  - 거래처: 사업자번호가 세 칸으로 나뉘어 있다 (앞 3자리 / 중간 2자리 / 뒤 5자리)
#  - 조회 버튼 trigger1. '인쇄' 버튼은 누르지 않는다.
#  - 결과: 표 하나(거래년월·거래건수·총금액·공급가액·세액·봉사료...) + 합계 줄(총합계).
#    분기가 -전체-이면 월별 줄 사이에 '상반기 합계', '하반기 합계' 소계 줄이 낀다.
#    '하반기 합계'가 곧 2기(7~12월) 값이다.
#  - ★이 표는 줄을 재사용해서, 새로 조회해도 이전 거래처 결과가 숨은 줄(display:none)로 남는다
#    (실제로 현재 거래처에 없는 10월 줄과 두 번째 '하반기 합계'가 숨은 줄에 남아 있었다).
#    그래서 "화면에 보이는 줄만" 읽는 방식(_TABLES_JS)이 맞다. 숨은 줄까지 읽으면 틀린 값이 섞인다.
#  - 조회 한 번에 서버 응답(wqAction.do POST)이 한 번 온다.
#  - 합계 검증: 월 줄의 합 = 총합계 = 상반기합계+하반기합계 가 모든 숫자 열(1~5)에서 맞는 것을 확인했다.
#    코드는 이 대조(verify_cash_totals)를 매번 하고, 안 맞으면 값을 내놓지 않고 그 거래처를 실패 처리한다.
#  - 자료가 없는 거래처는 표에 머리글 줄만 남고 합계 줄은 0이다 (읽기가 맞다는 것은 자료 있는 거래처로 증명함).
# ---------------------------------------------------------------------------

CASH_TITLE_HINT = "현금영수증 매출내역누계 조회"
CASH_YEAR_SELECT = "mf_txppWframe_selectYr"
CASH_QRT_SELECT = "mf_txppWframe_selectQrt"
CASH_BIZ_INPUTS = ("mf_txppWframe_txprDscmNo1", "mf_txppWframe_txprDscmNo2", "mf_txppWframe_txprDscmNo3")
CASH_QUERY_BUTTON = "mf_txppWframe_trigger1"
CASH_QUARTERS = ("전체",) + QUARTERS   # 화면의 '-전체-' (앞이 '-'라 명령줄에서 옵션으로 오인돼서 '전체'로 받는다)


def _cash_qrt_label(q):
    return "-전체-" if q == "전체" else q

_CASH_READBACK_JS = r"""
() => {
  const sel = id => { const s = document.getElementById(id); return s && s.options[s.selectedIndex] ? s.options[s.selectedIndex].text : null; };
  const v = id => document.getElementById(id).value.replace(/\D/g, '');
  return {year: sel('mf_txppWframe_selectYr'), qrt: sel('mf_txppWframe_selectQrt'),
          bizno: v('mf_txppWframe_txprDscmNo1') + v('mf_txppWframe_txprDscmNo2') + v('mf_txppWframe_txprDscmNo3')};
}
"""


def verify_cash_totals(tables):
    """월 줄의 합이 합계 줄(총합계)과 숫자 열 1~5에서 모두 같은지 대조한다. 다르면 ClientFailed.
    자료가 없는 거래처(본문 데이터 줄 0개)는 합계 줄이 전부 0일 때만 통과한다."""
    def num(s):
        return int(s.replace(",", "")) if re.fullmatch(r"-?[\d,]+", s or "") else None

    body = next((t for t in tables if t["id"].endswith("grdCshpt_body_table")), None)
    foot = next((t for t in tables if t["id"].endswith("grdCshpt_foot_table")), None)
    if body is None or foot is None or not foot["rows"]:
        raise ClientFailed("결과 표나 합계 줄을 화면에서 찾지 못했습니다 (값을 내놓지 않음)")
    rows = [[c["text"] for c in r] for r in body["rows"]][1:]          # 첫 줄은 머리글
    months = [r for r in rows if r and re.fullmatch(r"\d{4}-\d{2}", r[0])]
    foot_row = [c["text"] for c in foot["rows"][0]]
    for col in range(1, 6):
        total = num(foot_row[col]) if col < len(foot_row) else None
        if total is None:
            raise ClientFailed("합계 줄의 숫자를 읽지 못했습니다 (값을 내놓지 않음)")
        vals = [num(r[col]) for r in months if col < len(r)]
        if any(v is None for v in vals):
            raise ClientFailed("월별 줄에 숫자가 아닌 칸이 있어 합계를 대조하지 못했습니다 (값을 내놓지 않음)")
        if sum(vals) != total:
            raise ClientFailed("월별 줄의 합이 합계 줄과 맞지 않습니다 — 줄이 빠졌거나 섞였을 수 있어 값을 내놓지 않음")


_XLSX_NUM = re.compile(r"-?(0|[1-9]\d*|[1-9]\d{0,2}(,\d{3})+)")
_XML_BAD = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _xl_ref(r, c):
    s = ""
    c += 1
    while c:
        c, rem = divmod(c - 1, 26)
        s = chr(65 + rem) + s
    return f"{s}{r}"


def write_xlsx(path, sheet_name, tables):
    """화면에서 읽은 표(tables)를 엑셀(.xlsx) 한 장으로 저장한다. 외부 패키지 없이 zipfile 로 만든다.
    표는 위에서 아래로 쌓고(표 사이 한 줄 띄움), 병합 칸은 병합으로 옮긴다. 쉼표 있는 숫자는 숫자로 넣는다."""
    rows_xml, merges, r, ncols = [], [], 0, 1
    esc = lambda t: html.escape(_XML_BAD.sub("", t), quote=False)
    for tbl in tables:
        cap = tbl.get("caption")
        if cap:
            r += 1
            rows_xml.append(f'<row r="{r}"><c r="{_xl_ref(r, 0)}" s="4" t="inlineStr"><is><t xml:space="preserve">{esc(cap)}</t></is></c></row>')
        occupied = set()
        for row in tbl["rows"]:
            r += 1
            c, cells = 0, []
            for cell in row:
                while (r, c) in occupied:
                    c += 1
                cs, rs = max(cell["colspan"], 1), max(cell["rowspan"], 1)
                for dr in range(rs):
                    for dc in range(cs):
                        if dr or dc:
                            occupied.add((r + dr, c + dc))
                if cs > 1 or rs > 1:
                    merges.append(f"{_xl_ref(r, c)}:{_xl_ref(r + rs - 1, c + cs - 1)}")
                text = cell["text"] or ""
                ref = _xl_ref(r, c)
                if cell["tag"] != "th" and _XLSX_NUM.fullmatch(text):
                    cells.append(f'<c r="{ref}" s="2"><v>{int(text.replace(",", ""))}</v></c>')
                else:
                    style = 1 if cell["tag"] == "th" else 3
                    cells.append(f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{esc(text)}</t></is></c>')
                ncols = max(ncols, c + cs)
                c += cs
            rows_xml.append(f'<row r="{r}">{"".join(cells)}</row>')
        r += 1   # 표 사이 빈 줄
    cols = f'<cols><col min="1" max="{ncols}" width="16" customWidth="1"/></cols>'
    merge_xml = f'<mergeCells count="{len(merges)}">{"".join(f"<mergeCell ref=\"{m}\"/>" for m in merges)}</mergeCells>' if merges else ""
    sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
             f'{cols}<sheetData>{"".join(rows_xml)}</sheetData>{merge_xml}</worksheet>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
              '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0"/></numFmts>'
              '<fonts count="2"><font><sz val="10"/><name val="Malgun Gothic"/></font><font><b/><sz val="10"/><name val="Malgun Gothic"/></font></fonts>'
              '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
              '<fill><patternFill patternType="solid"><fgColor rgb="FFEEF2FA"/></patternFill></fill></fills>'
              '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
              '<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border></borders>'
              '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
              '<cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
              '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
              '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>'
              '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>'
              '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>'
              '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
              '</styleSheet>')
    safe_name = re.sub(r"[\[\]:*?/\\]", "", sheet_name)[:31] or "Sheet1"
    parts = {
        "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
        "_rels/.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="{html.escape(safe_name)}" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        "xl/styles.xml": styles,
        "xl/worksheets/sheet1.xml": sheet,
    }
    tmp = Path(str(path) + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
    os.replace(tmp, path)


class CashScreen:
    key = CASH_SUM_KEY
    title = dict(SCREENS)[CASH_SUM_KEY]
    menu_item_id = "menuAtag_0602070000"   # 세무대리/납세관리 > 현금영수증 매출총액 조회 (2026-11 메뉴 개편)
    arrival_heading = CASH_TITLE_HINT

    @staticmethod
    def add_arguments(p):
        p.add_argument("--year", required=True, help="조회년도, 예: 2026 (기본값 없음: 꼭 적는다)")
        p.add_argument("--quarter", required=True, choices=list(CASH_QUARTERS),
                       help="-전체- 이면 월별 + 상반기/하반기 합계가 나온다 ('하반기 합계'가 2기)")

    def conditions(self, args):
        if not re.fullmatch(r"20\d\d", str(args.year)):
            raise StopForSafety(f"연도 형식이 이상합니다: {args.year!r} (예: 2026)")
        if args.quarter not in CASH_QUARTERS:
            raise StopForSafety(f"분기 값을 모릅니다: {args.quarter!r} (가능: {list(CASH_QUARTERS)})")
        return {"year": str(args.year), "quarter": args.quarter}

    def describe(self, cond):
        return f"{cond['year']}년 {cond['quarter']}"

    def ensure_on(self, frame):
        if not work_heading_visible(frame, CASH_TITLE_HINT):
            raise StopForSafety(
                f"로그인은 돼 있지만 지금 화면이 '{CASH_TITLE_HINT}'이(가) 아닌 것 같습니다. "
                "메뉴에서 '현금영수증 매출총액 조회' 화면으로 옮겨 주세요."
            )

    def apply(self, page, frame, cond):
        _select_and_verify(frame, CASH_YEAR_SELECT, f"{cond['year']}년")
        _select_and_verify(frame, CASH_QRT_SELECT, _cash_qrt_label(cond["quarter"]))

    def query(self, page, frame, bizno, cond, ctx):
        for sel, part in zip(CASH_BIZ_INPUTS, (bizno[:3], bizno[3:5], bizno[5:])):
            _type_into(page, frame, f"#{sel}", part)
        if frame.evaluate(_CASH_READBACK_JS)["bizno"] != bizno:
            raise StopForSafety("사업자번호가 입력칸 세 개에 요청한 대로 들어가지 않았습니다. 멈춥니다.")

        ctx.dialogs.clear()
        ctx.watcher.reset()
        safe_click(page, frame, f"#{CASH_QUERY_BUTTON}")
        statuses = ctx.watcher.wait(page)

        raise_if_screen_message(page, frame, ctx)
        if not statuses:
            raise ClientFailed("조회 후 20초 안에 서버 응답이 오지 않았습니다 (값을 내놓지 않음)")
        bad = [s for s in statuses if s >= 400]
        if bad:
            raise ClientFailed(f"서버 응답이 오류입니다 (상태 {bad}). 값을 내놓지 않음")

        rb = frame.evaluate(_CASH_READBACK_JS)
        if (rb["year"], rb["qrt"], rb["bizno"]) != (f"{cond['year']}년", _cash_qrt_label(cond["quarter"]), bizno):
            raise StopForSafety(f"조회 뒤 화면의 조건이 요청과 다릅니다 (화면: {rb['year']} {rb['qrt']}). 멈춥니다.")
        tables = read_result_tables(frame)    # 보이는 줄만 읽는다 (숨은 줄은 이전 결과 찌꺼기)
        verify_cash_totals(tables)
        return tables, {"조회년도": rb["year"], "분기": rb["qrt"]}

    def after_save(self, page, frame, client_dir, tables, cond, ctx):
        # 이 화면에는 '엑셀 내려받기' 버튼이 없다(조회·인쇄뿐). 검증을 통과한 표를 엑셀 파일로 직접 만든다.
        name = f"{CASH_SUM_KEY}.xlsx"
        write_xlsx(client_dir / name, CASH_SUM_KEY, tables)
        return [name]


def run_cash_sum(args):
    run_batch(args, CashScreen())


# ---------------------------------------------------------------------------
# 수출실적명세서 조회 (세무대리 공통 조회 메뉴, 코드 4801200000)
# (일반 메뉴 쪽 '수출실적명세서 조회', 코드 4102060700 은 세무대리용이 아니다.)
#
# 2026-10-28 화면을 직접 조사해서 확인된 것:
#  - 기간: 기간 종류 라디오(월/분기/반기) + 선적년 `edtYear` + 종류별 선택칸
#    (월 `edtMon` / 분기 `edtQrt` / 반기 `edtHt` = 상반기·하반기). 2기는 '반기 + 하반기'.
#  - 거래처: 사업자등록번호 입력칸 `inputBsno` (최대 10자리). 조회 버튼 `trigger93`.
#  - ★`다운로드`, `파일생성`, `전자신고용 파일 생성방법`, `인쇄하기` 버튼은 파일을 만들거나 내려받으므로 절대 누르지 않는다.
#  - 결과 표 `grdVatKcsvExprArstMate` (수출신고번호, 선적일자, 통화코드, 환율, 외화금액, 원화환산금액).
#    합계 줄은 없다. (현금영수증 화면처럼 줄을 재사용해 숨은 줄이 남을 수 있어 '보이는 줄만' 읽는다.)
#  - 조회 한 번에 서버 응답(wqAction.do POST) 한 번. 그 JSON 의 `vatKcsvExprArstMateList` 가 서버가 알려준 수출 줄 목록이다.
#    → 화면 표와 별개의 독립 증거로, "서버 줄 수 = 화면에 보이는 줄 수"가 맞을 때만 결과를 인정한다.
#  - 사용자 요청(2026-10-28): 수출이 있는 사업장만 저장한다. 서버 0줄 + 화면 0줄이면 '자료 없음'으로 기록만 하고
#    폴더·파일을 만들지 않는다 (NoDataForClient).
#  - 명단 87곳 모두 2026년 하반기(87곳)·상반기(45곳 확인) 수출 자료가 없었다. 즉 '자료 있는 경우'로 읽기를 직접
#    증명하지는 못했고, 서버가 직접 0줄이라고 답한 것(독립 증거)에 기대고 있다.
# ---------------------------------------------------------------------------

EXPORT_TITLE_HINT = "수출실적명세서 조회"
EXPORT_TYPE_RADIO = "mf_txppWframe_shpnYmGubun_input"
EXPORT_TYPES = {"월": 0, "분기": 1, "반기": 2}
EXPORT_YEAR_SELECT = "mf_txppWframe_edtYear"
EXPORT_VALUE_SELECT = {"월": "mf_txppWframe_edtMon", "분기": "mf_txppWframe_edtQrt", "반기": "mf_txppWframe_edtHt"}
EXPORT_VALUES = {
    "월": tuple(f"{m}월" for m in range(1, 13)),
    "분기": QUARTERS,
    "반기": ("상반기", "하반기"),
}
EXPORT_BSNO_INPUT = "mf_txppWframe_inputBsno"
EXPORT_QUERY_BUTTON = "mf_txppWframe_trigger93"
EXPORT_GRID_BODY = "mf_txppWframe_grdVatKcsvExprArstMate_body_table"

_EXPORT_READBACK_JS = r"""
() => {
  const sel = id => { const s = document.getElementById(id); return s && (s.offsetWidth || s.offsetHeight) && s.options[s.selectedIndex] ? s.options[s.selectedIndex].text : null; };
  const checked = Array.from(document.querySelectorAll('input[name="mf_txppWframe_shpnYmGubun_input"]')).findIndex(r => r.checked);
  return {typeIdx: checked, year: sel('mf_txppWframe_edtYear'),
          mon: sel('mf_txppWframe_edtMon'), qrt: sel('mf_txppWframe_edtQrt'), ht: sel('mf_txppWframe_edtHt'),
          bizno: document.getElementById('mf_txppWframe_inputBsno').value.replace(/\D/g, '')};
}
"""

# 화면에 보이는 데이터 줄 수 (머리글 줄 제외). 숨은 줄(이전 조회 찌꺼기)은 세지 않는다.
_EXPORT_VISIBLE_ROWS_JS = r"""
() => {
  const t = document.getElementById('mf_txppWframe_grdVatKcsvExprArstMate_body_table');
  if (!t) return null;
  return Array.from(t.rows).slice(1).filter(r => (r.offsetWidth || r.offsetHeight)
    && Array.from(r.cells).some(c => (c.innerText || '').trim() !== '')).length;
}
"""


def _server_export_rows(ctx):
    """조회 응답(JSON)에서 서버가 알려준 수출 줄 수. 못 찾으면 None."""
    counts = []
    for body in ctx.watcher.json_bodies():
        if isinstance(body, dict) and isinstance(body.get("vatKcsvExprArstMateList"), list):
            counts.append(len(body["vatKcsvExprArstMateList"]))
    return counts[-1] if counts else None


class ExportScreen:
    key = EXPORT_SUM_KEY
    title = dict(SCREENS)[EXPORT_SUM_KEY]
    menu_item_id = "menuAtag_0602130000"   # 세무대리/납세관리 > 수출실적명세서 조회 (2026-11 메뉴 개편)
    arrival_heading = EXPORT_TITLE_HINT

    @staticmethod
    def add_arguments(p):
        p.add_argument("--year", required=True, help="선적년, 예: 2026 (기본값 없음: 꼭 적는다)")
        p.add_argument("--type", required=True, choices=list(EXPORT_TYPES), help="기간 종류: 월 / 분기 / 반기")
        p.add_argument("--value", required=True, help="종류에 맞는 값: 월 '10월' / 분기 '3분기' / 반기 '하반기'")

    def conditions(self, args):
        if not re.fullmatch(r"20\d\d", str(args.year)):
            raise StopForSafety(f"연도 형식이 이상합니다: {args.year!r} (예: 2026)")
        if args.type not in EXPORT_TYPES:
            raise StopForSafety(f"기간 종류를 모릅니다: {args.type!r} (가능: {list(EXPORT_TYPES)})")
        if args.value not in EXPORT_VALUES[args.type]:
            raise StopForSafety(f"{args.type} 기간 값을 모릅니다: {args.value!r} (가능: {list(EXPORT_VALUES[args.type])})")
        return {"year": str(args.year), "type": args.type, "value": args.value}

    def describe(self, cond):
        return f"{cond['year']}년 {cond['value']}({cond['type']} 기준)"

    def ensure_on(self, frame):
        if not work_heading_visible(frame, EXPORT_TITLE_HINT):
            raise StopForSafety(
                f"로그인은 돼 있지만 지금 화면이 '{EXPORT_TITLE_HINT}'이(가) 아닌 것 같습니다. 메뉴에서 그 화면으로 옮겨 주세요."
            )

    def apply(self, page, frame, cond):
        _set_radio_safe(page, frame, EXPORT_TYPE_RADIO, EXPORT_TYPES[cond["type"]])
        _select_and_verify(frame, EXPORT_YEAR_SELECT, cond["year"])
        _select_and_verify(frame, EXPORT_VALUE_SELECT[cond["type"]], cond["value"])

    def query(self, page, frame, bizno, cond, ctx):
        # 수출 자료가 있는 거래처를 조회한 직후(표가 늘어나 화면 배치가 바뀐 직후)에 번호가 안 바뀐 적이 있어서,
        # 요청한 대로 들어갔는지 읽어 확인하고 안 들어갔으면 몇 번 다시 시도한다. 끝내 안 되면 멈춘다 (틀린 번호로 조회하지 않는다).
        for _ in range(3):
            _type_into(page, frame, f"#{EXPORT_BSNO_INPUT}", bizno)
            if frame.evaluate(_EXPORT_READBACK_JS)["bizno"] == bizno:
                break
            page.wait_for_timeout(700)
        else:
            raise StopForSafety("사업자등록번호가 입력칸에 3번 시도해도 요청한 대로 들어가지 않았습니다. 멈춥니다.")

        ctx.dialogs.clear()
        server_n_early = None
        for attempt in (1, 2):
            ctx.watcher.reset()
            safe_click(page, frame, f"#{EXPORT_QUERY_BUTTON}")
            statuses = ctx.watcher.wait(page)

            raise_if_screen_message(page, frame, ctx)
            if not statuses:
                raise ClientFailed("조회 후 20초 안에 서버 응답이 오지 않았습니다 (값을 내놓지 않음)")
            bad = [s for s in statuses if s >= 400]
            if bad:
                raise ClientFailed(f"서버 응답이 오류입니다 (상태 {bad}). 값을 내놓지 않음")
            server_n_early = _server_export_rows(ctx)
            if server_n_early is not None:
                break
            # 거래처를 바꾼 뒤 첫 조회는 '거래처 기본정보 확인'(ATTABZAA001R01)만 하고 수출 목록은 가져오지 않는다.
            # 같은 조건으로 한 번 더 눌러야 수출 목록(ATERNABA125R01)이 온다. 이때 화면 표는 이전 결과 그대로일 수 있다.
            page.wait_for_timeout(2000)
        if server_n_early is None:
            raise ClientFailed("조회를 두 번 눌러도 서버가 수출 목록을 돌려주지 않았습니다 (값을 내놓지 않음)")

        rb = frame.evaluate(_EXPORT_READBACK_JS)
        got_value = {"월": rb["mon"], "분기": rb["qrt"], "반기": rb["ht"]}[cond["type"]]
        if (rb["typeIdx"], rb["year"], got_value, rb["bizno"]) != (EXPORT_TYPES[cond["type"]], cond["year"], cond["value"], bizno):
            raise StopForSafety(
                f"조회 뒤 화면의 조건이 요청과 다릅니다 (화면: {rb['year']}년 {got_value}). 멈춥니다."
            )

        server_n = server_n_early
        dom_n = frame.evaluate(_EXPORT_VISIBLE_ROWS_JS)
        if dom_n is None:
            raise ClientFailed("화면 표에서 수출 줄 수를 확인하지 못했습니다 (값을 내놓지 않음)")
        if server_n != dom_n:
            raise ClientFailed(
                f"서버가 알려준 수출 줄 수({server_n})와 화면에 보이는 줄 수({dom_n})가 다릅니다 (값을 내놓지 않음)")
        if server_n == 0:
            raise NoDataForClient("서버 응답의 수출 목록이 0줄이고 화면에도 보이는 줄이 없음")
        if frame.evaluate(_GRID_OVERFLOW_JS):
            raise ClientFailed("결과 표가 스크롤될 만큼 길어서 일부만 읽혔을 수 있습니다 (값을 내놓지 않음)")
        tables = read_result_tables(frame)
        if not tables:
            raise ClientFailed("결과 표를 화면에서 찾지 못했습니다 (값을 내놓지 않음)")
        return tables, {"선적년": cond["year"], "기간": f"{cond['value']} ({cond['type']} 기준)", "수출 줄 수": server_n}


def run_export_sum(args):
    run_batch(args, ExportScreen())


# ---------------------------------------------------------------------------
# invoice-sum: 전자(세금)계산서 합계표 조회 - 총합계 읽기 (세무대리 공통 조회 메뉴의 화면)
# ---------------------------------------------------------------------------

def run_invoice_sum(args):
    gubun_list = ["매출", "매입"] if args.gubun == "둘다" else [args.gubun]

    with sync_playwright() as p:
        browser = connect_chrome(p)
        ctx, page = find_hometax_page(browser)
        if page is None:
            log("[중단] 홈택스 탭을 찾지 못했습니다.")
            sys.exit(1)
        frame = page.main_frame

        try:
            log("[1/4] 로그인/화면 확인 중...")
            ensure_logged_in_and_on_screen(frame)
            log("      확인됨: 로그인돼 있고, 합계표 조회 화면이 맞습니다.")

            log(f"[2/4] 거래처를 {args.biz_no} 로 바꾸는 중...")
            switch_client(frame, args.biz_no)
            log("      확인됨: 거래처 전환이 반영됐습니다.")

            log(f"[3/4] {'/'.join(gubun_list)} 조건으로 이어서 조회하는 중...")
            results = run_multi_gubun(
                frame, args.clsf, args.period_mode, gubun_list,
                year=args.year, quarter=args.quarter, month=args.month,
            )
            log("      확인됨: 모든 조회를 마쳤습니다.")

        except StopForSafety as e:
            log(f"[중단] {e}")
            sys.exit(1)
        finally:
            browser.close()

    log("")
    for gubun, result, title in results:
        log(f"[{gubun}] 화면 제목: {title}")
    log("")
    log(format_combined_table(results))


# ---------------------------------------------------------------------------
# 예약 실행 (scheduled)
#
# 정해 둔 시각에 사람이 자리에 없어도 되는 부분(크롬 띄우기, 메뉴로 화면 이동, 두 화면 차례로 조회,
# 기록 남기기)은 코드가 하고, 로그인(공동인증서·세무대리인 로그인 포함)은 사람이 한다.
# 로그인 화면이 보이거나 로그인 전으로 보이면 아무것도 입력하지 않고 사람이 할 때까지 기다리며,
# 기다리는 시간이 지나면 멈춘다. 로그인은 홈택스가 1시간쯤 뒤 자동으로 풀어 버리므로,
# 미리 해 두지 말고 예약 시각 직전·직후에 하는 것이 좋다.
# ---------------------------------------------------------------------------

CHROME_EXE = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
CHROME_PROFILE = r"C:\hometax-axtax\chrome-profile"

_TOP_NAV_HOMEPAGE_JS = r"""
() => Array.from(document.querySelectorAll('#mf_wfHeader a'))
  .filter(a => (a.offsetWidth || a.offsetHeight) && (a.innerText || '').replace(/\s+/g, '') === '세무대리')
  .map(a => a.id)
"""

_TOP_NAV_CATEGORY_JS = r"""
() => Array.from(document.querySelectorAll('#mf_wfHeader a'))
  .filter(a => (a.offsetWidth || a.offsetHeight) && (a.innerText || '').replace(/[/\s]+/g, '') === '세무대리납세관리')
  .map(a => a.id)
"""

_CLOSE_CONFIRM_POPUP_JS = r"""
() => {
  const popups = Array.from(document.querySelectorAll('.w2window, .w2confirm'));
  for (const p of popups) {
    if (!(p.offsetWidth || p.offsetHeight)) continue;
    const btn = p.querySelector('.w2window_close') || Array.from(p.querySelectorAll('a,button')).find(b => (b.innerText||'').trim() === '취소');
    if (btn) { btn.click(); return true; }
  }
  return false;
}
"""

_LOGOUT_LINK_JS = r"""
() => Array.from(document.querySelectorAll('a'))
  .some(a => (a.offsetWidth || a.offsetHeight) && (a.innerText || '').trim() === '로그아웃')
"""


def _chrome_ready():
    try:
        with urllib.request.urlopen(CDP_URL + "/json/version", timeout=3) as r:
            return r.status == 200
    except Exception:
        return False


def ensure_chrome_running(launch):
    if _chrome_ready():
        return
    if not launch:
        raise StopForSafety("크롬(9333)이 꺼져 있습니다. --launch-chrome 을 붙이면 CLAUDE.md의 명령으로 띄웁니다.")
    log("크롬(9333)이 꺼져 있어 홈택스 전용 크롬을 띄웁니다.")
    subprocess.Popen([
        CHROME_EXE, "--remote-debugging-port=9333", f"--user-data-dir={CHROME_PROFILE}",
        "--disable-features=LocalNetworkAccessChecks", "--new-window", "https://hometax.go.kr/",
    ])
    for _ in range(30):
        time.sleep(2)
        if _chrome_ready():
            return
    raise StopForSafety("크롬을 띄웠지만 60초 안에 9333에 연결되지 않았습니다.")


def _looks_logged_in(frame):
    """로그인 여부를 머리말 글자 하나로 판단하지 않는다. 머리말 링크는 페이지가 아래로 스크롤되면 사라지므로
    맨 위로 올린 뒤 보고, '로그아웃 링크'가 없어도 조회 조건 칸+조회 버튼(작업 화면)이 있으면 로그인된 것으로 본다.
    애매하면 잠깐 뒤 한 번 더 본다."""
    for attempt in range(2):
        frame.evaluate("window.scrollTo(0, 0)")
        time.sleep(0.8)
        if frame.evaluate(_LOGOUT_LINK_JS) or _has_query_screen_elements(frame):
            return True
        if attempt == 0:
            time.sleep(1.5)
    return False


def _click_menu_item(page, frame, screen):
    """상단 메뉴를 펼치고 이 화면의 메뉴 항목을 누른다. 결과: 'arrived' / 'login' / 'stuck'.

    2026-11 메뉴 개편 대응: 홈페이지의 '세무대리'를 누르면 세무대리인 페이지로 이동한다.
    세무대리인 페이지에서 '세무대리/납세관리' 카테고리를 hover하면 드롭다운이 열린다."""
    frame.evaluate("window.scrollTo(0, 0)")
    time.sleep(0.8)
    frame.evaluate(_CLOSE_CONFIRM_POPUP_JS)
    time.sleep(0.3)

    if not frame.evaluate(_MENU_OVERLAY_JS):
        cats = frame.evaluate(_TOP_NAV_CATEGORY_JS)
        if len(cats) == 1:
            frame.evaluate("(id) => { const el = document.getElementById(id); if(el) { el.dispatchEvent(new MouseEvent('mouseover',{bubbles:true})); el.dispatchEvent(new MouseEvent('mouseenter',{bubbles:true})); } }", cats[0])
            time.sleep(1.5)
        else:
            navs = frame.evaluate(_TOP_NAV_HOMEPAGE_JS)
            if len(navs) == 1:
                safe_click(page, frame, "#" + navs[0])
                time.sleep(3)
                return "stuck"
            return "stuck"

    shown = frame.evaluate(
        "(id) => Array.from(document.querySelectorAll('#' + id)).filter(a => a.offsetWidth || a.offsetHeight).length",
        screen.menu_item_id,
    )
    if shown != 1:
        return "stuck"
    safe_click(page, frame, "#" + screen.menu_item_id)

    def state():
        if _is_login_screen_visible(frame):
            return "login"
        if work_heading_visible(frame, screen.arrival_heading) and not frame.evaluate(_MENU_OVERLAY_JS):
            return "arrived"
        return None

    return _wait_until(state, 20) or "stuck"


def navigate_to_screen(screen, deadline, notify_every=300):
    """메뉴로 그 화면까지 이동한다. 로그인 전이거나 로그인 화면이면 사람이 할 때까지 기다린다 (입력하지 않는다)."""
    stuck = 0
    last_notice = 0.0
    while True:
        reason = None
        with sync_playwright() as p:
            browser = connect_chrome(p)
            try:
                ctx, page = find_hometax_page(browser)
                if page is None:
                    reason = "홈택스 탭이 없습니다"
                else:
                    frame = page.main_frame
                    if _is_login_screen_visible(frame):
                        reason = "로그인 화면이 떠 있습니다 (사람이 로그인해 주세요)"
                    elif work_heading_visible(frame, screen.arrival_heading):
                        return
                    elif not _looks_logged_in(frame):
                        reason = "아직 로그인 전으로 보입니다 (사람이 로그인해 주세요)"
                    else:
                        manual_msg = f"자동 이동이 안 되는 화면입니다 — 사람이 메뉴로 '{screen.title}' 화면까지 이동해 주세요"
                        if stuck >= 3:
                            reason = manual_msg   # 더 누르지 않고 사람이 이동할 때까지 지켜만 본다
                        else:
                            outcome = _click_menu_item(page, frame, screen)
                            if outcome == "arrived":
                                return
                            if outcome == "login":
                                reason = "메뉴를 눌렀더니 로그인 화면이 떴습니다 (사람이 로그인해 주세요)"
                            else:
                                stuck += 1
                                reason = "메뉴를 눌렀지만 화면이 바뀌지 않았습니다 (다시 시도)" if stuck < 3 else manual_msg
            finally:
                browser.close()
        now = time.time()
        if now >= deadline:
            raise StopForSafety(f"기다리는 시간이 지났습니다 ({reason}).")
        if now - last_notice >= notify_every:
            log(f"[대기] {reason} — 남은 대기 {int((deadline - now) / 60)}분")
            try:
                ctypes.windll.user32.MessageBeep(0)
            except Exception:
                pass
            last_notice = now
        time.sleep(20)


def _apply_config(args):
    """설정 파일(JSON)의 값으로 명령줄에 안 준 항목을 채운다. 명령줄에 준 값이 우선이다."""
    if args.config:
        try:
            cfg = json.loads(Path(args.config).read_text(encoding="utf-8"))
        except (OSError, ValueError) as e:
            raise StopForSafety(f"설정 파일을 읽지 못했습니다 ({type(e).__name__}): {args.config}")
        for k, v in cfg.items():
            if k.startswith("_"):   # '_설명' 처럼 밑줄로 시작하면 사람이 읽는 메모
                continue
            if not hasattr(args, k):
                raise StopForSafety(f"설정 파일에 모르는 항목이 있습니다: {k}")
            if getattr(args, k) in (None, False):
                setattr(args, k, v)
    if args.wait_minutes is None:
        args.wait_minutes = 180
    need = ["list", "tonghap_year", "tonghap_ki", "tonghap_rtn", "card_year", "card_from_q", "card_to_q",
            "cash_year", "cash_quarter", "export_year", "export_type", "export_value"]
    missing = [k for k in need if not getattr(args, k)]
    if missing:
        raise StopForSafety(f"예약 실행에 꼭 필요한 값이 없습니다: {missing} (명령줄이나 --config 설정 파일로 주세요)")
    if args.tonghap_ki not in KI_OPTIONS or args.tonghap_rtn not in RTN_OPTIONS:
        raise StopForSafety(f"통합조회 기(1기/2기)나 신고구분 값이 이상합니다: {args.tonghap_ki!r}, {args.tonghap_rtn!r}")
    if args.card_from_q not in QUARTERS or args.card_to_q not in QUARTERS:
        raise StopForSafety(f"신용카드 분기 값이 이상합니다: {args.card_from_q!r}, {args.card_to_q!r}")
    if args.cash_quarter not in CASH_QUARTERS:
        raise StopForSafety(f"현금영수증 분기 값이 이상합니다: {args.cash_quarter!r}")
    if args.export_type not in EXPORT_TYPES or args.export_value not in EXPORT_VALUES.get(args.export_type, ()):
        raise StopForSafety(f"수출실적명세서 기간 값이 이상합니다: {args.export_type!r} {args.export_value!r}")


def run_scheduled(args):
    global _LOG_FILE
    _apply_config(args)
    base = dict(
        list=args.list, biz_no=None, name=None, only=args.only, limit=args.limit, redo=args.redo,
        redo_before=args.redo_before, dry_run=args.dry_run, out=args.out, folder_name=args.folder_name, allow_cloud_folder=args.allow_cloud_folder,
    )
    plan = [
        (TonghapScreen(), argparse.Namespace(**base, year=args.tonghap_year, ki=args.tonghap_ki, rtn=args.tonghap_rtn)),
        (CardScreen(), argparse.Namespace(**base, year=args.card_year, from_q=args.card_from_q, to_q=args.card_to_q)),
        (CashScreen(), argparse.Namespace(**base, year=args.cash_year, quarter=args.cash_quarter)),
        # 수출실적명세서는 마지막에 둔다: 수출 자료가 있는 거래처를 조회한 뒤 화면이 입력을 안 받게 된 적이 있어서(원인 모름),
        # 이 화면이 멈추더라도 앞의 화면들에는 영향이 없게 한다.
        (ExportScreen(), argparse.Namespace(**base, year=args.export_year, type=args.export_type, value=args.export_value)),
    ]
    conds = [screen.conditions(ns) for screen, ns in plan]   # 조건이 이상하면 브라우저를 건드리기 전에 멈춘다

    out_root = resolve_out_root(args.out, args.allow_cloud_folder)
    if not args.dry_run:
        out_root.mkdir(parents=True, exist_ok=True)
        _LOG_FILE = out_root / "_예약실행기록.txt"
    log("예약 실행 시작")
    for (screen, _), cond in zip(plan, conds):
        log(f"  - {screen.title}: {screen.describe(cond)}")
    pending = [title for key, title in SCREENS if not SCREEN_READY.get(key)]
    if pending:
        # 조용히 빠지지 않도록 실행 기록에 분명히 남긴다.
        log("  ※ 아직 조사·구현되지 않아 이번 실행에서 조회하지 않는 화면: " + ", ".join(pending))

    if args.dry_run:
        for screen, ns in plan:
            run_batch(ns, screen)
        return

    deadline = time.time() + args.wait_minutes * 60
    ensure_chrome_running(args.launch_chrome)
    results = []
    for screen, ns in plan:
        navigate_to_screen(screen, deadline)
        log(f"[{screen.title}] 조회 시작")
        try:
            run_batch(ns, screen)
            code = 0
        except SystemExit as e:
            code = e.code or 0
        results.append((screen.title, code))
        if code == 1:
            log(f"[{screen.title}] 중간에 멈췄습니다. 다음 화면은 하지 않고 예약 실행을 끝냅니다.")
            break
    label = {0: "정상", 2: "일부 거래처 실패", 1: "중단"}
    log("예약 실행 끝: " + " / ".join(f"{t} = {label.get(c, c)}" for t, c in results))
    sys.exit(max((c for _, c in results), default=0))


# ---------------------------------------------------------------------------
# 실행
# ---------------------------------------------------------------------------

def add_common_args(p):
    p.add_argument("--list", help="거래처 명단 xlsx 경로 (여러 거래처)")
    p.add_argument("--biz-no", help="사업자등록번호 하나 (명단 대신 한 곳만)")
    p.add_argument("--name", help="--biz-no 와 같이 쓸 때 폴더 이름에 넣을 상호 (생략 가능)")
    p.add_argument("--only", nargs="+", help="명단 중 이 사업자번호들만")
    p.add_argument("--limit", type=int, help="앞에서 N곳만 (시험용)")
    p.add_argument("--redo", action="store_true", help="이미 끝난 거래처도 다시")
    p.add_argument("--redo-before", help="이 시각(예: 2027-01-17T00:00:00) 이전에 끝낸 기록은 '끝난 것'으로 안 치고 다시 한다. "
                                         "그 이후에 끝낸 곳은 건너뛴다 (--redo 와 달리 같은 날 이어하기가 된다)")
    p.add_argument("--dry-run", action="store_true", help="조회하지 않고 대상 거래처 수·저장 이름만 점검 (크롬을 건드리지 않음)")
    p.add_argument("--out", help="저장 폴더 (기본: 바탕화면\\홈택스_조회결과)")
    p.add_argument("--folder-name", choices=["name", "both", "bizno"], default="name",
                   help="name: 상호만(겹치는 상호만 번호 덧붙임, 기본) / both: 상호_사업자번호 / bizno: 사업자번호만")
    p.add_argument("--allow-cloud-folder", action="store_true", help="OneDrive 등 클라우드 폴더에 저장 허용")


def build_parser():
    parser = argparse.ArgumentParser(description="홈택스 조회 도구 (조회만 한다)")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("check-list", help="거래처 명단(xlsx)의 개수·형식만 점검 (값은 출력 안 함)")
    p.add_argument("--list", required=True, help="거래처 명단 xlsx 경로")

    p = sub.add_parser("tonghap", help="부가가치세 신고자료 통합조회 -> 거래처별 폴더 (캡처 + 엑셀)")
    add_common_args(p)
    TonghapScreen.add_arguments(p)

    p = sub.add_parser("card-sum", help="신용카드/판매(결제)대행 매출자료 조회 -> 거래처별 폴더 (캡처 + 엑셀)")
    add_common_args(p)
    CardScreen.add_arguments(p)

    p = sub.add_parser("scheduled", help="예약 실행: 통합조회 -> 신용카드/판매대행을 차례로 (로그인만 사람이 한다)")
    add_common_args(p)
    p.add_argument("--config", help="설정 파일(JSON, UTF-8). 명령줄에 안 준 값은 여기서 읽는다 (윈도우 .cmd 는 한글을 못 읽어서 필요)")
    p.add_argument("--tonghap-year", help="통합조회 과세기간 연도")
    p.add_argument("--tonghap-ki", choices=list(KI_OPTIONS))
    p.add_argument("--tonghap-rtn", choices=list(RTN_OPTIONS))
    p.add_argument("--card-year", help="신용카드/판매대행 결제년도")
    p.add_argument("--card-from-q", choices=list(QUARTERS))
    p.add_argument("--card-to-q", choices=list(QUARTERS))
    p.add_argument("--cash-year", help="현금영수증 매출총액 조회년도")
    p.add_argument("--cash-quarter", choices=list(CASH_QUARTERS), help="전체 / 1~4분기 ('전체'의 '하반기 합계'가 2기)")
    p.add_argument("--export-year", help="수출실적명세서 선적년")
    p.add_argument("--export-type", choices=list(EXPORT_TYPES), help="수출실적명세서 기간 종류: 월 / 분기 / 반기")
    p.add_argument("--export-value", help="수출실적명세서 기간 값: 월 '10월' / 분기 '3분기' / 반기 '하반기'")
    p.add_argument("--wait-minutes", type=int, default=None, help="로그인·이동을 기다리는 최대 시간(분, 기본 180)")
    p.add_argument("--launch-chrome", action="store_true", help="크롬(9333)이 꺼져 있으면 띄운다")

    p = sub.add_parser("vat-sum", help="부가가치세신고용 합계표조회 (아직 화면 조사 전이라 멈춤)")
    p.add_argument("--list")
    p.add_argument("--biz-no")

    p = sub.add_parser("cash-sum", help="현금영수증 매출총액(매출내역누계) 조회 -> 거래처별 폴더 (캡처 + 엑셀)")
    add_common_args(p)
    CashScreen.add_arguments(p)

    p = sub.add_parser("export-sum", help="수출실적명세서 조회 -> 수출이 있는 거래처만 저장")
    add_common_args(p)
    ExportScreen.add_arguments(p)

    p = sub.add_parser("invoice-sum", help="전자(세금)계산서 합계표 조회 - 총합계 읽기")
    p.add_argument("--biz-no", required=True, help="거래처 사업자등록번호 (숫자 10자리, 대시 있어도 됨)")
    p.add_argument("--clsf", default="전자세금계산서", choices=list(CLSF_OPTIONS))
    p.add_argument("--gubun", default="매출", choices=list(GUBUN_OPTIONS) + ["둘다"])
    p.add_argument("--period-mode", default="분기별", choices=["분기별", "월별"])
    p.add_argument("--year", default=None, help="예: 2026")
    p.add_argument("--quarter", default=None, help="분기별 모드일 때, 예: '1기 확정'")
    p.add_argument("--month", default=None, help="월별 모드일 때, 예: '09월'")
    return parser


def main():
    # 예약 작업 창(콘솔)의 글자 설정 때문에 한글을 찍다가 죽지 않게 한다.
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    args = build_parser().parse_args()
    try:
        if args.cmd == "check-list":
            run_check_list(args)
        elif args.cmd == "tonghap":
            run_tonghap(args)
        elif args.cmd == "vat-sum":
            run_vat_sum(args)
        elif args.cmd == "cash-sum":
            run_cash_sum(args)
        elif args.cmd == "card-sum":
            run_card_sum(args)
        elif args.cmd == "scheduled":
            run_scheduled(args)
        elif args.cmd == "export-sum":
            run_export_sum(args)
        elif args.cmd == "invoice-sum":
            run_invoice_sum(args)
    except StopForSafety as e:
        log(f"[중단] {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
