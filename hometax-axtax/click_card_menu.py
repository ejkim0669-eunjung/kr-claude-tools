"""
신용카드/판매(결제)대행 매출자료 조회 메뉴를 한 번 클릭한다.
실제 마우스 이동(hover) 후 클릭 — JS 이벤트가 아닌 native mouse 사용.
"""
import sys
import time
from playwright.sync_api import sync_playwright

CDP_URL = "http://127.0.0.1:9333"
MENU_ITEM_ID = "menuAtag_0602130000"  # 수출실적명세서 조회
TOP_NAV_JS = r"""
() => Array.from(document.querySelectorAll('#mf_wfHeader a'))
  .filter(a => (a.offsetWidth || a.offsetHeight) && (a.innerText || '').replace(/[/\s]+/g, '') === '세무대리납세관리')
  .map(a => a.id)
"""

def get_center(frame, selector):
    r = frame.evaluate("""
    (sel) => {
        const el = document.querySelector(sel);
        if (!el) return null;
        const rect = el.getBoundingClientRect();
        if (!rect.width && !rect.height) return null;
        return {x: rect.left + rect.width/2, y: rect.top + rect.height/2};
    }
    """, selector)
    return r

def run():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(CDP_URL)
        page = None
        for ctx in browser.contexts:
            for pg in ctx.pages:
                if "hometax" in pg.url and "popup.html" not in pg.url:
                    page = pg
                    break
        if page is None:
            print("홈택스 탭을 찾지 못했습니다.")
            return 1

        frame = page.main_frame
        frame.evaluate("window.scrollTo(0, 0)")
        time.sleep(0.5)

        # 1) 이미 메뉴가 열려 있는지 확인
        overlay = frame.evaluate(
            "() => Array.from(document.querySelectorAll('.head_dim.on')).some(e => e.offsetWidth || e.offsetHeight)"
        )

        if not overlay:
            # 2) '세무대리/납세관리' 링크 찾아서 실제 마우스로 hover
            cat_ids = frame.evaluate(TOP_NAV_JS)
            if not cat_ids:
                print("'세무대리/납세관리' 메뉴를 찾지 못했습니다.")
                return 1
            pos = get_center(frame, "#" + cat_ids[0])
            if not pos:
                print(f"'세무대리/납세관리' 위치를 읽지 못했습니다.")
                return 1
            print(f"hover 이동: ({pos['x']:.0f}, {pos['y']:.0f})")
            page.mouse.move(pos["x"], pos["y"])
            time.sleep(1.5)

        # 3) 메뉴 항목이 보이는지 확인
        shown = frame.evaluate(
            "(id) => Array.from(document.querySelectorAll('#' + id)).filter(a => a.offsetWidth || a.offsetHeight).length",
            MENU_ITEM_ID,
        )
        if shown != 1:
            print(f"메뉴 항목 #{MENU_ITEM_ID} 이 보이지 않습니다 (found={shown}).")
            return 1

        pos2 = get_center(frame, "#" + MENU_ITEM_ID)
        if not pos2:
            print("메뉴 항목 위치를 읽지 못했습니다.")
            return 1
        print(f"클릭: #{MENU_ITEM_ID} ({pos2['x']:.0f}, {pos2['y']:.0f})")
        page.mouse.click(pos2["x"], pos2["y"])
        time.sleep(2)
        print("클릭 완료. 화면 전환 여부는 직접 확인하세요.")
        return 0

sys.exit(run())
