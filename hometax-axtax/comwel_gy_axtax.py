#!/usr/bin/env python -I
# -*- coding: utf-8 -*-
"""고용산재보험 토탈서비스 — 고용보험 당월보험료 + 지원금 엑셀 다운로드.

사용법:
  1) 크롬을 디버그 포트로 띄운다:
     start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" ^
       --remote-debugging-port=9333 --user-data-dir=C:\hometax-axtax\chrome-profile ^
       --disable-features=LocalNetworkAccessChecks --new-window https://total.comwel.or.kr/
  2) 크롬에서 로그인한다 (사람이 직접).
  3) 부과고지 보험료 조회(20209) 화면으로 이동한다.
  4) 이 스크립트를 실행한다:
     python -I comwel_gy_axtax.py
"""
import sys, time, os, random, subprocess, urllib.request
sys.stdout.reconfigure(encoding="utf-8")
from playwright.sync_api import sync_playwright

CDP_URL = "http://127.0.0.1:9333"
CHROME_EXE = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
CHROME_PROFILE = r"C:\hometax-axtax\chrome-profile"


def _chrome_ready():
    try:
        with urllib.request.urlopen(CDP_URL + "/json/version", timeout=3) as r:
            return r.status == 200
    except Exception:
        return False


def ensure_chrome_running():
    if _chrome_ready():
        return
    print("크롬(9333)이 꺼져 있어 띄웁니다.")
    subprocess.Popen([
        CHROME_EXE, "--remote-debugging-port=9333", f"--user-data-dir={CHROME_PROFILE}",
        "--disable-features=LocalNetworkAccessChecks", "--new-window", "https://total.comwel.or.kr/",
    ])
    for _ in range(30):
        time.sleep(2)
        if _chrome_ready():
            return
    raise RuntimeError("크롬을 띄웠지만 60초 안에 9333에 연결되지 않았습니다.")

# ── 설정 ──────────────────────────────────────────────
YEAR  = "2025"
MONTH = "10"

SAVE_BASE = r"C:\Users\admin\Desktop\사회보험"

FIRMS = [
    {"name": "(재)한국대학봉사회",           "mgr": "20382317340"},
    {"name": "(주)올리브트리 커뮤니케이션",   "mgr": "92540270681"},
    {"name": "(주)대광건화",                 "mgr": "10981730600"},
    {"name": "(주)에스랩엔지니어링",         "mgr": "88587023940"},
    {"name": "(주)정희농산",                 "mgr": "21587161620"},
    {"name": "남일스포츠",                   "mgr": "20350677930"},
    {"name": "유니베라신대방대리점",         "mgr": "91610209981"},
    {"name": "이남수부동산",                 "mgr": "12423849120"},
    {"name": "주식회사경상농산",             "mgr": "26781015420"},
    {"name": "주식회사비케이컨셉트",         "mgr": "57988003290"},
    {"name": "주식회사아리씨엠(RECMCorp.)",  "mgr": "17986033980"},
    {"name": "주식회사 에스에스피코리아",    "mgr": "84388027040"},
    {"name": "주식회사 에스에스피파트너스",  "mgr": "33187036680"},
    {"name": "큐플러스매니지먼트주식회사",   "mgr": "53988024790"},
    {"name": "현진프라퍼티주식회사",         "mgr": "19186017830"},
]

# ── 헬퍼 ──────────────────────────────────────────────
def fmt_mgr(mgr):
    if len(mgr) >= 11:
        return f"{mgr[:3]}-{mgr[3:5]}-{mgr[5:10]}-{mgr[10:]}"
    return mgr

def close_visible_popups(page):
    """보이는 w2window 팝업을 닫기 버튼으로 닫는다."""
    page.evaluate("""() => {
        document.querySelectorAll('.w2window').forEach(function(w) {
            if (w.offsetParent !== null || w.offsetWidth > 0) {
                var c = w.querySelector('.w2window_close');
                if (c) c.click();
            }
        });
    }""")
    time.sleep(0.5)

def close_samu_popup(page):
    """사무대행기관 팝업만 닫기 버튼으로 닫는다."""
    page.evaluate("""() => {
        document.querySelectorAll('.w2window').forEach(function(w) {
            if (w.id.indexOf('samuInfo') >= 0) {
                var c = w.querySelector('.w2window_close');
                if (c) c.click();
            }
        });
    }""")
    time.sleep(0.5)

def find_button_by_text(page, text, scope_id=None):
    """보이는 input[type=button] 중 텍스트가 포함된 것의 ID를 반환."""
    return page.evaluate("""(args) => {
        var scope = args.scopeId ? document.getElementById(args.scopeId) : document;
        if (!scope) return null;
        var btns = scope.querySelectorAll('input[type="button"]');
        for (var i = 0; i < btns.length; i++) {
            if (btns[i].value && btns[i].value.indexOf(args.text) >= 0 && btns[i].offsetWidth > 0) {
                return btns[i].id;
            }
        }
        return null;
    }""", {"text": text, "scopeId": scope_id})

def handle_pwd_popup(page, popup_prefix, download_list, timeout=30):
    """암호 팝업에서 비적용 선택 → 확인 → 다운로드 대기."""
    pwd_id = f"{popup_prefix}_excelPwd_popup"
    radio_comp = f"{pwd_id}_wframe_password_enabled"
    radio_dom = f"{pwd_id}_wframe_password_enabled_input_1"
    confirm_id = f"{pwd_id}_wframe_btn_confirm"

    start_len = len(download_list)

    for t in range(15):
        time.sleep(1)
        if len(download_list) > start_len:
            break
        result = page.evaluate(f"""() => {{
            var p = document.getElementById('{pwd_id}');
            if (!p || !(p.offsetParent !== null || p.offsetWidth > 0)) return {{found: false}};
            var r1 = document.getElementById('{radio_dom}');
            var btn = document.getElementById('{confirm_id}');
            return {{found: true, hasR1: !!r1, hasBtn: !!btn}};
        }}""")
        if result.get('found') and result.get('hasR1') and result.get('hasBtn'):
            page.evaluate(f"""() => {{
                try {{
                    var radio = $w.getComponentById('{radio_comp}');
                    if (radio && radio.setValue) radio.setValue('N');
                }} catch(e) {{}}
                var r = document.getElementById('{radio_dom}');
                if (r) {{ r.checked = true; r.click(); }}
            }}""")
            time.sleep(0.5)
            page.evaluate(f"document.getElementById('{confirm_id}').click()")
            break

    if len(download_list) <= start_len:
        for j in range(timeout):
            time.sleep(1)
            if len(download_list) > start_len:
                break
            if (j + 1) % 5 == 0:
                page.evaluate("""() => {
                    document.querySelectorAll('.w2window').forEach(function(w) {
                        if ((w.offsetParent !== null || w.offsetWidth > 0) &&
                            w.id.indexOf('excelPwd') < 0 && w.id.indexOf('WL0502') < 0) {
                            var btn = w.querySelector('input[type="button"], button');
                            if (btn) btn.click();
                        }
                    });
                }""")

    return download_list[-1] if len(download_list) > start_len else None

def is_login_page(page):
    """로그인 화면 판정 — 작업 화면 요소가 있는지도 함께 확인."""
    for _ in range(2):
        result = page.evaluate("""() => {
            var mgr = document.getElementById('mf_wfm_content_maeGwanriNo');
            var btn = document.getElementById('mf_wfm_content_btnSearch');
            var logout = document.body.innerText.indexOf('로그아웃') >= 0;
            return {hasMgr: !!mgr, hasBtn: !!btn, hasLogout: logout};
        }""")
        if result['hasMgr'] and result['hasBtn']:
            return False
        time.sleep(1)
    return True


# ── 메인 ──────────────────────────────────────────────
ensure_chrome_running()
with sync_playwright() as pw:
    browser = pw.chromium.connect_over_cdp(CDP_URL)
    ctx = browser.contexts[0]

    comwel = None
    for p in ctx.pages:
        if not p.is_closed() and "comwel" in p.url.lower():
            comwel = p
            break
    if not comwel:
        print("ERROR: 고용산재보험 토탈서비스 페이지를 찾을 수 없습니다.")
        print("  크롬에서 https://total.comwel.or.kr/ 에 접속하세요.")
        browser.close()
        sys.exit(1)

    if is_login_page(comwel):
        print("로그인이 필요합니다. 크롬에서 로그인 후 다시 실행하세요.")
        browser.close()
        sys.exit(1)

    close_visible_popups(comwel)
    print(f"저장 폴더: {SAVE_BASE}")
    print(f"조회 기간: {YEAR}년 {MONTH}월")
    print(f"대상 사업장: {len(FIRMS)}개\n")

    all_downloads = []
    comwel.on("download", lambda dl: all_downloads.append(dl))

    ok_count = 0
    skip_count = 0

    for idx, firm in enumerate(FIRMS):
        mgr = firm['mgr']
        mgr_fmt = fmt_mgr(mgr)
        print(f"[{idx+1}/{len(FIRMS)}] {firm['name']}  ({mgr_fmt})")

        if is_login_page(comwel):
            print("  !! 로그인 만료. 사람이 다시 로그인해야 합니다.")
            browser.close()
            sys.exit(1)

        close_visible_popups(comwel)

        comwel.evaluate(f"$w.getComponentById('mf_wfm_content_maeGwanriNo').setValue('{mgr_fmt}')")
        comwel.evaluate(f"""() => {{
            try {{ $w.getComponentById('mf_wfm_content_comYear').setValue('{YEAR} 년'); }} catch(e) {{}}
            try {{ $w.getComponentById('mf_wfm_content_comMM').setValue('{MONTH} 월'); }} catch(e) {{}}
        }}""")
        time.sleep(0.3)

        comwel.evaluate("document.getElementById('mf_wfm_content_btnSearch').click()")
        time.sleep(5)

        sjnm = comwel.evaluate(
            "(document.getElementById('mf_wfm_content_edtSaeopjangNm')||{}).value||''"
        )
        if not sjnm:
            print("  조회 실패 — 건너뜀")
            skip_count += 1
            time.sleep(3)
            continue
        print(f"  사업장명: {sjnm}")

        close_samu_popup(comwel)

        # 조회 후 버튼 ID 재탐색 (DOM 갱신으로 ID 변동)
        btn_dangwol = find_button_by_text(comwel, '당월보험료')
        btn_jiwon = find_button_by_text(comwel, '지원금')
        if not btn_dangwol or not btn_jiwon:
            print("  ERROR: 당월보험료/지원금 버튼 없음 — 건너뜀")
            skip_count += 1
            continue

        # ──────────────────────────────────────────────
        # A) 당월보험료 부과내역조회 (고용)
        # ──────────────────────────────────────────────
        print("  [당월보험료] 팝업 열기...")
        comwel.evaluate(f"document.getElementById('{btn_dangwol}').click()")
        time.sleep(5)

        comwel.evaluate("""() => {
            var wframe = $w.getComponentById('mf_wfm_content_WL0502_P04_wframe');
            var win = wframe.getWindow();
            var ds = win.dsInInfo;
            ds.setCellData(0, 'BOHEOM_FG', '3');
            win.imgBoheomName.setValue('고용보험');
            win.infoJumin.hide();
            var gojiYm = ds.getCellData(0, 'BUGWA_GOJI_YYMM') || '';
            if (gojiYm != '' && Number(gojiYm) >= 202010) {
                ds.setCellData(0, 'NEW_JY_YN', 'Y');
                win.grdGy2.show('');
                win.grdGy.hide();
            } else {
                win.grdGy.show('');
                win.grdGy2.hide();
            }
            win.grdSj.hide();
            win.grdSj2.hide();
            win.dsGySanjengInfo.removeAll();
            win.dsSjSanjengInfo.removeAll();
        }""")

        comwel.evaluate("""() => {
            var wframe = $w.getComponentById('mf_wfm_content_WL0502_P04_wframe');
            var win = wframe.getWindow();
            win.scwin.fn_dsInfoSet();
            win.scwin.fn_call('searchList');
        }""")
        time.sleep(5)

        gy_rows = comwel.evaluate("""() => {
            var wframe = $w.getComponentById('mf_wfm_content_WL0502_P04_wframe');
            var win = wframe.getWindow();
            return win.dsGySanjengInfo ? win.dsGySanjengInfo.getRowCount() : 0;
        }""")
        print(f"  [당월보험료] 고용 {gy_rows}건")

        if gy_rows > 0:
            excel_btn = find_button_by_text(comwel, '엑셀저장', 'mf_wfm_content_WL0502_P04')
            if excel_btn:
                comwel.evaluate(f"document.getElementById('{excel_btn}').click()")
                dl = handle_pwd_popup(comwel, 'mf_wfm_content_WL0502_P04_wframe', all_downloads, timeout=30)
                if dl:
                    fname = dl.suggested_filename or f"당월보험료_고용_{mgr}.xlsx"
                    firm_dir = os.path.join(SAVE_BASE, firm['name'])
                    os.makedirs(firm_dir, exist_ok=True)
                    save_path = os.path.join(firm_dir, fname)
                    dl.save_as(save_path)
                    fsize = os.path.getsize(save_path)
                    print(f"  [당월보험료] 저장 OK  ({fsize:,}바이트)")
                else:
                    print("  [당월보험료] 다운로드 실패")
            else:
                print("  [당월보험료] 엑셀저장 버튼 없음")
        else:
            print("  [당월보험료] 데이터 없음 — 건너뜀")

        close_visible_popups(comwel)
        time.sleep(1)

        # ──────────────────────────────────────────────
        # B) 고용보험료 지원금 정보
        # ──────────────────────────────────────────────
        print("  [지원금] 팝업 열기...")
        comwel.evaluate(f"document.getElementById('{btn_jiwon}').click()")
        time.sleep(5)

        jiwon_rows = comwel.evaluate("""() => {
            try {
                var wframe = $w.getComponentById('mf_wfm_content_WL0502_P02_wframe');
                var win = wframe.getWindow();
                return win.dsGyJiwonPrcInfo.getRowCount();
            } catch(e) { return -1; }
        }""")
        print(f"  [지원금] {jiwon_rows}건")

        excel_btn2 = find_button_by_text(comwel, '엑셀저장', 'mf_wfm_content_WL0502_P02')
        if excel_btn2:
            comwel.evaluate(f"document.getElementById('{excel_btn2}').click()")
            dl2 = handle_pwd_popup(comwel, 'mf_wfm_content_WL0502_P02_wframe', all_downloads, timeout=30)
            if dl2:
                fname2 = dl2.suggested_filename or f"지원금_{mgr}.xlsx"
                firm_dir = os.path.join(SAVE_BASE, firm['name'])
                os.makedirs(firm_dir, exist_ok=True)
                save_path2 = os.path.join(firm_dir, fname2)
                dl2.save_as(save_path2)
                fsize2 = os.path.getsize(save_path2)
                print(f"  [지원금] 저장 OK  ({fsize2:,}바이트)")
            else:
                print("  [지원금] 다운로드 실패")
        else:
            print("  [지원금] 엑셀저장 버튼 없음")

        close_visible_popups(comwel)
        ok_count += 1

        delay = random.uniform(2, 5)
        print(f"  완료 ({delay:.1f}초 대기)\n")
        time.sleep(delay)

    print("=" * 50)
    print(f"완료: {ok_count}개 처리, {skip_count}개 건너뜀")
    print(f"저장 위치: {SAVE_BASE}")
    browser.close()
