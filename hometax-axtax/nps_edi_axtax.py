#!/usr/bin/env python -I
# -*- coding: utf-8 -*-
"""국민연금 EDI - 사업장별 결정내역(2차) + 국고지원내역 PDF 다운로드.

크롬이 --remote-debugging-port=9333 으로 떠 있고,
edi.nps.or.kr 에 로그인 후 nexacro/index.html이 열려 있어야 한다.

워크플로 (사업장 당):
  사업장전환(ChangeBusi) → 결정내역(M08010000) → 2차(ED08) →
  M08010200 → Tabpage1: fnPrint(2차최종결정내역) PDF →
              Tabpage5: fnRdPrint01(지원대상자) + fnRdPrint02(지원예정자) PDF
"""
import sys, time, os, subprocess, urllib.request

sys.stdout.reconfigure(encoding="utf-8")

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
        "--disable-features=LocalNetworkAccessChecks", "--new-window", "https://edi.nps.or.kr/",
    ])
    for _ in range(30):
        time.sleep(2)
        if _chrome_ready():
            return
    raise RuntimeError("크롬을 띄웠지만 60초 안에 9333에 연결되지 않았습니다.")
from playwright.sync_api import sync_playwright

# ── 사업장 목록 (NHIS sym + NPS bplcMngNo 매핑) ──────────
FIRMS = [
    {"sym": "70888510", "bplc": "84887030210", "name": "세무법인 역삼"},
    {"sym": "70070950", "bplc": "20382317340", "name": "(재)한국대학봉사회"},
    {"sym": "73080892", "bplc": "92540270681", "name": "(주)올리브트리 커뮤니케이션"},
    {"sym": "72017243", "bplc": "10981730600", "name": "(주)대광건화"},
    {"sym": "79493259", "bplc": "88587023940", "name": "(주)에스랩엔지니어링"},
    {"sym": "72425585", "bplc": "21587161620", "name": "（주）정희농산"},
    {"sym": "71461360", "bplc": "20350677930", "name": "남일스포츠"},
    {"sym": "72688055", "bplc": "21521252910", "name": "베리포커스드"},
    {"sym": "77750031", "bplc": "49607010620", "name": "아이러브아트센터"},
    {"sym": "76277169", "bplc": "91610209981", "name": "유니베라신대방대리점"},
    {"sym": "76349721", "bplc": "20630678480", "name": "유비컴퍼니（ubcompany）"},
    {"sym": "77103547", "bplc": "12423849120", "name": "이남수부동산"},
    {"sym": "78313973", "bplc": "26781015420", "name": "주식회사경상농산"},
    {"sym": "76341510", "bplc": "57988003290", "name": "주식회사비케이컨셉트"},
    {"sym": "70913996", "bplc": "17986033980", "name": "주식회사아리씨엠(RECMCorp.)"},
    {"sym": "73277724", "bplc": "84388027040", "name": "주식회사 에스에스피코리아"},
    {"sym": "73392812", "bplc": "33187036680", "name": "주식회사 에스에스피파트너스"},
    {"sym": "73338646", "bplc": "48988029460", "name": "주식회사 케이엔태경이엔씨"},
    {"sym": "73681322", "bplc": "45681039460", "name": "주식회사 플레이모도클럽(Playmodo Club Co., Ltd.)"},
    {"sym": "72529728", "bplc": "14181142210", "name": "(주)한신소망사랑"},
    {"sym": "79666301", "bplc": "53988024790", "name": "큐플러스매니지먼트주식회사"},
    {"sym": "73563365", "bplc": "35414033140", "name": "폴라멜로"},
    {"sym": "78343496", "bplc": "19186017830", "name": "현진프라퍼티주식회사"},
]

SAVE_BASE = r"C:\Users\admin\Desktop\사회보험"
DELAY_BETWEEN = 3
BTN_CHANGE_X, BTN_CHANGE_Y = 1169, 40


def get_nps(ctx):
    for p in ctx.pages:
        if not p.is_closed() and "nexacro/index.html" in p.url:
            return p
    return None


def clear_overlays(nps):
    nps.evaluate(
        "() => {"
        "  document.querySelectorAll('.nexamodaloverlay')"
        "    .forEach(o => o.style.display = 'none');"
        "  document.querySelectorAll('[id*=\"16775\"]')"
        "    .forEach(el => el.style.display = 'none');"
        "}"
    )


def is_login_page(nps):
    url = nps.url
    if "login" in url.lower() and "nexacro" not in url:
        return True
    return False


def switch_firm(nps, bplc_mng_no):
    """사업장을 전환한다. 성공하면 True."""
    clear_overlays(nps)
    time.sleep(0.5)

    nps.mouse.click(BTN_CHANGE_X, BTN_CHANGE_Y)
    time.sleep(4)

    popup_ok = nps.evaluate(
        "() => {"
        "  try {"
        "    const cb = nexacro.getApplication()"
        "      .mainframe.VFrameSet.FrameSdi.ChangeBusi;"
        "    return !!cb;"
        "  } catch(e) { return false; }"
        "}"
    )
    if not popup_ok:
        clear_overlays(nps)
        time.sleep(1)
        nps.mouse.click(BTN_CHANGE_X, BTN_CHANGE_Y)
        time.sleep(4)
        popup_ok = nps.evaluate(
            "() => {"
            "  try {"
            "    return !!nexacro.getApplication()"
            "      .mainframe.VFrameSet.FrameSdi.ChangeBusi;"
            "  } catch(e) { return false; }"
            "}"
        )
    if not popup_ok:
        return False

    nps.evaluate(
        f"() => {{"
        f"  const cb = nexacro.getApplication()"
        f"    .mainframe.VFrameSet.FrameSdi.ChangeBusi;"
        f"  const pf = cb.form.divPopBg.form.divPopWork.form;"
        f"  pf.dsSearch.setColumn(0, 'bplcMngNo', '{bplc_mng_no}');"
        f"  pf.fnChangeLoginTran();"
        f"}}"
    )
    time.sleep(6)

    cur = nps.evaluate(
        "() => {"
        "  try {"
        "    const gu = JSON.parse(localStorage.getItem('gdsUser') || '{}');"
        "    return gu.bplcMngNo || '';"
        "  } catch(e) { return ''; }"
        "}"
    )
    if cur == bplc_mng_no:
        clear_overlays(nps)
        time.sleep(2)
        return True

    time.sleep(5)
    cur = nps.evaluate(
        "() => {"
        "  try {"
        "    return JSON.parse(localStorage.getItem('gdsUser') || '{}')"
        "      .bplcMngNo || '';"
        "  } catch(e) { return ''; }"
        "}"
    )
    if cur == bplc_mng_no:
        clear_overlays(nps)
        time.sleep(2)
        return True

    return False


def close_divworks(nps):
    try:
        nps.evaluate(
            "() => {"
            "  try {"
            "    const form = nexacro.getApplication()"
            "      .mainframe.VFrameSet.FrameSdi.form;"
            "    ['divWork_M08010200', 'divWork_M08010000'].forEach(k => {"
            "      if (form[k]) {"
            "        try { form[k].destroy(); } catch(e) {}"
            "      }"
            "    });"
            "  } catch(e) {}"
            "}"
        )
    except Exception:
        pass
    time.sleep(1)


def open_menu_m08010000(nps):
    close_divworks(nps)
    try:
        nps.evaluate(
            "() => {"
            "  const form = nexacro.getApplication()"
            "    .mainframe.VFrameSet.FrameSdi.form;"
            "  form.gfnOpenMenu('M08010000');"
            "}"
        )
    except Exception:
        nps.reload(wait_until="domcontentloaded", timeout=30000)
        time.sleep(10)
        nps.evaluate(
            "() => {"
            "  const form = nexacro.getApplication()"
            "    .mainframe.VFrameSet.FrameSdi.form;"
            "  form.gfnOpenMenu('M08010000');"
            "}"
        )
    time.sleep(5)


def get_ed08_rows(nps):
    return nps.evaluate(
        "() => {"
        "  const form = nexacro.getApplication()"
        "    .mainframe.VFrameSet.FrameSdi.form;"
        "  if (!form.divWork_M08010000) return [];"
        "  const wf = form.divWork_M08010000.form.divWork.form;"
        "  const rows = [];"
        "  for (let r = 0; r < wf.dsList1.rowcount; r++) {"
        "    if (wf.dsList1.getColumn(r, 'wediTmpltCd') === 'ED08') {"
        "      rows.push({row: r,"
        "        rcptnYmd: wf.dsList1.getColumn(r, 'rcptnYmd'),"
        "        wediDocNo: wf.dsList1.getColumn(r, 'wediDocNo')});"
        "    }"
        "  }"
        "  return rows;"
        "}"
    )


def open_ed08(nps, row_idx):
    nps.evaluate(
        f"() => {{"
        f"  const form = nexacro.getApplication()"
        f"    .mainframe.VFrameSet.FrameSdi.form;"
        f"  const wf = form.divWork_M08010000.form.divWork.form;"
        f"  var wediTmpltCd = 'ED08';"
        f"  var oParam = wf.cfnDocuMatchWediTmpltCd(wediTmpltCd);"
        f"  oParam.wediTmpltCd = wediTmpltCd;"
        f"  oParam.wediDocNo = wf.dsList1.getColumn({row_idx}, 'wediDocNo');"
        f"  oParam.rcptnYmd = wf.dsList1.getColumn({row_idx}, 'rcptnYmd');"
        f"  oParam.bplcMngNo = wf.bplcMngNo;"
        f"  oParam.bplcMngNm = wf.bplcMngNm;"
        f"  oParam.rcptnWediDocNo = wf.dsList1.getColumn({row_idx}, 'rcptnWediDocNo');"
        f"  oParam.issym = wf.dsList1.getColumn({row_idx}, 'issym');"
        f"  wf.gfnOpenMenuForm(oParam.menuId, oParam);"
        f"}}"
    )
    time.sleep(5)

    loaded = nps.evaluate(
        "() => {"
        "  const form = nexacro.getApplication()"
        "    .mainframe.VFrameSet.FrameSdi.form;"
        "  if (!form.divWork_M08010200) return null;"
        "  const wf = form.divWork_M08010200.form.divWork.form;"
        "  return {bplcMngNo: wf.bplcMngNo, dcsnYm: wf.dcsnYm,"
        "    bplcMngNm: wf.bplcMngNm};"
        "}"
    )
    return loaded


def _cleanup_extra_pages(ctx):
    for p in ctx.pages:
        if p.is_closed():
            continue
        u = p.url
        if "rdPreview.do" in u or u == "about:blank":
            try:
                p.close()
            except Exception:
                pass
    time.sleep(0.5)


def _ensure_nexacro(ctx):
    nps = get_nps(ctx)
    if not nps:
        return None
    try:
        ok = nps.evaluate(
            "() => { try { nexacro.getApplication().mainframe.VFrameSet.FrameSdi.form; return true; }"
            " catch(e) { return false; } }"
        )
        if not ok:
            raise Exception("nexacro broken")
    except Exception:
        nps.reload(wait_until="domcontentloaded", timeout=30000)
        time.sleep(10)
        nps = get_nps(ctx)
    return nps


def print_and_save_pdf(nps, ctx, firm_name, dcsn_ym):
    """popupReport 인터셉트 → iframe POST → CrownIX iframe 렌더링 → PDF 버튼 클릭."""
    for p in ctx.pages:
        if not p.is_closed() and p.url == "about:blank":
            try:
                p.close()
            except Exception:
                pass

    nps.evaluate("""() => {
        window.__rdCaptured = null;
        var f = nexacro.getApplication().mainframe.VFrameSet.FrameSdi.form;
        var tab = f.divWork_M08010200.form.divWork.form.tab00.Tabpage1.form;
        tab._orig = tab.popupReport;
        tab.popupReport = function(n, d) {
            window.__rdCaptured = {rdName: n, rdStr: '/rnl [!~#] /rdata [' + d + ']'};
        };
        var wf = f.divWork_M08010200.form.divWork.form;
        tab.fnPrint(wf.bplcMngNo, wf.dcsnYm, wf.bplcMngNm, '2');
    }""")

    captured = None
    for _ in range(20):
        time.sleep(1)
        captured = nps.evaluate("() => window.__rdCaptured")
        if captured:
            break

    nps.evaluate("""() => {
        var f = nexacro.getApplication().mainframe.VFrameSet.FrameSdi.form;
        var tab = f.divWork_M08010200.form.divWork.form.tab00.Tabpage1.form;
        if (tab._orig) { tab.popupReport = tab._orig; delete tab._orig; }
    }""")

    if not captured:
        return None, 0

    html = nps.evaluate("""(p) => {
        return new Promise((resolve) => {
            var ifr = document.createElement('iframe');
            ifr.name = '_rd_' + Date.now();
            ifr.style.cssText = 'position:absolute;left:-9999px;width:1200px;height:900px;';
            document.body.appendChild(ifr);
            var loaded = false;
            ifr.onload = function() {
                if (loaded) return;
                loaded = true;
                setTimeout(function() {
                    try { resolve(ifr.contentDocument.documentElement.outerHTML); }
                    catch(e) { resolve(''); }
                    finally { try { ifr.remove(); } catch(e) {} }
                }, 5000);
            };
            var f = document.createElement('form');
            f.method = 'POST'; f.action = '/comm/rdPreview.do'; f.target = ifr.name;
            var i1 = document.createElement('input');
            i1.type = 'hidden'; i1.name = 'rdName'; i1.value = p.rdName;
            f.appendChild(i1);
            var i2 = document.createElement('input');
            i2.type = 'hidden'; i2.name = 'rdData'; i2.value = p.rdStr;
            f.appendChild(i2);
            document.body.appendChild(f); f.submit(); f.remove();
            setTimeout(() => { try { ifr.remove(); } catch(e) {} if (!loaded) resolve(''); }, 30000);
        });
    }""", captured)

    if not html or len(html) < 500 or "ErrorCode" in html:
        return None, 0

    nps.evaluate("""(html) => {
        document.querySelectorAll('iframe[id="rdViewFrame"]').forEach(f => f.remove());
        var ifr = document.createElement('iframe');
        ifr.id = 'rdViewFrame';
        ifr.style.cssText = 'position:fixed;top:0;left:0;width:100vw;height:100vh;z-index:999999;border:none;background:white;';
        document.body.appendChild(ifr);
        var doc = ifr.contentDocument;
        doc.open(); doc.write(html); doc.close();
    }""", html)
    time.sleep(20)

    has_btn = nps.evaluate("""() => {
        var ifr = document.getElementById('rdViewFrame');
        if (!ifr) return false;
        return !!ifr.contentDocument.querySelector('button[title="PDF 파일로 저장"]');
    }""")
    if not has_btn:
        nps.evaluate("() => { var f = document.getElementById('rdViewFrame'); if (f) f.remove(); }")
        return None, 0

    year = dcsn_ym[:4]
    mon = str(int(dcsn_ym[4:6]))
    save_dir = os.path.join(SAVE_BASE, firm_name)
    os.makedirs(save_dir, exist_ok=True)
    filename = f"국민연금_2차최종결정내역_{firm_name}_{year}년{mon}월.pdf"
    save_path = os.path.join(save_dir, filename)

    try:
        with nps.expect_download(timeout=30000) as download_info:
            nps.evaluate("""() => {
                var ifr = document.getElementById('rdViewFrame');
                ifr.contentDocument.querySelector('button[title="PDF 파일로 저장"]').click();
            }""")
        download = download_info.value
        download.save_as(save_path)
    except Exception:
        nps.evaluate("() => { var f = document.getElementById('rdViewFrame'); if (f) f.remove(); }")
        _cleanup_extra_pages(ctx)
        return None, 0

    nps.evaluate("() => { var f = document.getElementById('rdViewFrame'); if (f) f.remove(); }")
    _cleanup_extra_pages(ctx)

    if not os.path.exists(save_path) or os.path.getsize(save_path) < 1000:
        return None, 0
    return save_path, os.path.getsize(save_path)


def print_tab5_pdf(nps, ctx, report_type, firm_name, dcsn_ym):
    """Tabpage5 국고지원내역 PDF. report_type: 'support' 또는 'planned'."""
    _cleanup_extra_pages(ctx)

    fn_name = "fnRdPrint01" if report_type == "support" else "fnRdPrint02"
    ds_name = "dsExcelList1" if report_type == "support" else "dsExcelList2"
    label = "지원대상자" if report_type == "support" else "지원예정자"

    nps.evaluate("""() => {
        var f = nexacro.getApplication().mainframe.VFrameSet.FrameSdi.form;
        var wf = f.divWork_M08010200.form.divWork.form;
        if (wf.tab00.tabindex !== 4) wf.tab00.set_tabindex(4);
    }""")
    time.sleep(3)

    row_count = nps.evaluate(f"""() => {{
        var f = nexacro.getApplication().mainframe.VFrameSet.FrameSdi.form;
        var tp = f.divWork_M08010200.form.divWork.form.tab00.Tabpage5;
        return tp.form.{ds_name} ? tp.form.{ds_name}.rowcount : 0;
    }}""")
    if row_count == 0:
        return None, 0

    nps.evaluate(f"""() => {{
        window.__rdCaptured = null;
        var f = nexacro.getApplication().mainframe.VFrameSet.FrameSdi.form;
        var tp = f.divWork_M08010200.form.divWork.form.tab00.Tabpage5.form;
        tp._origPopup = tp.popupReport;
        tp.popupReport = function(n, d) {{
            window.__rdCaptured = {{rdName: n, rdStr: '/rnl [!~#] /rdata [' + d + ']'}};
        }};
        tp.{fn_name}('', '1');
    }}""")

    captured = None
    for _ in range(20):
        time.sleep(1)
        captured = nps.evaluate("() => window.__rdCaptured")
        if captured:
            break

    nps.evaluate("""() => {
        var f = nexacro.getApplication().mainframe.VFrameSet.FrameSdi.form;
        var tp = f.divWork_M08010200.form.divWork.form.tab00.Tabpage5.form;
        if (tp._origPopup) { tp.popupReport = tp._origPopup; delete tp._origPopup; }
    }""")

    if not captured:
        return None, 0

    html = nps.evaluate("""(p) => {
        return new Promise((resolve) => {
            var ifr = document.createElement('iframe');
            ifr.name = '_rd_' + Date.now();
            ifr.style.cssText = 'position:absolute;left:-9999px;width:1200px;height:900px;';
            document.body.appendChild(ifr);
            var loaded = false;
            ifr.onload = function() {
                if (loaded) return;
                loaded = true;
                setTimeout(function() {
                    try { resolve(ifr.contentDocument.documentElement.outerHTML); }
                    catch(e) { resolve(''); }
                    finally { try { ifr.remove(); } catch(e) {} }
                }, 5000);
            };
            var f = document.createElement('form');
            f.method = 'POST'; f.action = '/comm/rdPreview.do'; f.target = ifr.name;
            var i1 = document.createElement('input');
            i1.type = 'hidden'; i1.name = 'rdName'; i1.value = p.rdName;
            f.appendChild(i1);
            var i2 = document.createElement('input');
            i2.type = 'hidden'; i2.name = 'rdData'; i2.value = p.rdStr;
            f.appendChild(i2);
            document.body.appendChild(f); f.submit(); f.remove();
            setTimeout(() => { try { ifr.remove(); } catch(e) {} if (!loaded) resolve(''); }, 30000);
        });
    }""", captured)

    if not html or len(html) < 500 or "ErrorCode" in html:
        return None, 0

    nps.evaluate("""(html) => {
        document.querySelectorAll('iframe[id="rdViewFrame"]').forEach(f => f.remove());
        var ifr = document.createElement('iframe');
        ifr.id = 'rdViewFrame';
        ifr.style.cssText = 'position:fixed;top:0;left:0;width:100vw;height:100vh;z-index:999999;border:none;background:white;';
        document.body.appendChild(ifr);
        var doc = ifr.contentDocument;
        doc.open(); doc.write(html); doc.close();
    }""", html)
    time.sleep(20)

    has_btn = nps.evaluate("""() => {
        var ifr = document.getElementById('rdViewFrame');
        if (!ifr) return false;
        return !!ifr.contentDocument.querySelector('button[title="PDF 파일로 저장"]');
    }""")
    if not has_btn:
        nps.evaluate("() => { var f = document.getElementById('rdViewFrame'); if (f) f.remove(); }")
        return None, 0

    year = dcsn_ym[:4]
    mon = str(int(dcsn_ym[4:6]))
    save_dir = os.path.join(SAVE_BASE, firm_name)
    os.makedirs(save_dir, exist_ok=True)
    filename = f"국민연금_국고지원내역_{label}_{firm_name}_{year}년{mon}월.pdf"
    save_path = os.path.join(save_dir, filename)

    try:
        with nps.expect_download(timeout=30000) as download_info:
            nps.evaluate("""() => {
                var ifr = document.getElementById('rdViewFrame');
                ifr.contentDocument.querySelector('button[title="PDF 파일로 저장"]').click();
            }""")
        download = download_info.value
        download.save_as(save_path)
    except Exception:
        nps.evaluate("() => { var f = document.getElementById('rdViewFrame'); if (f) f.remove(); }")
        _cleanup_extra_pages(ctx)
        return None, 0

    nps.evaluate("() => { var f = document.getElementById('rdViewFrame'); if (f) f.remove(); }")
    _cleanup_extra_pages(ctx)

    if not os.path.exists(save_path) or os.path.getsize(save_path) < 1000:
        return None, 0
    return save_path, os.path.getsize(save_path)


def main():
    ensure_chrome_running()
    with sync_playwright() as pw:
        browser = pw.chromium.connect_over_cdp(CDP_URL)
        ctx = browser.contexts[0]

        nps = get_nps(ctx)
        if not nps:
            print("ERROR: NPS Nexacro 페이지를 찾을 수 없습니다.")
            browser.close()
            return

        if is_login_page(nps):
            print("ERROR: 로그인 필요 — 사람이 로그인해야 합니다.")
            browser.close()
            return

        cur_bplc = nps.evaluate(
            "() => {"
            "  try {"
            "    return JSON.parse(localStorage.getItem('gdsUser') || '{}')"
            "      .bplcMngNo || '';"
            "  } catch(e) { return ''; }"
            "}"
        )
        print(f"현재 사업장: {cur_bplc}")

        ok_count = 0
        skip_count = 0
        fail_count = 0
        results = []

        for i, firm in enumerate(FIRMS):
            bplc = firm["bplc"]
            name = firm["name"]
            print(f"\n[{i + 1}/{len(FIRMS)}] {name}")

            try:
                _cleanup_extra_pages(ctx)
                nps = _ensure_nexacro(ctx)
                if not nps or is_login_page(nps):
                    print("  로그인 화면 감지 — 중단")
                    results.append({"name": name, "status": "login_required"})
                    break

                need_switch = nps.evaluate(
                    f"() => {{"
                    f"  try {{"
                    f"    return JSON.parse(localStorage.getItem('gdsUser') || '{{}}')"
                    f"      .bplcMngNo !== '{bplc}';"
                    f"  }} catch(e) {{ return true; }}"
                    f"}}"
                )

                if need_switch:
                    print("  사업장 전환...")
                    if not switch_firm(nps, bplc):
                        print("  전환 실패")
                        fail_count += 1
                        results.append({"name": name, "status": "switch_fail"})
                        continue

                    nps = get_nps(ctx)
                    if not nps:
                        print("  전환 후 페이지 없음")
                        fail_count += 1
                        results.append({"name": name, "status": "page_lost"})
                        continue

                clear_overlays(nps)
                time.sleep(1)

                print("  결정내역 조회...")
                open_menu_m08010000(nps)

                ed08_rows = get_ed08_rows(nps)
                if not ed08_rows:
                    print("  ED08(2차 결정내역) 없음 — 건너뜀")
                    skip_count += 1
                    results.append({"name": name, "status": "no_ed08"})
                    time.sleep(DELAY_BETWEEN)
                    continue

                best = max(ed08_rows, key=lambda x: x["rcptnYmd"])
                print(f"  ED08 발견: row={best['row']} ({best['rcptnYmd']})")

                detail = open_ed08(nps, best["row"])
                if not detail:
                    print("  2차 결정내역 열기 실패")
                    fail_count += 1
                    results.append({"name": name, "status": "open_fail"})
                    time.sleep(DELAY_BETWEEN)
                    continue

                dcsn_ym = detail.get("dcsnYm", "")
                print(f"  출력(fnPrint) → PDF...")
                path, size = print_and_save_pdf(nps, ctx, name, dcsn_ym)
                if not path:
                    print("  PDF 저장 실패")
                    fail_count += 1
                    results.append({"name": name, "status": "pdf_fail"})
                    time.sleep(DELAY_BETWEEN)
                    continue

                print(f"  저장: {os.path.basename(path)} ({size:,} bytes)")
                ok_count += 1
                result_entry = {"name": name, "status": "ok", "month": dcsn_ym, "size": size}

                # 국고지원내역 (Tabpage5) PDF
                for rtype, rlabel in [("support", "지원대상자"), ("planned", "지원예정자")]:
                    print(f"  국고지원내역({rlabel}) PDF...")
                    t5path, t5size = print_tab5_pdf(nps, ctx, rtype, name, dcsn_ym)
                    if t5path:
                        print(f"  저장: {os.path.basename(t5path)} ({t5size:,} bytes)")
                        result_entry[f"tab5_{rtype}"] = t5size
                    else:
                        print(f"  {rlabel}: 데이터 없음 또는 실패")
                    time.sleep(DELAY_BETWEEN)

                results.append(result_entry)

            except Exception as e:
                print(f"  ERROR: {e}")
                fail_count += 1
                results.append({"name": name, "status": "error", "msg": str(e)})

            time.sleep(DELAY_BETWEEN)

        browser.close()

        print(f"\n{'=' * 50}")
        print(f"완료: 성공 {ok_count} / 건너뜀 {skip_count} / 실패 {fail_count}")
        print(f"{'=' * 50}")
        for r in results:
            s = r["status"]
            if s == "ok":
                t5s = r.get("tab5_support", 0)
                t5p = r.get("tab5_planned", 0)
                t5_info = ""
                if t5s or t5p:
                    parts = []
                    if t5s:
                        parts.append(f"지원대상{t5s:,}B")
                    if t5p:
                        parts.append(f"지원예정{t5p:,}B")
                    t5_info = f" + 국고[{', '.join(parts)}]"
                print(f"  OK   {r['name']} ({r['month']}, {r['size']:,}B{t5_info})")
            elif s == "no_ed08":
                print(f"  SKIP {r['name']} — ED08 없음")
            elif s == "login_required":
                print(f"  STOP {r['name']} — 로그인 필요")
            else:
                print(f"  FAIL {r['name']} — {r.get('msg', s)}")


if __name__ == "__main__":
    main()
