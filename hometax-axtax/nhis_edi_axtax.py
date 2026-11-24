#!/usr/bin/env python -I
# -*- coding: utf-8 -*-
"""건강보험 EDI - 가입자고지내역서(건강) PDF 다운로드.

크롬이 --remote-debugging-port=9333 으로 떠 있고,
EDI 로그인 후 Nexacro(webedi/xui/index.jsp)가 열려 있어야 한다.

워크플로:
  WETC_110(받은문서) → BMBB_020(가입자고지내역서) 찾기/열기
  → fn_command_print() → ClipReport 팝업 → pdfDownLoad() → PDF 저장
"""
import sys, time, base64, os, glob, shutil, subprocess, urllib.request

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
        "--disable-features=LocalNetworkAccessChecks", "--new-window", "https://edi.nhis.or.kr/",
    ])
    for _ in range(30):
        time.sleep(2)
        if _chrome_ready():
            return
    raise RuntimeError("크롬을 띄웠지만 60초 안에 9333에 연결되지 않았습니다.")

# ── 사업장 목록 ──────────────────────────────────────────
FIRMS = [
    {"sym": "70888510", "name": "세무법인 역삼"},
    {"sym": "70070950", "name": "(재)한국대학봉사회"},
    {"sym": "73080892", "name": "(주)올리브트리 커뮤니케이션"},
    {"sym": "72017243", "name": "(주)대광건화"},
    {"sym": "79493259", "name": "(주)에스랩엔지니어링"},
    {"sym": "72425585", "name": "（주）정희농산"},
    {"sym": "71461360", "name": "남일스포츠"},
    {"sym": "72688055", "name": "베리포커스드"},
    {"sym": "77750031", "name": "아이러브아트센터"},
    {"sym": "76277169", "name": "유니베라신대방대리점"},
    {"sym": "76349721", "name": "유비컴퍼니（ubcompany）"},
    {"sym": "77103547", "name": "이남수부동산"},
    {"sym": "78313973", "name": "주식회사경상농산"},
    {"sym": "76341510", "name": "주식회사비케이컨셉트"},
    {"sym": "73352663", "name": "주식회사 송평"},
    {"sym": "70913996", "name": "주식회사아리씨엠(RECMCorp.)"},
    {"sym": "73277724", "name": "주식회사 에스에스피코리아"},
    {"sym": "73392812", "name": "주식회사 에스에스피파트너스"},
    {"sym": "73338646", "name": "주식회사 케이엔태경이엔씨"},
    {"sym": "73681322", "name": "주식회사 플레이모도클럽(Playmodo Club Co., Ltd.)"},
    {"sym": "72529728", "name": "(주)한신소망사랑"},
    {"sym": "79666301", "name": "큐플러스매니지먼트주식회사"},
    {"sym": "73563365", "name": "폴라멜로"},
    {"sym": "78343496", "name": "현진프라퍼티주식회사"},
]

SAVE_BASE = r"C:\Users\admin\Desktop\사회보험"
DOWNLOAD_TEMP = os.path.join(os.environ.get("TEMP", r"C:\Temp"), "nhis_pdf_dl")
DELAY_BETWEEN = 3


def get_nxa(ctx):
    """Nexacro 페이지 반환."""
    for p in ctx.pages:
        if not p.is_closed() and "webedi/xui/index.jsp" in p.url:
            return p
    return None


def get_main(ctx):
    """메인 페이지 반환."""
    for p in ctx.pages:
        if not p.is_closed() and "retrieveMain" in p.url:
            return p
    return None


def ensure_wetc110(nxa):
    """받은문서 목록(WETC_110)으로 이동."""
    url = nxa.evaluate(
        "() => window.application.mainframe.childframe.form.div_body.url || ''"
    )
    if "WETC_110" not in url:
        nxa.evaluate(
            "() => window.application.mainframe.childframe.form"
            ".div_body.set_url('WE::WETC_110.xfdl')"
        )
        time.sleep(4)


def find_doc(nxa, docid, target_month=None):
    """ds_list에서 docid에 해당하는 최신 문서 행 반환."""
    rows = nxa.evaluate(
        f"""() => {{
        const ds = window.application.mainframe.childframe.form.div_body.ds_list;
        const result = [];
        for (let r = 0; r < ds.rowcount; r++) {{
            if (ds.getColumn(r, 'DOCID') === '{docid}') {{
                result.push({{
                    row: r,
                    month: ds.getColumn(r, 'GOJI_YYYYMM'),
                    chasu: ds.getColumn(r, 'WRT_CHASU')
                }});
            }}
        }}
        return result;
    }}"""
    )
    if not rows:
        return None
    if target_month:
        rows = [r for r in rows if r["month"] == target_month]
    if not rows:
        return None
    return max(rows, key=lambda x: (x["month"], x["chasu"]))


def open_doc(nxa, row_idx):
    """fn_OpenDoc으로 문서를 연다."""
    nxa.evaluate(
        f"() => window.application.mainframe.childframe.form"
        f".div_body.fn_OpenDoc({row_idx})"
    )
    time.sleep(5)
    url = nxa.evaluate(
        "() => window.application.mainframe.childframe.form.div_body.url || ''"
    )
    return "BMBB_020" in url


def print_and_download_pdf(nxa, ctx, download_dir):
    """fn_command_print → ClipReport pdfDownLoad → 다운로드된 파일 경로 반환."""
    os.makedirs(download_dir, exist_ok=True)
    for f in glob.glob(os.path.join(download_dir, "*")):
        os.remove(f)

    new_pages = []

    def on_page(page):
        new_pages.append(page)

    ctx.on("page", on_page)

    nxa.evaluate(
        "() => window.application.mainframe.childframe.form"
        ".div_body.fn_command_print()"
    )
    time.sleep(10)

    popup = None
    for p in ctx.pages:
        if not p.is_closed() and "popup.html" in p.url:
            popup = p
            break

    ctx.remove_listener("page", on_page)

    if not popup:
        return None

    popup.wait_for_load_state("networkidle", timeout=15000)
    time.sleep(5)

    iframe_el = popup.query_selector(
        "#p_div_body_form_web_report_WebBrowser"
    )
    if not iframe_el:
        popup.close()
        return None

    frame = iframe_el.content_frame()
    if not frame:
        popup.close()
        return None

    frame.wait_for_load_state("networkidle", timeout=15000)
    time.sleep(3)

    has_report = frame.evaluate(
        "() => typeof m_reportHashMap !== 'undefined'"
        " && Object.keys(m_reportHashMap).length > 0"
    )
    if not has_report:
        popup.close()
        return None

    cdp = ctx.new_cdp_session(popup)
    cdp.send(
        "Browser.setDownloadBehavior",
        {
            "behavior": "allowAndName",
            "downloadPath": download_dir,
            "eventsEnabled": True,
        },
    )

    frame.evaluate(
        "() => { const k = Object.keys(m_reportHashMap)[0];"
        " m_reportHashMap[k].pdfDownLoad(); }"
    )
    time.sleep(12)
    cdp.detach()
    popup.close()

    files = os.listdir(download_dir)
    if not files:
        return None
    fpath = os.path.join(download_dir, files[0])
    if os.path.getsize(fpath) < 1000:
        return None
    return fpath


def save_pdf(src_path, firm_name, month_str):
    """다운로드된 파일을 대상 경로로 복사."""
    year = month_str[:4]
    mon = str(int(month_str[4:6]))
    save_dir = os.path.join(SAVE_BASE, firm_name)
    os.makedirs(save_dir, exist_ok=True)
    filename = f"건강보험개인별내역_{firm_name}_{year}년{mon}월.pdf"
    dest = os.path.join(save_dir, filename)
    shutil.copy2(src_path, dest)
    return dest, os.path.getsize(dest)


def main():
    ensure_chrome_running()
    with sync_playwright() as pw:
        browser = pw.chromium.connect_over_cdp(CDP_URL)
        ctx = browser.contexts[0]

        nxa = get_nxa(ctx)
        if not nxa:
            print("ERROR: Nexacro 페이지를 찾을 수 없습니다.")
            print("  EDI 로그인 후 Nexacro가 열려야 합니다.")
            browser.close()
            return

        firm_sym = nxa.evaluate("() => window.application.GV_FIRM_SYM")
        mdt = nxa.evaluate("() => window.application.GV_MDT_CHK_YN || '0'")
        print(f"현재 사업장: GV_FIRM_SYM={firm_sym}, MDT={mdt}")

        ensure_wetc110(nxa)
        time.sleep(2)

        doc = find_doc(nxa, "BMBB_020")
        if not doc:
            print("BMBB_020 문서 없음")
            browser.close()
            return

        month = doc["month"]
        row = doc["row"]
        print(f"BMBB_020 발견: row={row} month={month}")

        if not open_doc(nxa, row):
            print("문서 열기 실패")
            browser.close()
            return

        print("출력(fn_command_print) → PDF 다운로드...")
        dl_path = print_and_download_pdf(nxa, ctx, DOWNLOAD_TEMP)
        if not dl_path:
            print("PDF 다운로드 실패")
            browser.close()
            return

        matched = [f for f in FIRMS if f["sym"] == firm_sym]
        if matched:
            name = matched[0]["name"]
        else:
            name = f"사업장_{firm_sym}"

        dest, size = save_pdf(dl_path, name, month)
        print(f"저장: {os.path.basename(dest)} ({size:,} bytes)")

        browser.close()
        print("완료")


if __name__ == "__main__":
    main()
