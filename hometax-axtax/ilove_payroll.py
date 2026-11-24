"""
아이러브안과 급여 반영
바탕화면의 급여대장 → 양식에 매핑 → .xls(97-2003) 저장

사용법: python ilove_payroll.py
  바탕화면에서 '아이러브안과*급여*.xlsx' 와 '아이러브안과의원-*.xlsx' 를 자동 탐색
"""
import sys, os, glob, subprocess
import openpyxl
from openpyxl.utils import column_index_from_string as ci
from openpyxl.styles import Font

DESKTOP = os.path.join(os.environ["USERPROFILE"], "Desktop")

# ── 급여대장 → 양식 열 매핑 ──
# (급여대장 열번호, 이름행 기준 offset, 양식 열문자)
PAYROLL_MAP = [
    # offset 0: 이름행
    (3,  0, "F"),   # 기본급
    (4,  0, "H"),   # 식대
    (6,  0, "I"),   # 차량유지비
    (7,  0, "U"),   # 육아수당
    (8,  0, "M"),   # 직급수당
    (9,  0, "AN"),  # 지급액계
    (10, 0, "AO"),  # 국민연금
    (11, 0, "AP"),  # 건강보험
    (12, 0, "AR"),  # 장기요양보험료
    (13, 0, "AQ"),  # 고용보험
    (14, 0, "AS"),  # 소득세
    (15, 0, "BF"),  # 공제액계
    (16, 0, "BG"),  # 차인지급액
    # offset 1
    (3,  1, "Q"),   # 휴일대체근무수당
    (4,  1, "J"),   # 연장근로수당
    (5,  1, "J"),   # 연장근로수당(금액 열)
    (6,  1, "AF"),  # 성과인센티브
    (7,  1, "AD"),  # 수술팀인센티브
    (8,  1, "K"),   # 기타수당
    (10, 1, "AT"),  # 지방소득세
    (11, 1, "BD"),  # 고용보험연말정산
    (12, 1, "BA"),  # 연말정산소득세
    (13, 1, "BB"),  # 연말정산지방소득세
    (14, 1, "BE"),  # 건강보험연말정산
    # offset 2
    (3,  2, "O"),   # 소개인센티브
    (4,  2, "N"),   # 서류인센티브
    (6,  2, "AI"),  # 드림렌즈검사인센티브
    (7,  2, "AC"),  # 특별인센티브
    (8,  2, "T"),   # 휴가비
    (10, 2, "BC"),  # 장기요양보험정산
    # offset 3
    (3,  3, "AG"),  # 컬러링수당
    (4,  3, "AH"),  # 카카오주말응대수당
    (6,  3, "AL"),  # 콜지원수당
    (7,  3, "R"),   # 연차수당
    (8,  3, "AK"),  # 외래팀
    # offset 4
    (3,  4, "AE"),  # 퇴직정산수당
    (4,  4, "Y"),   # 급여정산차액
    (6,  4, "X"),   # 경조사
    (7,  4, "AM"),  # 육아휴직
]

COL_F = ci("F")
COL_BG = ci("BG")


def find_files():
    pay = sorted(glob.glob(os.path.join(DESKTOP, "아이러브안과*급여*.xlsx")),
                 key=os.path.getmtime)
    tmpl = sorted(glob.glob(os.path.join(DESKTOP, "아이러브안과의원-*.xlsx")),
                  key=os.path.getmtime)
    if not pay:
        sys.exit("급여대장 없음 (바탕화면에 '아이러브안과*급여*.xlsx')")
    if not tmpl:
        sys.exit("양식 없음 (바탕화면에 '아이러브안과의원-*.xlsx')")
    return pay[-1], tmpl[-1]


def find_name_rows(ws):
    """급여대장에서 직원 이름행 찾기: A=이름, I=지급총액(숫자)"""
    rows = []
    skip = {"성명", "사원번호", "총합계", ""}
    for r in range(1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        i = ws.cell(r, 9).value
        if (isinstance(a, str) and a not in skip
                and isinstance(i, (int, float)) and i > 0):
            rows.append(r)
    return rows


def parse_payroll(ws):
    employees = {}
    for nr in find_name_rows(ws):
        name = ws.cell(nr, 1).value
        data = {}
        for pcol, offset, tcol_letter in PAYROLL_MAP:
            v = ws.cell(nr + offset, pcol).value
            if isinstance(v, (int, float)):
                tc = ci(tcol_letter)
                data[tc] = data.get(tc, 0) + v
        employees[name] = data
    return employees


def fill_template(ws, employees):
    matched = 0
    total_row = None

    for row in range(3, ws.max_row + 1):
        if ws.cell(row, 1).value == "합계":
            total_row = row
            break
        name = ws.cell(row, 2).value
        if name and name in employees:
            for col_idx, val in employees[name].items():
                ws.cell(row, col_idx).value = val
            matched += 1

    if total_row:
        for c in range(COL_F, COL_BG + 1):
            ws.cell(total_row, c).value = sum(
                ws.cell(r, c).value or 0 for r in range(3, total_row))

    end = total_row or ws.max_row
    for r in range(3, end + 1):
        for c in range(COL_F, COL_BG + 1):
            cell = ws.cell(r, c)
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0"
            f = cell.font
            cell.font = Font(name="맑은 고딕", size=f.size or 9, bold=f.bold)

    return matched


def to_xls(xlsx_path):
    """PowerShell Excel COM으로 .xls 변환"""
    xls_path = xlsx_path.rsplit(".", 1)[0] + ".xls"
    abs_in = os.path.abspath(xlsx_path)
    abs_out = os.path.abspath(xls_path)
    ps = f'''
$e = New-Object -ComObject Excel.Application
$e.Visible = $false
$e.DisplayAlerts = $false
$wb = $e.Workbooks.Open("{abs_in}")
$wb.SaveAs("{abs_out}", 56)
$wb.Close()
$e.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($e) | Out-Null
'''
    r = subprocess.run(["powershell", "-Command", ps],
                       capture_output=True, text=True, timeout=30)
    if r.returncode == 0 and os.path.exists(xls_path):
        return xls_path
    print(f"xls 변환 실패: {r.stderr.strip()}")
    return xlsx_path


def main():
    pay_path, tmpl_path = find_files()
    print(f"급여대장: {os.path.basename(pay_path)}")
    print(f"양식:     {os.path.basename(tmpl_path)}")

    wb_pay = openpyxl.load_workbook(pay_path, data_only=True)
    employees = parse_payroll(wb_pay.active)
    print(f"파싱: {len(employees)}명")

    wb_tmpl = openpyxl.load_workbook(tmpl_path)
    matched = fill_template(wb_tmpl.active, employees)
    print(f"반영: {matched}명")

    wb_tmpl.save(tmpl_path)
    out = to_xls(tmpl_path)
    print(f"저장: {os.path.basename(out)}")
    return out


if __name__ == "__main__":
    main()
