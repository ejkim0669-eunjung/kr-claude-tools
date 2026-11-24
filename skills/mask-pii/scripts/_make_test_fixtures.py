"""Creates synthetic (fake) test files to sanity-check mask_pii.py.
These are NOT real user data -- purely fabricated for a smoke test."""
import pymupdf
import openpyxl
import docx
import csv
from pathlib import Path

OUT = Path(__file__).parent / "_test_fixtures"
OUT.mkdir(exist_ok=True)

KOREAN_FONT = r"C:\Windows\Fonts\malgun.ttf"

# ---- fake text PDF (form-style, like the tax support document) ----
doc = pymupdf.open()
page = doc.new_page()
page.insert_font(fontname="F0", fontfile=KOREAN_FONT)
lines = [
    "해외주식 양도소득 과세보조자료 (샘플/가상)",
    "(1)성명 홍길동",
    "(2)주민등록번호 830101-1234567",
    "(3)주소 서울특별시 강남구 테헤란로 123",
    "(4)금융기관명 미래에셋증권",
    "(5)사업자등록번호 123-45-67890",
    "(6)계좌번호 12345678901",
    "(7)관리번호 MGT-2026-001",
    "(8)일련번호 SN-000123",
    "이메일 hong@example.com",
    "여권번호 M12345678",
    "",
    "[거래내역]",
    "종목명: Apple Inc.  수량: 10  금액: 1,500,000원  일자: 2026-03-05  환율: 1350.5",
    "종목명: Tesla Inc.  수량: 5   금액: 900,000원  일자: 2026-05-12  환율: 1330.2",
]
y = 60
for line in lines:
    page.insert_text((50, y), line, fontsize=11, fontname="F0")
    y += 20
doc.save(str(OUT / "sample_text.pdf"))
doc.close()

# ---- rasterized (scanned-like) version of the same PDF ----
src = pymupdf.open(str(OUT / "sample_text.pdf"))
out = pymupdf.open()
for p in src:
    pix = p.get_pixmap(dpi=150)
    img_page = out.new_page(width=p.rect.width, height=p.rect.height)
    img_page.insert_image(p.rect, pixmap=pix)
out.save(str(OUT / "sample_scanned.pdf"))
out.close()
src.close()

# ---- fake DOCX ----
d = docx.Document()
d.core_properties.author = "김작성"
d.core_properties.company = "가상회사"
d.add_paragraph("해외주식 양도소득 과세보조자료 (샘플/가상)")
d.add_paragraph("(1)성명 이영희")
d.add_paragraph("(2)주민등록번호 900202-2345678")
d.add_paragraph("(3)주소 부산광역시 해운대구 센텀로 45")
d.add_paragraph("이메일 lee@example.com")
d.add_paragraph("종목명: Microsoft  수량: 20  금액: 3,000,000원  일자: 2026-02-01  환율: 1345.0")
d.sections[0].header.paragraphs[0].text = "관리번호 MGT-HEADER-999"
d.save(str(OUT / "sample.docx"))

# ---- fake XLSX ----
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "거래별명세"
ws.append(["성명", "주민등록번호", "계좌번호", "종목명", "수량", "금액", "일자"])
ws.append(["홍길동", "830101-1234567", "12345678901", "Apple Inc.", 10, 1500000, "2026-03-05"])
ws.append(["홍길동", "830101-1234567", "12345678901", "Tesla Inc.", 5, 900000, "2026-05-12"])
ws.append(["이영희", "900202-2345678", "98765432101", "Apple Inc.", 3, 450000, "2026-04-01"])
ws2 = wb.create_sheet("비고")
ws2.append(["메모"])
ws2.append(["담당자 연락처: 010-1234-5678, 김담당"])
wb.save(str(OUT / "sample.xlsx"))

# ---- fake CSV ----
with open(OUT / "sample.csv", "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f)
    w.writerow(["성명", "이메일", "종목명", "금액"])
    w.writerow(["홍길동", "hong@example.com", "Apple Inc.", 1500000])
    w.writerow(["이영희", "lee@example.com", "Tesla Inc.", 900000])
    w.writerow(["홍길동", "hong@example.com", "Tesla Inc.", 300000])

print("fixtures created in", OUT)
