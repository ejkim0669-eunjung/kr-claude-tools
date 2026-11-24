#!/usr/bin/env python3
"""
mask_pii.py - PII masking for PDF / DOCX / XLSX / CSV (first-pass tool)

IMPORTANT: this is a first-pass masking tool, not a certified anonymizer.
A verification count of 0 means "no known pattern survived a regex scan of
extractable text" -- it does NOT mean the file is safe to share. See the
human-review checklist this script prints, and SKILL.md.

Usage:
    py -3 mask_pii.py <input_file> [<input_file2> ...] [--out-dir DIR]

For every input file this script:
  1. Never modifies the original file. It copies the original into
     <out-dir>/originals_untouched/ and does all work on a separate copy.
  2. Produces a masked copy (<name>_masked.<ext>) in <out-dir>/.
  3. For PDF, also produces <name>_masked.txt (extracted text of the
     masked file, so numeric/tabular data can be read back programmatically
     without re-reading the image-like PDF).
  4. Runs a leftover-PII regex scan on the masked output and prints a count.
  5. Prints the fixed human-review checklist and caveats.
"""

import sys
import os
import re
import csv
import shutil
import hashlib
import argparse
from pathlib import Path

try:
    import pymupdf  # PyMuPDF
except ImportError:
    pymupdf = None

try:
    import openpyxl
except ImportError:
    openpyxl = None

try:
    import docx  # python-docx
except ImportError:
    docx = None


# --------------------------------------------------------------------------
# Shared config: labels + standalone regex patterns
# --------------------------------------------------------------------------

# Labels this tool treats as "personal-data field labels" in free-form
# documents (PDF / DOCX). The LABEL ITSELF IS NEVER REMOVED -- only the
# value that follows it. This list mirrors the fields requested for the
# original masking task, plus email and passport number.
LABELS = [
    "성명", "성 명",
    "주민등록번호", "주민등록번호(외국인등록번호)", "주민번호",
    "주소",
    "금융기관명", "금융기관",
    "사업자등록번호", "사업자번호",
    "계좌번호",
    "관리번호",
    "일련번호",
    "이메일", "전자우편", "전자우편주소", "e-mail", "E-mail", "email",
    "여권번호",
]

MASK_MARK = "[마스킹됨]"

# Standalone patterns: usable even without a nearby label, and used by the
# verification scan to count leftover PII in the MASKED output.
# NOTE on coverage gaps -- read before trusting a "0":
#   - RRN pattern matches the standard 6-1-6 digit resident registration
#     number shape. It will NOT catch a number that OCR misread, or one
#     typed with extra spaces/line breaks.
#   - Passport pattern is an approximation (1 letter + 8 digits), which is
#     the common ROK passport shape, but not a validated checksum -- it can
#     both miss real numbers and flag coincidental false positives.
#   - There is no reliable standalone pattern for bank account numbers
#     (format varies by bank/broker) or for Korean personal names in free
#     text. Both rely on the LABEL-anchored pass only, which means a name
#     or account number written WITHOUT a label next to it (e.g. inside a
#     sentence, a memo column, a signature block) will not be found by this
#     script. This is exactly why the human-review checklist below exists.
STANDALONE_PATTERNS = {
    "주민등록번호": re.compile(r"\d{6}\s?-?\s?[1-8]\d{6}"),
    "사업자등록번호": re.compile(r"\d{3}\s?-\s?\d{2}\s?-\s?\d{5}"),
    "이메일": re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"),
    "여권번호(추정)": re.compile(r"\b[MSRDmsrd]\d{8}\b"),
}

# Common Korean financial institution name fragments, used only to help the
# PDF/DOCX pass catch a bank/broker name that appears WITHOUT the label
# "금융기관명" next to it (e.g. in a free sentence). This list is NOT
# exhaustive -- treat it as a bonus catch, not a guarantee.
FINANCIAL_INSTITUTIONS = [
    "미래에셋증권", "한국투자증권", "삼성증권", "신한투자증권", "NH투자증권",
    "키움증권", "KB증권", "하나증권", "대신증권", "유안타증권", "교보증권",
    "메리츠증권", "IBK투자증권", "SK증권", "현대차증권", "토스증권",
    "신한은행", "국민은행", "우리은행", "하나은행", "농협은행", "기업은행",
    "SC제일은행", "카카오뱅크", "케이뱅크", "토스뱅크", "새마을금고", "신협",
]

# Spreadsheet header keywords -> PII field category.
# Order matters: first match wins, so put more specific keywords first.
XLSX_HEADER_MAP = [
    (re.compile(r"이름|성명|고객명|name", re.I), "성명"),
    (re.compile(r"주민등록번호|주민번호|생년월일.*번호|resident", re.I), "주민등록번호"),
    (re.compile(r"이메일|전자우편|e-?mail", re.I), "이메일"),
    (re.compile(r"여권번호|passport", re.I), "여권번호"),
    (re.compile(r"주소|address", re.I), "주소"),
    (re.compile(r"금융기관|증권사|은행명|broker", re.I), "금융기관명"),
    (re.compile(r"사업자등록번호|사업자번호", re.I), "사업자등록번호"),
    (re.compile(r"계좌번호|계좌", re.I), "계좌번호"),
    (re.compile(r"관리번호", re.I), "관리번호"),
    (re.compile(r"일련번호", re.I), "일련번호"),
]

# Categories whose values are tied to the same "person" pseudonym so a
# human/spreadsheet can still group/aggregate by person after masking.
PERSON_LINKED_CATEGORIES = {"성명", "주민등록번호", "주소"}
# Categories pseudonymized with their OWN consistent per-value mapping,
# but not tied to the person index (e.g. grouping by account or by
# institution should still work, independent of who the account belongs to).
VALUE_LINKED_CATEGORIES = {"금융기관명", "사업자등록번호", "계좌번호", "관리번호", "일련번호"}

TRANSACTION_COLUMN_HINTS = re.compile(
    r"종목|수량|금액|단가|일자|날짜|환율|가액|취득가|양도가|세율|세액|차익"
)

HUMAN_REVIEW_CHECKLIST = """\
[반드시 사람이 최종 확인해야 하는 항목 -- 이 스크립트는 아래를 확인하지 못합니다]
  1. 라벨 없이 문장 속에 흩어져 있는 이름ㆍ주소ㆍ상호 (예: "김철수 고객님께서...")
  2. 스캔ㆍ촬영 이미지 문서 (글자 추출이 0이면 이 스크립트는 애초에 아무것도 찾지 못합니다)
  3. 파일 이름ㆍ폴더 이름에 들어간 사람 이름
  4. 워드 문서의 머리말ㆍ꼬리말ㆍ텍스트 상자ㆍ메모(주석) -- 텍스트 상자는 python-docx가
     아예 읽지 못하는 영역이라 이 스크립트가 손대지 못합니다
  5. 엑셀의 숨긴 열ㆍ숨긴 시트ㆍ비고/메모 칸 (이 스크립트는 숨긴 열/시트도 값 자체는
     스캔하지만, 헤더명이 PII로 인식되지 않는 자유서술 칸의 개인정보는 놓칠 수 있습니다)
  6. 도장ㆍ서명ㆍ사진처럼 이미지 안에 박혀 있어 텍스트로 존재하지 않는 개인정보
  7. 파일 속성(작성자ㆍ회사ㆍ마지막 저장자 등 메타데이터) -- 이 스크립트가 자동으로
     지우지만, 지워졌는지 최종 확인은 사람이 해야 합니다
"""

SAFETY_NOTICE = """\
[중요] 이 스크립트는 "1차 처리 도구"입니다. 자동 완전 마스킹 도구가 아닙니다.
검증 결과가 0건이라는 것은 "이 스크립트가 아는 패턴으로는 못 찾았다"는 뜻일 뿐,
"안전하다"는 뜻이 아닙니다. 반드시 위 [사람 최종 확인] 체크리스트를 사람이 눈으로
확인한 뒤에 문서를 배포/공유하세요.
"""

SCANNED_DOC_WARNING = """\
[경고] 이 파일에서 추출된 글자 수가 사실상 0입니다. 스캔본/촬영본처럼 글자 층이
없는 이미지 PDF로 보입니다. 이 스크립트는 "글자"를 찾아서 가리는 방식이므로,
글자가 없으면 가릴 대상 자체가 없어 아무 것도 하지 못한 채 "검증 0건"이 나옵니다.
이 0건은 안전하다는 뜻이 아니라 "이 도구가 이 파일에서 아예 작동하지 않았다"는
뜻입니다. 이런 파일은 OCR로 글자층을 만든 뒤 사람이 직접 검토하거나, 이미지
편집 도구로 개인정보 영역을 직접 지워야 합니다. 이 스크립트로는 처리하지 마세요.
"""


class PersonPseudonymizer:
    """Consistent value -> pseudonym mapping, kept in memory only (never
    written to disk), so the same real person/value always gets the same
    pseudonym within one run -- which is what keeps per-person aggregation
    (sums, group-by) working after masking."""

    def __init__(self):
        self._map = {}
        self._counters = {}

    def get(self, category, raw_value):
        raw_value = (raw_value or "").strip()
        if not raw_value:
            return raw_value
        key = (category, raw_value)
        if key not in self._map:
            n = self._counters.get(category, 0) + 1
            self._counters[category] = n
            self._map[key] = n
        idx = self._map[key]
        return idx

    def person_key_for_row(self, row_values_by_category):
        """Pick the most reliable identifier available in a row to key the
        shared person-pseudonym on: RRN > name > address."""
        for cat in ("주민등록번호", "성명", "주소"):
            v = (row_values_by_category.get(cat) or "").strip()
            if v:
                return v
        return None


def sha256_of(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def ensure_out_dirs(out_dir):
    out_dir = Path(out_dir)
    (out_dir / "originals_untouched").mkdir(parents=True, exist_ok=True)
    return out_dir


def protect_original(input_path, out_dir):
    """Copy the original into originals_untouched/ before touching anything,
    and return that copy's path plus its hash (recorded so we can prove the
    real original was never modified)."""
    input_path = Path(input_path)
    dest = out_dir / "originals_untouched" / input_path.name
    if not dest.exists():
        shutil.copy2(input_path, dest)
    return dest, sha256_of(input_path)


# --------------------------------------------------------------------------
# PDF
# --------------------------------------------------------------------------

def pdf_extractable_char_count(doc):
    total = 0
    for page in doc:
        total += len(page.get_text("text").strip())
    return total


def find_label_values(full_text):
    """Very simple label -> value line parser for form-style PDFs/DOCX text:
    for each line containing a known label, take the text on the same line
    after the label (or, if empty, the very next non-empty line) as the
    value to redact. Designed for '(1)성명 홍길동' or '성명: 홍길동' style
    rows, not for prose."""
    lines = full_text.splitlines()
    targets = []
    sorted_labels = sorted(LABELS, key=len, reverse=True)
    for i, line in enumerate(lines):
        for label in sorted_labels:
            idx = line.find(label)
            if idx == -1:
                continue
            after = line[idx + len(label):]
            # Strip only separator punctuation/whitespace that can sit
            # between a label and its value ("성명: 홍길동", "성명 (홍길동)").
            # Digits must NOT be in this class: an all-numeric value (e.g.
            # an account number) would otherwise be stripped away entirely,
            # leaving nothing to redact -- a real leak, not a display bug.
            after = re.sub(r"^[\s:：\(\)\.]*", "", after).strip()
            if after:
                # Take the WHOLE rest of the line as the value, not just
                # the first word -- a multi-word value (e.g. a street
                # address) must not be truncated to its first token, or
                # everything after that token leaks unmasked.
                value = after.strip(" \t:：()[]")
                if value:
                    targets.append((label, value))
            elif (i + 1 < len(lines) and lines[i + 1].strip()
                  and not any(lbl in lines[i + 1] for lbl in sorted_labels)):
                # Label and value are on separate lines -- but only treat
                # the next line as this field's value if that next line
                # isn't itself another labeled field (otherwise a genuinely
                # empty field would swallow and mis-redact its neighbor).
                value = lines[i + 1].strip()
                targets.append((label, value))
            break  # longest matching label wins; don't also match a
            # shorter label that happens to be a substring of it
            # (e.g. "금융기관" inside "금융기관명") on the same line
    return targets


def mask_pdf(src_copy_path, out_path):
    if pymupdf is None:
        raise RuntimeError("pymupdf(PyMuPDF)가 설치되어 있지 않습니다. "
                            "'py -3 -m pip install pymupdf' 로 설치하세요.")

    doc = pymupdf.open(str(src_copy_path))
    full_text = "\n".join(page.get_text("text") for page in doc)
    extractable_chars = len(full_text.strip())

    report = {"scanned_warning": False, "targets_found": 0, "pages": len(doc)}

    if extractable_chars < 20 * max(len(doc), 1):
        report["scanned_warning"] = True
        doc.close()
        return report, extractable_chars

    redact_strings = set()

    for label, value in find_label_values(full_text):
        if len(value) >= 2:
            redact_strings.add(value)

    for pattern in STANDALONE_PATTERNS.values():
        for m in pattern.finditer(full_text):
            redact_strings.add(m.group(0))

    for inst in FINANCIAL_INSTITUTIONS:
        if inst in full_text:
            redact_strings.add(inst)

    n_redactions = 0
    for page in doc:
        for s in redact_strings:
            s_clean = s.strip()
            if not s_clean:
                continue
            try:
                rects = page.search_for(s_clean)
            except Exception:
                rects = []
            for r in rects:
                # Stamp text must be ASCII: the default annotation font
                # (Helvetica) cannot render Hangul, and an unsupported
                # glyph would silently draw as blank/garbled -- which would
                # leave the box looking redacted while actually failing.
                # "[MASKED]" is guaranteed to render regardless of the
                # document's embedded fonts.
                page.add_redact_annot(r, fill=(0, 0, 0), text="[MASKED]", fontsize=6)
                n_redactions += 1
        # true redaction: strips underlying text/image content, not just a
        # drawn box, so copy/search cannot recover the original value
        page.apply_redactions()

    doc.set_metadata({})  # clear author/company/etc. metadata
    doc.save(str(out_path), garbage=4, deflate=True, clean=True)
    doc.close()

    report["targets_found"] = n_redactions
    return report, extractable_chars


def verify_pdf(masked_path):
    doc = pymupdf.open(str(masked_path))
    text = "\n".join(page.get_text("text") for page in doc)
    doc.close()
    # check_label_tails=False: PyMuPDF's true-redaction stamps ("[MASKED]")
    # are inserted as new content and do NOT come back in the same reading
    # position as the label during text extraction (they tend to trail at
    # the end of the page's extracted text instead of sitting right after
    # the label). That is a property of how the redaction was drawn, not a
    # sign the value is still there -- the label-adjacency heuristic would
    # misfire on every label in a masked PDF, so it is only meaningful for
    # DOCX, where python-docx keeps everything in one inline run.
    return verify_text_blob(text, check_label_tails=False)


# --------------------------------------------------------------------------
# DOCX
# --------------------------------------------------------------------------

_LABEL_ALT = "|".join(re.escape(l) for l in sorted(LABELS, key=len, reverse=True))
# Matches "<label><separator><value>" as one unit, e.g. "성명: 홍길동" or
# "(2)주민등록번호 900202-2345678". Using ONE combined, longest-first
# alternation (instead of looping over LABELS and mutating the string) means
# a shorter label that happens to be a prefix of a longer one (e.g.
# "금융기관" inside "금융기관명") can never match a second time over the
# text a longer label + value pair already consumed -- re.sub scans
# left-to-right and does not revisit consumed spans.
_LABEL_VALUE_RE = re.compile(
    # value = everything up to the next known label or end of the
    # paragraph text -- NOT just the first whitespace-delimited token, or
    # a multi-word value like a street address would only be half-masked.
    rf"(?P<label>{_LABEL_ALT})(?P<sep>[\s:：\(\)\.]*)(?P<value>.+?)(?=\s*(?:{_LABEL_ALT})|$)"
)


def _mask_paragraph_text(text):
    def _replace(m):
        if MASK_MARK in m.group("value"):
            return m.group(0)
        return m.group("label") + m.group("sep") + MASK_MARK

    new_text, n = _LABEL_VALUE_RE.subn(_replace, text)
    changed = n > 0
    text = new_text
    for pattern in STANDALONE_PATTERNS.values():
        new_text, n = pattern.subn(MASK_MARK, text)
        if n:
            text = new_text
            changed = True
    for inst in FINANCIAL_INSTITUTIONS:
        if inst in text:
            text = text.replace(inst, MASK_MARK)
            changed = True
    return text, changed


def _mask_paragraphs(paragraphs):
    n = 0
    for p in paragraphs:
        if not p.text.strip():
            continue
        new_text, changed = _mask_paragraph_text(p.text)
        if changed:
            # python-docx has no clean "replace text, keep formatting"
            # primitive across runs, so we collapse to a single run. This
            # is a known limitation: complex per-run formatting on a masked
            # line will be simplified.
            for run in p.runs:
                run.text = ""
            if p.runs:
                p.runs[0].text = new_text
            else:
                p.add_run(new_text)
            n += 1
    return n


def mask_docx(src_copy_path, out_path):
    if docx is None:
        raise RuntimeError("python-docx가 설치되어 있지 않습니다. "
                            "'py -3 -m pip install python-docx' 로 설치하세요.")

    d = docx.Document(str(src_copy_path))
    n = 0
    n += _mask_paragraphs(d.paragraphs)

    for table in d.tables:
        for row in table.rows:
            for cell in row.cells:
                n += _mask_paragraphs(cell.paragraphs)

    for section in d.sections:
        n += _mask_paragraphs(section.header.paragraphs)
        n += _mask_paragraphs(section.footer.paragraphs)

    # clear author/company metadata
    cp = d.core_properties
    cp.author = ""
    cp.last_modified_by = ""
    try:
        cp.company = ""
    except Exception:
        pass

    d.save(str(out_path))
    return {
        "targets_found": n,
        "text_box_warning": True,  # python-docx cannot see text boxes at all
    }


def verify_docx(masked_path):
    d = docx.Document(str(masked_path))
    chunks = [p.text for p in d.paragraphs]
    for table in d.tables:
        for row in table.rows:
            for cell in row.cells:
                chunks.extend(p.text for p in cell.paragraphs)
    for section in d.sections:
        chunks.extend(p.text for p in section.header.paragraphs)
        chunks.extend(p.text for p in section.footer.paragraphs)
    return verify_text_blob("\n".join(chunks))


# --------------------------------------------------------------------------
# XLSX
# --------------------------------------------------------------------------

def classify_header(header_text):
    if not header_text:
        return None
    for pattern, category in XLSX_HEADER_MAP:
        if pattern.search(str(header_text)):
            return category
    return None


def mask_xlsx(src_copy_path, out_path):
    if openpyxl is None:
        raise RuntimeError("openpyxl이 설치되어 있지 않습니다. "
                            "'py -3 -m pip install openpyxl' 로 설치하세요.")

    wb = openpyxl.load_workbook(str(src_copy_path))
    pseudo = PersonPseudonymizer()
    hidden_notes = []
    n_masked_cells = 0

    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            hidden_notes.append(f"시트 '{ws.title}' (상태: {ws.sheet_state})")

        if ws.max_row < 1:
            continue

        header_row_idx = 1
        headers = {}
        for col_idx in range(1, ws.max_column + 1):
            header_val = ws.cell(row=header_row_idx, column=col_idx).value
            category = classify_header(header_val)
            if category:
                headers[col_idx] = category
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            dim = ws.column_dimensions.get(col_letter)
            if dim is not None and dim.hidden:
                hidden_notes.append(f"시트 '{ws.title}' 숨긴 열 {col_letter} "
                                     f"(헤더: {header_val!r})")

        if not headers:
            continue

        for row_idx in range(header_row_idx + 1, ws.max_row + 1):
            row_dim = ws.row_dimensions.get(row_idx)
            if row_dim is not None and row_dim.hidden:
                hidden_notes.append(f"시트 '{ws.title}' 숨긴 행 {row_idx}")

            row_values_by_category = {}
            for col_idx, category in headers.items():
                if category in PERSON_LINKED_CATEGORIES:
                    v = ws.cell(row=row_idx, column=col_idx).value
                    row_values_by_category[category] = "" if v is None else str(v)
            person_key = pseudo.person_key_for_row(row_values_by_category)

            for col_idx, category in headers.items():
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value in (None, ""):
                    continue
                raw = str(cell.value)

                if category in PERSON_LINKED_CATEGORIES:
                    key_value = person_key if person_key else raw
                    idx = pseudo.get("PERSON", key_value)
                    if category == "성명":
                        cell.value = f"고객{idx:03d}"
                    elif category == "주민등록번호":
                        cell.value = f"주민P{idx:03d}"
                    elif category == "주소":
                        cell.value = f"주소P{idx:03d}"
                elif category in VALUE_LINKED_CATEGORIES:
                    idx = pseudo.get(category, raw)
                    short = {
                        "금융기관명": "기관", "사업자등록번호": "사업자",
                        "계좌번호": "계좌", "관리번호": "관리번호",
                        "일련번호": "일련번호",
                    }[category]
                    cell.value = f"{short}{idx:03d}"
                elif category == "이메일":
                    key_value = person_key if person_key else raw
                    idx = pseudo.get("PERSON", key_value)
                    cell.value = f"MASKED-EMAIL-{idx:03d}"
                elif category == "여권번호":
                    key_value = person_key if person_key else raw
                    idx = pseudo.get("PERSON", key_value)
                    cell.value = f"MASKED-PASSPORT-{idx:03d}"
                n_masked_cells += 1

        # Safety-net scan: any OTHER (non-PII-header) cell that still looks
        # like a RRN / business-no / email is flagged, not silently trusted.
        # Never scans columns that look like transaction data by header name.
        for col_idx in range(1, ws.max_column + 1):
            if col_idx in headers:
                continue
            header_val = str(ws.cell(row=header_row_idx, column=col_idx).value or "")
            if TRANSACTION_COLUMN_HINTS.search(header_val):
                continue
            for row_idx in range(header_row_idx + 1, ws.max_row + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value is None:
                    continue
                raw = str(cell.value)
                for pat_name, pattern in STANDALONE_PATTERNS.items():
                    if pattern.search(raw):
                        col_letter = openpyxl.utils.get_column_letter(col_idx)
                        hidden_notes.append(
                            f"시트 '{ws.title}' {col_letter}{row_idx} "
                            f"(헤더 '{header_val}')에서 라벨 없는 {pat_name} 패턴 발견 "
                            f"-> 자동 마스킹함, 사람이 재확인 필요"
                        )
                        cell.value = f"[{pat_name} 마스킹됨]"
                        n_masked_cells += 1
                        break

    try:
        wb.properties.creator = None
        wb.properties.lastModifiedBy = None
        wb.properties.company = None
    except Exception:
        pass

    wb.save(str(out_path))
    return {"targets_found": n_masked_cells, "hidden_notes": hidden_notes}


def verify_xlsx(masked_path):
    wb = openpyxl.load_workbook(str(masked_path))
    chunks = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    chunks.append(str(cell.value))
    return verify_text_blob("\n".join(chunks), check_label_tails=False)


# --------------------------------------------------------------------------
# CSV
# --------------------------------------------------------------------------

def mask_csv(src_copy_path, out_path):
    pseudo = PersonPseudonymizer()
    n_masked_cells = 0

    with open(src_copy_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        rows = list(reader)

    if not rows:
        with open(out_path, "w", newline="", encoding="utf-8") as f:
            pass
        return {"targets_found": 0}

    header = rows[0]
    headers = {}
    for col_idx, h in enumerate(header):
        category = classify_header(h)
        if category:
            headers[col_idx] = category

    out_rows = [header]
    for row in rows[1:]:
        row = list(row) + [""] * (len(header) - len(row))
        row_values_by_category = {}
        for col_idx, category in headers.items():
            if category in PERSON_LINKED_CATEGORIES:
                row_values_by_category[category] = row[col_idx] if col_idx < len(row) else ""
        person_key = pseudo.person_key_for_row(row_values_by_category)

        for col_idx, category in headers.items():
            if col_idx >= len(row) or row[col_idx] == "":
                continue
            raw = row[col_idx]
            if category in PERSON_LINKED_CATEGORIES:
                key_value = person_key if person_key else raw
                idx = pseudo.get("PERSON", key_value)
                if category == "성명":
                    row[col_idx] = f"고객{idx:03d}"
                elif category == "주민등록번호":
                    row[col_idx] = f"주민P{idx:03d}"
                elif category == "주소":
                    row[col_idx] = f"주소P{idx:03d}"
            elif category in VALUE_LINKED_CATEGORIES:
                idx = pseudo.get(category, raw)
                short = {
                    "금융기관명": "기관", "사업자등록번호": "사업자",
                    "계좌번호": "계좌", "관리번호": "관리번호",
                    "일련번호": "일련번호",
                }[category]
                row[col_idx] = f"{short}{idx:03d}"
            elif category == "이메일":
                key_value = person_key if person_key else raw
                idx = pseudo.get("PERSON", key_value)
                row[col_idx] = f"MASKED-EMAIL-{idx:03d}"
            elif category == "여권번호":
                key_value = person_key if person_key else raw
                idx = pseudo.get("PERSON", key_value)
                row[col_idx] = f"MASKED-PASSPORT-{idx:03d}"
            n_masked_cells += 1

        # safety-net scan on non-PII columns
        for col_idx, cell_val in enumerate(row):
            if col_idx in headers or not cell_val:
                continue
            h = header[col_idx] if col_idx < len(header) else ""
            if TRANSACTION_COLUMN_HINTS.search(h):
                continue
            for pat_name, pattern in STANDALONE_PATTERNS.items():
                if pattern.search(cell_val):
                    row[col_idx] = f"[{pat_name} 마스킹됨]"
                    n_masked_cells += 1
                    break

        out_rows.append(row)

    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerows(out_rows)

    return {"targets_found": n_masked_cells}


def verify_csv(masked_path):
    with open(masked_path, "r", encoding="utf-8-sig") as f:
        text = f.read()
    return verify_text_blob(text, check_label_tails=False)


# --------------------------------------------------------------------------
# Shared verification
# --------------------------------------------------------------------------

def verify_text_blob(text, check_label_tails=True):
    """Verification for continuous prose text (PDF/DOCX): runs the
    standalone regexes, AND checks whether a known label is still followed
    by something that isn't the mask marker (a sign the value wasn't
    caught). check_label_tails must be OFF for tabular dumps (xlsx/csv),
    where cells are just concatenated and a header label cell sits right
    next to the FOLLOWING header label, not its own value -- that adjacency
    is not a leak and would otherwise be misreported as one."""
    counts = {}
    for name, pattern in STANDALONE_PATTERNS.items():
        matches = pattern.findall(text)
        if matches:
            counts[name] = len(matches)
    if check_label_tails:
        label_leak = 0
        for label in LABELS:
            for m in re.finditer(re.escape(label), text):
                tail = text[m.end(): m.end() + 30]
                if MASK_MARK not in tail and not re.match(r"^[\s:：\(\)\d\.]*$", tail):
                    label_leak += 1
        if label_leak:
            counts["라벨 뒤 마스킹 안 된 값으로 추정"] = label_leak
    return counts


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def process_one(input_path, out_dir):
    input_path = Path(input_path)
    out_dir = ensure_out_dirs(out_dir)
    ext = input_path.suffix.lower()

    print(f"\n{'=' * 70}\n입력 파일: {input_path.name}\n{'=' * 70}")

    if ext == ".hwp":
        print("[미지원] .hwp(한글) 바이너리 포맷은 이 스크립트가 안전하게 자동")
        print("마스킹할 수 없습니다 (텍스트 상자ㆍ표ㆍ개체가 섞인 독자 포맷이라")
        print("오탐/누락 위험이 큽니다). 워드(.docx)나 PDF로 다른 이름으로 저장한")
        print("뒤 이 스크립트를 돌리거나, 사람이 직접 검토ㆍ마스킹하세요.")
        return

    copy_path, orig_hash = protect_original(input_path, out_dir)

    if ext == ".pdf":
        out_path = out_dir / f"{input_path.stem}_masked.pdf"
        report, extractable_chars = mask_pdf(copy_path, out_path)
        if report.get("scanned_warning"):
            print(SCANNED_DOC_WARNING)
            print(f"(추출된 글자 수: {extractable_chars}, 페이지 수: {report['pages']})")
            print(HUMAN_REVIEW_CHECKLIST)
            print(SAFETY_NOTICE)
            return
        print(f"레이블/패턴 기반으로 마스킹한 항목 수: {report['targets_found']}")
        print(f"마스킹본 PDF: {out_path}")

        txt_out = out_dir / f"{input_path.stem}_masked.txt"
        doc = pymupdf.open(str(out_path))
        with open(txt_out, "w", encoding="utf-8") as f:
            for page in doc:
                f.write(page.get_text("text"))
                f.write("\n\n")
        doc.close()
        print(f"마스킹본 텍스트: {txt_out}")

        counts = verify_pdf(out_path)

    elif ext == ".docx":
        out_path = out_dir / f"{input_path.stem}_masked.docx"
        report = mask_docx(copy_path, out_path)
        print(f"레이블/패턴 기반으로 마스킹한 문단 수: {report['targets_found']}")
        print(f"마스킹본 DOCX: {out_path}")
        print("[주의] python-docx는 텍스트 상자(도형 안 글상자)를 읽지 못합니다.")
        print("       텍스트 상자가 있는 문서는 사람이 직접 확인하세요.")
        counts = verify_docx(out_path)

    elif ext in (".xlsx", ".xlsm"):
        out_path = out_dir / f"{input_path.stem}_masked.xlsx"
        report = mask_xlsx(copy_path, out_path)
        print(f"가명 처리ㆍ마스킹한 셀 수: {report['targets_found']}")
        print(f"마스킹본 XLSX: {out_path}")
        if report["hidden_notes"]:
            print("[주의] 숨긴 시트/행/열 또는 라벨 없는 패턴이 발견되어 처리했습니다:")
            for note in report["hidden_notes"]:
                print(f"   - {note}")
        counts = verify_xlsx(out_path)

    elif ext == ".csv":
        out_path = out_dir / f"{input_path.stem}_masked.csv"
        report = mask_csv(copy_path, out_path)
        print(f"가명 처리ㆍ마스킹한 셀 수: {report['targets_found']}")
        print(f"마스킹본 CSV: {out_path}")
        counts = verify_csv(out_path)

    else:
        print(f"[미지원 확장자] {ext} -- 이 스크립트는 pdf/docx/xlsx/xlsm/csv만 처리합니다.")
        return

    print(f"\n원본 보호 확인: 원본 SHA-256 = {orig_hash[:16]}... "
          f"(원본 파일은 손대지 않았고, {out_dir/'originals_untouched'} 에 별도 사본이 있습니다)")

    print("\n--- 잔존 개인정보 패턴 검증 ---")
    if counts:
        total = sum(counts.values())
        print(f"검증 결과: {total}건 발견 (자동 도구 기준 미흡 -- 규칙을 보강해서 다시 만드세요)")
        for name, c in counts.items():
            print(f"   - {name}: {c}건")
    else:
        print("검증 결과: 0건 (이 스크립트가 아는 패턴 기준으로는 못 찾음)")

    print("\n" + HUMAN_REVIEW_CHECKLIST)
    print(SAFETY_NOTICE)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("inputs", nargs="+", help="마스킹할 파일 경로 (여러 개 가능)")
    parser.add_argument("--out-dir", default=None,
                         help="출력 폴더 (기본값: 첫 입력 파일과 같은 폴더의 masked_output/)")
    args = parser.parse_args()

    first = Path(args.inputs[0])
    out_dir = Path(args.out_dir) if args.out_dir else first.parent / "masked_output"

    for inp in args.inputs:
        process_one(inp, out_dir)


if __name__ == "__main__":
    main()
