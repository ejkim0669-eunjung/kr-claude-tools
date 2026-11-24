---
name: pii-masking
description: Create a masked (personal-data-redacted) copy of a PDF, Word (.docx), Excel (.xlsx), or CSV file before doing anything else with it. Explicit triggers -- both full requests and short ones -- include "개인정보 마스킹 해줘", "이 파일 가려줘", "PII를 가려서 작업해줘", "개인정보 가려줘", "개인정보 안 보이게 해줘", "개인정보 읽지 말고 작업해줘", "마스킹", "마스킹본", "주민번호 가려", "이름 가려줘", "비식별", "비식별화", "익명화", "가명처리", "민감정보 가려줘". Also self-trigger (without being asked) before working on an uploaded file whose file name contains a person's name or a company/business name, or the words 고객/거래처/관여처/수임/명부/인사/급여/회원/환자/직원 -- regardless of file type (PDF/DOCX/XLSX/CSV/HWP). Covers 성명(name), 주민등록번호(resident registration number), 주소(address), 금융기관명(financial institution), 사업자등록번호(business registration number), 계좌번호(account number), 관리번호(management number), 일련번호(serial number), 이메일(email), 여권번호(passport number).
---

# PII masking (first-pass tool -- not a certified anonymizer)

## When this triggers

The user says something like "개인정보 마스킹 해줘", "이 파일 개인정보 가려줘",
"가린 사본으로 작업해줘", or attaches a document containing personal data and
asks you to work with it. The expected workflow every time:

1. Make a masked copy first.
2. Do the requested downstream work (reading, analysis, spreadsheet building,
   summarizing, etc.) using the **masked copy**, not the original.
3. Always surface the verification report and the human-review checklist to
   the user -- never silently swallow them.

## Self-trigger without being asked

Run this skill proactively -- before doing anything else with the file --
whenever an uploaded/referenced file matches either signal below. File
*type* does not matter (PDF/DOCX/XLSX/CSV/HWP all count); judge by name and
content, not extension.

**Strong signal -- mask immediately, do not ask first:**
- The file name contains a person's name, or a company/business name, OR
- The file name or a quick look at its content/headers contains any of:
  고객, 거래처, 관여처, 수임, 명부, 인사, 급여, 회원, 환자, 직원
- The content contains an RRN-shaped value (`\d{6}-\d{7}`) or similar
  unambiguous personal-data pattern

Any one of these is enough on its own -- don't wait for multiple signals,
and don't ask for confirmation when one of these is present. Proceed
straight to masking.

**Weak/ambiguous signal -- ask once, briefly, before masking:**
The file plausibly contains personal data but none of the strong signals
above are present (e.g. a generic filename, mixed business content where
you can't tell from a glance). In this case ask a short one-line question
("이 파일에 개인정보가 있을 것 같은데, 먼저 마스킹본을 만들고 진행할까요?")
instead of silently guessing either way.

**Skip self-triggering (proceed with the original file) when:**
- The file is public/official material with no personal data by nature:
  공시자료, 법령, 판례, 통계, 보도자료, 약관 등
- The user is actively composing/drafting the content themselves in the
  conversation (this isn't a third party's file being processed)
- The file is already a masked copy -- recognize this from a `_masked`
  filename suffix, or from the tool's own markers already present in the
  content (`[마스킹됨]`, `[MASKED]`, `고객001`-style pseudonyms, etc.).
  Do not re-run masking on an already-masked file.

**When you do self-trigger, always say so in one line before acting**, e.g.:
> "개인정보가 있을 수 있어 먼저 마스킹본을 만들고 그것으로 작업하겠습니다."

Never mask silently -- the user should always see that this happened and
why, even when they didn't ask for it. After that one line, proceed with
the Step-by-step flow below and report results the same way as an
explicitly-requested run.

## Never copy real PII values into your own output

While reading originals, running the script, or reporting results: never
transcribe an actual person's name, resident registration number, or
address into your chat reply, a summary, or any file you generate --
including inside code you write, commit messages, or filenames. Refer to
people by their masked pseudonym (`고객001`, etc.), by row/field label, or
generically ("해당 고객"). This applies even when quoting the script's own
console output -- if you need to show an example, use the pseudonym or a
placeholder, not the value you just saw in the original file.

## What this is, in one sentence

`scripts/mask_pii.py` is a **first-pass, best-effort masking tool**. It is
not an automatic, certified anonymizer, and it does not replace a human
final review. Every time you run it, relay its printed safety notice and
human-review checklist to the user verbatim (or a faithful translation) --
do not summarize them away, and do not tell the user the file is "safe" or
"done" based on the script's output alone.

## Step-by-step

1. **Identify the input file(s)** the user means (path they gave, or the
   most recently attached/relevant file in the working directory).

2. **Never touch the original.** Do not open it in write mode, do not
   overwrite it, do not delete it. The script itself also enforces this: it
   copies the original into `<out-dir>/originals_untouched/` before doing
   any work, and prints a SHA-256 hash you can point to as proof the
   original is unchanged.

3. **Run the script:**
   ```
   py -3 "<this skill's folder>/scripts/mask_pii.py" "<input file>" [more files...] --out-dir "<output folder>"
   ```
   - If `--out-dir` is omitted, output goes to a `masked_output/` folder
     next to the first input file.
   - You can pass multiple files in one call (PDF/DOCX/XLSX/CSV mixed is
     fine).
   - `.hwp` files are refused with an explanation (see Limitations) --
     don't attempt to work around this yourself; tell the user why.

4. **For PDF specifically**, the script first checks whether the file has
   extractable text at all. If it doesn't (a scanned/photographed page),
   the script does NOT attempt to mask it and prints a loud warning
   instead. Read that warning to the user rather than treating "0 leftover
   PII" as success -- for a scanned file, 0 means "this tool could not see
   any text," not "this file is clean." Point the user to OCR + manual
   review for those files; do not try to invent an OCR pipeline on the fly
   unless the user separately asks for that.

5. **Read the script's own report back to the user**, including:
   - how many fields/cells were masked
   - the leftover-pattern verification count (and what it did/didn't find)
   - the fixed safety notice ("0 is not the same as safe")
   - the fixed human-review checklist (unlabeled prose, scanned docs,
     filenames, headers/footers/text boxes/comments, hidden sheets/columns,
     stamps/signatures/images, file metadata)

6. **Use the masked copy** (`<name>_masked.pdf/docx/xlsx/csv`, and for PDF
   also `<name>_masked.txt` for programmatic reading of numbers/tables) for
   whatever the user actually asked you to do next. Do not re-open the
   original for that downstream task.

## What masking actually does, by format

- **PDF / DOCX** (documents meant to be read, not aggregated): each
  recognized field's *value* is replaced with a literal marker
  (`[MASKED]` in the PDF's visual redaction stamp, `[마스킹됨]` in DOCX
  text) so a human can see something was removed. The **label itself is
  always preserved** ("(1)성명" stays; the name after it does not).
  Redaction on PDF is a *true* redaction via PyMuPDF's
  `add_redact_annot`/`apply_redactions` -- it strips the underlying text
  object, not just a black box drawn on top, so copy-paste or text search
  cannot recover the original value. Document metadata (author/company) is
  cleared automatically.

- **XLSX / CSV** (tabular data meant to be aggregated/summed): PII columns
  are detected by header name (성명/이름, 주민등록번호, 이메일, 여권번호,
  주소, 금융기관명, 사업자등록번호, 계좌번호, 관리번호, 일련번호--
  Korean and common English variants). Values are replaced with a
  **consistent pseudonym per real-world value**, not a blanket "[MASKED]"
  for every row -- e.g. every row belonging to the same person gets the
  same `고객001`, and their 주민등록번호/주소 cells get the matching
  `주민P001`/`주소P001` so a `GROUP BY`/pivot/sum by person still works
  correctly after masking. Non-PII columns identified by header (종목명,
  수량, 금액, 단가, 일자, 환율, 세율, 세액, 취득가, 양도가 등) are never
  touched. The script also does a best-effort safety-net scan of every
  *other* cell for a bare RRN/business-no/email/passport-shaped value that
  landed in an unlabeled column (e.g. a "비고" notes column) -- if found,
  it's masked and flagged for you to relay to the user, not silently fixed.
  Hidden sheets/rows/columns are still masked (not skipped), and their
  existence is reported so a reviewer knows to check them.

- **The pseudonym mapping is never written to disk.** It only exists in
  memory for the duration of one run, so there's no new file that maps
  fake IDs back to real people.

## Verification (built into the script -- do not re-implement your own)

After masking, the script re-scans the masked output for: RRN shape,
business-registration-number shape, email shape, and an approximate
Korean-passport shape (1 letter + 8 digits). It prints how many matches
survived. **Always state this count to the user, and always pair it with
the fixed caveat**: a count of 0 means "no pattern this script knows about
survived," not "this file contains no personal data." There is no reliable
regex for a bare Korean name, a free-text address, or an account number
without a label next to it -- those depend entirely on the label-anchored
pass and the human-review checklist.

## Human-review checklist (always show this -- the script prints it, relay it)

The script prints these verbatim every run; make sure the user actually
sees them, don't compress them into "looks good":

1. Unlabeled personal data sitting inside a sentence (names, addresses,
   company names with no field label next to them)
2. Scanned/photographed documents (extractable text = 0 → this tool did
   nothing, not "found nothing")
3. Person names inside file names or folder names
4. DOCX headers/footers/text boxes/comments -- **text boxes specifically
   are invisible to python-docx**, the script cannot see or mask them at
   all
5. Excel hidden columns/sheets/notes columns -- scanned, but a free-text
   note without a recognizable pattern can still slip through
6. Personal data baked into an image (a stamp, a signature, a photo) --
   there is no text to find
7. File metadata (author/company) -- the script clears it automatically,
   but confirm it actually did

## Limitations to be upfront about

- `.hwp` (한글) binary files are not auto-masked. The format mixes text,
  tables, and floating objects in ways that make blind regex-and-redact
  unsafe (high risk of both missed PII and corrupted files). Tell the user
  to either save/export as .docx or PDF first, or review it by hand.
- Bank/brokerage account numbers have no universal shape, so they can only
  be caught when a label ("계좌번호") sits next to them. A bare account
  number in running prose will not be found.
- The Korean passport pattern (1 letter + 8 digits) and the financial
  institution name list are both approximations/samples, not authoritative
  -- they can miss real values and can also be wrong occasionally.
- The verification counter can only count things that have a predictable
  shape. It is a floor, not a ceiling, on how much PII might remain.

## Handling originals -- non-negotiable

- Never modify, move, or delete the user's original file.
- All work happens on a copy; the script enforces this by copying to
  `originals_untouched/` and hashing before touching anything.
- If the user asks you to "just edit the original" for this kind of task,
  push back and explain why a separate masked copy is the safer default --
  do this even if they seem impatient, since the whole point of the tool is
  to avoid ever accidentally shipping the real file.

## Setup (first run on a new machine)

The script needs `pymupdf`, `openpyxl`, and `python-docx`:
```
py -3 -m pip install pymupdf openpyxl python-docx
```
`scripts/_make_test_fixtures.py` generates small synthetic (fake) sample
files across all four formats and is useful for a smoke test after editing
the script -- it is not part of normal masking usage.
