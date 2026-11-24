# 제작: 김은정 · AI 세무실무 강의 실습 · 2026-09-18
# 개인 실습용 도구 — 재배포 금지
import hashlib
import html.parser
import json
import os
import re
import socket
import ssl
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET

sys.stdin.reconfigure(encoding="utf-8")
sys.stdout.reconfigure(encoding="utf-8", newline="\n")
sys.stderr.reconfigure(encoding="utf-8")


def log(*args):
    """Logs always go to stderr. (stdout carries protocol data only.)"""
    print(*args, file=sys.stderr, flush=True)


def read_key():
    """환경변수 LAW_OC -> law_oc.txt 순서로 찾는다. 값 자체는 절대 반환 로그에 남기지 않는다."""
    oc = (os.environ.get("LAW_OC") or "").strip()
    if oc:
        return oc
    try:
        with open(r"C:\krlaw\law_oc.txt", encoding="utf-8") as f:
            return f.read().strip()
    except OSError:
        return ""


HEADERS = {"User-Agent": "Mozilla/5.0", "Referer": "https://www.law.go.kr/"}


def http_get(url, params, timeout=20):
    full = url + "?" + urllib.parse.urlencode(params, encoding="utf-8")
    req = urllib.request.Request(full, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read().decode("utf-8", "replace")


def search_law(keyword):
    """Statute search -> human-readable list (str)."""
    raw = http_get("https://www.law.go.kr/DRF/lawSearch.do", {
        "OC": read_key(), "target": "law", "type": "XML",
        "query": keyword, "display": "100",   # wide, so the exact match lands inside the page
    })
    root = ET.fromstring(raw)
    items = []
    for child in list(root):              # a child with children = one result row
        if not list(child):
            continue
        d = {}
        for f in list(child):
            if f.text and f.text.strip():
                d[f.tag] = f.text.strip()
        if d:
            items.append(d)
    if not items:
        return ("[검색결과 없음] '%s' — 키워드를 줄여서 다시 검색하세요"
                "(공백은 AND 조건)." % keyword)

    # Fix for #1: exact-name matches go first.
    exact = [d for d in items if d.get("법령명한글", "") == keyword]
    rest = [d for d in items if d not in exact]
    ordered = exact + rest

    lines = []
    for d in ordered[:15]:                # show only the head of the list
        name = d.get("법령명한글", "?")
        lid = d.get("법령ID", "?")        # what the detail call in 05 needs
        stat = d.get("현행연혁코드", "")
        ef = d.get("시행일자", "")
        # Fix for #2: warn on non-current versions.
        mark = "" if stat == "현행" else "  ⚠️ %s — 현행 아님(과거·예정 버전)" % (stat or "상태불명")
        lines.append("%s [법령ID %s] 시행 %s%s" % (name, lid, ef, mark))
    if len(ordered) > 15:
        lines.append("… 외 %d건 — 키워드를 더 좁혀 보세요." % (len(ordered) - 15))
    return "\n".join(lines)


def find_law_id(keyword):
    """법령명 -> 법령ID (정확명 우선). 못 찾으면 None."""
    raw = http_get("https://www.law.go.kr/DRF/lawSearch.do", {
        "OC": read_key(), "target": "law", "type": "XML",
        "query": keyword, "display": "100",
    })
    root = ET.fromstring(raw)
    items = []
    for child in list(root):
        if not list(child):
            continue
        d = {}
        for f in list(child):
            if f.text and f.text.strip():
                d[f.tag] = f.text.strip()
        if d:
            items.append(d)
    exact = [d for d in items if d.get("법령명한글", "") == keyword]
    pick = (exact or items or [None])[0]
    return pick.get("법령ID") if pick else None


def jo_code(jo):
    """'제38조' / '제10조의2' / '38' -> 6자리 코드."""
    jo = str(jo).strip()
    if re.fullmatch(r"\d{6}", jo):
        return jo
    m = re.search(r"제?\s*(\d+)\s*조(?:\s*의\s*(\d+))?", jo)
    if m:
        return "%04d%02d" % (int(m.group(1)), int(m.group(2) or 0))
    return "%04d00" % int(jo) if jo.isdigit() else jo


def text_of(v):
    """조문내용/항내용/호내용 은 항상 문자열이 아니다(중첩 조문은 list/dict).
    어떤 모양이든 문서 순서대로 문자열을 복원한다."""
    if v is None:
        return ""
    if isinstance(v, str):
        return v.strip()
    if isinstance(v, list):
        return "\n".join(t for t in (text_of(x) for x in v) if t)
    if isinstance(v, dict):
        return "\n".join(t for t in (text_of(x) for x in v.values()) if t)
    return str(v)


def as_list(x):
    """API 는 항목이 하나면 dict, 여럿이면 list 로 준다 — 항상 list 로 맞춘다."""
    if x is None:
        return []
    return x if isinstance(x, list) else [x]


def _num(x):
    m = re.search(r"\d+", str(x or ""))
    return int(m.group()) if m else 0


def get_article(law_id, jo):
    """법령ID + 조문번호 -> 조문 본문.
    조문이 없으면 '[조문 없음]'(유효한 결과), 조회 자체가 실패하면 '[조회 실패]'."""
    law_id = re.sub(r"\D", "", str(law_id)) or str(law_id)
    raw = http_get("https://www.law.go.kr/DRF/lawService.do", {
        "OC": read_key(), "target": "eflaw", "type": "JSON", "ID": law_id,
    }, timeout=25)
    if (len(raw) < 8000
            and not raw.lstrip().startswith(("{", "["))
            and "미신청된" in raw):
        return ("[조회 실패] 법령ID %s — 조회 조합이 잘못됐습니다. "
                "인증키 문제가 아니니 키를 다시 입력하지 마세요." % law_id)
    try:
        data = json.loads(raw)
    except ValueError:
        return "[조회 실패] 법령ID %s — 응답이 JSON 이 아닙니다." % law_id
    law = data.get("법령")
    if not isinstance(law, dict):
        return "[조회 실패] 법령ID %s — 법령 전체본을 받지 못했습니다." % law_id

    jo_node = law.get("조문")
    units = as_list(jo_node.get("조문단위")) if isinstance(jo_node, dict) else as_list(jo_node)
    code = jo_code(jo)
    want_jo, want_ji = int(code[:4]), int(code[4:])
    sel = [u for u in units if isinstance(u, dict)
           and _num(u.get("조문번호")) == want_jo
           and _num(u.get("조문가지번호")) == want_ji]
    if not sel:
        return ("[조문 없음] %s — 이 법령(법령ID %s)에 해당 조문이 없습니다. "
                "폐지되었거나 조문번호가 틀렸을 수 있습니다." % (jo, law_id))

    out = []
    for u in sel:
        body = text_of(u.get("조문내용"))
        if body:
            out.append(body)
        for hang in as_list(u.get("항")):
            if not isinstance(hang, dict):
                continue
            t = text_of(hang.get("항내용"))
            if t:
                out.append("  " + t.replace("\n", "\n  "))
            for ho in as_list(hang.get("호")):
                if not isinstance(ho, dict):
                    continue
                t2 = text_of(ho.get("호내용"))
                if t2:
                    out.append("    " + t2.replace("\n", "\n    "))
                for mok in as_list(ho.get("목")):
                    if not isinstance(mok, dict):
                        continue
                    t3 = text_of(mok.get("목내용"))
                    if t3:
                        out.append("      " + t3.replace("\n", "\n      "))
            for mok in as_list(hang.get("목")):
                if not isinstance(mok, dict):
                    continue
                t3 = text_of(mok.get("목내용"))
                if t3:
                    out.append("      " + t3.replace("\n", "\n      "))
    return "\n".join(out)


def search_tribunal(keyword):
    """조세심판원 심판례 검색 -> human-readable list (str)."""
    raw = http_get("https://www.law.go.kr/DRF/lawSearch.do", {
        "OC": read_key(), "target": "ttSpecialDecc", "type": "XML",
        "query": keyword, "display": "10",
    })
    root = ET.fromstring(raw)
    items = []
    for child in list(root):
        if not list(child):
            continue
        d = {}
        for f in list(child):
            if f.text and f.text.strip():
                d[f.tag] = f.text.strip()
        if d:
            items.append(d)
    if not items:
        return "[검색결과 없음] '%s' — 키워드를 줄여 다시 검색하세요." % keyword

    lines = []
    for d in items:
        # Title/number field names vary by row — use whichever exists.
        title = d.get("사건명") or d.get("안건명") or d.get("재결례명") or "?"
        serial = ""
        for k, v in d.items():
            if k.endswith("일련번호"):       # *일련번호 = the ID key for the detail call
                serial = v
                break
        extra = " · ".join("%s %s" % (k, d[k])
                           for k in ("세목", "청구번호", "의결일자", "처분일자") if d.get(k))
        lines.append("[%s] %s%s" % (serial, title, ("\n   " + extra) if extra else ""))
    return "\n".join(lines)


def get_tribunal(serial):
    """조세심판원 심판례 일련번호 -> 전체 본문 (JSON, 최상위 키 이름은 target마다 달라 고정할 수 없다)."""
    raw = http_get("https://www.law.go.kr/DRF/lawService.do", {
        "OC": read_key(), "target": "ttSpecialDecc", "type": "JSON", "ID": str(serial),
    })
    data = json.loads(raw)
    # The top-level key name varies by target — grab the first dict value.
    node = next((v for v in data.values() if isinstance(v, dict)), None)
    if node is None:
        return "[본문 없음] 일련번호 %s — 응답에서 본문 노드를 찾지 못했습니다." % serial

    order = ("사건명", "청구번호", "세목", "의결일자", "처분일자",
             "주문", "재결요지", "이유")
    out = []
    shown = set()
    for k in order:                    # important fields first
        v = node.get(k)
        if isinstance(v, str) and v.strip():
            v = v.strip()
            out.append("%s:\n%s" % (k, v) if len(v) > 60 else "%s: %s" % (k, v))
            shown.add(k)
    for k, v in node.items():          # then any remaining string fields
        if k in shown or not isinstance(v, str) or not v.strip():
            continue
        v = v.strip()
        out.append("%s:\n%s" % (k, v) if len(v) > 60 else "%s: %s" % (k, v))
    return "\n\n".join(out) or "[본문 없음] 일련번호 %s" % serial


# -- 국세청 예규(세법해석) 검색은 법제처 API, 본문은 국세법령정보시스템 공개 페이지 ------
# 법제처는 목록만 준다. 본문은 공식 API가 아니므로 호출 속도를 낮추고,
# 인증서 문제가 나도 검증을 끄지 않는다 (아래 _pinned_context_if_trusted 참고).

TAXLAW_HOST = "taxlaw.nts.go.kr"
TAXLAW_TRUSTED_ROOT_SHA256 = "7BB647A62AEEAC88BF257AA522D01FFEA395E0AB45C73F93F65654EC38F25A06"
TAXLAW_MAX_BYTES = 4 * 1024 * 1024   # 응답은 2MB 안팎 — 여유를 두고 상한을 건다
_last_taxlaw_call = [0.0]


def search_tax_interpretation(keyword):
    """국세청 예규(세법해석) 목록 검색 -> human-readable list (str). 본문은 여기 없다."""
    raw = http_get("https://www.law.go.kr/DRF/lawSearch.do", {
        "OC": read_key(), "target": "ntsCgmExpc", "type": "XML",
        "query": keyword, "display": "20",
    })
    root = ET.fromstring(raw)
    items = []
    for child in list(root):
        if not list(child):
            continue
        d = {}
        for f in list(child):
            if f.text and f.text.strip():
                d[f.tag] = f.text.strip()
        if d:
            items.append(d)
    if not items:
        return "[검색결과 없음] '%s' — 키워드를 줄여 다시 검색하세요." % keyword

    lines = ["※ 같은 안건번호가 연도만 다른 여러 건일 수 있습니다. 인용할 때 해석일자를 함께 적으세요."]
    for d in items:
        title = d.get("안건명", "?")
        case_no = d.get("안건번호", "")
        date = d.get("해석일자", "")
        link = d.get("법령해석상세링크", "")
        m = re.search(r"ntstDcmId=(\d{18})", link)
        dcm_id = m.group(1) if m else "?"
        lines.append("%s [안건번호 %s · 해석일자 %s · ntstDcmId %s]"
                     % (title, case_no or "?", date or "?", dcm_id))
    return "\n".join(lines)


class _HtmlTextExtractor(html.parser.HTMLParser):
    """HTML -> 평문. 표는 셀=탭, 행/문단/줄바꿈=개행으로 남긴다."""
    _NEWLINE_TAGS = {"p", "div", "tr", "li"}
    _CELL_TAGS = {"td", "th"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []

    def handle_starttag(self, tag, attrs):
        if tag == "br":
            self.parts.append("\n")
        elif tag in self._CELL_TAGS:
            self.parts.append("\t")
        elif tag in self._NEWLINE_TAGS:
            self.parts.append("\n")

    def handle_startendtag(self, tag, attrs):
        if tag == "br":
            self.parts.append("\n")

    def handle_data(self, data):
        self.parts.append(data)


def _html_to_text(s):
    if not s:
        return ""
    p = _HtmlTextExtractor()
    p.feed(s)
    text = "".join(p.parts)
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _pinned_context_if_trusted(host):
    """서버 인증서 체인의 루트가 지정된 SHA-256과 정확히 같을 때만 그 루트를
    신뢰 루트로 추가한 SSLContext 를 돌려준다. 다르면 None (아무것도 추가하지 않는다).
    이 연결은 인증서 체인만 읽는다 — 데이터는 받지 않는다."""
    probe_ctx = ssl._create_unverified_context()
    with socket.create_connection((host, 443), timeout=10) as sock:
        with probe_ctx.wrap_socket(sock, server_hostname=host) as ssock:
            chain = ssock.get_unverified_chain()
    if not chain:
        return None
    root_der = bytes(chain[-1])
    if hashlib.sha256(root_der).hexdigest().upper() != TAXLAW_TRUSTED_ROOT_SHA256:
        return None
    ctx = ssl.create_default_context()
    ctx.load_verify_locations(cadata=root_der)
    return ctx


def _rate_limit_taxlaw():
    wait = 0.4 - (time.time() - _last_taxlaw_call[0])
    if wait > 0:
        time.sleep(wait)


def _read_capped(resp, limit=TAXLAW_MAX_BYTES):
    chunks, total = [], 0
    while True:
        chunk = resp.read(65536)
        if not chunk:
            break
        total += len(chunk)
        if total > limit:
            raise ValueError("응답이 상한(%d바이트)을 넘었습니다" % limit)
        chunks.append(chunk)
    return b"".join(chunks)


def _is_cert_verify_error(e):
    if isinstance(e, ssl.SSLCertVerificationError):
        return True
    reason = getattr(e, "reason", None)
    return isinstance(reason, ssl.SSLCertVerificationError) or "CERTIFICATE_VERIFY_FAILED" in str(e)


def _post_taxlaw(ntst_dcm_id):
    """POST https://taxlaw.nts.go.kr/action.do — 공식 API 가 아니므로 속도를 낮추고
    재시도는 1회까지만. 반환: (본문 str, 오류 메시지 str) — 성공이면 오류는 None."""
    url = "https://%s/action.do" % TAXLAW_HOST
    param_data = json.dumps({"dcmDVO": {"ntstDcmId": ntst_dcm_id}},
                            ensure_ascii=False, separators=(",", ":"))
    body = ("actionId=ASIQTB002PR01&paramData=" + urllib.parse.quote(param_data, safe="")).encode("utf-8")
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Content-Type": "application/x-www-form-urlencoded",
        "Referer": "https://%s/qt/USEQTA002P.do?ntstDcmId=%s" % (TAXLAW_HOST, ntst_dcm_id),
    }

    def attempt(context=None):
        _rate_limit_taxlaw()
        req = urllib.request.Request(url, data=body, headers=headers, method="POST")
        _last_taxlaw_call[0] = time.time()
        with urllib.request.urlopen(req, timeout=25, context=context) as resp:
            return _read_capped(resp).decode("utf-8", "replace")

    try:
        return attempt(), None
    except ValueError as e:
        return None, "[조회 실패] %s — 중단했습니다." % e
    except Exception as e:
        if not _is_cert_verify_error(e):
            try:
                return attempt(), None
            except Exception as e2:
                return None, "[조회 실패] 국세법령정보시스템 요청 오류: %s" % e2
        ctx = _pinned_context_if_trusted(TAXLAW_HOST)
        if ctx is None:
            return None, ("[조회 실패] 인증서 확인 실패 — 이 PC 가 신뢰하지 않는 루트이고, "
                          "지정된 지문과도 일치하지 않아 검증을 끄지 않고 중단했습니다.")
        try:
            return attempt(context=ctx), None
        except ValueError as e2:
            return None, "[조회 실패] %s — 중단했습니다." % e2
        except Exception as e2:
            return None, "[조회 실패] 인증서 추가 뒤 재시도도 실패: %s" % e2


def get_tax_interpretation(ntst_dcm_id):
    """ntstDcmId -> 국세법령정보시스템 공개 페이지의 예규 본문(평문).
    법제처 OPEN API 가 아니다 — 결과 끝에 출처와 경고를 반드시 붙인다."""
    ntst_dcm_id = re.sub(r"\D", "", str(ntst_dcm_id))
    if len(ntst_dcm_id) != 18:
        return "[오류] ntstDcmId 는 숫자 18자리여야 합니다."

    raw, err = _post_taxlaw(ntst_dcm_id)
    if err:
        return err
    try:
        data = json.loads(raw)
    except ValueError:
        return "[조회 실패] ntstDcmId %s — 응답이 JSON 이 아닙니다." % ntst_dcm_id

    node = (data.get("data") or {}).get("ASIQTB002PR01") or {}
    dvo = node.get("dcmDVO") or {}
    title = (dvo.get("ntstDcmTtl") or "").strip()
    doc_no = (dvo.get("ntstDcmDscmCntn") or "").strip()
    made_at = (dvo.get("ntstDcmRgtDt") or "").strip()
    gist = _html_to_text(dvo.get("ntstDcmGistCntn") or "")
    reply = _html_to_text(dvo.get("ntstDcmCntn") or "")

    editor_list = node.get("dcmHwpEditorDVOList") or []
    if isinstance(editor_list, dict):
        editor_list = [editor_list]
    body_html = "".join((item.get("dcmFleByte") or "") for item in editor_list if isinstance(item, dict))
    body = _html_to_text(body_html)

    if len(re.sub(r"\s", "", gist + reply + body)) < 20:
        return "[문서 없음] ntstDcmId %s — 본문을 찾지 못했습니다." % ntst_dcm_id

    parts = []
    if title:
        parts.append("제목: %s" % title)
    if doc_no:
        parts.append("문서번호: %s" % doc_no)
    if made_at:
        parts.append("생산일자: %s" % made_at)
    if gist:
        parts.append("요지:\n%s" % gist)
    if reply:
        parts.append("회신:\n%s" % reply)
    if body:
        parts.append("본문:\n%s" % body)

    footer = ("출처: https://%s/qt/USEQTA002P.do?ntstDcmId=%s\n"
             "⚠️ 법제처 OPEN API 가 아닌 국세법령정보시스템 공개 페이지에서 가져온 본문입니다.\n"
             "   인용 전 원문으로 대조하세요.") % (TAXLAW_HOST, ntst_dcm_id)
    return "\n\n".join(parts) + "\n\n" + footer


# -- 과거 판(연혁) — 판은 (MST, 시행일자) 짝으로만 정해진다 --------------------
# target=law 폴백은 넣지 않는다: 에러 없이 다른 연도의 판을 돌려줄 수 있어서다.

def digits(x):
    """숫자만 남긴다: '제21223호' -> '21223', '2019.01.01' -> '20190101'."""
    return re.sub(r"\D", "", str(x or ""))


def history_rows(name):
    """이름이 정확히 같은 법령의 모든 판, 최신순."""
    raw = http_get("https://www.law.go.kr/DRF/lawSearch.do", {
        "OC": read_key(), "target": "eflaw", "type": "XML",
        "query": name, "display": "100",
    }, timeout=30)
    rows = []
    for r in ET.fromstring(raw).iter("law"):
        if (r.findtext("법령명한글") or "").strip() != name:
            continue                    # 검색은 느슨하다 — 이름이 정확히 같은 것만 남긴다
        rows.append({
            "mst": (r.findtext("법령일련번호") or "").strip(),
            "efyd": digits(r.findtext("시행일자")),
            "prom_date": digits(r.findtext("공포일자")),
            "prom_no": (r.findtext("공포번호") or "").strip(),
            "kind": (r.findtext("제개정구분명") or "").strip(),
            "state": (r.findtext("현행연혁코드") or "").strip(),
        })
    # 같은 날 시행되는 여러 판이 있을 수 있다 — 그중 그 날 유효한 것은
    # 가장 나중에 공포된 것이므로 공포일자로 동점을 가른다.
    rows.sort(key=lambda x: (x["efyd"], x["prom_date"]), reverse=True)
    return rows


def list_law_history(name):
    rows = history_rows(name)
    if not rows:
        return "[연혁 없음] '%s' 와 이름이 정확히 같은 법령을 찾지 못했습니다." % name
    lines = ["%s — 시행일 기준 %d개 판 (최신순)" % (name, len(rows)),
             "※ 판은 (MST, 시행일자) 짝으로 고릅니다. 조회할 때 두 값을 함께 넘기세요."]
    for x in rows[:30]:
        lines.append("  시행 %s · 공포 %s 제%s호 · %s · %s · MST %s"
                     % (x["efyd"], x["prom_date"], x["prom_no"], x["kind"], x["state"], x["mst"]))
    if len(rows) > 30:
        lines.append("  … 외 %d개 판" % (len(rows) - 30))
    return "\n".join(lines)


def fetch_version(mst, efyd):
    """(MST, 시행일자) 판의 전체 조문/부칙. 요청한 시행일과 다른 판이 오면 [조회 실패]."""
    mst, efyd = digits(mst), digits(efyd)
    if not mst or len(efyd) != 8:
        return None, "[조회 실패] MST 와 시행일자(8자리)를 함께 넘기세요. 하나만으로는 판이 정해지지 않습니다."
    raw = http_get("https://www.law.go.kr/DRF/lawService.do", {
        "OC": read_key(), "target": "eflaw", "type": "JSON",
        "MST": mst, "efYd": efyd,
    }, timeout=30)
    if len(raw) < 8000 and not raw.lstrip().startswith(("{", "[")) and "미신청된" in raw:
        return None, ("[조회 실패] MST %s · 시행일 %s — 조회 조합이 잘못됐습니다. "
                      "인증키 문제가 아니니 키를 다시 입력하지 마세요." % (mst, efyd))
    try:
        law = json.loads(raw).get("법령")
    except ValueError:
        law = None
    if not isinstance(law, dict):
        return None, ("[조회 실패] MST %s · 시행일 %s — 그 짝의 판이 없습니다. 시행일자는 연혁 목록의 "
                      "그 행 값을 그대로 넣어야 합니다('그 무렵 날짜'는 거부됩니다)." % (mst, efyd))
    got = digits((law.get("기본정보") or {}).get("시행일자"))
    if got != efyd:                    # 다른 판이 조용히 오면 그대로 쓰지 않는다
        return None, "[조회 실패] 요청한 시행일 %s 가 아니라 %s 판이 왔습니다. 쓰지 마세요." % (efyd, got)
    return law, None


def get_article_at(mst, efyd, jo):
    law, err = fetch_version(mst, efyd)
    if err:
        return err
    code = jo_code(jo)
    want_jo, want_ji = int(code[:4]), int(code[4:])
    node = law.get("조문")
    units = as_list(node.get("조문단위")) if isinstance(node, dict) else as_list(node)
    sel = [u for u in units if isinstance(u, dict) and u.get("조문여부") != "전문"
           and _num(u.get("조문번호")) == want_jo and _num(u.get("조문가지번호")) == want_ji]
    if not sel:
        return "[조문 없음] %s — 시행일 %s 판에는 이 조문이 없습니다." % (jo, digits(efyd))
    out = ["=== %s [시행 %s] ===" % ((law.get("기본정보") or {}).get("법령명_한글", ""), digits(efyd))]
    for u in sel:
        body = text_of(u.get("조문내용"))
        if body:
            out.append(body)
        for hang in as_list(u.get("항")):
            if not isinstance(hang, dict):
                continue
            t = text_of(hang.get("항내용"))
            if t:
                out.append("  " + t.replace("\n", "\n  "))
            for ho in as_list(hang.get("호")):
                if not isinstance(ho, dict):
                    continue
                t2 = text_of(ho.get("호내용"))
                if t2:
                    out.append("    " + t2.replace("\n", "\n    "))
                for mok in as_list(ho.get("목")):
                    if isinstance(mok, dict) and text_of(mok.get("목내용")):
                        out.append("      " + text_of(mok.get("목내용")).replace("\n", "\n      "))
            for mok in as_list(hang.get("목")):       # 목 이 항에 바로 걸리는 경우 (05 참고)
                if isinstance(mok, dict) and text_of(mok.get("목내용")):
                    out.append("      " + text_of(mok.get("목내용")).replace("\n", "\n      "))
    return "\n".join(out)


def get_addenda(mst, efyd, prom_no):
    """그 판에 실린, 지정한 공포번호의 부칙."""
    law, err = fetch_version(mst, efyd)
    if err:
        return err
    want = digits(prom_no)
    units = as_list((law.get("부칙") or {}).get("부칙단위"))
    hit = [u for u in units if isinstance(u, dict)
           and digits(u.get("부칙공포번호")).lstrip("0") == want.lstrip("0")]
    if not hit:
        return "[부칙 없음] 시행일 %s 판에 제%s호 부칙이 없습니다." % (digits(efyd), want)
    out = []
    for u in hit:
        out.append("부칙 <제%s호, %s>" % (digits(u.get("부칙공포번호")), digits(u.get("부칙공포일자"))))
        out.append(text_of(u.get("부칙내용")))
    return "\n".join(out)


def search_law_interpretation(keyword):
    """법제처 법령해석례(expc) 검색 -> human-readable list (str).
    건강보험·사회보험 등 법령해석 조회에 활용한다."""
    raw = http_get("https://www.law.go.kr/DRF/lawSearch.do", {
        "OC": read_key(), "target": "expc", "type": "XML",
        "query": keyword, "display": "20",
    })
    root = ET.fromstring(raw)
    items = []
    for child in list(root):
        if not list(child):
            continue
        d = {}
        for f in list(child):
            if f.text and f.text.strip():
                d[f.tag] = f.text.strip()
        if d:
            items.append(d)
    if not items:
        return "[검색결과 없음] '%s' — 키워드를 줄여 다시 검색하세요." % keyword

    lines = []
    for d in items:
        title = d.get("안건명", "?")
        serial = d.get("법령해석례일련번호", "")
        case_no = d.get("안건번호", "")
        org = d.get("회신기관명", "")
        date = d.get("회신일자", "")
        extra = " · ".join(x for x in [org, "안건번호 " + case_no if case_no else "", "회신 " + date if date else ""] if x)
        lines.append("[%s] %s%s" % (serial or "?", title, ("\n   " + extra) if extra else ""))
    return "\n".join(lines)


def get_law_interpretation(serial):
    """법령해석례 일련번호 -> 전체 본문."""
    raw = http_get("https://www.law.go.kr/DRF/lawService.do", {
        "OC": read_key(), "target": "expc", "type": "JSON", "ID": str(serial),
    })
    data = json.loads(raw)
    node = next((v for v in data.values() if isinstance(v, dict)), None)
    if node is None:
        return "[본문 없음] 일련번호 %s — 응답에서 본문 노드를 찾지 못했습니다." % serial

    order = ("안건명", "안건번호", "회신기관명", "회신일자", "질의기관명",
             "질의요지", "회답", "이유", "법령해석례내용")
    out = []
    shown = set()
    for k in order:
        v = node.get(k)
        if isinstance(v, str) and v.strip():
            v = v.strip()
            out.append("%s:\n%s" % (k, v) if len(v) > 60 else "%s: %s" % (k, v))
            shown.add(k)
    for k, v in node.items():
        if k in shown or not isinstance(v, str) or not v.strip():
            continue
        v = v.strip()
        out.append("%s:\n%s" % (k, v) if len(v) > 60 else "%s: %s" % (k, v))
    return "\n\n".join(out) or "[본문 없음] 일련번호 %s" % serial


def search_precedent(keyword):
    """대법원 판례 검색 -> human-readable list (str)."""
    raw = http_get("https://www.law.go.kr/DRF/lawSearch.do", {
        "OC": read_key(), "target": "prec", "type": "XML",
        "query": keyword, "display": "20",
    })
    root = ET.fromstring(raw)
    items = []
    for child in list(root):
        if not list(child):
            continue
        d = {}
        for f in list(child):
            if f.text and f.text.strip():
                d[f.tag] = f.text.strip()
        if d:
            items.append(d)
    if not items:
        return "[검색결과 없음] '%s' — 키워드를 줄여 다시 검색하세요." % keyword

    lines = []
    for d in items:
        title = d.get("사건명", "?")
        case_no = d.get("사건번호", "")
        court = d.get("법원명", "")
        date = d.get("선고일자", "")
        serial = d.get("판례일련번호", "")
        lines.append("[%s] %s%s" % (serial, title,
                     ("\n   %s · %s · 선고 %s" % (court, case_no, date)) if case_no else ""))
    return "\n".join(lines)


def get_precedent(serial):
    """판례 일련번호 -> 전체 본문."""
    raw = http_get("https://www.law.go.kr/DRF/lawService.do", {
        "OC": read_key(), "target": "prec", "type": "JSON", "ID": str(serial),
    })
    data = json.loads(raw)
    node = next((v for v in data.values() if isinstance(v, dict)), None)
    if node is None:
        return "[본문 없음] 일련번호 %s — 응답에서 본문 노드를 찾지 못했습니다." % serial

    order = ("사건명", "사건번호", "법원명", "선고일자", "선고",
             "판시사항", "판결요지", "판례내용")
    out = []
    shown = set()
    for k in order:
        v = node.get(k)
        if isinstance(v, str) and v.strip():
            v = v.strip()
            out.append("%s:\n%s" % (k, v) if len(v) > 60 else "%s: %s" % (k, v))
            shown.add(k)
    for k, v in node.items():
        if k in shown or not isinstance(v, str) or not v.strip():
            continue
        v = v.strip()
        out.append("%s:\n%s" % (k, v) if len(v) > 60 else "%s: %s" % (k, v))
    return "\n\n".join(out) or "[본문 없음] 일련번호 %s" % serial


TOOLS = [
    {
        "name": "ping",
        "description": "서버가 살아 있는지 확인한다. 부르면 pong 을 돌려준다.",
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "search_law",
        "description": "법령명을 검색해서 목록을 돌려준다.",
        "inputSchema": {
            "type": "object",
            "properties": {"keyword": {"type": "string"}},
            "required": ["keyword"],
        },
    },
    {
        "name": "get_law_article",
        "description": "법령 이름과 조문 번호를 받아 그 조문의 본문을 돌려준다.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "law_name": {"type": "string"},
                "jo": {"type": "string"},
            },
            "required": ["law_name", "jo"],
        },
    },
    {
        "name": "search_tribunal",
        "description": "조세심판원 심판례를 검색해서 목록을 돌려준다.",
        "inputSchema": {
            "type": "object",
            "properties": {"keyword": {"type": "string"}},
            "required": ["keyword"],
        },
    },
    {
        "name": "get_tribunal_case",
        "description": "심판례 일련번호로 그 결정문 전체를 돌려준다.",
        "inputSchema": {
            "type": "object",
            "properties": {"serial": {"type": "string"}},
            "required": ["serial"],
        },
    },
    {
        "name": "search_tax_interpretation",
        "description": "국세청 예규(세법해석)를 검색해서 목록을 돌려준다. 본문은 없다.",
        "inputSchema": {
            "type": "object",
            "properties": {"keyword": {"type": "string"}},
            "required": ["keyword"],
        },
    },
    {
        "name": "get_tax_interpretation",
        "description": "ntstDcmId(18자리)로 국세청 예규 본문을 받아 평문으로 돌려준다. "
                       "법제처 OPEN API 가 아닌 국세법령정보시스템 공개 페이지에서 가져온다.",
        "inputSchema": {
            "type": "object",
            "properties": {"ntstDcmId": {"type": "string"}},
            "required": ["ntstDcmId"],
        },
    },
    {
        "name": "list_law_history",
        "description": "법령 이름으로 시행일별 판 목록을 돌려준다. 조회할 때는 그중 한 행의 MST·시행일자를 함께 쓴다.",
        "inputSchema": {
            "type": "object",
            "properties": {"law_name": {"type": "string"}},
            "required": ["law_name"],
        },
    },
    {
        "name": "get_article_at",
        "description": "(MST, 시행일자) 짝과 조문 번호로 그 시점 조문을 돌려준다. 다른 판이 오면 [조회 실패].",
        "inputSchema": {
            "type": "object",
            "properties": {
                "mst": {"type": "string"},
                "efyd": {"type": "string"},
                "jo": {"type": "string"},
            },
            "required": ["mst", "efyd", "jo"],
        },
    },
    {
        "name": "get_addenda",
        "description": "(MST, 시행일자) 짝과 공포번호로 그 판에 실린 부칙을 돌려준다.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "mst": {"type": "string"},
                "efyd": {"type": "string"},
                "prom_no": {"type": "string"},
            },
            "required": ["mst", "efyd", "prom_no"],
        },
    },
    {
        "name": "search_law_interpretation",
        "description": "법제처 법령해석례를 키워드로 검색해서 목록을 돌려준다. "
                       "건강보험·국민연금·고용보험 등 사회보험 관련 법령해석 조회에 활용한다.",
        "inputSchema": {
            "type": "object",
            "properties": {"keyword": {"type": "string"}},
            "required": ["keyword"],
        },
    },
    {
        "name": "get_law_interpretation",
        "description": "법령해석례 일련번호로 그 해석례 전체 본문을 돌려준다.",
        "inputSchema": {
            "type": "object",
            "properties": {"serial": {"type": "string"}},
            "required": ["serial"],
        },
    },
    {
        "name": "search_precedent",
        "description": "대법원 판례를 키워드로 검색해서 목록을 돌려준다.",
        "inputSchema": {
            "type": "object",
            "properties": {"keyword": {"type": "string"}},
            "required": ["keyword"],
        },
    },
    {
        "name": "get_precedent",
        "description": "판례 일련번호로 그 판결문 전체를 돌려준다.",
        "inputSchema": {
            "type": "object",
            "properties": {"serial": {"type": "string"}},
            "required": ["serial"],
        },
    },
]


def call_tool(name, args):
    """Run one tool. Same shape for success and failure; only isError differs."""
    if name == "ping":
        return {"content": [{"type": "text", "text": "pong"}], "isError": False}
    if name == "search_law":
        keyword = (args.get("keyword") or "").strip()
        if not keyword:
            return {"content": [{"type": "text", "text": "[오류] keyword 가 비어 있습니다"}],
                    "isError": True}
        try:
            text = search_law(keyword)
        except Exception as e:
            return {"content": [{"type": "text", "text": "[오류] 검색 실패: %s" % e}],
                    "isError": True}
        return {"content": [{"type": "text", "text": text}], "isError": False}
    if name == "get_law_article":
        law_name = (args.get("law_name") or "").strip()
        jo = (args.get("jo") or "").strip()
        if not law_name or not jo:
            return {"content": [{"type": "text", "text": "[오류] law_name 과 jo 가 모두 필요합니다"}],
                    "isError": True}
        try:
            law_id = find_law_id(law_name)
        except Exception as e:
            return {"content": [{"type": "text", "text": "[조회 실패] 법령 검색 오류: %s" % e}],
                    "isError": True}
        if not law_id:
            return {"content": [{"type": "text", "text": "[조회 실패] '%s' 법령을 찾지 못했습니다" % law_name}],
                    "isError": True}
        try:
            text = get_article(law_id, jo)
        except Exception as e:
            return {"content": [{"type": "text", "text": "[조회 실패] 조문 조회 오류: %s" % e}],
                    "isError": True}
        # 없는 조문은 도구 고장이 아니라 유효한 답 -> isError False.
        # 조회 자체가 깨진 경우([조회 실패])만 isError True.
        return {"content": [{"type": "text", "text": text}],
                "isError": text.startswith("[조회 실패]")}
    if name == "search_tribunal":
        keyword = (args.get("keyword") or "").strip()
        if not keyword:
            return {"content": [{"type": "text", "text": "[오류] keyword 가 비어 있습니다"}],
                    "isError": True}
        try:
            text = search_tribunal(keyword)
        except Exception as e:
            return {"content": [{"type": "text", "text": "[오류] 검색 실패: %s" % e}],
                    "isError": True}
        return {"content": [{"type": "text", "text": text}], "isError": False}
    if name == "get_tribunal_case":
        serial = (args.get("serial") or "").strip()
        if not serial:
            return {"content": [{"type": "text", "text": "[오류] serial 이 비어 있습니다"}],
                    "isError": True}
        try:
            text = get_tribunal(serial)
        except Exception as e:
            return {"content": [{"type": "text", "text": "[조회 실패] 본문 조회 오류: %s" % e}],
                    "isError": True}
        return {"content": [{"type": "text", "text": text}], "isError": False}
    if name == "search_tax_interpretation":
        keyword = (args.get("keyword") or "").strip()
        if not keyword:
            return {"content": [{"type": "text", "text": "[오류] keyword 가 비어 있습니다"}],
                    "isError": True}
        try:
            text = search_tax_interpretation(keyword)
        except Exception as e:
            return {"content": [{"type": "text", "text": "[오류] 검색 실패: %s" % e}],
                    "isError": True}
        return {"content": [{"type": "text", "text": text}], "isError": False}
    if name == "get_tax_interpretation":
        ntst_dcm_id = (args.get("ntstDcmId") or "").strip()
        if not ntst_dcm_id:
            return {"content": [{"type": "text", "text": "[오류] ntstDcmId 가 비어 있습니다"}],
                    "isError": True}
        try:
            text = get_tax_interpretation(ntst_dcm_id)
        except Exception as e:
            return {"content": [{"type": "text", "text": "[조회 실패] 예규 본문 조회 오류: %s" % e}],
                    "isError": True}
        # 문서 없음/입력 오류는 유효한 답 -> isError False. 조회 자체가 깨진 경우만 True.
        return {"content": [{"type": "text", "text": text}],
                "isError": text.startswith("[조회 실패]")}
    if name == "list_law_history":
        law_name = (args.get("law_name") or "").strip()
        if not law_name:
            return {"content": [{"type": "text", "text": "[오류] law_name 이 비어 있습니다"}],
                    "isError": True}
        try:
            text = list_law_history(law_name)
        except Exception as e:
            return {"content": [{"type": "text", "text": "[오류] 연혁 조회 실패: %s" % e}],
                    "isError": True}
        return {"content": [{"type": "text", "text": text}], "isError": False}
    if name == "get_article_at":
        mst = (args.get("mst") or "").strip()
        efyd = (args.get("efyd") or "").strip()
        jo = (args.get("jo") or "").strip()
        if not mst or not efyd or not jo:
            return {"content": [{"type": "text", "text": "[오류] mst · efyd · jo 가 모두 필요합니다"}],
                    "isError": True}
        try:
            text = get_article_at(mst, efyd, jo)
        except Exception as e:
            return {"content": [{"type": "text", "text": "[조회 실패] 조문 조회 오류: %s" % e}],
                    "isError": True}
        return {"content": [{"type": "text", "text": text}],
                "isError": text.startswith("[조회 실패]")}
    if name == "get_addenda":
        mst = (args.get("mst") or "").strip()
        efyd = (args.get("efyd") or "").strip()
        prom_no = (args.get("prom_no") or "").strip()
        if not mst or not efyd or not prom_no:
            return {"content": [{"type": "text", "text": "[오류] mst · efyd · prom_no 가 모두 필요합니다"}],
                    "isError": True}
        try:
            text = get_addenda(mst, efyd, prom_no)
        except Exception as e:
            return {"content": [{"type": "text", "text": "[조회 실패] 부칙 조회 오류: %s" % e}],
                    "isError": True}
        return {"content": [{"type": "text", "text": text}],
                "isError": text.startswith("[조회 실패]")}
    if name == "search_law_interpretation":
        keyword = (args.get("keyword") or "").strip()
        if not keyword:
            return {"content": [{"type": "text", "text": "[오류] keyword 가 비어 있습니다"}],
                    "isError": True}
        try:
            text = search_law_interpretation(keyword)
        except Exception as e:
            return {"content": [{"type": "text", "text": "[오류] 검색 실패: %s" % e}],
                    "isError": True}
        return {"content": [{"type": "text", "text": text}], "isError": False}
    if name == "get_law_interpretation":
        serial = (args.get("serial") or "").strip()
        if not serial:
            return {"content": [{"type": "text", "text": "[오류] serial 이 비어 있습니다"}],
                    "isError": True}
        try:
            text = get_law_interpretation(serial)
        except Exception as e:
            return {"content": [{"type": "text", "text": "[조회 실패] 법령해석례 조회 오류: %s" % e}],
                    "isError": True}
        return {"content": [{"type": "text", "text": text}], "isError": False}
    if name == "search_precedent":
        keyword = (args.get("keyword") or "").strip()
        if not keyword:
            return {"content": [{"type": "text", "text": "[오류] keyword 가 비어 있습니다"}],
                    "isError": True}
        try:
            text = search_precedent(keyword)
        except Exception as e:
            return {"content": [{"type": "text", "text": "[오류] 검색 실패: %s" % e}],
                    "isError": True}
        return {"content": [{"type": "text", "text": text}], "isError": False}
    if name == "get_precedent":
        serial = (args.get("serial") or "").strip()
        if not serial:
            return {"content": [{"type": "text", "text": "[오류] serial 이 비어 있습니다"}],
                    "isError": True}
        try:
            text = get_precedent(serial)
        except Exception as e:
            return {"content": [{"type": "text", "text": "[조회 실패] 본문 조회 오류: %s" % e}],
                    "isError": True}
        return {"content": [{"type": "text", "text": text}], "isError": False}
    return {"content": [{"type": "text", "text": "[오류] 없는 도구: %s" % name}],
            "isError": True}


def handle(req):
    """One request in -> result dict out. None = method not supported."""
    method = req.get("method", "")
    params = req.get("params") or {}
    if method == "initialize":
        return {
            "protocolVersion": params.get("protocolVersion", "2025-06-18"),
            "capabilities": {"tools": {}},
            "serverInfo": {"name": "my-law-server", "version": "0.1"},
        }
    if method == "ping":
        return {}
    if method == "tools/list":
        return {"tools": TOOLS}
    if method == "tools/call":
        return call_tool(params.get("name", ""), params.get("arguments") or {})
    return None


def main():
    log("개인 실습용 도구입니다 — 재배포 금지")
    if not read_key():
        log("setup_key.py 를 먼저 실행하세요")
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            req = json.loads(line)
        except ValueError:
            continue
        if "id" not in req:
            continue
        result = handle(req)
        if result is None:
            resp = {"jsonrpc": "2.0", "id": req["id"],
                    "error": {"code": -32601,
                              "message": "지원하지 않는 요청: %s" % req.get("method", "")}}
        else:
            resp = {"jsonrpc": "2.0", "id": req["id"], "result": result}
        sys.stdout.write(json.dumps(resp, ensure_ascii=False) + "\n")
        sys.stdout.flush()


if __name__ == "__main__":
    main()
