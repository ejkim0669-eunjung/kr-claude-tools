# 사회보험 법령해석례 조회 모음
# my_law_server.py 에서 사회보험(건강보험·국민연금·고용보험·산재보험) 관련 함수 발췌
# 출처 라인: my_law_server.py L619~L677

import json
import os
import sys
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET

HEADERS = {"User-Agent": "Mozilla/5.0", "Referer": "https://www.law.go.kr/"}


def read_key():
    oc = (os.environ.get("LAW_OC") or "").strip()
    if oc:
        return oc
    try:
        with open(r"C:\krlaw\law_oc.txt", encoding="utf-8") as f:
            return f.read().strip()
    except OSError:
        return ""


def http_get(url, params, timeout=20):
    full = url + "?" + urllib.parse.urlencode(params, encoding="utf-8")
    req = urllib.request.Request(full, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read().decode("utf-8", "replace")


# ── 법령해석례 검색 (건강보험·국민연금·고용보험 등) ──────────────────────────

def search_law_interpretation(keyword):
    """법제처 법령해석례(expc) 키워드 검색 -> human-readable list.
    사용 예) search_law_interpretation("건강보험료 부과")
    """
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
        title  = d.get("안건명", "?")
        serial = d.get("법령해석례일련번호", "")
        case_no = d.get("안건번호", "")
        org    = d.get("회신기관명", "")
        date   = d.get("회신일자", "")
        extra  = " · ".join(x for x in [
            org,
            "안건번호 " + case_no if case_no else "",
            "회신 " + date if date else "",
        ] if x)
        lines.append("[%s] %s%s" % (serial or "?", title,
                     ("\n   " + extra) if extra else ""))
    return "\n".join(lines)


def get_law_interpretation(serial):
    """법령해석례 일련번호 -> 전체 본문.
    사용 예) get_law_interpretation("100171")
    """
    raw = http_get("https://www.law.go.kr/DRF/lawService.do", {
        "OC": read_key(), "target": "expc", "type": "JSON", "ID": str(serial),
    })
    data = json.loads(raw)
    node = next((v for v in data.values() if isinstance(v, dict)), None)
    if node is None:
        return "[본문 없음] 일련번호 %s" % serial

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


# ── 간단 테스트 (직접 실행 시) ──────────────────────────────────────────────
if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    kw = sys.argv[1] if len(sys.argv) > 1 else "건강보험료"
    print(search_law_interpretation(kw))
