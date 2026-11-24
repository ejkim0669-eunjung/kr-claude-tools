"""
홈택스 조회 MCP 서버

Claude Desktop 에서 "홈택스 조회해줘" 라고 하면 이 서버의 도구를 불러
hometax_axtax.py 를 실행해 조회합니다.

사전 조건:
  1. 홈택스 전용 크롬이 디버깅 포트 9333 으로 열려 있어야 합니다
     (CLAUDE.md 의 '홈택스 크롬 띄우는 명령' 참고)
  2. 사람이 직접 홈택스에 로그인하고 조회 화면까지 옮겨 놓아야 합니다
  - 로그인은 이 서버가 절대 대신하지 않습니다

표준 라이브러리만 사용합니다 (json, os, subprocess, sys).
hometax_axtax.py 를 subprocess 로 호출합니다.
"""

import json
import os
import subprocess
import sys

# stdout 은 JSON-RPC 전용 채널. 로그는 전부 stderr 로.
sys.stdout = open(sys.stdout.fileno(), "w", encoding="utf-8", newline="\n", closefd=False)
sys.stderr = open(sys.stderr.fileno(), "w", encoding="utf-8", closefd=False)
sys.stdin  = open(sys.stdin.fileno(),  "r", encoding="utf-8", closefd=False)

_HERE         = os.path.dirname(os.path.abspath(__file__))
SERVER_SCRIPT = os.path.join(_HERE, "hometax_axtax.py")
DEFAULT_LIST  = os.path.join(
    os.path.expanduser("~"), "Desktop",
    "2. 거래처 인수인계", "김은정 부가세신고 거래처.xlsx"
)

# ---------------------------------------------------------------------------
# 도구 목록
# ---------------------------------------------------------------------------

TOOLS = [
    {
        "name": "hometax_tonghap",
        "description": (
            "홈택스 부가가치세 신고자료 통합조회를 실행합니다. "
            "year(연도), ki(1기/2기), rtn(예정/확정/예정+확정)이 필요합니다. "
            "list_path 를 생략하면 기본 거래처 명단(바탕화면/거래처.xlsx)을 사용합니다. "
            "결과는 바탕화면 홈택스_조회결과 폴더에 거래처별로 저장됩니다."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "year":      {"type": "string",  "description": "과세기간 연도 (예: 2026)"},
                "ki":        {"type": "string",  "enum": ["1기", "2기"]},
                "rtn":       {"type": "string",  "enum": ["예정", "확정", "예정+확정"]},
                "list_path": {"type": "string",  "description": "거래처 명단 xlsx 경로 (생략 시 기본값)"},
                "biz_no":    {"type": "string",  "description": "특정 사업자번호 하나만 조회"},
                "name":      {"type": "string",  "description": "biz_no 사용 시 상호명 (폴더 이름에 사용)"},
                "only":      {"type": "string",  "description": "특정 사업자번호만, 쉼표로 구분"},
                "limit":     {"type": "integer", "description": "앞에서 N개 거래처만"},
                "dry_run":   {"type": "boolean", "description": "점검만, 실제 조회 안 함"},
                "redo":      {"type": "boolean", "description": "이미 끝난 거래처도 다시 조회"},
                "out":       {"type": "string",  "description": "저장 폴더 (생략 시 바탕화면)"},
            },
            "required": ["year", "ki", "rtn"],
        },
    },
    {
        "name": "hometax_scheduled",
        "description": (
            "홈택스 예약 실행: 통합조회 → 신용카드/판매대행을 차례로 자동 실행합니다. "
            "year 과 ki 가 필요합니다. 크롬이 이미 열려 있고 로그인되어 있어야 합니다."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "year":      {"type": "string",  "description": "과세기간 연도 (예: 2026)"},
                "ki":        {"type": "string",  "enum": ["1기", "2기"]},
                "list_path": {"type": "string"},
                "only":      {"type": "string"},
                "limit":     {"type": "integer"},
                "redo":      {"type": "boolean"},
                "out":       {"type": "string"},
            },
            "required": ["year", "ki"],
        },
    },
    {
        "name": "hometax_card_sum",
        "description": (
            "홈택스 신용카드/판매(결제)대행 매출자료 조회. "
            "year, from_q(시작 분기), to_q(끝 분기)가 필요합니다. "
            "조회 가능 범위: 단일 분기, 상반기(1~2분기), 하반기(3~4분기)."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "year":      {"type": "string"},
                "from_q":    {"type": "string", "enum": ["1분기", "2분기", "3분기", "4분기"]},
                "to_q":      {"type": "string", "enum": ["1분기", "2분기", "3분기", "4분기"]},
                "list_path": {"type": "string"},
                "biz_no":    {"type": "string"},
                "name":      {"type": "string"},
                "only":      {"type": "string"},
                "limit":     {"type": "integer"},
                "redo":      {"type": "boolean"},
                "out":       {"type": "string"},
                "no_excel":  {"type": "boolean", "description": "엑셀 내려받기를 하지 않는다"},
            },
            "required": ["year", "from_q", "to_q"],
        },
    },
    {
        "name": "hometax_cash_sum",
        "description": "홈택스 현금영수증 매출총액(매출내역누계) 조회. year 과 ki 가 필요합니다.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "year":      {"type": "string"},
                "ki":        {"type": "string", "enum": ["1기", "2기"]},
                "list_path": {"type": "string"},
                "biz_no":    {"type": "string"},
                "name":      {"type": "string"},
                "only":      {"type": "string"},
                "limit":     {"type": "integer"},
                "redo":      {"type": "boolean"},
                "out":       {"type": "string"},
            },
            "required": ["year", "ki"],
        },
    },
    {
        "name": "hometax_export_sum",
        "description": "홈택스 수출실적명세서 조회. year 과 ki 가 필요합니다. 수출이 있는 거래처만 저장합니다.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "year":      {"type": "string"},
                "ki":        {"type": "string", "enum": ["1기", "2기"]},
                "list_path": {"type": "string"},
                "biz_no":    {"type": "string"},
                "name":      {"type": "string"},
                "only":      {"type": "string"},
                "limit":     {"type": "integer"},
                "redo":      {"type": "boolean"},
                "out":       {"type": "string"},
            },
            "required": ["year", "ki"],
        },
    },
    {
        "name": "hometax_check_list",
        "description": (
            "거래처 명단(xlsx) 형식을 점검합니다. "
            "상호·번호는 출력하지 않고 개수·문제 행만 보고합니다. "
            "list_path 를 생략하면 기본 거래처 명단을 사용합니다."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "list_path": {"type": "string"},
            },
            "required": [],
        },
    },
]


# ---------------------------------------------------------------------------
# 도구 실행
# ---------------------------------------------------------------------------

_CMD_MAP = {
    "hometax_tonghap":    "tonghap",
    "hometax_scheduled":  "scheduled",
    "hometax_card_sum":   "card-sum",
    "hometax_cash_sum":   "cash-sum",
    "hometax_export_sum": "export-sum",
    "hometax_check_list": "check-list",
}


def _build_cmd(hta_command, a):
    """hometax_axtax.py 명령행 조립. a = arguments dict."""
    cmd = [sys.executable, SERVER_SCRIPT, hta_command]

    # 거래처 소스: 번호 하나 vs 명단 파일
    if a.get("biz_no"):
        cmd += ["--biz-no", str(a["biz_no"])]
        if a.get("name"):
            cmd += ["--name", str(a["name"])]
    else:
        cmd += ["--list", str(a.get("list_path") or DEFAULT_LIST)]

    # 공통 기간 인수
    if a.get("year"):    cmd += ["--year",   str(a["year"])]
    if a.get("ki"):      cmd += ["--ki",     str(a["ki"])]
    if a.get("rtn"):     cmd += ["--rtn",    str(a["rtn"])]
    if a.get("from_q"):  cmd += ["--from-q", str(a["from_q"])]
    if a.get("to_q"):    cmd += ["--to-q",   str(a["to_q"])]

    # 필터
    if a.get("only"):
        for n in str(a["only"]).split(","):
            n = n.strip()
            if n:
                cmd += ["--only", n]
    if a.get("limit"):
        cmd += ["--limit", str(a["limit"])]

    # 옵션 플래그
    if a.get("dry_run"):  cmd.append("--dry-run")
    if a.get("redo"):     cmd.append("--redo")
    if a.get("no_excel"): cmd.append("--no-excel")
    if a.get("out"):      cmd += ["--out", str(a["out"])]

    return cmd


def run_tool(name, args):
    if not os.path.exists(SERVER_SCRIPT):
        return f"오류: hometax_axtax.py 를 찾지 못했습니다\n경로: {SERVER_SCRIPT}"

    hta_cmd = _CMD_MAP.get(name)
    if not hta_cmd:
        return f"알 수 없는 도구: {name}"

    if hta_cmd == "check-list":
        cmd = [sys.executable, SERVER_SCRIPT, "check-list",
               "--list", str(args.get("list_path") or DEFAULT_LIST)]
    else:
        cmd = _build_cmd(hta_cmd, args)

    print(f"[hometax-mcp] 실행: {cmd}", file=sys.stderr, flush=True)

    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=3600,       # 거래처가 많으면 최대 1시간까지
        )
        stdout = (proc.stdout or "").strip()
        stderr = (proc.stderr or "").strip()

        # hometax_axtax.py 는 로그를 stdout 에 출력한다
        result = stdout if stdout else "(출력 없음)"
        if stderr:
            result += "\n\n[stderr]\n" + stderr
        return result

    except subprocess.TimeoutExpired:
        return "오류: 1시간을 초과해 강제 종료했습니다."
    except FileNotFoundError:
        return "오류: python 또는 hometax_axtax.py 를 실행할 수 없습니다 (경로를 확인하세요)."
    except Exception as e:
        return f"오류: {type(e).__name__}: {e}"


# ---------------------------------------------------------------------------
# JSON-RPC 2.0 메인 루프
# ---------------------------------------------------------------------------

def _send(obj):
    print(json.dumps(obj, ensure_ascii=False), flush=True)


def _handle(msg):
    method = msg.get("method", "")
    id_    = msg.get("id")

    if method == "initialize":
        _send({"jsonrpc": "2.0", "id": id_, "result": {
            "protocolVersion": "2024-11-05",
            "capabilities": {"tools": {}},
            "serverInfo": {"name": "hometax-axtax", "version": "1.0.0"},
        }})

    elif method == "notifications/initialized":
        pass   # 알림 — 응답하지 않는다

    elif method == "tools/list":
        _send({"jsonrpc": "2.0", "id": id_, "result": {"tools": TOOLS}})

    elif method == "tools/call":
        params = msg.get("params", {})
        name   = params.get("name", "")
        args   = params.get("arguments") or {}
        text   = run_tool(name, args)
        _send({"jsonrpc": "2.0", "id": id_, "result": {
            "content": [{"type": "text", "text": text}],
        }})

    elif id_ is not None:
        _send({"jsonrpc": "2.0", "id": id_, "error": {
            "code": -32601,
            "message": f"Method not found: {method}",
        }})


def main():
    for raw in sys.stdin:
        raw = raw.strip()
        if not raw:
            continue
        try:
            msg = json.loads(raw)
        except json.JSONDecodeError as e:
            _send({"jsonrpc": "2.0", "id": None,
                   "error": {"code": -32700, "message": str(e)}})
            continue
        try:
            _handle(msg)
        except Exception as e:
            id_ = msg.get("id")
            if id_ is not None:
                _send({"jsonrpc": "2.0", "id": id_,
                       "error": {"code": -32603, "message": str(e)}})


if __name__ == "__main__":
    main()
