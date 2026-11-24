# ---------------------------------------------------------------
# 홈택스 MCP 서버를 Claude 데스크톱 앱에 등록합니다.
#
# Claude 앱이 완전히 꺼진 상태에서만 실행하세요.
# 앱이 켜진 채로 설정을 고치면 앱이 자기 설정을 다시 쓰면서
# 방금 넣은 항목을 지워 버립니다.
#
# 이 파일을 직접 실행하지 말고 등록하기_홈택스.bat 를 더블클릭하세요.
# ---------------------------------------------------------------

import datetime
import glob
import json
import os
import subprocess
import sys

SERVER_NAME = "hometax-axtax"
SERVER_PY   = r"C:\hometax-axtax\hometax_mcp_server.py"


def show(title, body, error=True):
    """더블클릭으로 실행하는 프로그램이므로 창으로 결과를 보여준다."""
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        (messagebox.showerror if error else messagebox.showinfo)(title, body)
        root.destroy()
    except Exception:
        print(title, "\n", body)


def claude_running():
    """Claude 프로세스가 살아 있으면 True."""
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq Claude.exe", "/NH"],
            capture_output=True, text=True, timeout=15
        ).stdout
        return "claude.exe" in (out or "").lower()
    except Exception:
        return False


def config_paths():
    """이 PC 에서 적용될 수 있는 모든 config 경로를 돌려준다.
    Store 버전은 LocalCache\\Packages 아래에, 일반 설치본은 %APPDATA% 에 있다.
    둘 다 쓴다 — 없는 쪽에 써도 무해하다."""
    found = []
    local = os.environ.get("LOCALAPPDATA", "")
    if local:
        for pkg in glob.glob(os.path.join(local, "Packages", "Claude_*")):
            found.append((
                os.path.join(pkg, "LocalCache", "Roaming", "Claude",
                             "claude_desktop_config.json"),
                "Store 버전"
            ))
    roaming = os.environ.get("APPDATA", "")
    if roaming:
        found.append((
            os.path.join(roaming, "Claude", "claude_desktop_config.json"),
            "일반 설치본"
        ))
    return found


def register_one(path):
    """config 파일 하나에 항목을 추가·덮어쓴다. 결과 한 줄을 돌려준다."""
    os.makedirs(os.path.dirname(path), exist_ok=True)

    data = {}
    if os.path.exists(path):
        raw   = open(path, "rb").read()
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        open(f"{path}.backup_{stamp}", "wb").write(raw)   # 백업 먼저
        body  = raw.decode("utf-8-sig", "replace").strip()
        if body:
            try:
                data = json.loads(body)
            except ValueError:
                return f"실패 — 기존 설정 파일이 깨져 있습니다:\n   {path}"

    if not isinstance(data, dict):
        data = {}
    servers = data.get("mcpServers")
    if not isinstance(servers, dict):
        servers = {}

    before = sorted(servers.keys())
    # 우리 항목만 교체하고 다른 서버는 그대로 둔다
    servers[SERVER_NAME] = {
        "command": sys.executable or "python",
        "args":    [SERVER_PY],
    }
    data["mcpServers"] = servers

    open(path, "wb").write(
        json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
    )

    # 저장 뒤 다시 읽어서 확인
    check = json.loads(open(path, "rb").read().decode("utf-8-sig", "replace"))
    if check.get("mcpServers", {}).get(SERVER_NAME, {}).get("args") != [SERVER_PY]:
        return f"실패 — 저장 뒤 확인했는데 항목이 없습니다:\n   {path}"

    kept = [k for k in before if k in check["mcpServers"] and k != SERVER_NAME]
    return "성공 — %s\n   함께 유지된 기존 항목 %d개%s" % (
        path, len(kept), (": " + ", ".join(kept)) if kept else ""
    )


def main():
    if claude_running():
        show(
            "Claude 앱이 아직 켜져 있습니다",
            "작업 관리자(Ctrl+Shift+Esc)에서 Claude 를 모두 끝내기 한 뒤\n"
            "이 프로그램을 다시 실행해 주세요.\n\n"
            "앱이 켜진 채로 설정을 고치면 앱이 자기 설정을 다시 쓰면서\n"
            "방금 넣은 항목을 지워 버립니다."
        )
        return 1

    if not os.path.exists(SERVER_PY):
        show(
            "서버 파일을 찾지 못했습니다",
            f"{SERVER_PY} 가 없습니다.\n"
            "hometax_mcp_server.py 가 같은 폴더에 있는지 확인하세요."
        )
        return 1

    lines = [
        "[%s]\n%s" % (kind, register_one(path))
        for path, kind in config_paths()
    ]
    if not lines:
        show("설정 위치를 찾지 못했습니다",
             "Claude 데스크톱 앱이 설치되어 있는지 확인하세요.")
        return 1

    body = "\n\n".join(lines)
    ok   = all("성공" in ln for ln in lines)
    if ok:
        show("등록되었습니다",
             "이제 Claude 앱을 다시 켜세요.\n\n"
             "Claude 에서 '홈택스 조회해줘' 라고 하면\n"
             "hometax-axtax 도구가 자동으로 실행됩니다.\n\n" + body,
             error=False)
    else:
        show("일부 실패했습니다", body)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
