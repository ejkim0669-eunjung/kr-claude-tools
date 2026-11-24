@echo off
chcp 65001 >nul
cd /d "%~dp0"
python social_insurance_monthly.py %*
if errorlevel 1 (
    echo [실패] 일부 사이트 조회 실패 — run.log 확인
    pause
)
