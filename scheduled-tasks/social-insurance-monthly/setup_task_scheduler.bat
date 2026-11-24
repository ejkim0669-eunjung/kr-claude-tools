@echo off
chcp 65001 >nul
:: Windows 작업 스케줄러에 사회보험료 조회 작업 등록
:: 매월 20, 21, 22, 23일 오전 11시 실행
:: 관리자 권한으로 실행 필요

set TASK_NAME=사회보험료_월간조회
set SCRIPT_DIR=%~dp0
set PYTHON_EXE=python

echo [설정] 작업 이름: %TASK_NAME%
echo [설정] 스크립트: %SCRIPT_DIR%social_insurance_monthly.py
echo.

:: 기존 작업 삭제 (있으면)
schtasks /Delete /TN "%TASK_NAME%" /F 2>nul

:: 매월 20일 11:00
schtasks /Create /TN "%TASK_NAME%" ^
    /TR "\"%PYTHON_EXE%\" \"%SCRIPT_DIR%social_insurance_monthly.py\"" ^
    /SC MONTHLY /D 20,21,22,23 ^
    /ST 11:00 ^
    /RL HIGHEST ^
    /F

if errorlevel 1 (
    echo [실패] 작업 등록 실패. 관리자 권한으로 다시 실행하세요.
    pause
    exit /b 1
)

echo.
echo [완료] Windows 작업 스케줄러에 등록됨
echo   - 작업 이름: %TASK_NAME%
echo   - 실행 일시: 매월 20~23일 오전 11:00
echo   - 확인: 작업 스케줄러(taskschd.msc) 열어 확인
pause
