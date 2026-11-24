@echo off
cd /d "%~dp0"
set "PYEXE="
py -c "import sys" >nul 2>nul
if not errorlevel 1 set "PYEXE=py"
if not defined PYEXE python -c "import sys" >nul 2>nul
if not defined PYEXE if not errorlevel 1 set "PYEXE=python"
if defined PYEXE %PYEXE% "%~dp007_register.py"
if not defined PYEXE echo [!] Python not found - see step 0-A (alias OFF + reinstall)
pause
