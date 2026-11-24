@echo off
rem Scheduled by Windows Task Scheduler (2027-01-17 10:00). ASCII only on purpose:
rem cmd.exe misreads Korean text in batch files. All Korean settings live in scheduled_2027-01-17.json.
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d C:\hometax-axtax
"C:\Users\admin\AppData\Local\Python\pythoncore-3.14-64\python.exe" hometax_axtax.py scheduled --config scheduled_2027-01-17.json %*
echo.
echo Finished. Results are in the folder on the Desktop; the run log is the txt file inside it.
pause
