@echo off
cd /d C:\kr-tools
cmdkey /delete:git:https://github.com > nul 2>&1
cmdkey /delete:https://github.com > nul 2>&1
git push -u origin master
pause
