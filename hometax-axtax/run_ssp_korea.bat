@echo off
chcp 65001 > nul
cd /d C:\hometax-axtax
python hometax_axtax.py scheduled --config 홈택스_조회.json --only 843-88-02704 --redo --launch-chrome
pause
