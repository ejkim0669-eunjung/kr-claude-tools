@echo off
echo 크롬 종료 중...
taskkill /f /im chrome.exe >nul 2>&1
timeout /t 3 /nobreak >nul
echo 크롬 시작...
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9333 --user-data-dir=C:\hometax-axtax\chrome-profile --disable-features=LocalNetworkAccessChecks --disable-popup-blocking --new-window https://edi.nhis.or.kr/
echo 완료!
timeout /t 2 /nobreak >nul
