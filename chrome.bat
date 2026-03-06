@echo off
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="%TEMP%\quickariba-profile" "https://s1-eu.ariba.com/gb/?realm=uva-1&locale=en_US"
exit /b 0