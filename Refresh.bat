@echo off
ping -n 1 127.0.01 > NUL 2>&1
taskkill /f /im explorer.exe
cls
ping -n 1 127.0.01 > NUL 2>&1
echo.
ping -n 1 127.0.01 > NUL 2>&1
start explorer.exe
exit