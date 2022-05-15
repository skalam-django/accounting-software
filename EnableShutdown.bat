@echo off
echo y | REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" /V NoClose /t REG_DWORD /d 00000000 
taskkill /f /im explorer.exe
cls
start explorer.exe
exit /b 