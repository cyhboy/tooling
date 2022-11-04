REM echo %~dp0
set str1=%~dp0

set substr1=%str1:~0,2%


IF %substr1%==C: (GOTO :CACTION)

IF %substr1%==M: (GOTO :MACTION)

:CACTION
set str2=%str1:C:=M:%
echo %str2%
IF NOT exist "%str2%" (md "%str2%")
REM explorer.exe "%str2%"
"M:\AppFiles\Beyond Compare 3\BCompare.exe" "%str1%" "%str2%"

GOTO :END

:MACTION
set str2=%str1:M:=C:%
echo %str2%
IF NOT exist "%str2%" (md "%str2%")
REM explorer.exe "%str2%"
"M:\AppFiles\Beyond Compare 3\BCompare.exe" "%str2%" "%str1%"

GOTO :END

:END