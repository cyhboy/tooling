cd %~dp0
REM copy Excel.officeUI Excel_%username%.officeUI
REM copy Word.officeUI Word_%username%.officeUI


IF EXIST "C:\Program Files\Microsoft Office\Office15\Library\" (
REM copy Justacro.xlam "C:\Program Files\Microsoft Office\Office15\Library\"
del "C:\Program Files\Microsoft Office\Office15\Library\Justacro.xlam"
)

IF EXIST "C:\Program Files (x86)\Microsoft Office\Office15\Library\" (
REM copy Justacro.xlam "C:\Program Files (x86)\Microsoft Office\Office15\Library\"
del "C:\Program Files (x86)\Microsoft Office\Office15\Library\Justacro.xlam"
)

IF EXIST "C:\Program Files\Microsoft Office\Office12\Library\" (
del "C:\Program Files\Microsoft Office\Office12\Library\Justacro.xlam"
)


REM copy Excel_%username%.officeUI "C:\Users\%username%\AppData\Local\Microsoft\Office\Excel.officeUI"
del "C:\Users\%username%\AppData\Local\Microsoft\Office\Excel.officeUI"
del "C:\Users\%username%\AppData\Local\Microsoft\Office\Word.officeUI"


REM mkdir "C:\AppFiles\"
REM copy Justacro.xlam "C:\AppFiles\"

del "C:\AppFiles\Justacro.xlam"
