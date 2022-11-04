cd %~dp0
copy Excel.officeUI Excel_%username%.officeUI
copy Word.officeUI Word_%username%.officeUI

IF EXIST "C:\Program Files\Microsoft Office\Office16\Library\" (
REM copy Justacro.xlam "C:\Program Files\Microsoft Office\Office16\Library\"
copy cst.xlam "C:\Program Files\Microsoft Office\Office16\Library\"
)

IF EXIST "C:\Program Files\Microsoft Office\Office15\Library\" (
REM copy Justacro.xlam "C:\Program Files\Microsoft Office\Office15\Library\"
copy cst.xlam "C:\Program Files\Microsoft Office\Office15\Library\"
)

IF EXIST "C:\Program Files (x86)\Microsoft Office\Office15\Library\" (
REM copy Justacro.xlam "C:\Program Files (x86)\Microsoft Office\Office15\Library\"
copy cst.xlam "C:\Program Files (x86)\Microsoft Office\Office15\Library\"
)

IF EXIST "C:\Program Files\Microsoft Office\Office12\Library\" (
REM copy Justacro.xlam "C:\Program Files\Microsoft Office\Office12\Library\"
copy cst.xlam "C:\Program Files\Microsoft Office\Office12\Library\"
)

REM copy Justacro.xlam "C:\Documents and Settings\%username%\Application Data\Microsoft\AddIns\"
copy cst.xlam "C:\Documents and Settings\%username%\Application Data\Microsoft\AddIns\"

copy Excel_%username%.officeUI "C:\Users\%username%\AppData\Local\Microsoft\Office\Excel.officeUI"
REM copy Word_%username%.officeUI "C:\Users\%username%\AppData\Local\Microsoft\Office\Word.officeUI"


mkdir "C:\AppFiles\"
REM copy Justacro.xlam "C:\AppFiles\"
copy cst.xlam "C:\AppFiles\"


REM copy Justacro.xlam "C:\Program Files (x86)\Microsoft Office\Office14\Library\"
REM copy Justacro.xlam "C:\Program Files\Microsoft Office\Office14\Library\"

REM copy Macro.xla "C:\Documents and Settings\%username%\Application Data\Microsoft\AddIns\"
REM del "C:\Documents and Settings\%username%\Application Data\Microsoft\AddIns\Macro.xla"
REM del "C:\Users\%username%\AppData\Roaming\Microsoft\AddIns\Macro.xla"
REM del "C:\Program Files (x86)\Microsoft Office\Office14\Library\Macro.xla"
REM del "C:\Program Files\Microsoft Office\Office14\Library\Macro.xla"
REM del "C:\AppFiles\Macro.xla"

REM copy XSFormatCleaner.xla "C:\Program Files (x86)\Microsoft Office\Office14\Library\"
REM copy XSFormatCleaner.xla "C:\Program Files\Microsoft Office\Office14\Library\"

REM copy PresentationMacro.ppam "C:\Users\%username%\AppData\Roaming\Microsoft\AddIns\"


REM copy PresentationMacro.ppam "C:\AppFiles\"

REM copy "WaitThenRunHiddenJob.vbs" "C:\AppFiles\"
REM copy "RUNAS.vbs" "C:\AppFiles\"


