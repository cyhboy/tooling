
cscript Rename_ByDate.vbs Excel_%username%.officeUI
cscript Rename_ByDate.vbs Justacro_%username%.xlam
cscript Rename_ByDate.vbs cst_%username%.xlam

copy "%USERPROFILE%\AppData\Local\Microsoft\Office\Excel.officeUI" Excel_%username%.officeUI

copy "%USERPROFILE%\AppData\Local\Microsoft\Office\Word.officeUI" Word_%username%.officeUI



copy "%USERPROFILE%\AppData\Roaming\Microsoft\AddIns\Justacro.xlam" "C:\AppFiles\SupportDeliver\Justacro_%username%.xlam"
copy "%USERPROFILE%\AppData\Roaming\Microsoft\AddIns\cst.xlam" "C:\AppFiles\SupportDeliver\cst_%username%.xlam"

IF EXIST "C:\Program Files (x86)\Microsoft Office\Office15\Library\" (
copy "C:\Program Files (x86)\Microsoft Office\Office15\Library\Justacro.xlam" "C:\AppFiles\SupportDeliver\Justacro_%username%.xlam"
copy "C:\Program Files (x86)\Microsoft Office\Office15\Library\cst.xlam" "C:\AppFiles\SupportDeliver\cst_%username%.xlam"
)

IF EXIST "C:\Program Files (x86)\Microsoft Office\Office16\Library\" (
copy "C:\Program Files (x86)\Microsoft Office\Office16\Library\Justacro.xlam" "C:\AppFiles\SupportDeliver\Justacro_%username%.xlam"
copy "C:\Program Files (x86)\Microsoft Office\Office16\Library\cst.xlam" "C:\AppFiles\SupportDeliver\cst_%username%.xlam"
)


copy Excel_%username%.officeUI "C:\AppFiles\SupportSetup\Excel.officeUI"
copy Excel_%username%.officeUI "M:\AppFiles\SupportSetup\Excel.officeUI"
copy Excel_%username%.officeUI "I:\AppFiles\SupportSetup\Excel.officeUI"


copy Word_%username%.officeUI "C:\AppFiles\SupportSetup\Word.officeUI"
copy Word_%username%.officeUI "M:\AppFiles\SupportSetup\Word.officeUI"
copy Word_%username%.officeUI "I:\AppFiles\SupportSetup\Word.officeUI"


copy Justacro_%username%.xlam "C:\AppFiles\SupportSetup\Justacro.xlam"
copy Justacro_%username%.xlam "M:\AppFiles\SupportSetup\Justacro.xlam"
copy Justacro_%username%.xlam "I:\AppFiles\SupportSetup\Justacro.xlam"

copy cst_%username%.xlam "C:\AppFiles\SupportSetup\cst.xlam"
copy cst_%username%.xlam "M:\AppFiles\SupportSetup\cst.xlam"
copy cst_%username%.xlam "I:\AppFiles\SupportSetup\cst.xlam"

copy Justacro_%username%.xlam "C:\AppFiles\Justacro.xlam"
copy Justacro_%username%.xlam "%USERPROFILE%\Documents\Justacro.xlam"
copy Justacro_%username%.xlam "%USERPROFILE%\Desktop\Justacro.xlam"

copy cst_%username%.xlam "C:\AppFiles\cst.xlam"
copy cst_%username%.xlam "%USERPROFILE%\Documents\cst.xlam"
copy cst_%username%.xlam "%USERPROFILE%\Desktop\cst.xlam"

REM cscript Rename_ByDate.vbs Svc_%username%.xlam

REM cscript Rename_ByDate.vbs PresentationMacro_%username%.ppam
REM cscript Rename_ByDate.vbs Word_%username%.officeUI


REM cscript replace.vbs "Excel_%username%.officeUI" "C:_Program_Files_Microsoft_Office_Office15_LIBRARY_" ""
REM cscript replace.vbs "Excel_%username%.officeUI" "C:\Program Files\Microsoft Office\Office15\LIBRARY\" ""


REM copy Macro.xla "C:\Documents and Settings\%username%\Application Data\Microsoft\AddIns\"
REM del "C:\Documents and Settings\%username%\Application Data\Microsoft\AddIns\Macro.xla"
REM copy Excel_%username%.officeUI "C:\Users\%username%\AppData\Local\Microsoft\Office\Excel.officeUI"
REM copy Excel_%username%.officeUI Excel.officeUI

REM del "C:\Program Files\Microsoft Office\Office14\Library\Macro.xla"

REM copy "C:\Program Files\Microsoft Office\Office15\Library\Justacro.xlam" "C:\AppFiles\SupportDeliver\Justacro_%username%.xlam"

REM copy "C:\Program Files\Microsoft Office\Office15\Library\Svc.xlam" "C:\AppFiles\SupportDeliver\Svc_%username%.xlam"

REM copy "C:\Program Files\Microsoft Office\Office14\Library\XSFormatCleaner.xla" "C:\AppFiles\SupportDeliver\"

REM copy "C:\Users\%username%\AppData\Roaming\Microsoft\AddIns\PresentationMacro.ppam" "C:\AppFiles\SupportDeliver\PresentationMacro_%username%.ppam"

REM copy Excel_%username%.officeUI "G:\AppFiles\SupportSetup\Excel.officeUI"
REM copy Excel_%username%.officeUI "X:\AppFiles\SupportSetup\Excel.officeUI"
REM copy Excel_%username%.officeUI "I:\AppFiles\SupportSetup\Excel.officeUI"
REM copy Excel_%username%.officeUI "J:\AppFiles\SupportSetup\Excel.officeUI"
REM copy Excel_%username%.officeUI "U:\AppFiles\SupportSetup\Excel.officeUI"
REM copy Excel_%username%.officeUI "Z:\AppFiles\SupportSetup\Excel.officeUI"

REM copy Word_%username%.officeUI "G:\AppFiles\SupportSetup\Word.officeUI"
REM copy Word_%username%.officeUI "X:\AppFiles\SupportSetup\Word.officeUI"
REM copy Word_%username%.officeUI "I:\AppFiles\SupportSetup\Word.officeUI"
REM copy Word_%username%.officeUI "J:\AppFiles\SupportSetup\Word.officeUI"
REM copy Word_%username%.officeUI "U:\AppFiles\SupportSetup\Word.officeUI"
REM copy Word_%username%.officeUI "Z:\AppFiles\SupportSetup\Word.officeUI"

REM copy Justacro_%username%.xlam "C:\AppFiles\SupportSetup\Svc.xlam"
REM copy Justacro_%username%.xlam "M:\AppFiles\SupportSetup\Svc.xlam"

REM copy Justacro_%username%.xlam "G:\AppFiles\SupportSetup\Justacro.xlam"
REM copy Justacro_%username%.xlam "X:\AppFiles\SupportSetup\Justacro.xlam"
REM copy Justacro_%username%.xlam "I:\AppFiles\SupportSetup\Justacro.xlam"
REM copy Justacro_%username%.xlam "J:\AppFiles\SupportSetup\Justacro.xlam"
REM copy Justacro_%username%.xlam "U:\AppFiles\SupportSetup\Justacro.xlam"
REM copy Justacro_%username%.xlam "Z:\AppFiles\SupportSetup\Justacro.xlam"

REM copy XSFormatCleaner.xla "C:\AppFiles\SupportSetup\"
REM copy XSFormatCleaner.xla "M:\AppFiles\SupportSetup\"
REM copy XSFormatCleaner.xla "G:\AppFiles\SupportSetup\"
REM copy XSFormatCleaner.xla "X:\AppFiles\SupportSetup\"
REM copy XSFormatCleaner.xla "I:\AppFiles\SupportSetup\"
REM copy XSFormatCleaner.xla "J:\AppFiles\SupportSetup\"
REM copy XSFormatCleaner.xla "U:\AppFiles\SupportSetup\"
REM copy XSFormatCleaner.xla "Z:\AppFiles\SupportSetup\"


REM copy "C:\AppFiles\hostdb.properties" "C:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\hostdb.properties" "M:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\hostdb.properties" "G:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\hostdb.properties" "X:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\hostdb.properties" "I:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\hostdb.properties" "J:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\hostdb.properties" "U:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\hostdb.properties" "Z:\AppFiles\SupportSetup\"

REM copy "C:\AppFiles\WaitThenRunHiddenJob.vbs" "C:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\WaitThenRunHiddenJob.vbs" "M:\AppFiles\SupportSetup\"

REM copy "C:\AppFiles\WaitThenRunHiddenJob.vbs" "G:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\WaitThenRunHiddenJob.vbs" "X:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\WaitThenRunHiddenJob.vbs" "I:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\WaitThenRunHiddenJob.vbs" "J:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\WaitThenRunHiddenJob.vbs" "U:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\WaitThenRunHiddenJob.vbs" "Z:\AppFiles\SupportSetup\"

REM copy "C:\AppFiles\AutoCommShts.vbs" "C:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\AutoCommShts.vbs" "M:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\AutoCommShts.vbs" "G:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\AutoCommShts.vbs" "X:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\AutoCommShts.vbs" "I:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\AutoCommShts.vbs" "J:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\AutoCommShts.vbs" "U:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\AutoCommShts.vbs" "Z:\AppFiles\SupportSetup\"

REM copy "C:\AppFiles\RUNAS.vbs" "C:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\RUNAS.vbs" "M:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\RUNAS.vbs" "G:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\RUNAS.vbs" "X:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\RUNAS.vbs" "I:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\RUNAS.vbs" "J:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\RUNAS.vbs" "U:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\RUNAS.vbs" "Z:\AppFiles\SupportSetup\"

REM copy "C:\AppFiles\fei.properties" "C:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\fei.properties" "M:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\fei.properties" "G:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\fei.properties" "X:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\fei.properties" "I:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\fei.properties" "J:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\fei.properties" "U:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\fei.properties" "Z:\AppFiles\SupportSetup\"

REM copy "C:\AppFiles\terminal.properties" "C:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\terminal.properties" "M:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\terminal.properties" "G:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\terminal.properties" "X:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\terminal.properties" "I:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\terminal.properties" "J:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\terminal.properties" "U:\AppFiles\SupportSetup\"
REM copy "C:\AppFiles\terminal.properties" "Z:\AppFiles\SupportSetup\"

REM copy "C:\Overall Procedure And Guideline\Simple Version Control Script For All Party (Microsoft Word).vb" "C:\AppFiles\SupportSetup\"
REM copy "C:\Overall Procedure And Guideline\Simple Version Control Script For All Party (Microsoft Word).vb" "M:\AppFiles\SupportSetup\"
REM copy "C:\Overall Procedure And Guideline\Simple Version Control Script For All Party (Microsoft Word).vb" "G:\AppFiles\SupportSetup\"
REM copy "C:\Overall Procedure And Guideline\Simple Version Control Script For All Party (Microsoft Word).vb" "X:\AppFiles\SupportSetup\"
REM copy "C:\Overall Procedure And Guideline\Simple Version Control Script For All Party (Microsoft Word).vb" "I:\AppFiles\SupportSetup\"
REM copy "C:\Overall Procedure And Guideline\Simple Version Control Script For All Party (Microsoft Word).vb" "J:\AppFiles\SupportSetup\"
REM copy "C:\Overall Procedure And Guideline\Simple Version Control Script For All Party (Microsoft Word).vb" "U:\AppFiles\SupportSetup\"
REM copy "C:\Overall Procedure And Guideline\Simple Version Control Script For All Party (Microsoft Word).vb" "Z:\AppFiles\SupportSetup\"

REM copy Justacro_%username%.xlam "C:\AppFiles\Svc.xlam"

REM copy PresentationMacro_%username%.ppam "C:\AppFiles\PresentationMacro.ppam"

REM ..\RunHiddenJob.vbs "java -jar ..\SupportInstallationCommand.jar ""{SupportDelivery}"""

