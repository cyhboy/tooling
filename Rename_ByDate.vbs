


If Wscript.Arguments.Count = 0 Then
    arrFiles = Array(".")
	WScript.Quit
Else
    Dim arrFiles()
    For i = 0 to Wscript.Arguments.Count - 1
        Redim Preserve arrFiles(i)
        arrFiles(i) = Wscript.Arguments(i)
    Next
End If

	Dim ddFMT
	Dim dd

Dim Fso, fo
Set Fso = WScript.CreateObject("Scripting.FileSystemObject")

Dim sourceFile

For Each arrFile in arrFiles
    sourceFile = arrFile
	Set fo = Fso.GetFile(sourceFile)

	arrFile = Right(arrFile, Len(arrFile) - InstrRev(arrFile, "\"))
	arrFile = Left(arrFile, InstrRev(arrFile, ".") - 1)

	Dim todayStr
	'todayStr = CDATE(DATE)
	todayStr = CDATE(fo.DateLastModified)
	'MsgBox todayStr
	todayStr = Replace(todayStr, "-", "")
	todayStr = Replace(todayStr, " ", "")
	todayStr = Replace(todayStr, ":", "")
	todayStr = Left(todayStr, 8)

	targetFile = Replace(sourceFile, arrFile, arrFile & "_" & todayStr)

	Set fo = Nothing
	Fso.MoveFile sourceFile, targetFile

Next


Set Fso = Nothing

WScript.Quit

