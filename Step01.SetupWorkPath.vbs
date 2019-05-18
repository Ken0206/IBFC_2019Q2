If WScript.Arguments.Named.Exists("elevated") = False Then
  CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 1
  WScript.Quit
  Else
  'Change the working directory from the system32 folder back to the script's folder.
  Set oShell = CreateObject("WScript.Shell")
  oShell.CurrentDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
  'MsgBox "Now running with elevated permissions"
  body
End If

Function body

Dim strWorkPath, mypo
strBeginTime = ConvertDate(Now())

Set WshNetwork = WScript.CreateObject("WScript.Network")
 'WScript.Echo "Domain = " & WshNetwork.UserDomain
 'WScript.Echo "Computer Name = " & WshNetwork.ComputerName
 'WScript.Echo "User Name = " & WshNetwork.UserName

Dim objFSO 'File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objTS 'Text Stream Object

Set objTS = objFSO.OpenTextFile("logpath.txt")
strLogPath = objTS.ReadLine
objTS.Close

Set objTS = objFSO.OpenTextFile("workpath.txt")
If objTS.AtEndOfLine <> True Then
strWorkPath = objTS.ReadLine
End if
objTS.Close

mypo =  Replace(strWorkPath,Chr(34),"")
'WScript.Echo mypo 

If InStr(mypo,"txt") <> 0 Then
'If mypo <> "0" Then
 WScript.Echo "注意. workpath.txt內容錯誤. 請清空該檔後重新執行Step01"
 WScript.Quit(1) 
End If

Const ForWriting = 2
Set objTS = objFSO.OpenTextFile("workpath.txt", ForWriting, True)
strWorkPath = Chr(34) & strLogPath &"\" & WshNetwork.ComputerName & ".txt" & Chr(34)

'WScript.Echo "開始維護. 記錄檔路徑為" & strWorkPath
'strWorkPath = objFSO.BuildPath(objFSO.GetAbsolutePathName("."), WshNetwork.ComputerName & ".txt")

objTS.Write(strWorkPath)
objTS.Close()

strIBMid = (InputBox("請輸入分公司代碼(IB2~IB9)"))
Set objTS = objFSO.OpenTextFile("workpath.txt")
strWorkPath = Replace(objTS.ReadLine,Chr(34),"")
'WScript.Echo strworkpath
objTS.Close

Set objTS = objFSO.OpenTextFile(strWorkPath,ForWriting, True)
objTS.writeline("Maintain_BeginTime|" & strBeginTime)
objTS.writeline("Maintain_OperatorID|" & strIBMid)
objTS.Close()

WScript.Echo "開始維護檢核." & vbcrlf & "記錄檔路徑為 " & strWorkPath

End function
function ConvertDate(dat)
	
	if isdate(dat) then
		dattmp = year(dat) & "/" & right("0" & month(dat),2) & "/" & right("0" & day(dat),2) & " " & right("0" & hour(dat),2) & ":" & right("0" & minute(dat),2) & ":" & right("0" & second(dat),2)
	else
		datTmp = ""
	end if
	ConvertDate = datTmp
end function