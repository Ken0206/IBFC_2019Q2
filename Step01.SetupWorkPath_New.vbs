strBeginTime = ConvertDate(Now())

Dim WshShell, WshNetwork, strWorkPath
Set WshShell = WScript.CreateObject("WScript.Shell")
Set WshNetwork = WScript.CreateObject("WScript.Network")
strWorkPath = Chr(34) & WshShell.CurrentDirectory & "\" & WshNetwork.ComputerName & ".txt" & Chr(34)

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set outputFile = objFSO.OpenTextFile("workpath.txt", 2, True)
outputFile.writeline(strWorkPath)
outputFile.Close

strIBMid = (InputBox("請輸入分公司代碼(IB2~IB9)"))
Set ReadFile = objFSO.OpenTextFile("workpath.txt")
strWorkPath = Replace(ReadFile.ReadLine,Chr(34),"")
ReadFile.Close

Set outputFile = objFSO.OpenTextFile(strWorkPath, 2, True)
outputFile.writeline("Maintain_BeginTime|" & strBeginTime)
outputFile.writeline("Maintain_OperatorID|" & strIBMid)
outputFile.Close()

WScript.Echo "開始維護檢核." & vbcrlf & "記錄檔路徑為 " & strWorkPath

function ConvertDate(dat)
	if isdate(dat) then
		dattmp = year(dat) & "/" & right("0" & month(dat),2) & "/" & right("0" & day(dat),2) & " " & right("0" & hour(dat),2) & ":" & right("0" & minute(dat),2) & ":" & right("0" & second(dat),2)
	else
		datTmp = ""
	end if
	ConvertDate = datTmp
end function