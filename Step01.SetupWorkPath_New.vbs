message = "執行此步驟將清空報告檔，重建新資料" & vbcrlf & vbcrlf & "確定要執行此步驟？"
If MsgBox(message, 4 + 48,"確認") = 6 Then
   Maintain
Else
   WScript.Quit
End If

Sub Maintain()
	strBeginTime = ConvertDate(Now())
	Dim WshShell, WshNetwork, strWorkPath
	Set WshShell = WScript.CreateObject("WScript.Shell")
	Set WshNetwork = WScript.CreateObject("WScript.Network")
	strWorkPath = Chr(34) & WshShell.CurrentDirectory & "\" & WshNetwork.ComputerName & ".txt" & Chr(34)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set outputFile = objFSO.OpenTextFile("workpath.txt", 2, True)
	outputFile.writeline(strWorkPath)
	outputFile.Close
	strIBMid = (InputBox("請輸入分公司代碼(IB2~IB9)", "分公司代碼"))
	Set ReadFile = objFSO.OpenTextFile("workpath.txt")
	strWorkPath = Replace(ReadFile.ReadLine,Chr(34),"")
	ReadFile.Close
	Set outputFile = objFSO.OpenTextFile(strWorkPath, 2, True)
	outputFile.writeline("Maintain_BeginTime|" & strBeginTime)
	outputFile.writeline("Maintain_OperatorID|" & strIBMid)
	outputFile.Close()
	message = "開始維護檢核." & vbcrlf & vbcrlf & "記錄檔路徑為 " & strWorkPath
	MsgBox message, 0 + 64,"完成"
end Sub

function ConvertDate(dat)
	if isdate(dat) then
		dattmp = year(dat) & "/" & right("0" & month(dat),2) & "/" & right("0" & day(dat),2) & " " & right("0" & hour(dat),2) & ":" & right("0" & minute(dat),2) & ":" & right("0" & second(dat),2)
	else
		datTmp = ""
	end if
	ConvertDate = datTmp
end function