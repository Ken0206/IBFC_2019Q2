strBeginTime = ConvertDate(Now())
Dim WshNetwork, objFSO, objTS
Set objFSO = CreateObject("Scripting.FileSystemObject")
If not objFSO.FileExists("logpath.txt") Then
  MsgBox "logpath.txt 不存在", 0 + 32,"錯誤"
  WScript.Quit
End If

message = "執行此步驟將清空報告檔，重建新資料" & vbcrlf & vbcrlf & "確定要執行此步驟？"
If MsgBox(message, 4 + 48,"確認") = 6 Then
  Maintain
Else
  WScript.Quit
End If

Sub Maintain()
  Set WshNetwork = WScript.CreateObject("WScript.Network")
  Set objTS = objFSO.OpenTextFile("logpath.txt")
  strLogPath = objTS.ReadLine
  objTS.Close
  
  BuildFullPath strLogPath
 
  strWorkPath = Chr(34) & strLogPath &"\" & WshNetwork.ComputerName & ".txt" & Chr(34)
  OpenWorkPath = strLogPath &"\" & WshNetwork.ComputerName & ".txt"
  Set objTS = objFSO.OpenTextFile("workpath.txt", 2, True)
  objTS.writeline(strWorkPath)
  objTS.Close
  
  strIBMid = (InputBox("請輸入分公司代碼(IB2~IB9)", "分公司代碼"))
  'Set objTS = objFSO.OpenTextFile("workpath.txt")
  strWorkPath = Replace(strWorkPath,Chr(34),"")
  'objTS.Close
  
  Set objTS = objFSO.OpenTextFile(strWorkPath, 2, True)
  objTS.writeline("Maintain_BeginTime|" & strBeginTime)
  objTS.writeline("Maintain_OperatorID|" & strIBMid)
  objTS.Close()
  
  message = "開始維護檢核." & vbcrlf & vbcrlf & "記錄檔路徑為 " & strWorkPath
  MsgBox message, 0 + 64,"完成"

end Sub

Sub BuildFullPath(ByVal FullPath)
  If Not objFSO.FolderExists(FullPath) Then
    BuildFullPath objFSO.GetParentFolderName(FullPath)
    objFSO.CreateFolder FullPath
  End If
End Sub

function ConvertDate(dat)
	if isdate(dat) then
		dattmp = year(dat) & "/" & right("0" & month(dat),2) & "/" & right("0" & day(dat),2) & " " & right("0" & hour(dat),2) & ":" & right("0" & minute(dat),2) & ":" & right("0" & second(dat),2)
	else
		datTmp = ""
	end if
	ConvertDate = datTmp
end function