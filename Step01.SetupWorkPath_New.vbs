strBeginTime = ConvertDate(Now())
Dim WshNetwork, objFSO, objTS
Set objFSO = CreateObject("Scripting.FileSystemObject")
If not objFSO.FileExists("logpath.txt") Then
  MsgBox "logpath.txt ���s�b", 0 + 32,"���~"
  WScript.Quit
End If

message = "���榹�B�J�N�M�ų��i�ɡA���طs���" & vbcrlf & vbcrlf & "�T�w�n���榹�B�J�H"
If MsgBox(message, 4 + 48,"�T�{") = 6 Then
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
  
  strIBMid = (InputBox("�п�J�����q�N�X(IB2~IB9)", "�����q�N�X"))
  'Set objTS = objFSO.OpenTextFile("workpath.txt")
  strWorkPath = Replace(strWorkPath,Chr(34),"")
  'objTS.Close
  
  Set objTS = objFSO.OpenTextFile(strWorkPath, 2, True)
  objTS.writeline("Maintain_BeginTime|" & strBeginTime)
  objTS.writeline("Maintain_OperatorID|" & strIBMid)
  objTS.Close()
  
  message = "�}�l���@�ˮ�." & vbcrlf & vbcrlf & "�O���ɸ��|�� " & strWorkPath
  MsgBox message, 0 + 64,"����"

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