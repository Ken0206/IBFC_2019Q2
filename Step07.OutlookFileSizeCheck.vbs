strBeginTime = ConvertDate(Now())

'Set objShell = CreateObject("Wscript.Shell")
CreateObject("WScript.Shell").Run "control.exe "
'strPath = "c:\"

'strPath = "explorer.exe /e," & strPath
'objShell.Run strPath
'WScript.Sleep 3000

intAns = InputBox("�п�JOutlook�ɮ�SIZE:(MB)" & vbcrlf & vbcrlf & vbcrlf ,"Outlook mail PST",,100,100)

'Write Text
Set objFS = CreateObject("Scripting.FileSystemObject")
Set f = objFS.openTextFile("workpath.txt")
strWorkPath = Replace(f.ReadLine,Chr(34),"")
f.close
Set f = objFS.OpenTextFile(strWorkPath, 8,true)

msgbox("�Y�l���ɶW�L 4GB" & vbcrlf & "�Ш̤��Ѥ� Outlook�s�Wpst�ɳ]�w�B�J�s�Wpst��" & vbcrlf & "�Y user �ڵ�, ���p���`���q��T�ǳ��H�e����")

f.writeline("Outlook_BeginTime|" & strBeginTime)
f.writeline("Outlook_MailSize|" & intAns)
f.writeline("Outlook_TooLargeMailSize|")
f.writeline("Outlook_EndTime|" & ConvertDate(now()))
f.Close
function ConvertDate(dat)
	
	if isdate(dat) then
		dattmp = year(dat) & "/" & right("0" & month(dat),2) & "/" & right("0" & day(dat),2) & " " & right("0" & hour(dat),2) & ":" & right("0" & minute(dat),2) & ":" & right("0" & second(dat),2)
	else
		datTmp = ""
	end if
	ConvertDate = datTmp
end Function
Function FormatCapacity(SizeInBytes)
    If SizeInBytes >= 1073741824 Then
        FormatCapacity = abs(SizeInBytes/1073741824) & "GB"
    Else
        FormatCapacity = abs(SizeInBytes/1048576) & "MB"
    End If
End Function