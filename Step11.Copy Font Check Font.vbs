strBeginTime = ConvertDate(Now())
intAns = msgbox("�U�C�z���ӭn�ݨ�5�ӳy�r,�Y�L�k�ݨ��ܱz���r�X��s�����D" & vbcrlf & vbcrlf & "�S�T�U�V�W" & vbcrlf & vbcrlf & "�W�C�z���ӭn�ݨ�5�ӳy�r,�Y�L�k�ݨ��ܱz���r�X��s�����D" & vbcrlf & "�O�_�ݨ�y�r�H", vbyesno)
if intAns = vbyes then
	strCheck = "CheckCreatedCharacters_States|Y"
else
	strCheck = "CheckCreatedCharacters_States|N"	
end if

'Write Text
Set objFS = CreateObject("Scripting.FileSystemObject")
Set f = objFS.openTextFile("workpath.txt")
strWorkPath = Replace(f.ReadLine,chr(34),"")
f.close
Set f = objFS.OpenTextFile(strWorkPath, 8,true)

f.writeline("CheckCreatedCharacters_BeginTime|" & strBeginTime)
f.writeline(strCheck )
f.writeline("CheckCreatedCharacters_EndTime|" & ConvertDate(now()))
f.Close
function ConvertDate(dat)
	
	if isdate(dat) then
		dattmp = year(dat) & "/" & right("0" & month(dat),2) & "/" & right("0" & day(dat),2) & " " & right("0" & hour(dat),2) & ":" & right("0" & minute(dat),2) & ":" & right("0" & second(dat),2)
	else
		datTmp = ""
	end if
	ConvertDate = datTmp
end function