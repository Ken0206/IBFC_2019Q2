strBeginTime = ConvertDate(Now())
intAns = msgbox("下列您應該要看到5個造字,若無法看到表示您的字碼更新有問題" & vbcrlf & vbcrlf & "�S�T�U�V�W" & vbcrlf & vbcrlf & "上列您應該要看到5個造字,若無法看到表示您的字碼更新有問題" & vbcrlf & "是否看到造字？", vbyesno)
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