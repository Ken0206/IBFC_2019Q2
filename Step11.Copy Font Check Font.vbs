strBeginTime = ConvertDate(Now())
intAns = msgbox("¤U¦C±zÀ³¸Ó­n¬Ý¨ì5­Ó³y¦r,­YµLªk¬Ý¨ìªí¥Ü±zªº¦r½X§ó·s¦³°ÝÃD" & vbcrlf & vbcrlf & "–S–T–U–V–W" & vbcrlf & vbcrlf & "¤W¦C±zÀ³¸Ó­n¬Ý¨ì5­Ó³y¦r,­YµLªk¬Ý¨ìªí¥Ü±zªº¦r½X§ó·s¦³°ÝÃD" & vbcrlf & "¬O§_¬Ý¨ì³y¦r¡H", vbyesno)
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