strBeginTime = ConvertDate(Now())

WScript.Echo "ÀË®Ö§¹²¦" 

'Write Text
Set objFS = CreateObject("Scripting.FileSystemObject")
Set f = objFS.openTextFile("workpath.txt")
strWorkPath = Replace(f.ReadLine,chr(34),"")
f.close


Set f = objFS.OpenTextFile(strWorkPath, 8,true)

f.writeline("Maintain_EndTime|" & ConvertDate(now()))
f.Close
function ConvertDate(dat)
	
	if isdate(dat) then
		dattmp = year(dat) & "/" & right("0" & month(dat),2) & "/" & right("0" & day(dat),2) & " " & right("0" & hour(dat),2) & ":" & right("0" & minute(dat),2) & ":" & right("0" & second(dat),2)
	else
		datTmp = ""
	end if
	ConvertDate = datTmp
end function