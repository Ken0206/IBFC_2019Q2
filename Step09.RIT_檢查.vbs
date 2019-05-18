On Error Resume Next
set WshShell = WScript.CreateObject("WScript.Shell")
rfile = String(1,34) & "C:\Program Files (x86)\RestoreIT 10\IBP\RestoreIT.exe" & String(1,34)
WshShell.Run rfile
rfile = String(1,34) & "C:\Program Files (x86)\RestoreIT 2014\IBP\RestoreIT.exe" & String(1,34)
WshShell.Run rfile
WScript.Sleep 2500
function ConvertDate(dat)
  if isdate(dat) then
    dattmp = year(dat) & "/" & right("0" & month(dat),2) & "/" & right("0" & day(dat),2) & " " & right("0" & hour(dat),2) & ":" & right("0" & minute(dat),2) & ":" & right("0" & second(dat),2)
  else
    datTmp = ""
  end if
	ConvertDate = datTmp
end function
Set objFS = CreateObject("Scripting.FileSystemObject")
Set f = objFS.openTextFile("workpath.txt")
strWorkPath = Replace(f.ReadLine,Chr(34),"")
f.close
Set f = objFS.OpenTextFile(strWorkPath, 8, true)
f.writeline("RIT_snapshots_BeginTime|" & ConvertDate(Now()))
Answer = InputBox("請紀錄 RestoreIT 快照日期時間" & vbcrlf & vbcrlf & _
				  "例如︰" & vbcrlf & _
				  "2018/11/05 14：30：12", "RestoreIT snapshots")
f.writeline("RIT_snapshots|" & Answer)
f.writeline("RIT_snapshots_EndTime|" & ConvertDate(now()))
f.Close
