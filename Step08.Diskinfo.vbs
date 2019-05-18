Set objShell = CreateObject("WScript.Shell") 
objShell.Run """Diskinfo\diskinfo.exe""" 
Set objShell = Nothing
strBeginTime = ConvertDate(Now())

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colDiskDrives = objWMIService.ExecQuery _    
    ("Select * from Win32_DiskDrive")

For each objDiskDrive in colDiskDrives
 
    For i = Lbound(objDiskDrive.Capabilities) to _
        Ubound(objDiskDrive.Capabilities)
        
    Next    
   'Wscript.Echo "Caption: " & vbTab &  objDiskDrive.Caption
   strModel = objDiskDrive.Model
   strSize = objDiskDrive.Size
   strSerial = objDiskDrive.SerialNumber
Next

Set objShell = CreateObject("WScript.Shell") 
objShell.Run """Diskinfo\diskinfo.exe""" 
Set objShell = Nothing
strBeginTime = ConvertDate(Now())

intAns = InputBox("請填入硬碟狀態" & vbcrlf & "1.良好" & vbcrlf & "2.警告" & vbcrlf & "3.錯誤","Disk Info Status","1",100,100)

Select Case intAns
Case "1"
   strCheck = "Diskinfo_HealthStatus|良好"
Case "2"
   strCheck = "Diskinfo_HealthStatus|警告"
Case "3"
   strCheck = "Diskinfo_HealthStatus|錯誤"
End Select
'Write Text
Set objFS = CreateObject("Scripting.FileSystemObject")
Set f = objFS.openTextFile("workpath.txt")
strWorkPath = Replace(f.ReadLine,chr(34),"")
f.close
Set f = objFS.OpenTextFile(strWorkPath, 8,true)

f.writeline("Diskinfo_BeginTime|" & strBeginTime)
f.writeline(strCheck)
f.writeline("Diskinfo_HDDManufacturer|" & strModel)
f.writeline("Diskinfo_HDDSize|" & FormatCapacity(strSize))
f.writeline("Diskinfo_HDDSerial|" & strSerialNumber)
f.writeline("Diskinfo_EndTime|" & ConvertDate(now()))
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