Dim WshShell
Set objFSO = CreateObject("Scripting.FileSystemObject")
set WshShell = WScript.CreateObject("WScript.Shell")
check_flag = 1
strBeginTime = ConvertDate(Now())

function ConvertDate(dat)
	if isdate(dat) then
		dattmp = year(dat) & "/" & right("0" & month(dat),2) & "/" & right("0" & day(dat),2) & " " & right("0" & hour(dat),2) & ":" & right("0" & minute(dat),2) & ":" & right("0" & second(dat),2)
	else
		datTmp = ""
	end if
	ConvertDate = datTmp
end function

'Write Text
Set f = objFSO.openTextFile("workpath.txt")
strWorkPath = Replace(f.ReadLine,chr(34),"")
f.close
Set f = objFSO.OpenTextFile(strWorkPath, 8,true)
f.writeline("OutlookConfig_BeginTime|" & strBeginTime)

'outlook 365 32-bit
rfile = "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"
dfile = String(1,34) & rfile & String(1,34)
If objFSO.FileExists(rfile) Then
  WshShell.RegWrite "HKCU\software\policies\microsoft\office\16.0\outlook\options\mail\blockextcontent", "00000001", "REG_DWORD"
  WshShell.RegWrite "HKCU\software\policies\microsoft\office\16.0\outlook\options\mail\readasplain", "00000001", "REG_DWORD"
  WshShell.RegWrite "HKCU\software\policies\microsoft\office\16.0\outlook\options\disablereadingpane", "00000001", "REG_DWORD"
  wscript.echo "Outlook 365 32-bit 設定完成！"
  check_flag = 0
  f.writeline("Outlook 365 32-bit|Y")
End If

'outlook 2016 32-bit
rfile = "C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE"
dfile = String(1,34) & rfile & String(1,34)
If objFSO.FileExists(rfile) Then
  WshShell.RegWrite "HKCU\software\policies\microsoft\office\16.0\outlook\options\mail\blockextcontent", "00000001", "REG_DWORD"
  WshShell.RegWrite "HKCU\software\policies\microsoft\office\16.0\outlook\options\mail\readasplain", "00000001", "REG_DWORD"
  WshShell.RegWrite "HKCU\software\policies\microsoft\office\16.0\outlook\options\disablereadingpane", "00000001", "REG_DWORD"
  wscript.echo "Outlook 2016 32-bit 設定完成！"
  check_flag = 0
  f.writeline("Outlook 2016 32-bit|Y")
End If

'outlook 2016 64-bit
rfile = "C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE"
dfile = String(1,34) & rfile & String(1,34)
If objFSO.FileExists(rfile) Then
  WshShell.RegWrite "HKCU\software\policies\microsoft\office\16.0\outlook\options\mail\blockextcontent", "00000001", "REG_DWORD"
  WshShell.RegWrite "HKCU\software\policies\microsoft\office\16.0\outlook\options\mail\readasplain", "00000001", "REG_DWORD"
  WshShell.RegWrite "HKCU\software\policies\microsoft\office\16.0\outlook\options\disablereadingpane", "00000001", "REG_DWORD"
  wscript.echo "Outlook 2016 64-bit 設定完成！"
  check_flag = 0
  f.writeline("Outlook 2016 64-bit|Y")
End If

'outlook 2010 32-bit
rfile = "C:\Program Files (x86)\Microsoft Office\Office14\OUTLOOK.EXE"
dfile = String(1,34) & rfile & String(1,34)
If objFSO.FileExists(rfile) Then
  WshShell.RegWrite "HKCU\software\policies\microsoft\office\14.0\outlook\options\mail\blockextcontent", "00000001", "REG_DWORD"
  WshShell.RegWrite "HKCU\software\policies\microsoft\office\14.0\outlook\options\mail\readasplain", "00000001", "REG_DWORD"
  WshShell.RegWrite "HKCU\software\policies\microsoft\office\14.0\outlook\options\disablereadingpane", "00000001", "REG_DWORD"
  wscript.echo "Outlook 2010 32-bit 設定完成！"
  check_flag = 0
  f.writeline("Outlook 2010 32-bit|Y")
End If

'outlook 2010 64-bit
rfile = "C:\Program Files\Microsoft Office\Office14\OUTLOOK.EXE"
dfile = String(1,34) & rfile & String(1,34)
If objFSO.FileExists(rfile) Then
  WshShell.RegWrite "HKCU\software\policies\microsoft\office\14.0\outlook\options\mail\blockextcontent", "00000001", "REG_DWORD"
  WshShell.RegWrite "HKCU\software\policies\microsoft\office\14.0\outlook\options\mail\readasplain", "00000001", "REG_DWORD"
  WshShell.RegWrite "HKCU\software\policies\microsoft\office\14.0\outlook\options\disablereadingpane", "00000001", "REG_DWORD"
  wscript.echo "Outlook 2010 64-bit 設定完成！"
  check_flag = 0
  f.writeline("Outlook 2010 64-bit|Y")
End If

'outlook 2007
rfile = "C:\Program Files (x86)\Microsoft Office\Office12\OUTLOOK.EXE"
dfile = String(1,34) & rfile & String(1,34)
If objFSO.FileExists(rfile) Then
  WshShell.RegWrite "HKCU\software\policies\microsoft\office\12.0\outlook\options\mail\blockextcontent", "00000001", "REG_DWORD"
  WshShell.RegWrite "HKCU\software\policies\microsoft\office\12.0\outlook\options\mail\readasplain", "00000001", "REG_DWORD"
  WshShell.RegWrite "HKCU\software\policies\microsoft\office\12.0\outlook\options\disablereadingpane", "00000001", "REG_DWORD"
  wscript.echo "Outlook 2007 設定完成！"
  check_flag = 0
  f.writeline("Outlook 2007|Y")
End If

If check_flag = 1 Then
  wscript.echo "找不到 Outlook ！"
  f.writeline("找不到 Outlook")
End If

f.writeline("OutlookConfig_EndTime|" & ConvertDate(now()))
f.Close
