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

'outlook 2016 32-bit
rfile = "C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE"
dfile = String(1,34) & rfile & String(1,34)
If objFSO.FileExists(rfile) Then
  wscript.echo "將自動開啟 Outlook 2016 32-bit ，" & vbcrlf & "在設定完成前請勿操作電腦！" & vbcrlf & vbcrlf & "設定完成後會自動關閉 Outlook！"
  WshShell.RegWrite "HKCU\Software\Microsoft\Office\16.0\Outlook\Options\Mail\BlockExtContent", "00000001", "REG_DWORD"
  WshShell.RegWrite "HKCU\Software\Microsoft\Office\16.0\Outlook\Options\Mail\ReadAsPlain", "00000001", "REG_DWORD"
  WScript.Sleep 500
  WshShell.Run dfile
  WScript.Sleep 5000
  WshShell.SendKeys "{TAB 8}"
  WScript.Sleep 500
  WshShell.SendKeys "{DOWN}~"
  WScript.Sleep 500
  WshShell.SendKeys "%v"
  WScript.Sleep 500
  WshShell.SendKeys "plo"
  WScript.Sleep 500
  WshShell.SendKeys "%a"
  WScript.Sleep 500
  WshShell.SendKeys "%v"
  WScript.Sleep 500
  WshShell.SendKeys "cv{HOME}~"
  WScript.Sleep 500
  WshShell.SendKeys "%v"
  WScript.Sleep 500
  WshShell.SendKeys "pno"
  WScript.Sleep 500
  WshShell.SendKeys "%{F4}"
  wscript.echo "Outlook 2016 32-bit 設定完成！"
  check_flag = 0
  f.writeline("Outlook 2016 32-bit|Y")
End If

'outlook 2016 64-bit
rfile = "C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE"
dfile = String(1,34) & rfile & String(1,34)
If objFSO.FileExists(rfile) Then
  wscript.echo "將自動開啟 Outlook 2016 64-bit ，" & vbcrlf & "在設定完成前請勿操作電腦！" & vbcrlf & vbcrlf & "設定完成後會自動關閉 Outlook！"
  WshShell.RegWrite "HKCU\Software\Microsoft\Office\16.0\Outlook\Options\Mail\BlockExtContent", "00000001", "REG_DWORD"
  WshShell.RegWrite "HKCU\Software\Microsoft\Office\16.0\Outlook\Options\Mail\ReadAsPlain", "00000001", "REG_DWORD"
  WScript.Sleep 500
  WshShell.Run dfile
  WScript.Sleep 5000
  WshShell.SendKeys "{TAB 8}"
  WScript.Sleep 500
  WshShell.SendKeys "{DOWN}~"
  WScript.Sleep 500
  WshShell.SendKeys "%v"
  WScript.Sleep 500
  WshShell.SendKeys "plo"
  WScript.Sleep 500
  WshShell.SendKeys "%a"
  WScript.Sleep 500
  WshShell.SendKeys "%v"
  WScript.Sleep 500
  WshShell.SendKeys "cv{HOME}~"
  WScript.Sleep 500
  WshShell.SendKeys "%v"
  WScript.Sleep 500
  WshShell.SendKeys "pno"
  WScript.Sleep 500
  WshShell.SendKeys "%{F4}"
  wscript.echo "Outlook 2016 64-bit 設定完成！"
  check_flag = 0
  f.writeline("Outlook 2016 64-bit|Y")
End If

'outlook 2010 32-bit
rfile = "C:\Program Files (x86)\Microsoft Office\Office14\OUTLOOK.EXE"
dfile = String(1,34) & rfile & String(1,34)
If objFSO.FileExists(rfile) Then
  wscript.echo "將自動開啟 Outlook 2010 32-bit ，" & vbcrlf & "在設定完成前請勿操作電腦！" & vbcrlf & vbcrlf & "設定完成後會自動關閉 Outlook！"
  WshShell.RegWrite "HKCU\Software\Microsoft\Office\14.0\Outlook\Options\Mail\BlockExtContent", "00000001", "REG_DWORD"
  WshShell.RegWrite "HKCU\Software\Microsoft\Office\14.0\Outlook\Options\Mail\ReadAsPlain", "00000001", "REG_DWORD"
  WScript.Sleep 500
  WshShell.Run dfile
  WScript.Sleep 5000
  WshShell.SendKeys "{TAB 8}"
  WScript.Sleep 500
  WshShell.SendKeys "{DOWN}~"
  WScript.Sleep 500
  WshShell.SendKeys "%v"
  WScript.Sleep 500
  WshShell.SendKeys "cv{HOME}~"
  WScript.Sleep 500
  WshShell.SendKeys "%v"
  WScript.Sleep 500
  WshShell.SendKeys "pno"
  WScript.Sleep 500
  WshShell.SendKeys "%{F4}"
  wscript.echo "Outlook 2010 32-bit 設定完成！"
  check_flag = 0
  f.writeline("Outlook 2010 32-bit|Y")
End If

'outlook 2010 64-bit
rfile = "C:\Program Files\Microsoft Office\Office14\OUTLOOK.EXE"
dfile = String(1,34) & rfile & String(1,34)
If objFSO.FileExists(rfile) Then
  wscript.echo "將自動開啟 Outlook 2010 64-bit ，" & vbcrlf & "在設定完成前請勿操作電腦！" & vbcrlf & vbcrlf & "設定完成後會自動關閉 Outlook！"
  WshShell.RegWrite "HKCU\Software\Microsoft\Office\14.0\Outlook\Options\Mail\BlockExtContent", "00000001", "REG_DWORD"
  WshShell.RegWrite "HKCU\Software\Microsoft\Office\14.0\Outlook\Options\Mail\ReadAsPlain", "00000001", "REG_DWORD"
  WScript.Sleep 500
  WshShell.Run dfile
  WScript.Sleep 5000
  WshShell.SendKeys "{TAB 8}"
  WScript.Sleep 500
  WshShell.SendKeys "{DOWN}~"
  WScript.Sleep 500
  WshShell.SendKeys "%v"
  WScript.Sleep 500
  WshShell.SendKeys "cv{HOME}~"
  WScript.Sleep 500
  WshShell.SendKeys "%v"
  WScript.Sleep 500
  WshShell.SendKeys "pno"
  WScript.Sleep 500
  WshShell.SendKeys "%{F4}"
  wscript.echo "Outlook 2010 64-bit 設定完成！"
  check_flag = 0
  f.writeline("Outlook 2010 64-bit|Y")
End If

'outlook 2007
rfile = "C:\Program Files (x86)\Microsoft Office\Office12\OUTLOOK.EXE"
dfile = String(1,34) & rfile & String(1,34)
If objFSO.FileExists(rfile) Then
  wscript.echo "將自動開啟 Outlook 2007 ，" & vbcrlf & "在設定完成前請勿操作電腦！" & vbcrlf & vbcrlf & "設定完成後會自動關閉 Outlook！"
  WshShell.RegWrite "HKCU\Software\Microsoft\Office\12.0\Outlook\Options\Mail\BlockExtContent", "00000001", "REG_DWORD"
  WshShell.RegWrite "HKCU\Software\Microsoft\Office\12.0\Outlook\Options\Mail\ReadAsPlain", "00000001", "REG_DWORD"
  WScript.Sleep 500
  WshShell.Run dfile
  WScript.Sleep 5000
  WshShell.SendKeys "{TAB 8}"
  WScript.Sleep 500
  WshShell.SendKeys "{DOWN 2}~"
  WScript.Sleep 500
  WshShell.SendKeys "%v"
  WScript.Sleep 500
  WshShell.SendKeys "ro"
  WScript.Sleep 500
  WshShell.SendKeys "%{F4}"
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
