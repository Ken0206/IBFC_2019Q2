On Error Resume Next

strBeginTime = ConvertDate(Now())
Set objFS = CreateObject("Scripting.FileSystemObject")
Set f = objFS.openTextFile("workpath.txt")
strWorkPath = Replace(f.ReadLine,Chr(34),"")
f.close
Set f = objFS.OpenTextFile(strWorkPath, 8,true)
f.writeline("SymantecEP_BeginTime|" & strBeginTime)

check_install

strBeginTime = ConvertDate(Now())
f.writeline("SymantecEP_EndTime|" & strBeginTime)
f.Close


function ConvertDate(dat)
  if isdate(dat) then
    dattmp = year(dat) & "/" & right("0" & month(dat),2) & "/" & right("0" & day(dat),2) & " " & right("0" & hour(dat),2) & ":" & right("0" & minute(dat),2) & ":" & right("0" & second(dat),2)
  else
    datTmp = ""
  end if
	ConvertDate = datTmp
end function

function check_install()
	strComputer = "."
	HKEY_LOCAL_MACHINE = &H80000002
	Dim objReg, buildDetailNames, buildDetailRegValNames
	buildDetailNames = Array("Display Name")
	buildDetailRegValNames = Array("DisplayName")
	Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
						 strComputer & "\root\default:StdRegProv")
	strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{C2CC34D6-103E-4065-97EA-021F03B43C29}"
	for I = 0 to UBound(buildDetailNames)
	  objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, buildDetailRegValNames(I), info
	  If info="Symantec Endpoint Protection" Then
		f.writeline("Symantec Endpoint Protection install|Y")
		
		serviceName ="SepMasterService"		'sc query 查詢 service name
		Set wmi = GetObject("winmgmts://./root/cimv2")
		state = wmi.Get("Win32_Service.Name='" & serviceName &"'").State
		If state="Running" Then
			f.writeline("Symantec Endpoint Protection service|Y")
			wscript.echo "Symantec Endpoint Protection 有安裝且服務已啟動"
		else
			f.writeline("Symantec Endpoint Protection service|N")
			wscript.echo "Symantec Endpoint Protection 有安裝且服務未啟動"
		end if
		
		CreateObject("WScript.Shell").Run "files\02.png"
		WScript.Sleep 1000
		message = "如顯示的範例圖片，請檢查..." & vbcrlf & "Symantec Endpoint Protection 工作列上顯示的圖示" & vbcrlf & "是否亮綠燈"
		If MsgBox(message, 4,"確認") = 6 Then
		   f.writeline("Symantec Endpoint Protection green|Y")
		Else
		   f.writeline("Symantec Endpoint Protection green|N")
		   wscript.echo "請通知總公司資訊室陳信呈先生" & vbcrlf & vbcrlf & "Symantec Endpoint Protection 沒有亮綠燈."
		End If
	  Else
		f.writeline("Symantec Endpoint Protection install|N")
		wscript.echo "請通知總公司資訊室陳信呈先生" & vbcrlf & vbcrlf & "Symantec Endpoint Protection 沒有安裝."
	  End If
	Next
end function

