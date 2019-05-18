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
		
		serviceName ="SepMasterService"		'sc query �d�� service name
		Set wmi = GetObject("winmgmts://./root/cimv2")
		state = wmi.Get("Win32_Service.Name='" & serviceName &"'").State
		If state="Running" Then
			f.writeline("Symantec Endpoint Protection service|Y")
			wscript.echo "Symantec Endpoint Protection ���w�˥B�A�Ȥw�Ұ�"
		else
			f.writeline("Symantec Endpoint Protection service|N")
			wscript.echo "Symantec Endpoint Protection ���w�˥B�A�ȥ��Ұ�"
		end if
		
		CreateObject("WScript.Shell").Run "files\02.png"
		WScript.Sleep 1000
		message = "�p��ܪ��d�ҹϤ��A���ˬd..." & vbcrlf & "Symantec Endpoint Protection �u�@�C�W��ܪ��ϥ�" & vbcrlf & "�O�_�G��O"
		If MsgBox(message, 4,"�T�{") = 6 Then
		   f.writeline("Symantec Endpoint Protection green|Y")
		Else
		   f.writeline("Symantec Endpoint Protection green|N")
		   wscript.echo "�гq���`���q��T�ǳ��H�e����" & vbcrlf & vbcrlf & "Symantec Endpoint Protection �S���G��O."
		End If
	  Else
		f.writeline("Symantec Endpoint Protection install|N")
		wscript.echo "�гq���`���q��T�ǳ��H�e����" & vbcrlf & vbcrlf & "Symantec Endpoint Protection �S���w��."
	  End If
	Next
end function

