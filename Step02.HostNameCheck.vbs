If WScript.Arguments.Named.Exists("elevated") = False Then
  CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 1
  WScript.Quit
  Else
  'Change the working directory from the system32 folder back to the script's folder.
  Set oShell = CreateObject("WScript.Shell")
  oShell.CurrentDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
  'MsgBox "Now running with elevated permissions"
  body
End If

Function body
strBeginTime = ConvertDate(Now())
Dim objFSO 'File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objTS 'Text Stream Object
Const ForAppending = 8
Set objTS = objFSO.OpenTextFile("workpath.txt")
strWorkPath = Replace(objTS.ReadLine,Chr(34),"")
objTS.Close()

Set objTS = objFSO.OpenTextFile(strWorkPath, ForAppending, True)
objTS.writeline("Host_BeginTime|" & strBeginTime)

' List Computer Manufacturer and Model
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
For Each objItem In colItems
  'WScript.Echo "Name: " & objItem.Name
  objTS.writeline("Host_ComputerName|" & objItem.Name)
  objTS.writeline("Host_ComputerManufacturer|" & objItem.Manufacturer)
  objTS.writeline("Host_ComputerModel|" & objItem.Model)
  objTS.writeline("Host_ComputerWorkgroup|" & objItem.Workgroup)
   strModel = objItem.Model
  ' WScript.Echo "Workgroup: " & objItem.Workgroup 'post-Windows 2000 only
    
Next

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_PhysicalMemoryArray")
For Each objItem in colItems
    'Wscript.Echo "Maximum Capacity: " & FormatCapacity(objItem.MaxCapacity)
    objTS.writeline("Host_ComputerMemory|" & FormatCapacity(objItem.MaxCapacity))
Next

Set colItems = objWMIService.ExecQuery("Select * from Win32_BaseBoard")
For Each objItem in colItems
    objTS.writeline("Host_ComputerSerial| " & strModel & " SN:" & objItem.SerialNumber)
    objTS.WriteLine("Host_BoardManufacturer|" & objItem.Manufacturer)
    objTS.writeline("Host_BoardModel|" & objItem.Model)
    objTS.writeline("Host_BoradProduct|" & objItem.Product)
    objTS.writeline("Host_BoradVersion|" & objItem.Version)
Next

Set colProcessors = objWMIService.ExecQuery("Select * from Win32_Processor")
For Each objProcessor in colProcessors
  objTS.writeline("Host_CPUName|" & objProcessor.Name)
Next

' List IP Addresses for a Computer
Set IPConfigSet = objWMIService.ExecQuery _
    ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE and DHCPEnabled=TRUE")
For Each IPConfig in IPConfigSet
    If Not IsNull(IPConfig.IPAddress) Then 
         objTS.writeline("Host_Lan_CardName|" & IPConfig.Description)
         objTS.writeline("Host_MACAddress|" & IPConfig.MACAddress)
         For i=LBound(IPConfig.IPAddress) to UBound(IPConfig.IPAddress)
             If Not Instr(IPConfig.IPAddress(i), ":") > 0 Then
              strIPaddress = IPConfig.IPAddress(i)
           ' strMsg = strMsg & IPConfig.IPAddress(i) & vbcrlf
             End If
         Next
         objTS.WriteLine("Host_IPAddress|" & strIPaddress)
    End If
Next

Set colOSes = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
For Each objOS in colOSes
  objTS.writeline("Host_OS|" & objOS.Caption) 'OS Name
  objTS.writeline("Host_OS_InstallDate|" & ConvertDateOS(objOS.InstallDate)) 

Next

' List Available Disk Space
Const HARD_DISK = 3
strComputer = "."

Set colDisks = objWMIService.ExecQuery _
    ("Select * from Win32_LogicalDisk Where DriveType = " & HARD_DISK & "")
For Each objDisk in colDisks
    'Wscript.Echo "DeviceID: "& vbTab &  objDisk.DeviceID 
    sizeGB = FormatNumber(objDisk.FreeSpace /(1024^3), 3 )
    sizeGB = left(sizeGB, len(sizeGB) -1 )
   objTS.writeline("Host_Disk"& objdisk.DeviceID &"Free_Space: "& vbTab & sizeGB & " GB")      
Next

objTS.writeline("Host_EndTime|" & ConvertDate(now()))
objTS.Close()
WScript.Echo "Step02 °õ¦æ§¹²¦"
End Function

function ConvertDate(dat)
	
	if isdate(dat) then
		dattmp = year(dat) & "/" & right("0" & month(dat),2) & "/" & right("0" & day(dat),2) & " " & right("0" & hour(dat),2) & ":" & right("0" & minute(dat),2) & ":" & right("0" & second(dat),2)
	else
		datTmp = ""
	end if
	ConvertDate = datTmp
End Function


Function FormatCapacity(SizeInBytes)
    If SizeInBytes >= 1073741824 Then
        FormatCapacity = abs(SizeInBytes/1073741824) & "TB"
    Else
        FormatCapacity = abs(SizeInBytes/1048576) & "GB"
    End If
End Function

function ConvertDateOS(dat)
	    
		dattmp = Left(dat,4) & "/" & Mid(dat,5 ,2) & "/" & Mid(dat,7,2) & " " 
	ConvertDateOS = datTmp
End Function