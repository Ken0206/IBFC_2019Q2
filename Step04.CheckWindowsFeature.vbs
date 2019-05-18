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
If objTS.AtEndOfLine <> True Then
   strWorkPath = Replace(objTS.ReadLine,Chr(34),"")
End if
objTS.Close()

Set objTS = objFSO.OpenTextFile(strWorkPath, ForAppending, True)
objTS.writeline("CheckFeature_BeginTime|" & strBeginTime)

On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20
DIM wshShell : Set wshShell = CreateObject("WScript.shell")

strComputer = "." 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OptionalFeature", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

For Each objItem In colItems
    'WScript.Echo "Caption: " & objItem.Caption
    'objTS.writeline("Caption: " & objItem.Caption)
    'WScript.Echo "Description: " & objItem.Description
    'objTS.writeline("Description: " & objItem.Description)
    'WScript.Echo "InstallDate: " & WMIDateStringToDate(objItem.InstallDate)
    'WScript.Echo "InstallState: " & objItem.InstallState
    'objTS.writeline("InstallState: " & objItem.InstallState)
    'WScript.Echo "Name: " & objItem.Name
    'objTS.writeline("Name: " & objItem.Name)
    Select Case objItem.Name
    Case "IIS-FTPSvc"
      If objItem.InstallState = "2" then
       objTS.writeline("CheckFeature_IIS-FTPSvc|N")
      Else
       objTS.writeline("CheckFeature_IIS-FTPSvc|Y")
      End If
    Case "IIS-FTPExtensibility"
      If objItem.InstallState = "2" then
       objTS.writeline("CheckFeature_IIS-FTPExtensibility|N")
      Else
       objTS.writeline("CheckFeature_IIS-FTPExtensibility|Y")
      End If
    Case "IIS-ManagementConsole"
      If objItem.InstallState = "2" then
       objTS.writeline("CheckFeature_IIS-ManagementConsole|N")
      Else
       objTS.writeline("CheckFeature_IIS-ManagementConsole|Y")
      End If
    Case "TelnetClient"
      If objItem.InstallState = "2" then
       objTS.writeline("CheckFeature_TelnetClient|N")
        WScript.Echo "偵測到必要元件未安裝.按確定後將自動安裝. 請稍候..." 
        EnableDISM("TelnetClient")
      Else
       objTS.writeline("CheckFeature_TelnetClient|Y")
      End If
    Case "SimpleTCP"
      If objItem.InstallState = "2" then
       objTS.writeline("CheckFeature_SimpleTCP|N")
       WScript.Echo "偵測到必要元件未安裝.按確定後將自動安裝. 請稍候..." 
       EnableDISM("SimpleTCP")
      Else
       objTS.writeline("CheckFeature_SimpleTCP|Y")
      End If 
     
    End Select
    
Next
WScript.Echo "Windwos元件確認完畢"
objTS.writeline("CheckFeature_State|Y")
objTS.writeline("CheckFeature_EndTime|" & ConvertDate(now()))
End Function

Function WMIDateStringToDate(dtmDate)
WScript.Echo dtm: 
    WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
    Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
    & " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function 

function ConvertDate(dat)
	if isdate(dat) then
		dattmp = year(dat) & "/" & right("0" & month(dat),2) & "/" & right("0" & day(dat),2) & " " & right("0" & hour(dat),2) & ":" & right("0" & minute(dat),2) & ":" & right("0" & second(dat),2)
	else
		datTmp = ""
	end if
	ConvertDate = datTmp
End Function

Function checkDISM(feature)
DIM DISMTest, strResult
set DISMTest = wshShell.exec("DISM /online /get-featureinfo:" & feature)
do until DISMTest.status = 1 : wscript.sleep 10 : loop
strResult = DISMTest.stdout.Readall
If Instr(strResult, "Enabled") Then
'checkDISM = TRUE
'wscript.echo feature & " is installed"
Else
'checkDISM = FALSE
'wscript.echo feature & " is not installed"
EnableDISM(feature)
End If
End Function

Sub EnableDISM(feature)
DIM DISMInstall, strResult
SET DISMInstall = oShell.exec("DISM /online /enable-feature /featurename:" &_
feature )
do until DISMInstall.status = 1 : wscript.sleep 10 : loop
strResult = DISMInstall.stdout.readall
WScript.echo feature & " 安裝成功"
'wscript.echo strResult
End Sub