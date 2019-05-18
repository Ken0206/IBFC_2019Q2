strBeginTime = ConvertDate(Now())
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("Wscript.Shell")
'objFSO.CopyFile "c:\TPE\1.jpg", "c:\"
strDocuments = ws.SpecialFolders("MyDocuments")

If  objFSO.FileExists("C:\Detect-File.vbs") Then
msgbox "C:\Detect-File.vbs OK"
else
objfso.copyfile "Don't_Delete\Detect-File.vbs","c:\Detect-File.vbs",true	
'msgbox "C:\Don't_Delete.jpg Error!!!"
End If

If  objFSO.FileExists("C:\Don't_Delete.jpg") Then
msgbox "C:\Don't_Delete.jpg OK"
else
objfso.copyfile "Don't_Delete\Don't_Delete.jpg","c:\Don't_Delete.jpg",true	
'msgbox "C:\Don't_Delete.jpg Error!!!"
End If

if objfso.DriveExists("D:") then
	If  objFSO.FileExists("D:\Don't_Delete.jpg") Then
	msgbox "D:\Don't_Delete.jpg OK"
	else
	objfso.copyfile "Don't_Delete\Don't_Delete.jpg","D:\Don't_Delete.jpg",true			
	'msgbox "D:\Don't_Delete.jpg Error!!!"
	End If
end if

If  objFSO.FileExists(strDocuments & "\Don't_Delete.jpg") Then
	msgbox strDocuments & "\Don't_Delete.jpg OK"
else
	'msgbox strDocuments & "\Don't_Delete.jpg Error!!"
	objfso.copyfile "Don't_Delete\Don't_Delete.jpg",strDocuments & "\Don't_Delete.jpg",true
End If


'Copy to All Users Mydocument Folder
strNoDocumentsAccount = ""
UsersFolder = objFSO.GetParentFolderName(objFSO.GetParentFolderName(ws.SpecialFolders("MyDocuments")))
dim sf
For Each sf in objfso.GetFolder(UsersFolder).SubFolders
	if objfso.FolderExists(sf.path & "\Documents") then
		objfso.copyfile "Don't_Delete\Don't_Delete.jpg",sf.path & "\Documents\Don't_Delete.jpg",true
	else
		msgbox sf.path & "\Documents" & " path not found!"
		strNoDocumentsAccount = strNoDocumentsAccount & "," & sf.path
	end if
	msgbox sf.path
Next
if strNoDocumentsAccount <> "" then
	strNoDocumentsAccount = mid(strNoDocumentsAccount,2)
end if

'strComputer = "." 
'Set objWMIService = GetObject("winmgmts:" _ 
'    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
' 
'Set colItems = objWMIService.ExecQuery _ 
'    ("Select * from Win32_UserAccount Where LocalAccount = True") 
' 
'For Each objItem in colItems 
'     UserDocumentFolder = UsersFolder & "\" & objItem.Name & "\Documents"
'     if objFSO.FolderExists(UserDocumentFolder ) then'
'	objfso.copyfile "Don't_Delete\Don't_Delete.jpg",UserDocumentFolder & "\Don't_Delete.jpg",true
'     end if
'Next


'copy to Allusersprofile Documents

objfso.copyfile "Don't_Delete\Don't_Delete.jpg","c:\ProgramData\Documents\Don't_Delete.jpg",true


strStartUP = Ws.SpecialFolders("AllUsersStartup" )
strStartUP = Ws.SpecialFolders("Startup" )
If  objFSO.FileExists(strStartUP & "\Detect-File.lnk") Then
	msgbox strStartUP & "\Detect-File.lnk OK"
else
	set oShellLink = Ws.CreateShortcut(strStartUP & "\Detect-File.lnk" )
	oShellLink.TargetPath = "c:\Detect-File.vbs"
	'oShellLink.WindowStyle = 1
	'oShellLink.IconLocation = "%SystemRoot%\explorer.exe"
	oShellLink.Description = "Detect Don't_Delete Jpg File"
	oShellLink.WorkingDirectory = "C:\"
	oShellLink.Save	
End If

'Write Text
Set objFS = CreateObject("Scripting.FileSystemObject")
Set f = objFS.openTextFile("workpath.txt")
strWorkPath = Replace(f.ReadLine,Chr(34),"")

f.close
Set f = objFS.OpenTextFile(strWorkPath, 8,true)

f.writeline("CheckDon'tDelete_BeginTime|" & strBeginTime)
f.writeline("CheckDon'tDelete_State|Y")
f.writeline("CheckDon'tDelete_NoDocumentPath|" & strNoDocumentsAccount)
f.writeline("CheckDon'tDelete_EndTime|" & ConvertDate(now()))
f.Close
MsgBox "¿À¨dßπ≤¶!!" 

function ConvertDate(dat)
	
	if isdate(dat) then
		dattmp = year(dat) & "/" & right("0" & month(dat),2) & "/" & right("0" & day(dat),2) & " " & right("0" & hour(dat),2) & ":" & right("0" & minute(dat),2) & ":" & right("0" & second(dat),2)
	else
		datTmp = ""
	end if
	ConvertDate = datTmp
end function
