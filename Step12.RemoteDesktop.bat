set /p workpath=<workpath.txt
echo RemoteDesktopSetup_BeginTime^|%date:~0,10% %time:~0,8%>>%workpath%
regedit /s files\RemoteDesktop.reg
echo RemoteDesktopSetup_State^|Y>>%workpath%
echo RemoteDesktopSetup_EndTime^|%date:~0,10% %time:~0,8%>>%workpath%
cscript files\MessageBox.vbs "���ݮୱ�]�w����."
exit /b
