set /p workpath=<workpath.txt
echo CopyFont_BeginTime^|%date:~0,10% %time:~0,8%>>%workpath%
rmdir "d:\�y�r" /s /q
xcopy /s /h /q /y "�y�r" "d:\�y�r\"
copy "d:\�y�r\system32\*" %windir%\system32\
copy "d:\�y�r\syswow64\*" %windir%\syswow64\
copy "d:\�y�r\winime.chm" %windir%\help\mui\0404
regedit /s "d:\�y�r\���X��J�k.reg"
copy "d:\�y�r\*.euf" %windir%\fonts
copy "d:\�y�r\*.tte" %windir%\fonts
regedit /s "d:\�y�r\�y�r201902.reg"
%windir%\system32\eudcedit.exe
control input.dll
echo CopyFont_State^|Y>>%workpath%
echo CopyFont_EndTime^|%date:~0,10% %time:~0,8%>>%workpath%

cscript files\MessageBox.vbs "�y�r�w�˧���."
exit /b
