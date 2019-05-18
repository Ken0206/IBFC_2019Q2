set /p workpath=<workpath.txt
echo CopyFont_BeginTime^|%date:~0,10% %time:~0,8%>>%workpath%
rmdir "d:\造字" /s /q
xcopy /s /h /q /y "造字" "d:\造字\"
copy "d:\造字\system32\*" %windir%\system32\
copy "d:\造字\syswow64\*" %windir%\syswow64\
copy "d:\造字\winime.chm" %windir%\help\mui\0404
regedit /s "d:\造字\內碼輸入法.reg"
copy "d:\造字\*.euf" %windir%\fonts
copy "d:\造字\*.tte" %windir%\fonts
regedit /s "d:\造字\造字201902.reg"
%windir%\system32\eudcedit.exe
control input.dll
echo CopyFont_State^|Y>>%workpath%
echo CopyFont_EndTime^|%date:~0,10% %time:~0,8%>>%workpath%

cscript files\MessageBox.vbs "造字安裝完成."
exit /b
