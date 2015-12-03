@echo off

set CONVERTER=bin\Release\xlsConverter.exe

if not exist %CONVERTER% ( 
	echo 先にReleaseでexeをコンパイルしてください。
	pause
	exit /b 0
)

rem ヘルプを出力します.
%CONVERTER%

rem templateIni.xmlを出力します.
%CONVERTER% /INI

rem Mstの出力をします.
%CONVERTER% sample/MstTest.xlsx sample/MstIni.xml イベントカード MstEventCard /V /H /C /J /M

rem Msgの出力をします.
%CONVERTER% sample/MsgTest.xlsx sample/MsgIni.xml String MsgGuiString /S /SX /HT

rem Enumの出力をします.
%CONVERTER% sample/EnumTest.xlsx sample/EnumIni.xml SVCode SVCode /E

pause