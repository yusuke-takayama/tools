@echo off

set CONVERTER=bin\Release\xlsConverter.exe

if not exist %CONVERTER% ( 
	echo ���Release��exe���R���p�C�����Ă��������B
	pause
	exit /b 0
)

rem �w���v���o�͂��܂�.
%CONVERTER%

rem templateIni.xml���o�͂��܂�.
%CONVERTER% /INI

rem Mst�̏o�͂����܂�.
%CONVERTER% sample/MstTest.xlsx sample/MstIni.xml �C�x���g�J�[�h MstEventCard /V /H /C /J /M

rem Msg�̏o�͂����܂�.
%CONVERTER% sample/MsgTest.xlsx sample/MsgIni.xml String MsgGuiString /S /SX /HT

rem Enum�̏o�͂����܂�.
%CONVERTER% sample/EnumTest.xlsx sample/EnumIni.xml SVCode SVCode /E

pause