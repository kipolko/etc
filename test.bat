rem echo off

set /a "MM=%date:~5,2%"

echo %ERRORLEVEL%

dir | findstr accdb

if %ERRORLEVEL% equ 1 (
	set /a "MM=%date:~5,1%*10 + %date:~6,1%-1"
)
echo %MM%

pause
