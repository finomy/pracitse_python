@echo off
set /p n="input file name"
if not dir /a /s /b d:\xxxx\%n% | findstr . >nul (
    xcopy d:\xxx\%n%\*.* d:xxx\xxx\%n%
	)

pause
