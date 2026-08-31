@echo off
rem Launcher: runs the PowerShell script with -ExecutionPolicy Bypass so the
rem default "Restricted" policy does not block it. This affects only this one
rem call and does NOT change any system setting.
rem NOTE: keep this file pure ASCII. cmd.exe parses batch files using the
rem console code page, and non-ASCII text here breaks parsing.
chcp 65001 >nul
setlocal
set "SCRIPT=%~dp0Get-VideoTranscript.ps1"

if not "%~1"=="" goto run

set /p "URL=Paste YouTube URL and press Enter: "
if "%URL%"=="" goto end
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Url "%URL%"
goto end

:run
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*

:end
echo.
pause
