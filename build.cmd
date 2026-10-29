@echo off
rem Double-click to build dist\FS2Tools.xlam
if not exist "%~dp0build.ps1" (
    echo.
    echo It looks like build.cmd was opened from inside the zip file.
    echo Close this window, right-click the zip, choose "Extract All...",
    echo then open the extracted folder and double-click build.cmd there.
    echo.
    pause
    exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build.ps1"
pause
