@echo off
rem Double-click to build dist\FS2Tools.xlam
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build.ps1"
pause
