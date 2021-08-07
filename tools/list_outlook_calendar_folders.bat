@echo off
cd %~dp0
powershell.exe -ExecutionPolicy Bypass "..\src\list_outlook_calendar_folders.ps1"
pause
