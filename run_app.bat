@echo off
setlocal

set "scriptdir=%~dp0"

powershell.exe -ExecutionPolicy Bypass -NoExit -File "%scriptdir%main.ps1"

endlocal
