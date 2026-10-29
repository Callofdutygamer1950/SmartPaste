@echo off
rem Starts SmartPaste without needing to change the PowerShell execution policy.
start "" powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0SmartPaste.ps1"
