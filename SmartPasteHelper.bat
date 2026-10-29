@echo off
rem Starts the SmartPaste Helper (tray icon and the button on the Raiser's Edge
rem one-of box) without needing to change the PowerShell execution policy.
start "" powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0SmartPasteHelper.ps1"
