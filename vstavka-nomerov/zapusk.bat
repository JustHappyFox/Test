@echo off
start "" /min powershell.exe -NoProfile -NoExit -ExecutionPolicy Bypass -STA -File "%~dp0vstavka.ps1"
