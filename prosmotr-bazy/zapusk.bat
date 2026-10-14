@echo off
start "" /min powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0prosmotr.ps1"
