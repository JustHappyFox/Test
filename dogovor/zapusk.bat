@echo off
rem Start without the PowerShell window (it hides right after start)
start "" /min powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0dogovor.ps1"
