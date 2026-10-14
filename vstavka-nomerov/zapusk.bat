@echo off
if exist "%~dp0programma.txt" (
    start "" /min powershell.exe -NoProfile -NoExit -ExecutionPolicy Bypass -STA -File "%~dp0vstavka.ps1"
) else (
    start "" powershell.exe -NoProfile -NoExit -ExecutionPolicy Bypass -STA -File "%~dp0vstavka.ps1"
)
