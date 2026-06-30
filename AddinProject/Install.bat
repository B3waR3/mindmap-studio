@echo off
:: Mind Map Studio — Installer
:: Double-click this file to install.

:: Check if already running as Administrator
net session >nul 2>&1
if %errorLevel% == 0 (
    :: Already admin — run the PowerShell installer
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install.ps1"
) else (
    :: Not admin — re-launch this bat file elevated
    powershell.exe -Command "Start-Process cmd.exe -ArgumentList '/c cd /d ""%~dp0"" && powershell.exe -NoProfile -ExecutionPolicy Bypass -File Install.ps1' -Verb RunAs"
)
