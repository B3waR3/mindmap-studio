@echo off
:: Mind Map Studio — Installer
:: Right-click this file and choose "Run as administrator"

net session >nul 2>&1
if %errorLevel% == 0 (
    :: Already running as admin
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install.ps1"
) else (
    :: Re-launch elevated
    powershell.exe -Command "Start-Process cmd.exe -ArgumentList '/c cd /d \"%~dp0\" && powershell.exe -NoProfile -ExecutionPolicy Bypass -File Install.ps1' -Verb RunAs -Wait"
)
