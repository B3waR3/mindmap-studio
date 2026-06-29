#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Unregisters and removes Mind Map Studio from OneNote.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

$scriptDir   = $PSScriptRoot
$buildOutput = Join-Path $scriptDir "bin\Release\net48"
$dllPath     = Join-Path $buildOutput "MindMapStudio.dll"
$regAsm      = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
$addinProgId = "MindMapStudio.Connect"
$regKeyPath  = "HKCU:\Software\Microsoft\Office\OneNote\AddIns\$addinProgId"

Write-Host "`n=== Mind Map Studio Uninstaller ===" -ForegroundColor Cyan

# Remove OneNote registry key
if (Test-Path $regKeyPath) {
    Remove-Item -Path $regKeyPath -Force -Recurse
    Write-Host "[OK] OneNote registry key removed." -ForegroundColor Green
} else {
    Write-Host "[SKIP] OneNote registry key not found." -ForegroundColor Yellow
}

# Unregister COM component
if (Test-Path $dllPath) {
    & $regAsm $dllPath /unregister /nologo 2>&1 | Write-Host
    Write-Host "[OK] COM component unregistered." -ForegroundColor Green
} else {
    Write-Host "[SKIP] DLL not found, skipping unregister." -ForegroundColor Yellow
}

Write-Host "`nMind Map Studio has been removed." -ForegroundColor Green
Write-Host "Please restart OneNote for changes to take effect.`n"
