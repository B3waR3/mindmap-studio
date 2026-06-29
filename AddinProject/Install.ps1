#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Builds and installs Mind Map Studio as a OneNote COM add-in.

.DESCRIPTION
    1. Checks prerequisites (.NET SDK, WebView2 Runtime)
    2. Builds the C# project in Release mode
    3. Registers the DLL as a COM component with RegAsm
    4. Writes the OneNote add-in registry keys
    5. Reminds the user to restart OneNote

.NOTES
    Run once from an elevated (Administrator) PowerShell prompt.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir = $PSScriptRoot
$projectFile = Join-Path $scriptDir "MindMapStudio.csproj"
$buildOutput = Join-Path $scriptDir "bin\Release\net48"
$dllPath     = Join-Path $buildOutput "MindMapStudio.dll"
$regAsm      = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
$addinProgId = "MindMapStudio.Connect"
$regKeyPath  = "HKCU:\Software\Microsoft\Office\OneNote\AddIns\$addinProgId"

# ── Step 1: Verify prerequisites ───────────────────────────────────────────────

Write-Host "`n=== Mind Map Studio Installer ===" -ForegroundColor Cyan

# .NET SDK
if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    Write-Host "`n[ERROR] .NET SDK not found." -ForegroundColor Red
    Write-Host "Download it from: https://dotnet.microsoft.com/download" -ForegroundColor Yellow
    Write-Host "Install .NET SDK 6 or later, then re-run this script." -ForegroundColor Yellow
    exit 1
}
$dotnetVersion = (dotnet --version)
Write-Host "[OK] .NET SDK $dotnetVersion found." -ForegroundColor Green

# RegAsm
if (-not (Test-Path $regAsm)) {
    Write-Host "`n[ERROR] RegAsm.exe not found at: $regAsm" -ForegroundColor Red
    Write-Host ".NET Framework 4.8 may not be installed." -ForegroundColor Yellow
    exit 1
}
Write-Host "[OK] RegAsm.exe found." -ForegroundColor Green

# WebView2 Runtime (check registry)
$wv2Key = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
if (-not (Test-Path $wv2Key)) {
    Write-Host "`n[WARNING] WebView2 Runtime may not be installed." -ForegroundColor Yellow
    Write-Host "If the add-in fails to open, install WebView2 from:" -ForegroundColor Yellow
    Write-Host "https://go.microsoft.com/fwlink/p/?LinkId=2124703" -ForegroundColor Yellow
    Write-Host "Continuing installation..." -ForegroundColor Yellow
} else {
    Write-Host "[OK] WebView2 Runtime found." -ForegroundColor Green
}

# ── Step 2: Build ──────────────────────────────────────────────────────────────

Write-Host "`n[1/3] Building Mind Map Studio..." -ForegroundColor Cyan
Push-Location $scriptDir
try {
    & dotnet build $projectFile -c Release --nologo 2>&1 | Write-Host
    if ($LASTEXITCODE -ne 0) { throw "Build failed." }
} finally {
    Pop-Location
}
Write-Host "[OK] Build succeeded." -ForegroundColor Green

# ── Step 3: Register COM DLL ───────────────────────────────────────────────────

Write-Host "`n[2/3] Registering COM component..." -ForegroundColor Cyan
if (-not (Test-Path $dllPath)) {
    throw "DLL not found at: $dllPath"
}

# /codebase embeds the full path so the DLL doesn't need to be in the GAC
& $regAsm $dllPath /codebase /nologo 2>&1 | Write-Host
if ($LASTEXITCODE -ne 0) { throw "RegAsm registration failed." }
Write-Host "[OK] COM component registered." -ForegroundColor Green

# ── Step 4: Write OneNote add-in registry keys ─────────────────────────────────

Write-Host "`n[3/3] Writing OneNote registry keys..." -ForegroundColor Cyan
if (-not (Test-Path $regKeyPath)) {
    New-Item -Path $regKeyPath -Force | Out-Null
}
Set-ItemProperty -Path $regKeyPath -Name "FriendlyName"  -Value "Mind Map Studio"
Set-ItemProperty -Path $regKeyPath -Name "Description"   -Value "Create and insert mind maps into OneNote pages"
Set-ItemProperty -Path $regKeyPath -Name "LoadBehavior"  -Type DWord -Value 3
Write-Host "[OK] Registry keys written." -ForegroundColor Green

# ── Done ───────────────────────────────────────────────────────────────────────

Write-Host "`n=====================================" -ForegroundColor Cyan
Write-Host " Mind Map Studio installed!" -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next step: Restart OneNote (desktop app)."
Write-Host "A new 'Mind Map' tab will appear in the OneNote ribbon."
Write-Host ""
