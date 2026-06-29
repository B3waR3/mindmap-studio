<#
.SYNOPSIS
    Builds and installs Mind Map Studio as a OneNote COM add-in.
    Double-click this file OR run from PowerShell to install.
#>

# ── Auto-elevate to Administrator if needed ────────────────────────────────────
if (-not ([Security.Principal.WindowsPrincipal] `
          [Security.Principal.WindowsIdentity]::GetCurrent()
         ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
    Write-Host "Requesting administrator privileges..." -ForegroundColor Yellow
    Start-Process powershell.exe `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" `
        -Verb RunAs
    exit
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Pause-AndExit($code) {
    Write-Host ""
    Write-Host "Press Enter to close this window..." -ForegroundColor Gray
    Read-Host | Out-Null
    exit $code
}

try {
    $scriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
    $projectFile = Join-Path $scriptDir "MindMapStudio.csproj"
    $buildOutput = Join-Path $scriptDir "bin\Release\net48"
    $dllPath     = Join-Path $buildOutput "MindMapStudio.dll"
    $regAsm      = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
    $addinProgId = "MindMapStudio.Connect"
    $regKeyPath  = "HKCU:\Software\Microsoft\Office\OneNote\AddIns\$addinProgId"

    Write-Host ""
    Write-Host "=====================================" -ForegroundColor Cyan
    Write-Host "   Mind Map Studio — Installer" -ForegroundColor Cyan
    Write-Host "=====================================" -ForegroundColor Cyan

    # ── Check: .NET SDK ────────────────────────────────────────────────────────
    Write-Host ""
    Write-Host "Checking prerequisites..." -ForegroundColor White

    if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
        Write-Host ""
        Write-Host "[ERROR] .NET SDK not found." -ForegroundColor Red
        Write-Host "Please download and install it from:" -ForegroundColor Yellow
        Write-Host "  https://dotnet.microsoft.com/download" -ForegroundColor Cyan
        Write-Host "Install .NET SDK 6 or later, then run this installer again." -ForegroundColor Yellow
        Pause-AndExit 1
    }
    $dotnetVer = dotnet --version
    Write-Host "  [OK] .NET SDK $dotnetVer" -ForegroundColor Green

    # ── Check: RegAsm ─────────────────────────────────────────────────────────
    if (-not (Test-Path $regAsm)) {
        Write-Host "  [ERROR] .NET Framework 4.8 not found (RegAsm missing)." -ForegroundColor Red
        Write-Host "  This is normally pre-installed on Windows 10/11." -ForegroundColor Yellow
        Pause-AndExit 1
    }
    Write-Host "  [OK] .NET Framework 4.8" -ForegroundColor Green

    # ── Check: WebView2 Runtime ───────────────────────────────────────────────
    $wv2Key = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
    if (Test-Path $wv2Key) {
        Write-Host "  [OK] WebView2 Runtime" -ForegroundColor Green
    } else {
        Write-Host "  [WARN] WebView2 Runtime not detected." -ForegroundColor Yellow
        Write-Host "  If the add-in window doesn't open, install WebView2 from:" -ForegroundColor Yellow
        Write-Host "    https://go.microsoft.com/fwlink/p/?LinkId=2124703" -ForegroundColor Cyan
    }

    # ── Step 1: Build ─────────────────────────────────────────────────────────
    Write-Host ""
    Write-Host "[1/3] Building Mind Map Studio..." -ForegroundColor Cyan
    Push-Location $scriptDir
    & dotnet build $projectFile -c Release --nologo
    if ($LASTEXITCODE -ne 0) { throw "Build failed (exit code $LASTEXITCODE)." }
    Pop-Location
    Write-Host "  [OK] Build succeeded." -ForegroundColor Green

    # ── Step 2: Register COM DLL ──────────────────────────────────────────────
    Write-Host ""
    Write-Host "[2/3] Registering COM component..." -ForegroundColor Cyan
    if (-not (Test-Path $dllPath)) { throw "DLL not found: $dllPath" }
    & $regAsm $dllPath /codebase /nologo
    if ($LASTEXITCODE -ne 0) { throw "RegAsm failed (exit code $LASTEXITCODE)." }
    Write-Host "  [OK] COM registered." -ForegroundColor Green

    # ── Step 3: Write OneNote registry keys ───────────────────────────────────
    Write-Host ""
    Write-Host "[3/3] Writing OneNote registry keys..." -ForegroundColor Cyan
    if (-not (Test-Path $regKeyPath)) { New-Item -Path $regKeyPath -Force | Out-Null }
    Set-ItemProperty -Path $regKeyPath -Name "FriendlyName"  -Value "Mind Map Studio"
    Set-ItemProperty -Path $regKeyPath -Name "Description"   -Value "Create and insert mind maps into OneNote"
    Set-ItemProperty -Path $regKeyPath -Name "LoadBehavior"  -Type DWord -Value 3
    Write-Host "  [OK] Registry keys written." -ForegroundColor Green

    # ── Done ──────────────────────────────────────────────────────────────────
    Write-Host ""
    Write-Host "=====================================" -ForegroundColor Green
    Write-Host "  Installation complete!" -ForegroundColor Green
    Write-Host "=====================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next step:" -ForegroundColor White
    Write-Host "  Restart OneNote (desktop app)." -ForegroundColor White
    Write-Host "  A new 'Mind Map' tab will appear in the ribbon." -ForegroundColor White

} catch {
    Write-Host ""
    Write-Host "=====================================" -ForegroundColor Red
    Write-Host "  Installation FAILED" -ForegroundColor Red
    Write-Host "=====================================" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}

Pause-AndExit 0
