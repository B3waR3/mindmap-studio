<#
.SYNOPSIS
    Builds and installs Mind Map Studio as a OneNote COM add-in.
    Double-click this file to install.
#>

$ErrorActionPreference = 'Stop'

function Pause-Script {
    Write-Host ""
    Write-Host "Press Enter to close this window..." -ForegroundColor Gray
    Read-Host | Out-Null
}

Write-Host ""
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "   Mind Map Studio - Installer" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan

# ── Resolve paths ──────────────────────────────────────────────────────────────
$scriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectFile = Join-Path $scriptDir "MindMapStudio.csproj"
$buildOutput = Join-Path $scriptDir "bin\Release\net48"
$dllPath     = Join-Path $buildOutput "MindMapStudio.dll"
$regAsm      = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
$addinProgId = "MindMapStudio.Connect"
$regKeyPath  = "HKCU:\Software\Microsoft\Office\OneNote\AddIns\$addinProgId"

# ── Check: .NET SDK ────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "Checking prerequisites..." -ForegroundColor White

# Check that the .NET SDK (not just the runtime) is installed
$sdkList = ""
try { $sdkList = & dotnet --list-sdks 2>&1 } catch { $sdkList = "" }

if (-not $sdkList -or $sdkList -match "No .NET SDKs") {
    Write-Host ""
    Write-Host "[ERROR] .NET SDK is not installed (only the Runtime was found)." -ForegroundColor Red
    Write-Host ""
    Write-Host "  You need the SDK (which includes build tools)." -ForegroundColor Yellow
    Write-Host "  Download the .NET 8 SDK installer (Windows x64) from:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "    https://dotnet.microsoft.com/en-us/download/dotnet/8.0" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  After installing, run Install.bat again." -ForegroundColor Yellow
    Pause-Script
    exit 1
}

Write-Host "  [OK] .NET SDK found" -ForegroundColor Green

# ── Check: RegAsm (.NET Framework 4.8) ────────────────────────────────────────
if (-not (Test-Path $regAsm)) {
    Write-Host "  [ERROR] .NET Framework 4.8 not found (RegAsm.exe missing)." -ForegroundColor Red
    Write-Host "  This is normally pre-installed on Windows 10/11." -ForegroundColor Yellow
    Pause-Script
    exit 1
}
Write-Host "  [OK] .NET Framework 4.8" -ForegroundColor Green

# ── Check: WebView2 Runtime ───────────────────────────────────────────────────
$wv2Key = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
if (Test-Path $wv2Key) {
    Write-Host "  [OK] WebView2 Runtime" -ForegroundColor Green
} else {
    Write-Host "  [WARN] WebView2 not detected - install if the window won't open:" -ForegroundColor Yellow
    Write-Host "    https://go.microsoft.com/fwlink/p/?LinkId=2124703" -ForegroundColor Cyan
}

# ── Step 1: Build ─────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "[1/3] Building Mind Map Studio..." -ForegroundColor Cyan

Push-Location $scriptDir
try {
    & dotnet build $projectFile -c Release --nologo
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ERROR] Build failed (exit code $LASTEXITCODE)." -ForegroundColor Red
        Pop-Location
        Pause-Script
        exit 1
    }
} catch {
    Write-Host "[ERROR] Build threw an exception: $($_.Exception.Message)" -ForegroundColor Red
    Pop-Location
    Pause-Script
    exit 1
}
Pop-Location
Write-Host "  [OK] Build succeeded." -ForegroundColor Green

# ── Step 2: Register COM DLL ──────────────────────────────────────────────────
Write-Host ""
Write-Host "[2/3] Registering COM component..." -ForegroundColor Cyan

if (-not (Test-Path $dllPath)) {
    Write-Host "[ERROR] DLL not found: $dllPath" -ForegroundColor Red
    Pause-Script
    exit 1
}

try {
    & $regAsm $dllPath /codebase /nologo
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ERROR] RegAsm failed (exit code $LASTEXITCODE)." -ForegroundColor Red
        Pause-Script
        exit 1
    }
} catch {
    Write-Host "[ERROR] RegAsm threw: $($_.Exception.Message)" -ForegroundColor Red
    Pause-Script
    exit 1
}
Write-Host "  [OK] COM registered." -ForegroundColor Green

# ── Step 3: Write OneNote registry keys ───────────────────────────────────────
Write-Host ""
Write-Host "[3/3] Writing OneNote registry keys..." -ForegroundColor Cyan

if (-not (Test-Path $regKeyPath)) { New-Item -Path $regKeyPath -Force | Out-Null }
Set-ItemProperty -Path $regKeyPath -Name "FriendlyName" -Value "Mind Map Studio"
Set-ItemProperty -Path $regKeyPath -Name "Description"  -Value "Create and insert mind maps into OneNote"
Set-ItemProperty -Path $regKeyPath -Name "LoadBehavior" -Type DWord -Value 3
Write-Host "  [OK] Registry keys written." -ForegroundColor Green

# ── Done ───────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "=====================================" -ForegroundColor Green
Write-Host "  Installation complete!" -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor Green
Write-Host ""
Write-Host "Next step: Restart OneNote (desktop app)." -ForegroundColor White
Write-Host "A new 'Mind Map' tab will appear in the OneNote ribbon." -ForegroundColor White

Pause-Script
