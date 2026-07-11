<#
.SYNOPSIS
    Builds and installs Mind Map Studio as a OneNote COM add-in.
    No administrator rights required.
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

# ── Paths ──────────────────────────────────────────────────────────────────────
$scriptDir    = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectFile  = Join-Path $scriptDir "MindMapStudio.csproj"
$buildOutput  = Join-Path $scriptDir "bin\Release\net48"
$dllName      = "MindMapStudio.dll"
$installDir   = Join-Path $env:LOCALAPPDATA "MindMapStudio"
$dllInstalled = Join-Path $installDir $dllName

$clsidGuid   = "C1A2B3D4-E5F6-7A8B-9C0D-E1F2A3B4C5D6"
$progId      = "MindMapStudio.Connect"
$addinRegKey = "HKCU:\Software\Microsoft\Office\OneNote\AddIns\$progId"

# ── Check: .NET SDK ────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "Checking prerequisites..." -ForegroundColor White

$sdkList = ""
try { $sdkList = & dotnet --list-sdks 2>&1 } catch { $sdkList = "" }

if (-not $sdkList -or "$sdkList" -match "No .NET SDKs") {
    Write-Host "  [ERROR] .NET SDK not installed." -ForegroundColor Red
    Write-Host "  Download .NET 8 SDK (Windows x64) from:" -ForegroundColor Yellow
    Write-Host "    https://dotnet.microsoft.com/en-us/download/dotnet/8.0" -ForegroundColor Cyan
    Pause-Script; exit 1
}
Write-Host "  [OK] .NET SDK found" -ForegroundColor Green

# .NET Framework 4.8 (RegAsm not used but framework runtime must be present)
$netfxPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\mscorlib.dll"
if (-not (Test-Path $netfxPath)) {
    Write-Host "  [ERROR] .NET Framework 4.8 runtime not found." -ForegroundColor Red
    Pause-Script; exit 1
}
Write-Host "  [OK] .NET Framework 4.8" -ForegroundColor Green

# WebView2
$wv2Key = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
if (Test-Path $wv2Key) {
    Write-Host "  [OK] WebView2 Runtime" -ForegroundColor Green
} else {
    Write-Host "  [WARN] WebView2 not detected. Install if window won't open:" -ForegroundColor Yellow
    Write-Host "    https://go.microsoft.com/fwlink/p/?LinkId=2124703" -ForegroundColor Cyan
}

# ── Step 1: Build ─────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "[1/4] Building Mind Map Studio..." -ForegroundColor Cyan

Push-Location $scriptDir
try {
    & dotnet build $projectFile -c Release --nologo
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  [ERROR] Build failed (exit $LASTEXITCODE)." -ForegroundColor Red
        Pop-Location; Pause-Script; exit 1
    }
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    Pop-Location; Pause-Script; exit 1
}
Pop-Location
Write-Host "  [OK] Build succeeded." -ForegroundColor Green

# ── Step 2: Copy DLL + dependencies to install folder ─────────────────────────
Write-Host ""
Write-Host "[2/4] Copying files to $installDir ..." -ForegroundColor Cyan

if (-not (Test-Path $installDir)) { New-Item -ItemType Directory -Path $installDir -Force | Out-Null }

Get-ChildItem -Path $buildOutput -File | ForEach-Object {
    Copy-Item $_.FullName -Destination $installDir -Force
}
# Copy WebView2Loader.dll from x64 subfolder if present
$wv2Loader = Join-Path $buildOutput "x64\WebView2Loader.dll"
if (Test-Path $wv2Loader) {
    $x64Dir = Join-Path $installDir "x64"
    if (-not (Test-Path $x64Dir)) { New-Item -ItemType Directory -Path $x64Dir -Force | Out-Null }
    Copy-Item $wv2Loader -Destination $x64Dir -Force
}

Write-Host "  [OK] Files copied." -ForegroundColor Green

# ── Step 3: Register COM per-user (HKCU — no admin needed) ────────────────────
Write-Host ""
Write-Host "[3/4] Registering COM component (per-user)..." -ForegroundColor Cyan

$dllUri   = "file:///" + $dllInstalled.Replace("\", "/")
$assembly = "MindMapStudio, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"

# HKCU\Software\Classes\CLSID\{guid}\InprocServer32
$clsidBase  = "HKCU:\Software\Classes\CLSID\{$clsidGuid}"
$inprocPath = "$clsidBase\InprocServer32"
New-Item -Path $inprocPath -Force | Out-Null
Set-ItemProperty -Path $inprocPath -Name "(Default)"       -Value "mscoree.dll"
Set-ItemProperty -Path $inprocPath -Name "Class"           -Value "MindMapStudio.Connect"
Set-ItemProperty -Path $inprocPath -Name "Assembly"        -Value $assembly
Set-ItemProperty -Path $inprocPath -Name "RuntimeVersion"  -Value "v4.0.30319"
Set-ItemProperty -Path $inprocPath -Name "ThreadingModel"  -Value "Both"
Set-ItemProperty -Path $inprocPath -Name "CodeBase"        -Value $dllUri

# CLSID -> ProgId
New-Item -Path "$clsidBase\ProgId" -Force | Out-Null
Set-ItemProperty -Path "$clsidBase\ProgId" -Name "(Default)" -Value $progId

# HKCU\Software\Classes\<ProgId>\CLSID
$progIdBase = "HKCU:\Software\Classes\$progId"
New-Item -Path "$progIdBase\CLSID" -Force | Out-Null
Set-ItemProperty -Path $progIdBase          -Name "(Default)" -Value "Mind Map Studio"
Set-ItemProperty -Path "$progIdBase\CLSID"  -Name "(Default)" -Value "{$clsidGuid}"

Write-Host "  [OK] COM registered (per-user)." -ForegroundColor Green

# ── Step 4: Write OneNote add-in registry keys ────────────────────────────────
Write-Host ""
Write-Host "[4/4] Writing OneNote registry keys..." -ForegroundColor Cyan

if (-not (Test-Path $addinRegKey)) { New-Item -Path $addinRegKey -Force | Out-Null }
Set-ItemProperty -Path $addinRegKey -Name "FriendlyName" -Value "Mind Map Studio"
Set-ItemProperty -Path $addinRegKey -Name "Description"  -Value "Create and insert mind maps into OneNote"
Set-ItemProperty -Path $addinRegKey -Name "LoadBehavior" -Type DWord -Value 3
Write-Host "  [OK] Registry keys written." -ForegroundColor Green

# ── Done ───────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "=====================================" -ForegroundColor Green
Write-Host "  Installation complete!" -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor Green
Write-Host ""
Write-Host "Restart OneNote (desktop app)." -ForegroundColor White
Write-Host "A new 'Mind Map' tab will appear in the ribbon." -ForegroundColor White

Pause-Script
