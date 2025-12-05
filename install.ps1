#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Installs the Elgato Audio Reset tool.
.DESCRIPTION
    Downloads elgato_reset.exe, sets up a scheduled task for silent admin execution,
    and creates a trigger batch file.
.LINK
    https://github.com/coylemichael/audio-reset
#>

$ErrorActionPreference = 'Stop'

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Elgato Audio Reset - Installer" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Configuration
$installDir = "$env:LOCALAPPDATA\ElgatoReset"
$exePath = "$installDir\elgato_reset.exe"
$batPath = "$installDir\elgato_reset.bat"
$taskName = "ElgatoReset"
$downloadUrl = "https://github.com/coylemichael/audio-reset/releases/latest/download/elgato_reset.exe"

# Create install directory
Write-Host "[1/4] Creating install directory..." -ForegroundColor Yellow
if (-not (Test-Path $installDir)) {
    New-Item -ItemType Directory -Path $installDir -Force | Out-Null
}
Write-Host "      $installDir" -ForegroundColor Gray

# Download the executable
Write-Host "[2/4] Downloading elgato_reset.exe..." -ForegroundColor Yellow
try {
    Invoke-WebRequest -Uri $downloadUrl -OutFile $exePath -UseBasicParsing
    Write-Host "      Downloaded successfully" -ForegroundColor Gray
} catch {
    Write-Host "      ERROR: Failed to download. Make sure a release exists at:" -ForegroundColor Red
    Write-Host "      $downloadUrl" -ForegroundColor Red
    Write-Host ""
    Write-Host "      You may need to compile and create a release first." -ForegroundColor Yellow
    Write-Host "      See: https://github.com/coylemichael/audio-reset#compiling-on-windows" -ForegroundColor Yellow
    exit 1
}

# Create scheduled task
Write-Host "[3/4] Creating scheduled task (for silent admin execution)..." -ForegroundColor Yellow
$action = New-ScheduledTaskAction -Execute $exePath
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -RunLevel Highest
Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal -Force | Out-Null
Write-Host "      Task '$taskName' created" -ForegroundColor Gray

# Create trigger batch file
Write-Host "[4/4] Creating trigger batch file..." -ForegroundColor Yellow
@"
@echo off
schtasks /run /tn "$taskName"
"@ | Out-File -FilePath $batPath -Encoding ASCII
Write-Host "      $batPath" -ForegroundColor Gray

# Done
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Installation Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Files installed to: $installDir" -ForegroundColor White
Write-Host ""
Write-Host "To run the reset:" -ForegroundColor White
Write-Host "  - Double-click: $batPath" -ForegroundColor Gray
Write-Host "  - Or run: schtasks /run /tn `"$taskName`"" -ForegroundColor Gray
Write-Host ""
Write-Host "For Stream Deck / macro buttons:" -ForegroundColor White
Write-Host "  Point to: $batPath" -ForegroundColor Gray
Write-Host ""
Write-Host "NOTE: Edit the C source and recompile to change audio device settings." -ForegroundColor Yellow
Write-Host "See: https://github.com/coylemichael/audio-reset#configuration" -ForegroundColor Yellow
Write-Host ""
