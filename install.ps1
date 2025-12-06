#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Installs the Elgato Audio Reset tool.
.DESCRIPTION
    Downloads elgato_reset.exe, launches it for device/folder configuration,
    then sets up a scheduled task for silent admin execution.
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
$defaultInstallDir = "$env:LOCALAPPDATA\ElgatoReset"
$tempExePath = "$env:TEMP\elgato_audio_reset_temp.exe"
$taskName = "ElgatoReset"
$downloadUrl = "https://github.com/coylemichael/audio-reset/releases/latest/download/elgato_audio_reset.exe"

# Download the executable to temp
Write-Host "[1/3] Downloading elgato_audio_reset.exe..." -ForegroundColor Yellow
try {
    Invoke-WebRequest -Uri $downloadUrl -OutFile $tempExePath -UseBasicParsing
    Write-Host "      Downloaded successfully" -ForegroundColor Gray
} catch {
    Write-Host "      ERROR: Failed to download. Make sure a release exists at:" -ForegroundColor Red
    Write-Host "      $downloadUrl" -ForegroundColor Red
    Write-Host ""
    Write-Host "      You may need to compile and create a release first." -ForegroundColor Yellow
    Write-Host "      See: https://github.com/coylemichael/audio-reset#advanced-method" -ForegroundColor Yellow
    exit 1
}

# Set environment variable with suggested install dir
Write-Host "[2/3] Launching configuration GUI..." -ForegroundColor Yellow
Write-Host "      Select your install folder and audio devices" -ForegroundColor Gray
$env:ELGATO_INSTALL_DIR = $defaultInstallDir

# Run the exe - it will show the GUI and handle everything
$process = Start-Process -FilePath $tempExePath -PassThru -Wait
$exitCode = $process.ExitCode

# Clean up temp file
Remove-Item $tempExePath -Force -ErrorAction SilentlyContinue

# Check if config was created (means user completed setup)
$installDir = $env:ELGATO_INSTALL_DIR
$configPath = "$installDir\config.txt"
$exePath = "$installDir\elgato_audio_reset.exe"
$batPath = "$installDir\elgato_audio_reset.bat"

if (-not (Test-Path $configPath)) {
    Write-Host ""
    Write-Host "Setup was cancelled or incomplete." -ForegroundColor Yellow
    Write-Host "No files were installed." -ForegroundColor Gray
    exit 0
}

# Create scheduled task
Write-Host "[3/3] Creating scheduled task..." -ForegroundColor Yellow
$action = New-ScheduledTaskAction -Execute $exePath
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -RunLevel Highest
Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal -Force | Out-Null
Write-Host "      Task '$taskName' created" -ForegroundColor Gray

# Create trigger batch file
@"
@echo off
schtasks /run /tn "$taskName"
"@ | Out-File -FilePath $batPath -Encoding ASCII
Write-Host "      Trigger batch file created" -ForegroundColor Gray

# Done
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Installation Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Files installed to: $installDir" -ForegroundColor White
Write-Host ""
Write-Host "For Stream Deck / macro buttons, point to:" -ForegroundColor White
Write-Host "  $batPath" -ForegroundColor Cyan
Write-Host ""
Write-Host "To reconfigure devices, delete config.txt and run the exe again:" -ForegroundColor Gray
Write-Host "  $configPath" -ForegroundColor Gray
Write-Host ""
