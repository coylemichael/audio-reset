#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Installs the Elgato Audio Reset tool.
.DESCRIPTION
    Downloads elgato_reset.exe, shows a GUI to select audio devices,
    sets up a scheduled task for silent admin execution,
    and creates a trigger batch file.
.LINK
    https://github.com/coylemichael/audio-reset
#>

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Elgato Audio Reset - Installer" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Configuration
$installDir = "$env:LOCALAPPDATA\ElgatoReset"
$exePath = "$installDir\elgato_reset.exe"
$configPath = "$installDir\config.txt"
$batPath = "$installDir\elgato_reset.bat"
$taskName = "ElgatoReset"
$downloadUrl = "https://github.com/coylemichael/audio-reset/releases/latest/download/elgato_reset.exe"

# Create install directory
Write-Host "[1/5] Creating install directory..." -ForegroundColor Yellow
if (-not (Test-Path $installDir)) {
    New-Item -ItemType Directory -Path $installDir -Force | Out-Null
}
Write-Host "      $installDir" -ForegroundColor Gray

# Download the executable
Write-Host "[2/5] Downloading elgato_reset.exe..." -ForegroundColor Yellow
try {
    Invoke-WebRequest -Uri $downloadUrl -OutFile $exePath -UseBasicParsing
    Write-Host "      Downloaded successfully" -ForegroundColor Gray
} catch {
    Write-Host "      ERROR: Failed to download. Make sure a release exists at:" -ForegroundColor Red
    Write-Host "      $downloadUrl" -ForegroundColor Red
    Write-Host ""
    Write-Host "      You may need to compile and create a release first." -ForegroundColor Yellow
    Write-Host "      See: https://github.com/coylemichael/audio-reset#advanced-method" -ForegroundColor Yellow
    exit 1
}

# Audio device detection using COM
Write-Host "[3/5] Detecting audio devices..." -ForegroundColor Yellow

$audioCode = @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public class AudioDevices {
    [DllImport("ole32.dll")]
    static extern int CoInitialize(IntPtr pvReserved);
    
    [DllImport("ole32.dll")]
    static extern void CoUninitialize();
    
    [ComImport, Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMMDeviceEnumerator {
        int EnumAudioEndpoints(int dataFlow, int stateMask, out IMMDeviceCollection devices);
        int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice device);
    }
    
    [ComImport, Guid("D666063F-1587-4E43-81F1-B948E807363F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMMDevice {
        int Activate([MarshalAs(UnmanagedType.LPStruct)] Guid iid, int clsCtx, IntPtr activationParams, [MarshalAs(UnmanagedType.IUnknown)] out object interfacePointer);
        int OpenPropertyStore(int stgmAccess, out IPropertyStore properties);
        int GetId([MarshalAs(UnmanagedType.LPWStr)] out string id);
        int GetState(out int state);
    }
    
    [ComImport, Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMMDeviceCollection {
        int GetCount(out int count);
        int Item(int index, out IMMDevice device);
    }
    
    [ComImport, Guid("886d8eeb-8cf2-4446-8d02-cdba1dbdcf99")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IPropertyStore {
        int GetCount(out int count);
        int GetAt(int index, out PropertyKey key);
        int GetValue(ref PropertyKey key, out PropVariant value);
    }
    
    [StructLayout(LayoutKind.Sequential)]
    public struct PropertyKey {
        public Guid fmtid;
        public int pid;
        public PropertyKey(Guid fmtid, int pid) { this.fmtid = fmtid; this.pid = pid; }
    }
    
    [StructLayout(LayoutKind.Sequential)]
    public struct PropVariant {
        public ushort vt;
        public ushort r1, r2, r3;
        public IntPtr data1;
        public IntPtr data2;
    }
    
    static readonly Guid CLSID_MMDeviceEnumerator = new Guid("BCDE0395-E52F-467C-8E3D-C4579291692E");
    static readonly PropertyKey PKEY_Device_FriendlyName = new PropertyKey(new Guid("a45c254e-df1c-4efd-8020-67d146a850e0"), 14);
    
    static string GetDeviceName(IMMDevice device) {
        try {
            IPropertyStore props;
            device.OpenPropertyStore(0, out props);
            PropVariant pv;
            var key = PKEY_Device_FriendlyName;
            props.GetValue(ref key, out pv);
            if (pv.data1 != IntPtr.Zero)
                return Marshal.PtrToStringUni(pv.data1);
        } catch { }
        return "";
    }
    
    public static string[] GetDevices(int dataFlow) {
        var result = new List<string>();
        CoInitialize(IntPtr.Zero);
        try {
            var type = Type.GetTypeFromCLSID(CLSID_MMDeviceEnumerator);
            var enumerator = (IMMDeviceEnumerator)Activator.CreateInstance(type);
            IMMDeviceCollection devices;
            // stateMask 1 = DEVICE_STATE_ACTIVE
            enumerator.EnumAudioEndpoints(dataFlow, 1, out devices);
            int count;
            devices.GetCount(out count);
            for (int i = 0; i < count; i++) {
                IMMDevice device;
                devices.Item(i, out device);
                string name = GetDeviceName(device);
                if (!string.IsNullOrEmpty(name))
                    result.Add(name);
            }
        } catch { }
        finally { CoUninitialize(); }
        return result.ToArray();
    }
    
    public static string GetDefaultDevice(int dataFlow, int role) {
        CoInitialize(IntPtr.Zero);
        try {
            var type = Type.GetTypeFromCLSID(CLSID_MMDeviceEnumerator);
            var enumerator = (IMMDeviceEnumerator)Activator.CreateInstance(type);
            IMMDevice device;
            int hr = enumerator.GetDefaultAudioEndpoint(dataFlow, role, out device);
            if (hr == 0 && device != null)
                return GetDeviceName(device);
        } catch { }
        finally { CoUninitialize(); }
        return "";
    }
}
"@

try {
    Add-Type -TypeDefinition $audioCode -Language CSharp -ErrorAction Stop
} catch {
    # Type may already be loaded
}

# Get all devices and defaults
$playbackDevices = [AudioDevices]::GetDevices(0)  # eRender
$recordingDevices = [AudioDevices]::GetDevices(1) # eCapture

$defaultPlayback = [AudioDevices]::GetDefaultDevice(0, 0)      # eRender, eConsole
$defaultPlaybackComm = [AudioDevices]::GetDefaultDevice(0, 2)  # eRender, eCommunications
$defaultRecording = [AudioDevices]::GetDefaultDevice(1, 0)     # eCapture, eConsole
$defaultRecordingComm = [AudioDevices]::GetDefaultDevice(1, 2) # eCapture, eCommunications

Write-Host "      Found $($playbackDevices.Count) playback devices" -ForegroundColor Gray
Write-Host "      Found $($recordingDevices.Count) recording devices" -ForegroundColor Gray

# Create the GUI
Write-Host "[4/5] Opening device selection..." -ForegroundColor Yellow

$form = New-Object System.Windows.Forms.Form
$form.Text = "Elgato Audio Reset - Device Selection"
$form.Size = New-Object System.Drawing.Size(520, 470)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$form.ForeColor = [System.Drawing.Color]::White

$font = New-Object System.Drawing.Font("Segoe UI", 9)
$smallFont = New-Object System.Drawing.Font("Segoe UI", 9)
$headerFont = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)

# Header
$header = New-Object System.Windows.Forms.Label
$header.Text = "Select your preferred audio devices"
$header.Font = $headerFont
$header.Location = New-Object System.Drawing.Point(20, 15)
$header.Size = New-Object System.Drawing.Size(470, 25)
$header.ForeColor = [System.Drawing.Color]::FromArgb(100, 200, 255)
$form.Controls.Add($header)

$subheader = New-Object System.Windows.Forms.Label
$subheader.Text = "These are your current default audio devices. No changes are needed`nif you're happy with your current input/output configuration."
$subheader.Font = $font
$subheader.Location = New-Object System.Drawing.Point(20, 42)
$subheader.Size = New-Object System.Drawing.Size(470, 40)
$subheader.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($subheader)

$yPos = 95

# Helper function to create dropdown with description
function Add-DeviceDropdown($labelText, $description, $devices, $defaultValue, $yPosition) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $labelText
    $label.Font = $font
    $label.Location = New-Object System.Drawing.Point(20, $yPosition)
    $label.Size = New-Object System.Drawing.Size(470, 20)
    $label.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($label)
    
    $desc = New-Object System.Windows.Forms.Label
    $desc.Text = $description
    $desc.Font = $smallFont
    $desc.Location = New-Object System.Drawing.Point(20, ($yPosition + 18))
    $desc.Size = New-Object System.Drawing.Size(470, 16)
    $desc.ForeColor = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $form.Controls.Add($desc)
    
    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.Font = $font
    $combo.Location = New-Object System.Drawing.Point(20, ($yPosition + 40))
    $combo.Size = New-Object System.Drawing.Size(460, 25)
    $combo.DropDownStyle = "DropDownList"
    $combo.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
    $combo.ForeColor = [System.Drawing.Color]::White
    $combo.FlatStyle = "Flat"
    
    foreach ($device in $devices) {
        $combo.Items.Add($device) | Out-Null
    }
    
    # Select the default
    $idx = $combo.Items.IndexOf($defaultValue)
    if ($idx -ge 0) { $combo.SelectedIndex = $idx }
    elseif ($combo.Items.Count -gt 0) { $combo.SelectedIndex = 0 }
    
    $form.Controls.Add($combo)
    return $combo
}

$comboPlaybackDefault = Add-DeviceDropdown "Default Playback Device" "Where your system audio plays (games, music, videos, notifications)" $playbackDevices $defaultPlayback $yPos
$yPos += 70
$comboRecordDefault = Add-DeviceDropdown "Default Recording Device" "Default microphone for apps that don't specify one" $recordingDevices $defaultRecording $yPos
$yPos += 70
$comboPlaybackComm = Add-DeviceDropdown "Communications Playback Device" "Where you hear voice chat (Discord, Teams, Zoom, etc.)" $playbackDevices $defaultPlaybackComm $yPos
$yPos += 70
$comboRecordComm = Add-DeviceDropdown "Communications Recording Device" "Microphone used for voice chat (Discord, Teams, Zoom, etc.)" $recordingDevices $defaultRecordingComm $yPos

# Save button
$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = "Save"
$saveButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$saveButton.Location = New-Object System.Drawing.Point(160, 382)
$saveButton.Size = New-Object System.Drawing.Size(180, 35)
$saveButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$saveButton.ForeColor = [System.Drawing.Color]::White
$saveButton.FlatStyle = "Flat"
$saveButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.Controls.Add($saveButton)
$form.AcceptButton = $saveButton

$result = $form.ShowDialog()

if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host ""
    Write-Host "Installation cancelled." -ForegroundColor Yellow
    exit 0
}

# Get selected values
$selectedPlaybackDefault = $comboPlaybackDefault.SelectedItem
$selectedPlaybackComm = $comboPlaybackComm.SelectedItem
$selectedRecordDefault = $comboRecordDefault.SelectedItem
$selectedRecordComm = $comboRecordComm.SelectedItem

Write-Host ""
Write-Host "      Selected devices:" -ForegroundColor Gray
Write-Host "        Playback Default: $selectedPlaybackDefault" -ForegroundColor White
Write-Host "        Playback Comms:   $selectedPlaybackComm" -ForegroundColor White
Write-Host "        Recording Default: $selectedRecordDefault" -ForegroundColor White
Write-Host "        Recording Comms:   $selectedRecordComm" -ForegroundColor White

# Write config file
$configContent = @"
# Elgato Audio Reset Configuration
# Created on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
# Edit these if your preferred devices change

PLAYBACK_DEFAULT=$selectedPlaybackDefault
PLAYBACK_COMM=$selectedPlaybackComm
RECORD_DEFAULT=$selectedRecordDefault
RECORD_COMM=$selectedRecordComm
"@
$configContent | Out-File -FilePath $configPath -Encoding UTF8
Write-Host ""
Write-Host "      Config saved to: $configPath" -ForegroundColor Gray

# Create scheduled task
Write-Host "[5/5] Creating scheduled task..." -ForegroundColor Yellow
$action = New-ScheduledTaskAction -Execute $exePath
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -RunLevel Highest
Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal -Force | Out-Null
Write-Host "      Task '$taskName' created" -ForegroundColor Gray

# Create trigger batch file
@"
@echo off
schtasks /run /tn "$taskName"
"@ | Out-File -FilePath $batPath -Encoding ASCII

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
Write-Host "To change devices later, edit:" -ForegroundColor Gray
Write-Host "  $configPath" -ForegroundColor Gray
Write-Host ""
