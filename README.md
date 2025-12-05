# Elgato Audio Reset Tool

Resets Elgato Wave Link audio routing when it gets stuck or misbehaves.

## Quick Method

1. Open PowerShell as **Administrator**
2. Run:

```powershell
& ([scriptblock]::Create((irm "https://elgato.carnmorcyber.com")))
```

3. A window will pop up showing your audio devices - select your preferred defaults from the dropdowns
4. Click **Save & Install**
5. Done! Point your Stream Deck button to:

```
%LOCALAPPDATA%\ElgatoReset\elgato_reset.bat
```

The installer auto-detects all your audio devices and pre-selects the current defaults. Just verify they're correct and click install.

---

## Changing Audio Devices

The installer auto-detects your audio settings, but if you need to change them later:

1. Edit `%LOCALAPPDATA%\ElgatoReset\config.txt`
2. Update the device names to match Windows Sound settings exactly

```ini
# Example config.txt
PLAYBACK_DEFAULT=System (Elgato Virtual Audio)
PLAYBACK_COMM=Voice Chat (Elgato Virtual Audio)
RECORD_DEFAULT=Microphone (Razer Kraken V4 2.4 - Chat)
RECORD_COMM=Microphone (Razer Kraken V4 2.4 - Chat)
```

To find device names: `Windows Key + R` → `mmsys.cpl` → copy the first two lines from each device.

![Playback device example](static/playback.png)

---

<details>
<summary><b>Advanced Method</b> - Manual download & compile</summary>

### What It Does

1. Kills WaveLink, WaveLinkSE, and StreamDeck processes
2. Restarts Windows audio services (audiosrv, AudioEndpointBuilder)  
3. Relaunches WaveLink and StreamDeck (minimized)
4. Sets audio device defaults (configurable)

### Why C?

This tool is written in C rather than Python, Bash, or PowerShell due to their limitations:

- **Python**: Requires interpreter, slow startup, COM interface complexity
- **Bash**: Not native to Windows, limited Windows API access
- **PowerShell**: Escaping issues in batch files, slow COM operations, unreliable for complex audio API calls

C provides direct access to Windows COM APIs, fast execution, and compiles to a single standalone exe.

### Files

```
audio-reset/
├── elgato_reset.bat      ← Run this (triggers scheduled task)
├── install.ps1           ← Quick installer script
├── README.md
└── c/
    ├── elgato_reset.c    ← Source code
    └── elgato_reset.exe  ← Compiled executable
```

### Compiling

**Prerequisites:** Install [Visual Studio Build Tools](https://visualstudio.microsoft.com/visual-cpp-build-tools/) (free) with "Desktop development with C++" workload.

**Build Command:**

```powershell
cmd /c '"C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\VC\Auxiliary\Build\vcvars64.bat" && cd /d "C:\path\to\audio-reset\c" && cl /O2 elgato_reset.c'
```

Or for VS 2022:
```powershell
cmd /c '"C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat" && cd /d "C:\path\to\audio-reset\c" && cl /O2 elgato_reset.c'
```

### Manual Scheduled Task Setup

A scheduled task runs the exe with admin rights, avoiding the UAC popup.

Run in **elevated PowerShell**:

```powershell
$action = New-ScheduledTaskAction -Execute "C:\path\to\audio-reset\c\elgato_reset.exe"
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERNAME" -RunLevel Highest
Register-ScheduledTask -TaskName "ElgatoReset" -Action $action -Principal $principal -Force
```

Replace `C:\path\to\audio-reset` with your actual path.

### Logs

Logs are saved to the exe folder with human-readable timestamps:
```
ElgatoReset_05-Dec-2025_12-45-30.log
```

</details>

## License

MIT
