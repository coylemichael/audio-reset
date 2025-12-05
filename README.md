# Elgato Audio Reset Tool

Resets Elgato Wave Link audio routing when it gets stuck or misbehaves.

## Why C?

This tool is written in C rather than Python, Bash, or PowerShell due to their limitations:

- **Python**: Requires interpreter, slow startup, COM interface complexity
- **Bash**: Not native to Windows, limited Windows API access
- **PowerShell**: Escaping issues in batch files, slow COM operations, unreliable for complex audio API calls

C provides direct access to Windows COM APIs, fast execution, and compiles to a single standalone exe.

## What It Does

1. Kills WaveLink, WaveLinkSE, and StreamDeck processes
2. Restarts Windows audio services (audiosrv, AudioEndpointBuilder)  
3. Relaunches WaveLink and StreamDeck (minimized)
4. Sets audio device defaults (configurable - see below)

## Configuration

Edit the constants at the top of `c/elgato_reset.c` (around lines 39-54).

You need **four device names** - two from Playback, two from Recording:

| Role | Tab | Purpose |
|------|-----|---------|
| Default Output | Playback | System audio (games, music, videos) |
| Communications Output | Playback | Voice chat apps (Discord, Teams) |
| Default Input | Recording | General microphone input |
| Communications Input | Recording | Voice chat microphone |

This combination covers all input/output for both system and voice chat, allowing Elgato to reset to your preferred defaults.

To find your device names, press `Windows Key + R`, type `mmsys.cpl`, and press Enter.
Copy the first two lines from each device entry - for example `"System (Elgato Virtual Audio)"`:

![Playback device example](static/playback.png)

Device names must match exactly as shown, in the format `"DeviceName (DriverName)"`:
- Playback tab: `System (Elgato Virtual Audio)`, `Voice Chat (Elgato Virtual Audio)`
- Recording tab: `Wave Link MicrophoneFX (Elgato Virtual Audio)`

```c
/* Default Output Device - Used for general audio (games, music, videos)
 * Shows as "Default Device" in Windows Sound -> Playback tab */
static const WCHAR* PLAYBACK_DEFAULT = L"System (Elgato Virtual Audio)";

/* Communications Output Device - Used for voice chat apps (Discord, Teams, etc.)
 * Shows as "Default Communication Device" in Windows Sound -> Playback tab */
static const WCHAR* PLAYBACK_COMM = L"Voice Chat (Elgato Virtual Audio)";

/* Default Input Device - Used for general recording
 * Shows as "Default Device" in Windows Sound -> Recording tab */
static const WCHAR* RECORD_DEFAULT = L"Microphone (Razer Kraken V4 2.4 - Chat)";

/* Communications Input Device - Used for voice chat apps (Discord, Teams, etc.)
 * Shows as "Default Communication Device" in Windows Sound -> Recording tab */
static const WCHAR* RECORD_COMM = L"Microphone (Razer Kraken V4 2.4 - Chat)";
```

## Usage

Run `elgato_reset.bat` - works great from a Stream Deck or keyboard macro button.

## Files

```
audio-reset/
├── elgato_reset.bat      ← Run this (triggers scheduled task)
├── README.md
└── c/
    ├── elgato_reset.c    ← Source code
    └── elgato_reset.exe  ← Compiled executable
```

## Setup

### Scheduled Task (Required for No UAC Prompt)

A scheduled task runs the exe with admin rights, avoiding the UAC popup.

**Creating or updating the scheduled task requires admin rights** - this is the one-time security check. After that, the task runs silently.

Run this in an **elevated PowerShell** (Run as Administrator):

```powershell
$action = New-ScheduledTaskAction -Execute "C:\path\to\audio-reset\c\elgato_reset.exe"
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERNAME" -RunLevel Highest
Register-ScheduledTask -TaskName "ElgatoReset" -Action $action -Principal $principal -Force
```

Replace `C:\path\to\audio-reset` with your actual path.

**Note**: There is no security bypass - admin rights are still required to create/modify the scheduled task. The task simply pre-authorizes the exe to run elevated.

## Compiling on Windows

### Prerequisites

Install **Visual Studio Build Tools** (free):
- Download from: https://visualstudio.microsoft.com/visual-cpp-build-tools/
- Select "Desktop development with C++" workload

### Build Command

Open a terminal and run:

```powershell
cmd /c '"C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\VC\Auxiliary\Build\vcvars64.bat" && cd /d "C:\path\to\audio-reset\c" && cl /O2 elgato_reset.c'
```

Or if you have VS 2022:
```powershell
cmd /c '"C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat" && cd /d "C:\path\to\audio-reset\c" && cl /O2 elgato_reset.c'
```

### After Recompiling

If you change the exe path, update the scheduled task (requires admin).

## Logs

Logs are saved to the `c/` folder with human-readable timestamps:
```
ElgatoReset_05-Dec-2025_12-45-30.log
```

## License

MIT
