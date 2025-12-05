# Elgato Audio Reset Tool

Resets Elgato Wave Link audio routing when it gets stuck or misbehaves.

## Quick Start

The recommended method - creates a scheduled task so no UAC prompt appears when you use it.

> **Why a scheduled task?** Windows requires admin privileges to restart audio services. The scheduled task runs with elevated permissions you grant once during setup, so subsequent runs don't need UAC approval. The task only executes the specific `elgato_reset.exe` you installed—it doesn't grant blanket admin access to anything else. Any attempt to modify, delete, or replace the scheduled task requires admin approval again, so nothing can change it without your knowledge.

1. Open PowerShell as **Administrator**
2. Run:

```powershell
& ([scriptblock]::Create((irm "https://elgato.carnmorcyber.com")))
```

3. A device selection window will appear - select your preferred audio devices from the dropdowns
4. Click **Save**
5. Done! Point your Stream Deck button to:

```
%LOCALAPPDATA%\ElgatoReset\elgato_reset.bat
```

---

<details>
<summary><b>Advanced Method</b> - Manual setup without scheduled task</summary>

This method doesn't require admin during setup, but you'll see a UAC prompt every time you run the tool.

### Option A: Download the Release

1. Download `elgato_reset.exe` from the [latest release](https://github.com/coylemichael/audio-reset/releases/latest)
2. Place it wherever you like (Desktop, OneDrive, etc.)
3. Run it once - a device selection GUI will appear
4. Select your preferred audio devices and click **Save**
5. Point your Stream Deck or Keyboard/Mouse Macro launcher directly to `elgato_reset.exe` (UAC prompt will appear each time).

### Option B: Clone and Build

1. Clone the repo: `git clone https://github.com/coylemichael/audio-reset.git`
2. Install [Visual Studio Build Tools](https://visualstudio.microsoft.com/visual-cpp-build-tools/) with "Desktop development with C++"
3. Build:
   ```powershell
   cmd /c '"C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\VC\Auxiliary\Build\vcvars64.bat" && cd /d "C:\path\to\audio-reset\c" && cl /O2 elgato_reset.c'
   ```
4. Run `elgato_reset.exe` - select your devices when prompted
5. Subsequent runs will use the saved config (UAC prompt will appear each time).

</details>

---

## What It Does

1. Kills WaveLink, WaveLinkSE, and StreamDeck processes
2. Restarts Windows audio services (audiosrv, AudioEndpointBuilder)  
3. Relaunches WaveLink and StreamDeck (minimized)
4. Sets your configured audio device defaults

> **Note:** Admin privileges are required because the tool uses `OpenSCManager` and `ControlService` to stop/restart Windows audio services (`audiosrv`, `AudioEndpointBuilder`) and `TerminateProcess` to kill WaveLink/StreamDeck processes.

## Configuration

On first run, the exe shows a device selection GUI and saves your choices to `config.txt` (in the same folder as the exe).

| Setting | Description |
|---------|-------------|
| `PLAYBACK_DEFAULT` | Where your system audio plays (games, music, videos, notifications) |
| `PLAYBACK_COMM` | Where you hear voice chat (Discord, Teams, Zoom, etc.) |
| `RECORD_DEFAULT` | Default microphone for apps that don't specify one |
| `RECORD_COMM` | Microphone used for voice chat (Discord, Teams, Zoom, etc.) |

### Changing Devices Later

**Option 1:** Delete `config.txt` and run the exe again - the GUI will reappear.

**Option 2:** Edit `config.txt` directly:

```ini
# Example config.txt
PLAYBACK_DEFAULT=System (Elgato Virtual Audio)
PLAYBACK_COMM=Voice Chat (Elgato Virtual Audio)
RECORD_DEFAULT=Wave Link MicrophoneFX (Elgato Virtual Audio)
RECORD_COMM=Microphone (Razer Kraken V4 - Chat)
```

To find device names: `Windows Key + R` → `mmsys.cpl` → copy the exact name from each device.

## Files

After Quick Start installation, files are located at:

```
%LOCALAPPDATA%\ElgatoReset\
├── elgato_reset.exe    ← Main executable (self-contained with GUI)
├── elgato_reset.bat    ← Trigger script for Stream Deck (runs via scheduled task)
├── config.txt          ← Your device configuration (auto-generated)
└── logs/               ← Log files from each run
```

## License

[MIT](LICENSE)
