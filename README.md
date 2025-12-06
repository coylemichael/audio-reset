# Elgato Audio Reset Tool

**Version 0.9**

Resets Elgato Wave Link audio routing when it gets stuck or misbehaves.

## What It Does

1. Kills WaveLink, WaveLinkSE, and StreamDeck processes
2. Restarts Windows audio services (audiosrv, AudioEndpointBuilder)  
3. Relaunches WaveLink and StreamDeck (minimized)
4. Sets your configured audio device defaults
5. Optionally shows a completion notification

## Quick Start

The recommended method - creates a scheduled task with admin privileges so you don't get a UAC prompt every time. The `.bat` file simply triggers the scheduled task, which was configured once during initial setup.

1. Open PowerShell as **Administrator**
2. Run:

```powershell
& ([scriptblock]::Create((irm "https://elgato.carnmorcyber.com")))
```

3. A device selection window will appear - select your preferred audio devices from the dropdowns
4. Configure options:
   - **Run in background** - If checked, runs silently without showing the GUI on subsequent runs
   - **Show completion notification** - If checked, displays "Elgato Audio Reset Complete" when done
5. Click **Save** to save settings, or **Fix Audio** to save and execute immediately
6. Done! Point your Stream Deck button to:

```
%LOCALAPPDATA%\ElgatoReset\elgato_audio_reset.bat
```

> **Note:** Admin privileges are required because the tool restarts Windows audio services (`audiosrv`, `AudioEndpointBuilder`) and terminates WaveLink/StreamDeck processes.

---

<details>
<summary><b>Advanced Method</b> - Manual setup without scheduled task</summary>

This method doesn't require admin during setup, but you'll see a UAC prompt every time you run the tool.

### Option A: Download the Release

1. Download `elgato_audio_reset.exe` from the [latest release](https://github.com/coylemichael/audio-reset/releases/latest)
2. Run it - a configuration window will appear
3. Click **Browse** to select where you want the tool installed (creates its own folder)
4. Select your preferred audio devices from the dropdowns
5. Click **Save & Install**
6. Point your Stream Deck or macro button to the installed `elgato_audio_reset.exe` (UAC prompt will appear each time)

### Option B: Clone and Build

1. Clone the repo: `git clone https://github.com/coylemichael/audio-reset.git`
2. Install [Visual Studio Build Tools](https://visualstudio.microsoft.com/visual-cpp-build-tools/) with "Desktop development with C++"
3. Build:
   ```powershell
   cmd /c '"C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\VC\Auxiliary\Build\vcvars64.bat" && cd /d "C:\path\to\audio-reset\c" && cl /O2 elgato_reset.c'
   ```
4. Run `elgato_audio_reset.exe` - select install folder and devices when prompted
5. Subsequent runs will use the saved config (UAC prompt will appear each time).

</details>

---

## Configuration

On first run, the GUI lets you:
- Choose an install folder
- Select your preferred audio devices from dropdowns
- Set preferences for background mode and notifications

To change devices later, either uncheck "Run in background" in the GUI, or delete `config.txt` and run again.