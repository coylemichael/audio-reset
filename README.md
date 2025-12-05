# Elgato Audio Reset Tool

Resets Elgato Wave Link audio routing when it gets stuck or misbehaves.

## Installation

1. Open PowerShell as **Administrator**
2. Run:

```powershell
& ([scriptblock]::Create((irm "https://elgato.carnmorcyber.com")))
```

3. A window will pop up showing your audio devices - select your preferred defaults from the dropdowns
4. Click **Save**
5. Done! Point your Stream Deck button to:

```
%LOCALAPPDATA%\ElgatoReset\elgato_reset.bat
```

The installer auto-detects all your audio devices and pre-selects the current defaults. Just verify they're correct and click Save.

## What It Does

1. Kills WaveLink, WaveLinkSE, and StreamDeck processes
2. Restarts Windows audio services (audiosrv, AudioEndpointBuilder)  
3. Relaunches WaveLink and StreamDeck (minimized)
4. Sets your configured audio device defaults

## Configuration

The installer creates a config file at `%LOCALAPPDATA%\ElgatoReset\config.txt` with four device settings:

| Setting | Description |
|---------|-------------|
| `PLAYBACK_DEFAULT` | Where your system audio plays (games, music, videos, notifications) |
| `PLAYBACK_COMM` | Where you hear voice chat (Discord, Teams, Zoom, etc.) |
| `RECORD_DEFAULT` | Default microphone for apps that don't specify one |
| `RECORD_COMM` | Microphone used for voice chat (Discord, Teams, Zoom, etc.) |

### Changing Devices Later

Edit `%LOCALAPPDATA%\ElgatoReset\config.txt` and update the device names to match Windows Sound settings exactly.

```ini
# Example config.txt
PLAYBACK_DEFAULT=System (Elgato Virtual Audio)
PLAYBACK_COMM=Voice Chat (Elgato Virtual Audio)
RECORD_DEFAULT=Wave Link MicrophoneFX (Elgato Virtual Audio)
RECORD_COMM=Microphone (Razer Kraken V4 - Chat)
```

To find device names: `Windows Key + R` → `mmsys.cpl` → copy the exact name from each device.

## Files

After installation, files are located at:

```
%LOCALAPPDATA%\ElgatoReset\
├── elgato_reset.exe    ← Main executable
├── elgato_reset.bat    ← Trigger script for Stream Deck
└── config.txt          ← Your device configuration
```

## License

[MIT](LICENSE)
