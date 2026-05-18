# MkvOpusEnc

MkvOpusEnc is a batch MKV audio processing toolchain with **Python as the primary implementation**.

The main script is [MkvOpusEnc.py](MkvOpusEnc.py). It scans the current folder for `.mkv` files, re-encodes non-AAC/Opus audio to Opus, preserves metadata, and organizes outputs automatically.

The PowerShell script [MkvOpusEnc.ps1](MkvOpusEnc.ps1) is a **legacy compatibility version** for Windows users who do not have Python installed.

## Project Status

- Primary script: [MkvOpusEnc.py](MkvOpusEnc.py)
- Legacy fallback: [MkvOpusEnc.ps1](MkvOpusEnc.ps1)

If both are available in your environment, use the Python script.

## What The Python Script Does

- Processes all `.mkv` files in the current directory (except files starting with `temp-output-`).
- Keeps `aac` and `opus` audio tracks as-is (remux).
- Re-encodes other audio codecs to Opus.
- Preserves language, track title, and delay for re-encoded tracks.
- Supports optional downmix to stereo for 6+ channel tracks.
- Writes one log file per input to `conv_logs`.
- Moves processed outputs to `completed` and originals to `original`.

## Requirements

### Primary (Python)

1. Python 3.8+
2. ffmpeg
3. ffprobe
4. mkvmerge (MKVToolNix)
5. sox_ng
6. opusenc (opus-tools)
7. mediainfo

All tools must be available in your `PATH`.

### Legacy (PowerShell)

For [MkvOpusEnc.ps1](MkvOpusEnc.ps1), install PowerShell 7+ plus the same external tools above.

## Installation Notes

Install the media tools using your package manager, then verify they are on `PATH`.

### Post-Install Check (PATH Validation)

Run the following commands to confirm each required tool is available:

```powershell
ffmpeg -version
mkvmerge --version
sox_ng --version
opusenc --version
mediainfo --Version
```

### Windows (winget)

```powershell
winget install Gyan.FFmpeg MoritzBunkus.MKVToolNix sox_ng.sox_ng MediaArea.MediaInfo
```

For Opus tools (`opusenc`), install from:

- https://github.com/Chocobo1/opus-tools_win32-build

### Linux (Debian / Ubuntu example)

```bash
sudo apt-get update
sudo apt-get install ffmpeg mkvtoolnix opus-tools mediainfo
```

Install `sox_ng` from your distro/package source if unavailable in the default repo.

## Usage (Primary)

Run from a folder containing your `.mkv` files:

```bash
python MkvOpusEnc.py
```

With downmix enabled:

```bash
python MkvOpusEnc.py --downmix
```

### Python CLI

- `--downmix`: downmixes 6+ channel tracks to stereo before Opus encoding.

## Output Layout

When at least one input file is found, the script creates:

- `completed/` for processed MKV files
- `original/` for original source files
- `conv_logs/` for per-file logs

## Legacy PowerShell Usage

Use [MkvOpusEnc.ps1](MkvOpusEnc.ps1) only when Python is unavailable:

```powershell
./MkvOpusEnc.ps1
./MkvOpusEnc.ps1 -Downmix
```

The PowerShell version follows the same batch behavior as the Python script.
