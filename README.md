# MkvOpusEnc

MkvOpusEnc is a batch MKV audio processing toolchain with **Python as the primary implementation**.

The main script is [MkvOpusEnc.py](MkvOpusEnc.py). It scans the current folder for `.mkv` files, re-encodes non-AAC/Opus audio to Opus, preserves metadata, and organizes outputs automatically.

The PowerShell script [MkvOpusEnc.ps1](MkvOpusEnc.ps1) is a **Windows compatibility version** for users who do not have Python installed. It follows the same ffmpeg-based pipeline.

## Project Status

- Primary script: [MkvOpusEnc.py](MkvOpusEnc.py)
- Windows fallback: [MkvOpusEnc.ps1](MkvOpusEnc.ps1)

If both are available in your environment, use the Python script.

## What The Scripts Do

- Process all `.mkv` files in the current directory (except files starting with `temp-output-`).
- Keep `aac` and `opus` audio tracks as-is (remux).
- Re-encode other audio codecs to Opus.
- Normalize re-encoded tracks with ffmpeg `loudnorm` (two-pass linear, true-peak aware). No SoX.
- Preserve language, track title, and delay for re-encoded tracks.
- Optional Nightmode Dialogue downmix to stereo for 6+ channel tracks (`pan '<'`, so the mix cannot clip).
- Write one log file per input to `conv_logs`.
- Move processed outputs to `completed` and originals to `original`.

## Requirements

### Primary (Python)

1. Python 3.8+
2. ffmpeg
3. ffprobe
4. mkvmerge (MKVToolNix)
5. opusenc (opus-tools)
6. mediainfo

All tools must be available in your `PATH`.

### Windows (PowerShell)

For [MkvOpusEnc.ps1](MkvOpusEnc.ps1), install PowerShell 7+ plus the same external tools above.

## Installation Notes

Install the media tools using your package manager, then verify they are on `PATH`.

### Post-Install Check (PATH Validation)

Run the following commands to confirm each required tool is available:

```powershell
ffmpeg -version
ffprobe -version
mkvmerge --version
opusenc --version
mediainfo --Version
```

### Windows (winget)

```powershell
winget install Gyan.FFmpeg MoritzBunkus.MKVToolNix MediaArea.MediaInfo
```

For Opus tools (`opusenc`), install from:

- https://github.com/Chocobo1/opus-tools_win32-build

### Linux (Debian / Ubuntu example)

```bash
sudo apt-get update
sudo apt-get install ffmpeg mkvtoolnix opus-tools mediainfo
```

## Usage (Primary)

Run from a folder containing your `.mkv` files:

```bash
python MkvOpusEnc.py
```

With Nightmode Dialogue downmix enabled:

```bash
python MkvOpusEnc.py --downmix
```

Custom loudness targets:

```bash
python MkvOpusEnc.py --norm-i -18 --norm-tp -1.5
```

### Python CLI

- `--downmix`: Nightmode Dialogue downmix of 5.1/7.1 to stereo before Opus encoding.
- `--norm-i LUFS`: target integrated loudness (default: `-18.0`).
- `--norm-tp dBTP`: true-peak ceiling (default: `-1.5`).

## Output Layout

When at least one input file is found, the script creates:

- `completed/` for processed MKV files
- `original/` for original source files
- `conv_logs/` for per-file logs

## Windows PowerShell Usage

Use [MkvOpusEnc.ps1](MkvOpusEnc.ps1) when Python is unavailable:

```powershell
./MkvOpusEnc.ps1
./MkvOpusEnc.ps1 -Downmix
./MkvOpusEnc.ps1 -Downmix -NormI -18 -NormTp -1.5
```

The PowerShell version follows the same batch behavior and ffmpeg loudnorm pipeline as the Python script.
