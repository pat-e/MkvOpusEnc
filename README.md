# MkvOpusEnc

A powerful PowerShell script for advanced audio track processing in MKV files. This script intelligently re-encodes or remuxes audio tracks based on their codec, while preserving all video, subtitle, attachment, and metadata information.

It includes a flexible downmixing feature to convert multi-channel surround sound into a dialogue-focused stereo track, perfect for "night mode" listening or playback on stereo devices.

## Key Features

- **Selective Re-encoding:** Automatically re-encodes common high-bitrate codecs like DTS, AC3, E-AC3, and FLAC into the efficient Opus format.
- **Lossless Remuxing:** Preserves audio quality by remuxing (copying) tracks that are already in a modern format like AAC or Opus.
- **Full Metadata Preservation:** Carefully reads and re-applies all critical metadata, ensuring nothing is lost:
    - **Video, Subtitle, and Font Attachments** are all carried over.
    - **Audio Track Titles** (e.g., "Commentary by Director") are preserved.
    - **Audio Track Language** tags are preserved.
    - **Audio Delay (A/V Sync)** is read accurately and re-applied to maintain perfect synchronization.
- **Optional Downmixing:** A `-Downmix` switch transforms the script's behavior:
    - **Standard Mode:** Preserves the original channel layout (e.g., 5.1 DTS becomes 5.1 Opus).
    - **Downmix Mode:** Converts 5.1 or 7.1 surround sound into a high-quality stereo track.
- **High-Quality Downmix Formulas:** Uses specific, fine-tuned `ffmpeg` `pan` filters for downmixing, including a "dialogue boost" formula for clear speech.
- **Cross-Platform:** Fully compatible with PowerShell 7 on **Windows** and **Linux**.
- **Clean Operation:** Creates and automatically cleans up a temporary directory for all intermediate files.

## Requirements

1.  **PowerShell 7+** (for cross-platform compatibility)
2.  **FFmpeg**: For audio extraction and downmixing.
3.  **MKVToolNix**: For reading container data and final muxing (`mkvmerge`).
4.  **SoX (Sound eXchange)**: For audio normalization.
5.  **opus-tools**: For Opus encoding (`opusenc`).
6.  **MediaInfo**: For accurate audio delay detection.

## Installation

First, ensure you have PowerShell 7 or newer installed. Then, install the required command-line tools. They must be available in your system's `PATH`.

#### Windows (using [Chocolatey](https://chocolatey.org/))
```powershell
choco install ffmpeg mkvtoolnix sox opus-tools mediainfo
```

#### Linux (Debian / Ubuntu)
```bash
sudo apt-get update
sudo apt-get install ffmpeg mkvtoolnix sox opus-tools mediainfo
```

#### Linux (Arch Linux)
```bash
sudo pacman -Syu ffmpeg mkvtoolnix-cli sox opus-tools mediainfo
```
*(Note: on Arch, the package is `mkvtoolnix-cli`)*

## Usage

The script is run from the command line, providing an input file, an output file, and an optional switch to control downmixing.

```powershell
./MkvOpusEnc.ps1 -InputFile <path> -OutputFile <path> [-Downmix]
```

### Parameters

* `-InputFile <string>`
    * **(Required)** The full path to the source MKV file.
* `-OutputFile <string>`
    * **(Required)** The full path for the processed output MKV file.
* `-Downmix`
    * **(Optional Switch)** If this flag is present, any 5.1 or 7.1 audio tracks that are being re-encoded will be downmixed to stereo. If omitted, their original channel layout will be preserved.

### Examples

#### Example 1: Standard Mode (Preserving Channels)

This command will re-encode a 5.1 DTS track into a 5.1 Opus track.

```powershell
./MkvOpusEnc.ps1 -InputFile "C:\Movies\MyMovie.mkv" -OutputFile "C:\Movies\MyMovie-Opus.mkv"
```

#### Example 2: Downmix Mode (Creating Stereo)

This command will re-encode a 5.1 DTS track into a **stereo** Opus track using the "dialogue boost" formula.

```powershell
./MkvOpusEnc.ps1 -InputFile "C:\Movies\MyMovie.mkv" -OutputFile "C:\Movies\MyMovie-Stereo.mkv" -Downmix
```

#### Example 3: Batch Mode

This command will re-encode multiple files to Opus.

```powershell
foreach ($file in gci *.mkv ) {$new = $file.BaseName + "_opus" + $file.Extension ; MkvOpusEnc.ps1 -InputFile $file -OutputFile $new }
```
