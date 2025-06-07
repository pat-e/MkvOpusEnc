<#
.SYNOPSIS
    Processes or downmixes an MKV file's audio tracks sequentially using a specific toolchain.
    This script is cross-platform and optimized for correctness and clean output.

.DESCRIPTION
    This script intelligently handles audio streams in an MKV file one by one.
    - AAC/Opus audio is remuxed.
    - Multi-channel audio (DTS, AC3, etc.) can be re-encoded or optionally downmixed to stereo.
    - All other streams and metadata (title, language, delay) are preserved.

.PARAMETER InputFile
    The full path to the source MKV file.

.PARAMETER OutputFile
    The full path for the processed output MKV file.

.PARAMETER Downmix
    An optional switch. If present, multi-channel audio will be downmixed to stereo.

.EXAMPLE
    # Regular mode - preserves channel layouts
    ./MkvOpusEnc.ps1 -InputFile "C:\Movies\My Movie (2025) [1080p].mkv" -OutputFile "C:\Movies\regular.mkv"

.EXAMPLE
    # Downmix mode
    ./MkvOpusEnc.ps1 -InputFile "C:\Movies\My Movie (2025) [1080p].mkv" -OutputFile "C:\Movies\downmixed.mkv" -Downmix
#>

# This makes the script behave like a compiled cmdlet with proper parameter handling.
[CmdletBinding()]
# The param block is now at the top level of the script.
param (
    [string]$InputFile,
    [string]$OutputFile,
    [switch]$Downmix
)

# Sanitize input path to remove PowerShell's escape characters (`).
if ($PSBoundParameters.ContainsKey('InputFile')) {
    $InputFile = $InputFile -replace '`',''
}
if ($PSBoundParameters.ContainsKey('OutputFile')) {
    $OutputFile = $OutputFile -replace '`',''
}


# Manual check for parameters.
if ([string]::IsNullOrWhiteSpace($InputFile) -or [string]::IsNullOrWhiteSpace($OutputFile)) {
    Write-Host "Usage Example:" -ForegroundColor Yellow
    Write-Host "  ./MkvOpusEnc.ps1 -InputFile `"C:\path\to\movie.mkv`" -OutputFile `"C:\path\to\movie.mkv`""
    Write-Host '  (Add -Downmix switch to downmix multi-channel audio to stereo)'
    return
}

# 1. --- Prerequisite Check ---
$requiredTools = "ffmpeg", "ffprobe", "mkvmerge", "sox", "opusenc", "mediainfo"
foreach ($tool in $requiredTools) {
    if (-not (Get-Command $tool -ErrorAction SilentlyContinue)) {
        Write-Error "Required tool '$tool' not found. Please install it and ensure it's in your system's PATH."
        return
    }
}

# --- MODIFIED: Use -LiteralPath to prevent PowerShell from interpreting [ and ] as wildcards.
if (-not (Test-Path -LiteralPath $InputFile)) {
    Write-Error "Input file not found: $InputFile"
    return
}

# Helper function is defined in the script scope.
function Convert-AudioTrack($index, $ch, $lang, $tempDirFullName, $sourceFile, [bool]$shouldDownmix) {
    $tempExtracted = Join-Path $tempDirFullName "track_${index}_extracted.flac"
    $tempNormalized = Join-Path $tempDirFullName "track_${index}_normalized.flac"
    $finalOpus = Join-Path $tempDirFullName "track_${index}_final.opus"

    # Step 1: Extract audio, with conditional downmixing
    Write-Host "    - Extracting to FLAC..."
    $ffmpegArgs = @("-v", "quiet", "-stats", "-i", $sourceFile, "-map", "0:$($index)")
    if ($shouldDownmix -and $ch -ge 6) {
        if ($ch -eq 6) { # 5.1 Channels
            Write-Host "      (Downmixing 5.1 to Stereo with dialogue boost)"
            $ffmpegArgs += "-af", "pan=stereo|c0=c2+0.30*c0+0.30*c4|c1=c2+0.30*c1+0.30*c5"
        }
        elseif ($ch -eq 8) { # 7.1 Channels
            Write-Host "      (Downmixing 7.1 to Stereo with dialogue boost)"
            $ffmpegArgs += "-af", "pan=stereo|c0=c2+0.30*c0+0.30*c4+0.30*c6|c1=c2+0.30*c1+0.30*c5+0.30*c7"
        }
        else { # Other multi-channel layouts
             Write-Host "      ($($ch)-channel source, downmixing to stereo using default -ac 2)"
             $ffmpegArgs += "-ac", "2"
        }
    } else {
        Write-Host "      (Preserving $($ch)-channel layout)"
    }
    $ffmpegArgs += "-c:a", "flac", $tempExtracted
    & ffmpeg $ffmpegArgs
    
    # Step 2: Normalize the track with SoX
    Write-Host "    - Normalizing with SoX..."
    & sox $tempExtracted $tempNormalized -S --temp $tempDirFullName --guard gain -n
    
    # Step 3: Encode to Opus with the correct bitrate
    $bitrate = "192k" # A fallback bitrate
    if ($shouldDownmix) {
        $bitrate = "128k"
    } else {
        switch ($ch) {
            2 { $bitrate = "128k" }
            6 { $bitrate = "256k" }
            8 { $bitrate = "384k" }
        }
    }
    Write-Host "    - Encoding to Opus at $bitrate..."
    & opusenc --vbr --bitrate $bitrate $tempNormalized $finalOpus

    return $finalOpus
}

# 2. --- Setup Temporary Environment ---
$tempDir = New-Item -ItemType Directory -Path (Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString()))
Write-Host "Temporary directory created at: $($tempDir.FullName)"

try {
    # 3. --- Get Media Information using robust argument handling ---
    Write-Host "Analyzing file: $InputFile"

    $ffprobeArgs = @("-v", "quiet", "-print_format", "json", "-show_streams", "-show_format", $InputFile)
    $ffprobeInfoJson = & ffprobe $ffprobeArgs
    $ffprobeInfo = $ffprobeInfoJson | ConvertFrom-Json -AsHashtable
    
    $mkvmergeJsonArgs = @("-J", $InputFile)
    $mkvInfoJson = & mkvmerge $mkvmergeJsonArgs
    $mkvInfo = $mkvInfoJson | ConvertFrom-Json
    
    $mediainfoArgs = @("--Output=JSON", "-f", $InputFile)
    $mediaInfoJsonString = & mediainfo $mediainfoArgs
    $mediaInfoJson = $mediaInfoJsonString | ConvertFrom-Json

    # 4. --- Prepare for Final mkvmerge Command ---
    $processedAudioFiles = [System.Collections.ArrayList]@()
    $audioTracksToRemux = [System.Collections.ArrayList]@()
    
    $videoTIDs = ($mkvInfo.tracks | Where-Object { $_.type -eq 'video' }).id
    $subtitleTIDs = ($mkvInfo.tracks | Where-Object { $_.type -eq 'subtitles' }).id
    $attachmentTIDs = ($mkvInfo.tracks | Where-Object { $_.type -eq 'attachments' }).id

    # 5. --- Process Each Audio Stream ---
    $audioStreams = $ffprobeInfo.streams | Where-Object { $_.codec_type -eq 'audio' }
    
    foreach ($stream in $audioStreams) {
        $streamIndex = $stream.index
        $codec = $stream.codec_name
        $channels = $stream.channels
        $language = if ($stream.tags.language) { $stream.tags.language } else { "und" }
        
        $mkvTrack = $mkvInfo.tracks[$streamIndex]
        $trackID = $mkvTrack.id
        $trackTitle = $mkvTrack.properties.track_name
        
        $trackDelay = 0
        $mediaInfoTrack = $mediaInfoJson.media.track | Where-Object { $_.'@type' -eq 'Audio' -and $_.StreamOrder -eq $streamIndex }
        $delayInSeconds = $mediaInfoTrack.Video_Delay
        
        if ($null -ne $delayInSeconds) {
            $trackDelay = [math]::Round( ([double]$delayInSeconds) * 1000 )
        }

        Write-Host "Processing Audio Stream #${streamIndex} (TID: $trackID, Title: $trackTitle, Delay: $($trackDelay)ms, Lang: $language, Codec: $codec, Channels: $channels)"

        switch ($codec) {
            'aac' { Write-Host "  -> Action: Remuxing."; $null = $audioTracksToRemux.Add($trackID) }
            'opus' { Write-Host "  -> Action: Remuxing."; $null = $audioTracksToRemux.Add($trackID) }
            { @('dts', 'ac3', 'eac3', 'flac') -contains $_ } {
                Write-Host "  -> Action: Re-encoding."
                $opusFile = Convert-AudioTrack $streamIndex $channels $language $tempDir.FullName $InputFile $Downmix.IsPresent
                $null = $processedAudioFiles.Add(@{ Path = $opusFile; Language = $language; Title = $trackTitle; Delay = $trackDelay })
            }
            default {
                Write-Warning "  -> Action: Unsupported codec '$codec'. Remuxing as fallback."
                $null = $audioTracksToRemux.Add($trackID)
            }
        }
    }

    # 6. --- Construct and Execute Final mkvmerge Command using Best Practices ---
    Write-Host "`nAssembling final mkvmerge command..."
    $mkvmergeArgs = @(
        "-o",
        $OutputFile
    )

    if ($videoTIDs.Count -gt 0) { $mkvmergeArgs += "-d", ($videoTIDs -join ',') }
    if ($subtitleTIDs.Count -gt 0) { $mkvmergeArgs += "-s", ($subtitleTIDs -join ',') }
    if ($attachmentTIDs.Count -gt 0) { $mkvmergeArgs += "-t", ($attachmentTIDs -join ',') }
    
    if ($audioStreams.Count -gt 0) { 
        if ($audioTracksToRemux.Count -gt 0) {
            $mkvmergeArgs += "-a", ($audioTracksToRemux -join ',')
        } else {
            $mkvmergeArgs += "--no-audio"
        }
    }
    
    $mkvmergeArgs += $InputFile

    foreach ($fileInfo in $processedAudioFiles) {
        $mkvmergeArgs += "--language", "0:$($fileInfo.Language)"
        $mkvmergeArgs += "--track-name", "0:$($fileInfo.Title)"

        if ($null -ne $fileInfo.Delay -and $fileInfo.Delay -ne 0) {
            $mkvmergeArgs += "--sync", "0:$($fileInfo.Delay)"
        }

        $mkvmergeArgs += $fileInfo.Path
    }

    Write-Host "`nExecuting command:"
    Write-Host ("mkvmerge " + ($mkvmergeArgs | ForEach-Object { if ($_ -match '\s|\[|\]|\(|\)') { "`"$_`"" } else { $_ } }) -join " ")
    
    & mkvmerge $mkvmergeArgs
}
catch {
    Write-Error "An error occurred during processing: $($_.Exception.Message)"
}
finally {
    # 7. --- Cleanup ---
    Write-Host "`nCleaning up temporary files..."
    if ($tempDir -and (Test-Path -LiteralPath $tempDir.FullName)) {
        # Using -LiteralPath here too for maximum robustness
        Remove-Item -Recurse -Force -LiteralPath $tempDir.FullName
    }
}
