<#
.SYNOPSIS
    Processes or downmixes an MKV file's audio tracks, with special "dialogue boost" downmix formulas.

.DESCRIPTION
    This script intelligently handles audio streams in an MKV file one by one.
    - Default mode preserves channel layouts (e.g., 5.1 -> 5.1 Opus).
    - Using the optional -Downmix switch, 5.1/7.1 audio is downmixed to stereo using dialogue-focused formulas.
    - All other streams and metadata are preserved.

.PARAMETER InputFile
    The full path to the source MKV file.

.PARAMETER OutputFile
    The full path for the processed output MKV file.

.PARAMETER Downmix
    An optional switch. If present, multi-channel audio will be downmixed to stereo.

.EXAMPLE
    # Regular mode - preserves channel layouts
    .\MkvOpusEnc.ps1 -InputFile "C:\path\to\movie.mkv" -OutputFile "C:\path\to\regular.mkv"

.EXAMPLE
    # Downmix mode - converts multi-channel audio to a dialogue-boosted stereo track
    .\MkvOpusEnc.ps1 -InputFile "C:\path\to\movie.mkv" -OutputFile "C:\path\to\downmixed.mkv" -Downmix
#>

# This makes the script behave like a compiled cmdlet with proper parameter handling.
[CmdletBinding()]
# The param block is now at the top level of the script.
param (
    [string]$InputFile,
    [string]$OutputFile,
    [switch]$Downmix
)

# Manual check for parameters.
if ([string]::IsNullOrWhiteSpace($InputFile) -or [string]::IsNullOrWhiteSpace($OutputFile)) {
    Write-Host "Usage Example:" -ForegroundColor Yellow
    Write-Host '  .\MkvOpusEnc.ps1 -InputFile "C:\path\to\movie.mkv" -OutputFile "C:\path\to\movie.mkv"'
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

if (-not (Test-Path $InputFile)) {
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
    if ($shouldDownmix -and $ch -ge 6) {
        # DOWNMIX PATH
        if ($ch -eq 6) { # 5.1 Channels
            Write-Host "      (Downmixing 5.1 to Stereo with dialogue boost)"
            ffmpeg -v quiet -stats -i "$sourceFile" -map "0:$($index)" -af "pan=stereo|c0=c2+0.30*c0+0.30*c4|c1=c2+0.30*c1+0.30*c5" -c:a flac "$tempExtracted"
        }
        elseif ($ch -eq 8) { # 7.1 Channels
            Write-Host "      (Downmixing 7.1 to Stereo with dialogue boost)"
            ffmpeg -v quiet -stats -i "$sourceFile" -map "0:$($index)" -af "pan=stereo|c0=c2+0.30*c0+0.30*c4+0.30*c6|c1=c2+0.30*c1+0.30*c5+0.30*c7" -c:a flac "$tempExtracted"
        }
        else { # Other multi-channel layouts (e.g., 4.0, 6.1)
             Write-Host "      ($($ch)-channel source, downmixing to stereo using default -ac 2)"
             ffmpeg -v quiet -stats -i "$sourceFile" -map "0:$($index)" -ac 2 -c:a flac "$tempExtracted"
        }
    }
    else {
        # REGULAR PATH (no downmix)
        Write-Host "      (Preserving $($ch)-channel layout)"
        ffmpeg -v quiet -stats -i "$sourceFile" -map "0:$($index)" -c:a flac "$tempExtracted"
    }
    
    # Step 2: Normalize the track with SoX
    Write-Host "    - Normalizing with SoX..."
    sox "$tempExtracted" "$tempNormalized" -S --temp $tempDirFullName --guard gain -n
    
    # Step 3: Encode to Opus with the correct bitrate
    $bitrate = "192k" # A fallback bitrate
    if ($shouldDownmix) {
        $bitrate = "128k"
    }
    else {
        switch ($ch) {
            2 { $bitrate = "128k" }
            6 { $bitrate = "256k" }
            8 { $bitrate = "384k" }
        }
    }

    Write-Host "    - Encoding to Opus at $bitrate..."
    opusenc --vbr --bitrate $bitrate "$tempNormalized" "$finalOpus"

    return $finalOpus
}

# 2. --- Setup Temporary Environment ---
$tempDir = New-Item -ItemType Directory -Path (Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString()))
Write-Host "Temporary directory created at: $($tempDir.FullName)"

try {
    # 3. --- Get Media Information ---
    Write-Host "Analyzing file: $InputFile"
    $ffprobeInfoJson = ffprobe -v quiet -print_format json -show_streams -show_format "$InputFile"
    $ffprobeInfo = $ffprobeInfoJson | ConvertFrom-Json
    
    $mkvInfo = (mkvmerge -J "$InputFile") | ConvertFrom-Json -AsHashtable
    
    $mediaInfoJson = (mediainfo --Output=JSON -f "$InputFile") | ConvertFrom-Json

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

    # 6. --- Construct and Execute Final mkvmerge Command ---
    Write-Host "`nAssembling final mkvmerge command..."
    $mkvmergeCommand = "mkvmerge -o `"$OutputFile`""

    if ($videoTIDs.Count -gt 0) { $mkvmergeCommand += " -d $($videoTIDs -join ',')" }
    if ($subtitleTIDs.Count -gt 0) { $mkvmergeCommand += " -s $($subtitleTIDs -join ',')" }
    if ($attachmentTIDs.Count -gt 0) { $mkvmergeCommand += " -t $($attachmentTIDs -join ',')" }
    
    if ($audioStreams.Count -gt 0) { 
        if ($audioTracksToRemux.Count -gt 0) {
            $mkvmergeCommand += " -a $($audioTracksToRemux -join ',')"
        } else {
            $mkvmergeCommand += " --no-audio"
        }
    }
    
    $mkvmergeCommand += " `"$InputFile`""

    foreach ($fileInfo in $processedAudioFiles) {
        $syncSwitch = ""
        if ($null -ne $fileInfo.Delay -and $fileInfo.Delay -ne 0) {
            $syncSwitch = " --sync `"0:$($fileInfo.Delay)`""
        }
        $mkvmergeCommand += " --language `"0:$($fileInfo.Language)`" --track-name `"0:$($fileInfo.Title)`"$syncSwitch `"$($fileInfo.Path)`""
    }

    Write-Host "`nExecuting command:"
    Write-Host $mkvmergeCommand
    Invoke-Expression $mkvmergeCommand
}
catch {
    Write-Error "An error occurred during processing: $($_.Exception.Message)"
}
finally {
    # 7. --- Cleanup ---
    Write-Host "`nCleaning up temporary files..."
    if ($tempDir -and (Test-Path $tempDir.FullName)) {
        Remove-Item -Recurse -Force -Path $tempDir.FullName
    }
}
