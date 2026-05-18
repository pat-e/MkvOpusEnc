<#
.SYNOPSIS
    Batch-processes MKV files in the current directory and converts non-AAC/Opus audio tracks to Opus.

.DESCRIPTION
    Mirrors the behavior of MkvOpusEnc.py:
    - Scans for *.mkv files (excluding temp-output-*).
    - AAC/Opus tracks are remuxed unchanged.
    - All other audio codecs are extracted, normalized, and encoded to Opus.
    - Optional downmix for 5.1/7.1 and other 6+ channel layouts.
    - Preserves language, title, and delay metadata for re-encoded tracks.
    - Writes per-file logs to conv_logs, moves processed files to completed, originals to original.
#>

[CmdletBinding()]
param (
    [switch]$Downmix
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Invoke-ExternalCommand {
    param (
        [Parameter(Mandatory = $true)][string]$Command,
        [Parameter(Mandatory = $true)][string[]]$Arguments,
        [switch]$CaptureOutput,
        [switch]$NoCheck
    )

    if ($CaptureOutput) {
        $output = & $Command @Arguments
        $exitCode = $LASTEXITCODE
        if (-not $NoCheck -and $exitCode -ne 0) {
            throw "Command failed (exit code $exitCode): $Command $($Arguments -join ' ')"
        }
        return ($output -join [Environment]::NewLine)
    }

    & $Command @Arguments
    $exitCode = $LASTEXITCODE
    if (-not $NoCheck -and $exitCode -ne 0) {
        throw "Command failed (exit code $exitCode): $Command $($Arguments -join ' ')"
    }

    return $null
}

function Test-RequiredTools {
    $requiredTools = @('ffmpeg', 'ffprobe', 'mkvmerge', 'sox_ng', 'opusenc', 'mediainfo')
    Write-Host '--- Prerequisite Check ---'

    $allFound = $true
    foreach ($tool in $requiredTools) {
        if (-not (Get-Command $tool -ErrorAction SilentlyContinue)) {
            Write-Error "Required tool '$tool' not found."
            $allFound = $false
        }
    }

    if (-not $allFound) {
        throw "Please install the missing tools and ensure they are in your system's PATH."
    }

    Write-Host 'All required tools found.'
}

function Convert-AudioTrack {
    param (
        [Parameter(Mandatory = $true)][int]$StreamIndex,
        [Parameter(Mandatory = $true)][int]$Channels,
        [Parameter(Mandatory = $true)][string]$TempDir,
        [Parameter(Mandatory = $true)][string]$SourceFile,
        [Parameter(Mandatory = $true)][bool]$ShouldDownmix,
        [Parameter(Mandatory = $true)][string]$BitrateInfo
    )

    $tempExtracted = Join-Path $TempDir "track_${StreamIndex}_extracted.flac"
    $tempNormalized = Join-Path $TempDir "track_${StreamIndex}_normalized.flac"
    $finalOpus = Join-Path $TempDir "track_${StreamIndex}_final.opus"

    Write-Host '    - Extracting to FLAC...'
    $ffmpegArgs = @('-v', 'quiet', '-stats', '-y', '-i', $SourceFile, '-map', "0:$StreamIndex")

    $finalChannels = $Channels
    if ($ShouldDownmix -and $Channels -ge 6) {
        if ($Channels -eq 6) {
            Write-Host '      (Downmixing 5.1 to Stereo with dialogue boost)'
            $ffmpegArgs += @('-af', 'pan=stereo|c0=c2+0.30*c0+0.30*c4|c1=c2+0.30*c1+0.30*c5')
            $finalChannels = 2
        }
        elseif ($Channels -eq 8) {
            Write-Host '      (Downmixing 7.1 to Stereo with dialogue boost)'
            $ffmpegArgs += @('-af', 'pan=stereo|c0=c2+0.30*c0+0.30*c4+0.30*c6|c1=c2+0.30*c1+0.30*c5+0.30*c7')
            $finalChannels = 2
        }
        else {
            Write-Host "      ($Channels-channel source, downmixing to stereo using default -ac 2)"
            $ffmpegArgs += @('-ac', '2')
            $finalChannels = 2
        }
    }
    else {
        Write-Host "      (Preserving $Channels-channel layout)"
    }

    $ffmpegArgs += @('-c:a', 'flac', $tempExtracted)
    Invoke-ExternalCommand -Command 'ffmpeg' -Arguments $ffmpegArgs | Out-Null

    Write-Host '    - Normalizing with SoX...'
    Invoke-ExternalCommand -Command 'sox_ng' -Arguments @($tempExtracted, $tempNormalized, '-S', '--temp', $TempDir, '--guard', 'gain', '-n') | Out-Null

    $bitrate = '192k'
    if ($finalChannels -eq 1) {
        $bitrate = '64k'
    }
    elseif ($finalChannels -eq 2) {
        $bitrate = '128k'
    }
    elseif ($finalChannels -eq 6) {
        $bitrate = '256k'
    }
    elseif ($finalChannels -eq 8) {
        $bitrate = '384k'
    }

    Write-Host "    - Encoding to Opus at $bitrate..."
    Write-Host "      Source: $BitrateInfo -> Destination: Opus $bitrate ($finalChannels channels)"
    Invoke-ExternalCommand -Command 'opusenc' -Arguments @('--vbr', '--bitrate', $bitrate, $tempNormalized, $finalOpus) | Out-Null

    return @{
        Path = $finalOpus
        FinalChannels = $finalChannels
        Bitrate = $bitrate
    }
}

Test-RequiredTools

$DIR_COMPLETED = Join-Path (Get-Location) 'completed'
$DIR_ORIGINAL = Join-Path (Get-Location) 'original'
$DIR_LOGS = Join-Path (Get-Location) 'conv_logs'

$filesToProcess = Get-ChildItem -File -Filter '*.mkv' |
    Where-Object { $_.Name -notlike 'temp-output-*' } |
    Sort-Object Name

if (-not $filesToProcess) {
    Write-Host 'No MKV files found to process. Exiting.'
    return
}

$null = New-Item -ItemType Directory -Path $DIR_COMPLETED -Force
$null = New-Item -ItemType Directory -Path $DIR_ORIGINAL -Force
$null = New-Item -ItemType Directory -Path $DIR_LOGS -Force

foreach ($file in $filesToProcess) {
    $logFilePath = Join-Path $DIR_LOGS "$($file.Name).log"
    $intermediateOutputFile = Join-Path (Get-Location) "temp-output-$($file.Name)"
    $tempDir = $null

    Start-Transcript -Path $logFilePath -Force | Out-Null
    try {
        Write-Host ('-' * 80)
        Write-Host "Starting processing for: $($file.Name)"
        Write-Host "Log file: $logFilePath"
        $startTime = Get-Date

        $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("mkvopusenc_{0}" -f ([System.Guid]::NewGuid().ToString('N')))
        $null = New-Item -ItemType Directory -Path $tempDir -Force
        Write-Host "Temporary directory for audio created at: $tempDir"

        Write-Host "Analyzing file: $($file.FullName)"
        $ffprobeInfoJson = Invoke-ExternalCommand -Command 'ffprobe' -Arguments @('-v', 'quiet', '-print_format', 'json', '-show_streams', '-show_format', $file.FullName) -CaptureOutput
        $ffprobeInfo = $ffprobeInfoJson | ConvertFrom-Json -AsHashtable

        $mkvmergeInfoJson = Invoke-ExternalCommand -Command 'mkvmerge' -Arguments @('-J', $file.FullName) -CaptureOutput
        $mkvInfo = $mkvmergeInfoJson | ConvertFrom-Json -AsHashtable

        $mediainfoJson = Invoke-ExternalCommand -Command 'mediainfo' -Arguments @('--Output=JSON', '-f', $file.FullName) -CaptureOutput
        $mediaInfo = $mediainfoJson | ConvertFrom-Json -AsHashtable

        $processedAudioFiles = [System.Collections.ArrayList]@()
        $tidsOfReencodedTracks = [System.Collections.ArrayList]@()

        $audioStreams = @($ffprobeInfo.streams | Where-Object { $_.codec_type -eq 'audio' })
        if ($audioStreams.Count -eq 0) {
            Write-Warning "No audio streams found in '$($file.Name)'. Skipping file."
            continue
        }

        $mkvTracksList = @($mkvInfo.tracks)
        $mkvAudioTracks = @($mkvTracksList | Where-Object { $_.type -eq 'audio' })

        $mediaTracksData = @($mediaInfo.media.track)
        $mediainfoAudioTracks = @{}
        foreach ($track in $mediaTracksData) {
            if ($track.'@type' -ne 'Audio') {
                continue
            }

            $streamOrder = -1
            if ($null -ne $track.StreamOrder) {
                $parsed = 0
                if ([int]::TryParse([string]$track.StreamOrder, [ref]$parsed)) {
                    $streamOrder = $parsed
                }
            }

            $mediainfoAudioTracks[$streamOrder] = $track
        }

        Write-Host "`n=== Audio Track Analysis ==="
        for ($audioStreamIdx = 0; $audioStreamIdx -lt $audioStreams.Count; $audioStreamIdx++) {
            $stream = $audioStreams[$audioStreamIdx]
            $streamIndex = [int]$stream.index
            $codec = [string]$stream.codec_name
            $channels = 2
            if ($null -ne $stream.channels) {
                $channels = [int]$stream.channels
            }

            $language = 'und'
            if ($stream.ContainsKey('tags') -and $null -ne $stream.tags -and $stream.tags.ContainsKey('language')) {
                $language = [string]$stream.tags.language
            }

            $trackId = -1
            $mkvTrack = @{}
            if ($audioStreamIdx -lt $mkvAudioTracks.Count) {
                $mkvTrack = $mkvAudioTracks[$audioStreamIdx]
                if ($mkvTrack.ContainsKey('id')) {
                    $trackId = [int]$mkvTrack.id
                }
            }

            if ($trackId -eq -1) {
                Write-Warning "Could not map ffprobe audio stream index $streamIndex to an mkvmerge track ID. Skipping this track."
                continue
            }

            $trackTitle = ''
            if ($mkvTrack.ContainsKey('properties') -and $null -ne $mkvTrack.properties -and $mkvTrack.properties.ContainsKey('track_name')) {
                $trackTitle = [string]$mkvTrack.properties.track_name
            }

            $trackDelay = 0
            $audioTrackInfo = $null
            if ($mediainfoAudioTracks.ContainsKey($streamIndex)) {
                $audioTrackInfo = $mediainfoAudioTracks[$streamIndex]
            }

            $bitrate = 'Unknown'
            if ($null -ne $audioTrackInfo) {
                if ($audioTrackInfo.ContainsKey('BitRate')) {
                    $brValue = 0
                    if ([int]::TryParse([string]$audioTrackInfo.BitRate, [ref]$brValue)) {
                        $bitrate = "{0}k" -f [int]($brValue / 1000)
                    }
                }
                elseif ($audioTrackInfo.ContainsKey('BitRate_Nominal')) {
                    $brValue = 0
                    if ([int]::TryParse([string]$audioTrackInfo.BitRate_Nominal, [ref]$brValue)) {
                        $bitrate = "{0}k" -f [int]($brValue / 1000)
                    }
                }

                if ($audioTrackInfo.ContainsKey('Video_Delay')) {
                    $delayRaw = $audioTrackInfo.Video_Delay
                    if ($null -ne $delayRaw) {
                        $delayVal = 0.0
                        if ([double]::TryParse([string]$delayRaw, [ref]$delayVal)) {
                            if ($delayVal -lt 1 -and $delayVal -gt -1) {
                                $trackDelay = [int][math]::Round($delayVal * 1000)
                            }
                            else {
                                $trackDelay = [int][math]::Round($delayVal)
                            }
                        }
                    }
                }
            }

            $trackInfo = "Audio Stream #$streamIndex (TID: $trackId, Codec: $codec, Bitrate: $bitrate, Channels: $channels)"
            if (-not [string]::IsNullOrWhiteSpace($trackTitle)) {
                $trackInfo += ", Title: '$trackTitle'"
            }
            if ($language -ne 'und') {
                $trackInfo += ", Language: $language"
            }
            if ($trackDelay -ne 0) {
                $trackInfo += ", Delay: ${trackDelay}ms"
            }

            Write-Host "`nProcessing $trackInfo"

            if ($codec -in @('aac', 'opus')) {
                Write-Host "  -> Action: Remuxing track (keeping original $($codec.ToUpperInvariant()) $bitrate)"
            }
            else {
                $bitrateInfo = "$($codec.ToUpperInvariant()) $bitrate"
                Write-Host "  -> Action: Re-encoding codec '$codec' to Opus"
                $converted = Convert-AudioTrack -StreamIndex $streamIndex -Channels $channels -TempDir $tempDir -SourceFile $file.FullName -ShouldDownmix $Downmix.IsPresent -BitrateInfo $bitrateInfo

                $null = $processedAudioFiles.Add(@{
                    Path = $converted.Path
                    Language = $language
                    Title = $trackTitle
                    Delay = $trackDelay
                })
                $null = $tidsOfReencodedTracks.Add([string]$trackId)
            }
        }

        Write-Host "`n=== Final MKV Creation ==="
        Write-Host 'Assembling final mkvmerge command...'
        $mkvmergeArgs = @('-o', $intermediateOutputFile)

        if ($processedAudioFiles.Count -eq 0) {
            Write-Host '  -> All audio tracks are in the desired format. Performing a full remux.'
            $mkvmergeArgs += $file.FullName
        }
        else {
            $mkvmergeArgs += @('--audio-tracks', '!' + ($tidsOfReencodedTracks -join ','))
            $mkvmergeArgs += $file.FullName

            foreach ($fileInfo in $processedAudioFiles) {
                $mkvmergeArgs += @('--language', "0:$($fileInfo.Language)")
                if (-not [string]::IsNullOrWhiteSpace([string]$fileInfo.Title)) {
                    $mkvmergeArgs += @('--track-name', "0:$($fileInfo.Title)")
                }
                if ([int]$fileInfo.Delay -ne 0) {
                    $mkvmergeArgs += @('--sync', "0:$($fileInfo.Delay)")
                }
                $mkvmergeArgs += [string]$fileInfo.Path
            }
        }

        Write-Host 'Executing mkvmerge...'
        Invoke-ExternalCommand -Command 'mkvmerge' -Arguments $mkvmergeArgs | Out-Null
        Write-Host 'MKV creation complete'

        Write-Host "`n=== File Management ==="
        $completedTarget = Join-Path $DIR_COMPLETED $file.Name
        $originalTarget = Join-Path $DIR_ORIGINAL $file.Name

        Write-Host "Moving processed file to: $completedTarget"
        Move-Item -LiteralPath $intermediateOutputFile -Destination $completedTarget -Force

        Write-Host "Moving original file to: $originalTarget"
        Move-Item -LiteralPath $file.FullName -Destination $originalTarget -Force

        $runtime = (Get-Date) - $startTime
        $hours = [int][math]::Floor($runtime.TotalHours)
        $runtimeStr = '{0:00}:{1:00}:{2:00}' -f $hours, $runtime.Minutes, $runtime.Seconds
        Write-Host "`nTotal processing time: $runtimeStr"
    }
    catch {
        Write-Error "An error occurred while processing '$($file.Name)': $($_.Exception.Message)"
        if (Test-Path -LiteralPath $intermediateOutputFile) {
            Remove-Item -LiteralPath $intermediateOutputFile -Force -ErrorAction SilentlyContinue
        }
    }
    finally {
        Write-Host "`n=== Cleanup ==="
        Write-Host 'Cleaning up temporary files...'
        if ($null -ne $tempDir -and (Test-Path -LiteralPath $tempDir)) {
            Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host 'Temporary directory removed.'
        }

        try {
            Stop-Transcript | Out-Null
        }
        catch {
            # Ignore transcript stop failures to preserve main flow.
        }
    }
}
