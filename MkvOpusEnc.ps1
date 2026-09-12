<#
.SYNOPSIS
    Batch-processes MKV files in the current directory and converts non-AAC/Opus audio tracks to Opus.

.DESCRIPTION
    Mirrors the behavior of MkvOpusEnc.py:
    - Scans for *.mkv files (excluding temp-output-*).
    - AAC/Opus tracks are remuxed unchanged.
    - All other audio codecs are extracted, Nightmode-downmixed if requested,
      loudness-normalized with ffmpeg loudnorm (two-pass linear), and encoded to Opus.
    - Optional Nightmode Dialogue downmix for 5.1/7.1 and other 6+ channel layouts.
    - Preserves language, title, and delay metadata for re-encoded tracks.
    - Writes per-file logs to conv_logs, moves processed files to completed, originals to original.

    Audio normalization uses ffmpeg loudnorm two-pass linear (constant gain, true-peak aware).
    Downmix is Nightmode Dialogue (Collier / Harrelson) with pan '<' so the mix cannot clip.
    No asoftclip, no sox_ng.

.PARAMETER Downmix
    Nightmode Dialogue downmix of 5.1/7.1 to stereo (pan '<', no mix clip).

.PARAMETER NormI
    Target integrated loudness in LUFS (default: -18.0).

.PARAMETER NormTp
    True-peak ceiling in dBTP (default: -1.5).
#>

[CmdletBinding()]
param (
    [switch]$Downmix,
    [double]$NormI = -18.0,
    [double]$NormTp = -1.5
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# loudnorm max. If target LRA < measured LRA, it silently switches to dynamic (compresses).
$script:LoudnessLra = 20.0

function Invoke-ExternalCommand {
    param (
        [Parameter(Mandatory = $true)][string]$Command,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Arguments,
        [switch]$CaptureOutput,
        [switch]$CaptureStdErr,
        [switch]$NoCheck
    )

    if ($CaptureStdErr) {
        $resolved = (Get-Command $Command -ErrorAction Stop).Source
        $psi = [System.Diagnostics.ProcessStartInfo]::new()
        $psi.FileName = $resolved
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true
        foreach ($arg in $Arguments) {
            [void]$psi.ArgumentList.Add($arg)
        }

        $proc = [System.Diagnostics.Process]::new()
        $proc.StartInfo = $psi
        [void]$proc.Start()
        $stdoutTask = $proc.StandardOutput.ReadToEndAsync()
        $stderrTask = $proc.StandardError.ReadToEndAsync()
        $proc.WaitForExit()
        $stdout = $stdoutTask.GetAwaiter().GetResult()
        $stderr = $stderrTask.GetAwaiter().GetResult()

        if (-not $NoCheck -and $proc.ExitCode -ne 0) {
            throw "Command failed (exit code $($proc.ExitCode)): $Command $($Arguments -join ' ')`n$stderr"
        }

        return $stderr
    }

    if ($CaptureOutput) {
        $output = & $Command @Arguments
        $exitCode = $LASTEXITCODE
        if (-not $NoCheck -and $exitCode -ne 0) {
            throw "Command failed (exit code $exitCode): $Command $($Arguments -join ' ')"
        }
        # Wrap in @() so a single-line native result is not joined character-by-character.
        return (@($output) -join [Environment]::NewLine)
    }

    & $Command @Arguments
    $exitCode = $LASTEXITCODE
    if (-not $NoCheck -and $exitCode -ne 0) {
        throw "Command failed (exit code $exitCode): $Command $($Arguments -join ' ')"
    }

    return $null
}

function Test-RequiredTools {
    $requiredTools = @('ffmpeg', 'ffprobe', 'mkvmerge', 'opusenc', 'mediainfo')
    Write-Host '--- Prerequisite Check ---'

    $allFound = $true
    foreach ($tool in $requiredTools) {
        if (-not (Get-Command $tool -ErrorAction SilentlyContinue)) {
            Write-Host "Error: Required tool '$tool' not found." -ForegroundColor Red
            $allFound = $false
        }
    }

    if (-not $allFound) {
        throw "Please install the missing tools and ensure they are in your system's PATH."
    }

    Write-Host 'All required tools found.'
}

function Get-LoudnormJson {
    param (
        [Parameter(Mandatory = $true)][string]$StdErrOutput
    )

    $jsonStartIndex = $StdErrOutput.IndexOf('{')
    if ($jsonStartIndex -lt 0) {
        throw 'No JSON block found in ffmpeg stderr output.'
    }

    $braceLevel = 0
    $jsonEndIndex = -1
    for ($i = $jsonStartIndex; $i -lt $StdErrOutput.Length; $i++) {
        $char = $StdErrOutput[$i]
        if ($char -eq '{') {
            $braceLevel++
        }
        elseif ($char -eq '}') {
            $braceLevel--
            if ($braceLevel -eq 0) {
                $jsonEndIndex = $i + 1
                break
            }
        }
    }

    if ($jsonEndIndex -lt 0) {
        throw 'No JSON block found in ffmpeg stderr output.'
    }

    $jsonText = $StdErrOutput.Substring($jsonStartIndex, $jsonEndIndex - $jsonStartIndex)
    return ($jsonText | ConvertFrom-Json -AsHashtable)
}

function Get-FiniteFloat {
    param (
        $Value,
        $Fallback
    )

    if ($null -eq $Value) {
        return $Fallback
    }

    $parsed = 0.0
    if (-not [double]::TryParse(
            [string]$Value,
            [System.Globalization.NumberStyles]::Float,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [ref]$parsed
        )) {
        return $Fallback
    }

    if ([double]::IsNaN($parsed) -or [double]::IsInfinity($parsed)) {
        return $Fallback
    }

    return $parsed
}

function Format-InvariantFixed {
    param (
        [double]$Value,
        [int]$Decimals = 2
    )

    $format = '0.' + ('0' * $Decimals)
    return $Value.ToString($format, [System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-StreamSampleRate {
    param (
        [Parameter(Mandatory = $true)][string]$SourceFile,
        [Parameter(Mandatory = $true)][int]$StreamIndex
    )

    try {
        $raw = Invoke-ExternalCommand -Command 'ffprobe' -Arguments @(
            '-v', 'quiet', '-print_format', 'json', '-show_streams', $SourceFile
        ) -CaptureOutput
        $info = $raw | ConvertFrom-Json -AsHashtable
        foreach ($stream in @($info.streams)) {
            if (-not $stream.ContainsKey('index')) {
                continue
            }
            if ([int]$stream.index -ne $StreamIndex) {
                continue
            }

            $rateText = ''
            if ($stream.ContainsKey('sample_rate') -and $null -ne $stream.sample_rate) {
                $rateText = [string]$stream.sample_rate
            }
            if ([string]::IsNullOrWhiteSpace($rateText)) {
                break
            }

            $rate = 0
            if ([int]::TryParse($rateText.Split()[0], [ref]$rate) -and $rate -gt 0) {
                return $rate
            }
        }
    }
    catch {
        # Fall through to default.
    }

    return 48000
}

function Invoke-ConstantGainLoudness {
    param (
        [Parameter(Mandatory = $true)][string]$InputPath,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][double]$NormI,
        [Parameter(Mandatory = $true)][double]$NormTp,
        [int]$SampleRate = 48000
    )

    Write-Host '   [Norm] Pass 1 — measuring integrated loudness and true peak...'
    Write-Host (
        "   [Norm] Targets: I=$NormI LUFS, TP=$NormTp dBTP, " +
        "LRA=$($script:LoudnessLra) LU (linear; not a compressor)"
    )
    Write-Host "   [Norm] Restore sample rate after loudnorm: $SampleRate Hz (source; Opus encode stays 48 kHz)"

    $measureFilter = "loudnorm=I=${NormI}:LRA=$($script:LoudnessLra):tp=${NormTp}:print_format=json"
    $stderr = Invoke-ExternalCommand -Command 'ffmpeg' -Arguments @(
        '-hide_banner', '-v', 'info', '-i', $InputPath,
        '-af', $measureFilter,
        '-f', 'null', '-'
    ) -CaptureStdErr

    $stats = $null
    $measuredI = $null
    try {
        $stats = Get-LoudnormJson -StdErrOutput $stderr
        $measuredI = Get-FiniteFloat -Value $stats['input_i'] -Fallback $null
    }
    catch {
        Write-Host "   [Norm] WARNING: Could not parse loudnorm JSON ($($_.Exception.Message)). Falling back to copy." -ForegroundColor Yellow
        $measuredI = $null
    }

    if ($null -eq $measuredI) {
        Write-Host '   [Norm] Fallback — copying without gain adjustment.'
        Invoke-ExternalCommand -Command 'ffmpeg' -Arguments @(
            '-v', 'quiet', '-y', '-i', $InputPath, '-c:a', 'flac', $OutputPath
        ) | Out-Null
        return
    }

    $measuredTp = Get-FiniteFloat -Value $stats['input_tp'] -Fallback -99.0
    $measuredLra = Get-FiniteFloat -Value $stats['input_lra'] -Fallback 0.0
    $measuredThresh = Get-FiniteFloat -Value $stats['input_thresh'] -Fallback -70.0
    $offset = Get-FiniteFloat -Value $stats['target_offset'] -Fallback 0.0
    $gainDb = $NormI - $measuredI

    $gainText = $gainDb.ToString('+0.00;-0.00;+0.00', [System.Globalization.CultureInfo]::InvariantCulture)
    $offsetText = $offset.ToString('+0.00;-0.00;+0.00', [System.Globalization.CultureInfo]::InvariantCulture)
    Write-Host (
        "   [Norm] Measured I=$(Format-InvariantFixed $measuredI) LUFS, " +
        "TP=$(Format-InvariantFixed $measuredTp) dBTP, " +
        "LRA=$(Format-InvariantFixed $measuredLra) LU → $gainText dB (offset $offsetText)"
    )

    if ($measuredLra -gt $script:LoudnessLra) {
        Write-Host (
            "   [Norm] WARNING: source LRA $(Format-InvariantFixed $measuredLra) > $($script:LoudnessLra); " +
            'loudnorm may use dynamic mode.'
        ) -ForegroundColor Yellow
    }

    Write-Host '   [Norm] Pass 2 — loudnorm linear=true (true-peak aware, not hard clip)...'
    $loudnormApply = (
        "loudnorm=I=${NormI}:LRA=$($script:LoudnessLra):tp=${NormTp}" +
        ":measured_I=$(Format-InvariantFixed $measuredI)" +
        ":measured_LRA=$(Format-InvariantFixed $measuredLra)" +
        ":measured_TP=$(Format-InvariantFixed $measuredTp)" +
        ":measured_thresh=$(Format-InvariantFixed $measuredThresh)" +
        ":offset=$(Format-InvariantFixed $offset)" +
        ':linear=true' +
        ':print_format=summary'
    )

    Invoke-ExternalCommand -Command 'ffmpeg' -Arguments @(
        '-hide_banner', '-v', 'error', '-stats', '-y',
        '-i', $InputPath,
        '-af', "$loudnormApply,aformat=sample_fmts=s32:sample_rates=$SampleRate",
        '-ar', [string]$SampleRate,
        '-c:a', 'flac', '-sample_fmt', 's32',
        $OutputPath
    ) | Out-Null
}

function Get-DownmixFilters {
    param (
        [Parameter(Mandatory = $true)][int]$Channels
    )

    if ($Channels -eq 6) {
        return @(
            'pan=stereo|FL<FC+0.30*FL+0.30*SL|FR<FC+0.30*FR+0.30*SR',
            'pan=stereo|FL<FC+0.30*FL+0.30*BL|FR<FC+0.30*FR+0.30*BR',
            'aformat=ch_layouts=5.1,pan=stereo|FL<FC+0.30*FL+0.30*BL|FR<FC+0.30*FR+0.30*BR',
            'pan=stereo|c0<c2+0.30*c0+0.30*c4|c1<c2+0.30*c1+0.30*c5'
        )
    }
    if ($Channels -eq 8) {
        return @(
            'pan=stereo|FL<FC+0.30*FL+0.30*SL+0.30*BL|FR<FC+0.30*FR+0.30*SR+0.30*BR',
            'pan=stereo|c0<c2+0.30*c0+0.30*c4+0.30*c6|c1<c2+0.30*c1+0.30*c5+0.30*c7'
        )
    }
    return @()
}

function Convert-AudioTrack {
    param (
        [Parameter(Mandatory = $true)][int]$StreamIndex,
        [Parameter(Mandatory = $true)][int]$Channels,
        [Parameter(Mandatory = $true)][string]$TempDir,
        [Parameter(Mandatory = $true)][string]$SourceFile,
        [Parameter(Mandatory = $true)][bool]$ShouldDownmix,
        [Parameter(Mandatory = $true)][string]$BitrateInfo,
        [Parameter(Mandatory = $true)][double]$NormI,
        [Parameter(Mandatory = $true)][double]$NormTp
    )

    $tempExtracted = Join-Path $TempDir "track_${StreamIndex}_extracted.flac"
    $tempNormalized = Join-Path $TempDir "track_${StreamIndex}_normalized.flac"
    $finalOpus = Join-Path $TempDir "track_${StreamIndex}_final.opus"

    Write-Host ' - Extracting to FLAC...'
    $baseArgs = @(
        '-hide_banner', '-v', 'error', '-stats', '-y',
        '-drc_scale', '0',
        '-i', $SourceFile,
        '-map', "0:$StreamIndex",
        '-map_metadata', '-1'
    )

    $finalChannels = $Channels
    $attempts = [System.Collections.ArrayList]@()
    if ($ShouldDownmix -and $Channels -ge 6) {
        foreach ($filt in (Get-DownmixFilters -Channels $Channels)) {
            [void]$attempts.Add($filt)
        }
        [void]$attempts.Add([string]::Empty)
        $finalChannels = 2
        Write-Host " (Nightmode Dialogue downmix ${Channels}ch → stereo, pan '<')"
    }
    else {
        [void]$attempts.Add('keep')
        Write-Host " (Preserving $Channels-channel layout)"
    }

    $lastError = $null
    $extracted = $false
    $n = 0
    foreach ($filt in $attempts) {
        $n++
        $ffmpegArgs = [System.Collections.Generic.List[string]]::new()
        foreach ($arg in $baseArgs) {
            $ffmpegArgs.Add($arg)
        }

        if ($filt -eq 'keep') {
            # Preserve channel layout.
        }
        elseif ([string]::IsNullOrEmpty($filt)) {
            $ffmpegArgs.Add('-ac')
            $ffmpegArgs.Add('2')
            Write-Host '   - Downmix fallback: -ac 2'
        }
        else {
            $ffmpegArgs.Add('-af')
            $ffmpegArgs.Add($filt)
            Write-Host "   - Downmix filter (try ${n}): $filt"
        }

        $ffmpegArgs.Add('-c:a')
        $ffmpegArgs.Add('flac')
        $ffmpegArgs.Add($tempExtracted)

        try {
            Invoke-ExternalCommand -Command 'ffmpeg' -Arguments $ffmpegArgs.ToArray() | Out-Null
            $extracted = $true
            break
        }
        catch {
            $lastError = $_
            Write-Host "   - Downmix try $n failed, trying next option..."
        }
    }

    if (-not $extracted) {
        if ($null -ne $lastError) {
            throw $lastError
        }
        throw "Failed to extract audio stream $StreamIndex"
    }

    Write-Host ' - Normalizing with ffmpeg loudnorm 2-pass linear...'
    Invoke-ConstantGainLoudness `
        -InputPath $tempExtracted `
        -OutputPath $tempNormalized `
        -NormI $NormI `
        -NormTp $NormTp `
        -SampleRate (Get-StreamSampleRate -SourceFile $SourceFile -StreamIndex $StreamIndex)

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

    Write-Host " - Encoding to Opus at $bitrate..."
    Write-Host " Source: $BitrateInfo -> Destination: Opus $bitrate ($finalChannels channels)"
    Invoke-ExternalCommand -Command 'opusenc' -Arguments @(
        '--vbr', '--bitrate', $bitrate, $tempNormalized, $finalOpus
    ) | Out-Null

    return @{
        Path           = $finalOpus
        FinalChannels  = $finalChannels
        Bitrate        = $bitrate
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
        Write-Host "Normalization target: $NormI LUFS  |  True-peak ceiling: $NormTp dBTP"
        $startTime = Get-Date

        $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("mkvopusenc_{0}" -f ([System.Guid]::NewGuid().ToString('N')))
        $null = New-Item -ItemType Directory -Path $tempDir -Force
        Write-Host "Temporary directory for audio created at: $tempDir"

        Write-Host "Analyzing file: $($file.FullName)"
        $ffprobeInfoJson = Invoke-ExternalCommand -Command 'ffprobe' -Arguments @(
            '-v', 'quiet', '-print_format', 'json', '-show_streams', '-show_format', $file.FullName
        ) -CaptureOutput
        $ffprobeInfo = $ffprobeInfoJson | ConvertFrom-Json -AsHashtable

        $mkvmergeInfoJson = Invoke-ExternalCommand -Command 'mkvmerge' -Arguments @('-J', $file.FullName) -CaptureOutput
        $mkvInfo = $mkvmergeInfoJson | ConvertFrom-Json -AsHashtable

        $mediainfoJson = Invoke-ExternalCommand -Command 'mediainfo' -Arguments @('--Output=JSON', '-f', $file.FullName) -CaptureOutput
        $mediaInfo = $mediainfoJson | ConvertFrom-Json -AsHashtable

        $processedAudioFiles = [System.Collections.ArrayList]@()
        $tidsOfReencodedTracks = [System.Collections.ArrayList]@()

        $audioStreams = @($ffprobeInfo.streams | Where-Object { $_.codec_type -eq 'audio' })
        if ($audioStreams.Count -eq 0) {
            Write-Host "Warning: No audio streams found in '$($file.Name)'. Skipping file."
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
                Write-Host " -> Warning: Could not map ffprobe audio stream index $streamIndex to an mkvmerge track ID. Skipping this track."
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
                foreach ($key in @('BitRate', 'BitRate_Nominal')) {
                    if ($audioTrackInfo.ContainsKey($key)) {
                        $brValue = 0
                        if ([int]::TryParse([string]$audioTrackInfo[$key], [ref]$brValue)) {
                            $bitrate = "{0}k" -f [int][math]::Floor($brValue / 1000)
                            break
                        }
                    }
                }

                if ($audioTrackInfo.ContainsKey('Video_Delay')) {
                    $delayRaw = $audioTrackInfo.Video_Delay
                    if ($null -ne $delayRaw) {
                        $delayVal = 0.0
                        if ([double]::TryParse(
                                [string]$delayRaw,
                                [System.Globalization.NumberStyles]::Float,
                                [System.Globalization.CultureInfo]::InvariantCulture,
                                [ref]$delayVal
                            )) {
                            if ([math]::Abs($delayVal) -lt 1) {
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
                Write-Host " -> Action: Remuxing track (keeping original $($codec.ToUpperInvariant()) $bitrate)"
            }
            else {
                $bitrateInfo = "$($codec.ToUpperInvariant()) $bitrate"
                Write-Host " -> Action: Re-encoding codec '$codec' to Opus"
                $converted = Convert-AudioTrack `
                    -StreamIndex $streamIndex `
                    -Channels $channels `
                    -TempDir $tempDir `
                    -SourceFile $file.FullName `
                    -ShouldDownmix $Downmix.IsPresent `
                    -BitrateInfo $bitrateInfo `
                    -NormI $NormI `
                    -NormTp $NormTp

                $null = $processedAudioFiles.Add(@{
                    Path     = $converted.Path
                    Language = $language
                    Title    = $trackTitle
                    Delay    = $trackDelay
                })
                $null = $tidsOfReencodedTracks.Add([string]$trackId)
            }
        }

        Write-Host "`n=== Final MKV Creation ==="
        Write-Host 'Assembling final mkvmerge command...'
        $mkvmergeArgs = @('-o', $intermediateOutputFile)

        if ($processedAudioFiles.Count -eq 0) {
            Write-Host ' -> All audio tracks are in the desired format. Performing a full remux.'
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
        Write-Host "`nAn error occurred while processing '$($file.Name)': $($_.Exception.Message)" -ForegroundColor Red
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
