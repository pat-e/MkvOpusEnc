# Process-MkvAdvanced
MKV Audio to Opus with Downmix-Option in Powershell

This script will encode all audio tracks inside a MKV to Opus and adjust the bitrate based on the channel layout:

2 channels (stereo) = 128 kbit/s

6 channels (5.1) = 256 kbit/s

8 channels (7.1) = 384 kbit/s

It will also do audio normalization using "sox" with clipping-guard.

It also includes a downmix - option to downmix all channels to stereo with a "Dialogue Boost Filter" based on "Robert Collier's Nightmode Dialogue" (https://superuser.com/questions/852400/properly-downmix-5-1-to-stereo-using-ffmpeg).

# Requirements:

For this to work, you need the following tools in our PATH variable:
"ffmpeg", "ffprobe", "mkvmerge", "sox", "opusenc", "mediainfo"
- "ffmpeg" and "ffprobe" - https://ffmpeg.org/
- "mkvmerge" - https://mkvtoolnix.org/
- "sox" - http://sox.sourceforge.net/
- "opusenc" - https://github.com/Chocobo1/opus-tools_win32-build
- "mediainfo" - https://mediaarea.net/en/MediaInfo

# Usage:

     Regular mode - preserves channel layouts
    .\Process-MkvAdvanced.ps1 -InputFile "C:\path\to\movie.mkv" -OutputFile "C:\path\to\regular.mkv"

      Downmix mode - converts multi-channel audio to a dialogue-boosted stereo track
    .\Process-MkvAdvanced.ps1 -InputFile "C:\path\to\movie.mkv" -OutputFile "C:\path\to\downmixed.mkv" -Downmix

# Batch - Mode:

     foreach ($file in gci *.mkv ) {$new = $file.BaseName + "_opus" + $file.Extension ; Process-MkvAdvanced.ps1 -InputFile $file -OutputFile $new }

     
# Batch - Mode down-mixing

     foreach ($file in gci *.mkv ) {$new = $file.BaseName + "_opus_downmixed" + $file.Extension ; Process-MkvAdvanced.ps1 -InputFile $file -OutputFile $new -Downmix }
