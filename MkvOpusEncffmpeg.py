#!/usr/bin/env python3
"""
Processes or downmixes an MKV file's audio tracks sequentially using a specific toolchain.
This script is cross-platform and optimized for correctness and clean output.

This script intelligently handles audio streams in an MKV file one by one.
- AAC/Opus audio is remuxed.
- Multi-channel audio (DTS, AC3, etc.) can be re-encoded or optionally downmixed to stereo.
- All other streams and metadata (title, language, delay) are preserved.

Audio normalization uses ffmpeg loudnorm (two-pass constant-gain, no dynamic compression),
matching the xav_automation.py philosophy: measure integrated LUFS, apply one linear gain,
brickwall-clamp true peak via asoftclip. No sox_ng required.
"""

import argparse
import json
import math
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Loudness normalization constants (can be overridden via --norm-i / --norm-tp)
# ---------------------------------------------------------------------------
LOUDNESS_I  = -16.0   # Target integrated loudness (LUFS)
LOUDNESS_TP = -1.5    # True-peak ceiling (dBTP)
# No LRA target — constant-gain only, no dynamic compression


class Tee:
    def __init__(self, *files):
        self.files = files
    def write(self, obj):
        for f in self.files:
            f.write(obj)
            f.flush()
    def flush(self):
        for f in self.files:
            f.flush()


def check_tools():
    """Checks if all required command-line tools are in the system's PATH."""
    required_tools = ["ffmpeg", "ffprobe", "mkvmerge", "opusenc", "mediainfo"]
    print("--- Prerequisite Check ---")
    all_found = True
    for tool in required_tools:
        if not shutil.which(tool):
            print(f"Error: Required tool '{tool}' not found.", file=sys.stderr)
            all_found = False
    if not all_found:
        sys.exit("Please install the missing tools and ensure they are in your system's PATH.")
    print("All required tools found.")


def run_cmd(args, capture_output=False, check=True):
    """Helper function to run a command and return its output."""
    process = subprocess.run(args, capture_output=capture_output, text=True, encoding="utf-8", check=check)
    return process.stdout


# ---------------------------------------------------------------------------
# Loudness normalization helpers (ported from xav_automation.py)
# ---------------------------------------------------------------------------

def _parse_loudnorm_json(stderr_output):
    """Finds and extracts the JSON block ffmpeg's loudnorm filter writes to stderr."""
    json_start_index = stderr_output.find("{")
    if json_start_index == -1:
        raise ValueError("No JSON block found in ffmpeg stderr output.")
    brace_level = 0
    json_end_index = -1
    for i, char in enumerate(stderr_output[json_start_index:]):
        if char == "{":
            brace_level += 1
        elif char == "}":
            brace_level -= 1
        if brace_level == 0:
            json_end_index = json_start_index + i + 1
            break
    return json.loads(stderr_output[json_start_index:json_end_index])


def _finite_float(value, fallback):
    """Returns fallback if value is NaN, ±inf, or unparseable."""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return fallback
    return number if math.isfinite(number) else fallback


def apply_constant_gain_loudness(input_path, output_path, norm_i, norm_tp):
    """
    Measure integrated LUFS, apply one linear gain, brickwall-clamp true peak.
    Constant-gain only — no LRA / dynamic range compression.

    Pass 1: ffmpeg loudnorm in analysis-only mode measures input_i (integrated LUFS).
    Pass 2: volume= applies the gain, asoftclip=type=hard clamps true peaks,
            output is written as 32-bit signed FLAC.

    Falls back to a plain FLAC copy if measurement fails.
    """
    # --- Pass 1: Measure integrated loudness ---
    print(f"   [Norm] Pass 1 — measuring integrated loudness...")
    result = subprocess.run(
        [
            "ffmpeg", "-v", "info", "-i", str(input_path),
            "-af", f"loudnorm=I={norm_i}:LRA=20:tp={norm_tp}:print_format=json",
            "-f", "null", "-",
        ],
        capture_output=True, text=True, encoding="utf-8", check=True,
    )

    try:
        stats = _parse_loudnorm_json(result.stderr)
        measured_i = _finite_float(stats.get("input_i"), None)
    except Exception as exc:
        print(f"   [Norm] WARNING: Could not parse loudnorm JSON ({exc}). Falling back to copy.", file=sys.stderr)
        measured_i = None

    if measured_i is None:
        # Fallback: copy without gain if measurement fails
        print(f"   [Norm] Fallback — copying without gain adjustment.")
        run_cmd(["ffmpeg", "-v", "quiet", "-y", "-i", str(input_path), "-c:a", "flac", str(output_path)])
        return

    # --- Compute gain ---
    gain_db   = norm_i - measured_i           # e.g. -16.0 - (-20.0) = +4.0 dB
    gain      = 10 ** (gain_db / 20.0)        # linear gain factor
    tp_linear = 10 ** (norm_tp / 20.0)        # brickwall ceiling in linear scale
    print(f"   [Norm] Measured: {measured_i:.2f} LUFS  |  Gain: {gain_db:+.2f} dB  |  TP ceiling: {norm_tp} dBTP")

    # --- Pass 2: Apply gain + brickwall true-peak clamp ---
    print(f"   [Norm] Pass 2 — applying gain and true-peak clamp...")
    run_cmd([
        "ffmpeg", "-hide_banner", "-v", "error", "-stats", "-y",
        "-i", str(input_path),
        "-af", (
            f"volume={gain:.10f},"
            f"asoftclip=type=hard:threshold={tp_linear:.10f},"
            f"aformat=sample_fmts=s32"
        ),
        "-c:a", "flac", "-sample_fmt", "s32",
        str(output_path),
    ])


# ---------------------------------------------------------------------------
# Audio track conversion
# ---------------------------------------------------------------------------

def convert_audio_track(stream_index, channels, temp_dir, source_file,
                        should_downmix, bitrate_info, norm_i, norm_tp):
    """Extracts, normalizes, and encodes a single audio track to Opus."""
    temp_extracted  = temp_dir / f"track_{stream_index}_extracted.flac"
    temp_normalized = temp_dir / f"track_{stream_index}_normalized.flac"
    final_opus      = temp_dir / f"track_{stream_index}_final.opus"

    # --- Step 1: Extract audio, with conditional downmixing ---
    print(" - Extracting to FLAC...")
    ffmpeg_args = [
        "ffmpeg", "-v", "quiet", "-stats", "-y",
        "-i", str(source_file),
        "-map", f"0:{stream_index}",
    ]

    final_channels = channels
    if should_downmix and channels >= 6:
        if channels == 6:   # 5.1
            print(" (Downmixing 5.1 to Stereo with dialogue boost)")
            ffmpeg_args.extend(["-af", "pan=stereo|c0=c2+0.30*c0+0.30*c4|c1=c2+0.30*c1+0.30*c5"])
            final_channels = 2
        elif channels == 8:  # 7.1
            print(" (Downmixing 7.1 to Stereo with dialogue boost)")
            ffmpeg_args.extend(["-af", "pan=stereo|c0=c2+0.30*c0+0.30*c4+0.30*c6|c1=c2+0.30*c1+0.30*c5+0.30*c7"])
            final_channels = 2
        else:
            print(f" ({channels}-channel source, downmixing to stereo using -ac 2)")
            ffmpeg_args.extend(["-ac", "2"])
            final_channels = 2
    else:
        print(f" (Preserving {channels}-channel layout)")

    # Output codec + path are common to all branches (fixes missing output in downmix paths)
    ffmpeg_args.extend(["-c:a", "flac", str(temp_extracted)])
    run_cmd(ffmpeg_args)

    # --- Step 2: Normalize with ffmpeg loudnorm (constant-gain, no LRA) ---
    print(" - Normalizing with ffmpeg (constant-gain LUFS normalization)...")
    apply_constant_gain_loudness(temp_extracted, temp_normalized, norm_i, norm_tp)

    # --- Step 3: Encode to Opus with the correct bitrate ---
    bitrate = "192k"  # Fallback
    if final_channels == 1:
        bitrate = "64k"
    elif final_channels == 2:
        bitrate = "128k"
    elif final_channels == 6:
        bitrate = "256k"
    elif final_channels == 8:
        bitrate = "384k"

    print(f" - Encoding to Opus at {bitrate}...")
    print(f" Source: {bitrate_info} -> Destination: Opus {bitrate} ({final_channels} channels)")
    run_cmd(["opusenc", "--vbr", "--bitrate", bitrate, str(temp_normalized), str(final_opus)])
    return final_opus, final_channels, bitrate


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    """Main script logic."""
    parser = argparse.ArgumentParser(
        description="Batch processes MKV file audio tracks to Opus using ffmpeg loudnorm normalization."
    )
    parser.add_argument(
        "--downmix", action="store_true",
        help="If present, multi-channel audio will be downmixed to stereo.",
    )
    parser.add_argument(
        "--norm-i", type=float, default=LOUDNESS_I,
        metavar="LUFS",
        help=f"Target integrated loudness in LUFS (default: {LOUDNESS_I}).",
    )
    parser.add_argument(
        "--norm-tp", type=float, default=LOUDNESS_TP,
        metavar="dBTP",
        help=f"True-peak ceiling in dBTP (default: {LOUDNESS_TP}).",
    )
    args = parser.parse_args()

    check_tools()

    # Define directory paths but don't create them yet
    DIR_COMPLETED = Path("completed")
    DIR_ORIGINAL  = Path("original")
    DIR_LOGS      = Path("conv_logs")
    current_dir   = Path(".")

    # Check if there are any MKV files to process
    files_to_process = sorted(
        f for f in current_dir.glob("*.mkv") if not f.name.startswith("temp-output-")
    )

    if not files_to_process:
        print("No MKV files found to process. Exiting.")
        return

    # Create directories only when we actually have files to process
    DIR_COMPLETED.mkdir(exist_ok=True)
    DIR_ORIGINAL.mkdir(exist_ok=True)
    DIR_LOGS.mkdir(exist_ok=True)

    for file_path in files_to_process:
        log_file_path = DIR_LOGS / f"{file_path.name}.log"
        log_file = open(log_file_path, "w", encoding="utf-8")
        original_stdout = sys.stdout
        original_stderr = sys.stderr
        sys.stdout = Tee(original_stdout, log_file)
        sys.stderr = Tee(original_stderr, log_file)

        intermediate_output_file = current_dir / f"temp-output-{file_path.name}"
        temp_dir = None

        try:
            print("-" * shutil.get_terminal_size(fallback=(80, 24)).columns)
            print(f"Starting processing for: {file_path.name}")
            print(f"Log file: {log_file_path}")
            print(f"Normalization target: {args.norm_i} LUFS  |  True-peak ceiling: {args.norm_tp} dBTP")
            start_time = datetime.now()
            temp_dir = Path(tempfile.mkdtemp(prefix="mkvopusenc_"))
            print(f"Temporary directory for audio created at: {temp_dir}")

            # --- Get Media Information ---
            print(f"Analyzing file: {file_path}")
            ffprobe_info_json  = run_cmd(["ffprobe", "-v", "quiet", "-print_format", "json", "-show_streams", "-show_format", str(file_path)], capture_output=True)
            ffprobe_info       = json.loads(ffprobe_info_json)
            mkvmerge_info_json = run_cmd(["mkvmerge", "-J", str(file_path)], capture_output=True)
            mkv_info           = json.loads(mkvmerge_info_json)
            mediainfo_json_str = run_cmd(["mediainfo", "--Output=JSON", "-f", str(file_path)], capture_output=True)
            media_info         = json.loads(mediainfo_json_str)

            # --- Prepare for Final mkvmerge Command ---
            processed_audio_files      = []
            tids_of_reencoded_tracks   = []

            # --- Process Each Audio Stream ---
            audio_streams = [s for s in ffprobe_info.get("streams", []) if s.get("codec_type") == "audio"]

            if not audio_streams:
                print(f"Warning: No audio streams found in '{file_path.name}'. Skipping file.")
                continue

            mkv_tracks_list      = mkv_info.get("tracks", [])
            mkv_audio_tracks     = [t for t in mkv_tracks_list if t.get("type") == "audio"]
            media_tracks_data    = media_info.get("media", {}).get("track", [])
            mediainfo_audio_tracks = {
                int(t.get("StreamOrder", -1)): t
                for t in media_tracks_data if t.get("@type") == "Audio"
            }

            print("\n=== Audio Track Analysis ===")
            for audio_stream_idx, stream in enumerate(audio_streams):
                stream_index = stream["index"]
                codec        = stream.get("codec_name")
                channels     = stream.get("channels", 2)
                language     = stream.get("tags", {}).get("language", "und")
                track_id     = -1
                mkv_track    = {}

                if audio_stream_idx < len(mkv_audio_tracks):
                    mkv_track = mkv_audio_tracks[audio_stream_idx]
                    track_id  = mkv_track.get("id", -1)

                if track_id == -1:
                    print(f" -> Warning: Could not map ffprobe audio stream index {stream_index} to an mkvmerge track ID. Skipping this track.")
                    continue

                track_title = mkv_track.get("properties", {}).get("track_name", "")
                track_delay = 0
                audio_track_info = mediainfo_audio_tracks.get(stream_index)

                # Get bitrate information from mediainfo
                bitrate = "Unknown"
                if audio_track_info:
                    for key in ("BitRate", "BitRate_Nominal"):
                        if key in audio_track_info:
                            try:
                                bitrate = f"{int(audio_track_info[key]) // 1000}k"
                                break
                            except (ValueError, TypeError):
                                pass

                delay_raw = audio_track_info.get("Video_Delay") if audio_track_info else None
                if delay_raw is not None:
                    try:
                        delay_val = float(delay_raw)
                        track_delay = int(round(delay_val * 1000)) if abs(delay_val) < 1 else int(round(delay_val))
                    except Exception:
                        track_delay = 0

                track_info = f"Audio Stream #{stream_index} (TID: {track_id}, Codec: {codec}, Bitrate: {bitrate}, Channels: {channels})"
                if track_title:
                    track_info += f", Title: '{track_title}'"
                if language != "und":
                    track_info += f", Language: {language}"
                if track_delay != 0:
                    track_info += f", Delay: {track_delay}ms"

                print(f"\nProcessing {track_info}")

                if codec in {"aac", "opus"}:
                    print(f" -> Action: Remuxing track (keeping original {codec.upper()} {bitrate})")
                else:
                    bitrate_info = f"{codec.upper()} {bitrate}"
                    print(f" -> Action: Re-encoding codec '{codec}' to Opus")
                    opus_file, final_channels, final_bitrate = convert_audio_track(
                        stream_index, channels, temp_dir, file_path,
                        args.downmix, bitrate_info,
                        args.norm_i, args.norm_tp,
                    )
                    processed_audio_files.append({
                        "Path":     opus_file,
                        "Language": language,
                        "Title":    track_title,
                        "Delay":    track_delay,
                    })
                    tids_of_reencoded_tracks.append(str(track_id))

            # --- Construct and Execute Final mkvmerge Command ---
            print("\n=== Final MKV Creation ===")
            print("Assembling final mkvmerge command...")
            mkvmerge_args = ["mkvmerge", "-o", str(intermediate_output_file)]

            if not processed_audio_files:
                print(" -> All audio tracks are in the desired format. Performing a full remux.")
                mkvmerge_args.append(str(file_path))
            else:
                mkvmerge_args.extend(["--audio-tracks", "!" + ",".join(tids_of_reencoded_tracks)])
                mkvmerge_args.append(str(file_path))

                for file_info in processed_audio_files:
                    mkvmerge_args.extend(["--language", f"0:{file_info['Language']}"])
                    if file_info["Title"]:
                        mkvmerge_args.extend(["--track-name", f"0:{file_info['Title']}"])
                    if file_info["Delay"] != 0:
                        mkvmerge_args.extend(["--sync", f"0:{file_info['Delay']}"])
                    mkvmerge_args.append(str(file_info["Path"]))

            print("Executing mkvmerge...")
            run_cmd(mkvmerge_args)
            print("MKV creation complete")

            # Move files to their final destinations
            print("\n=== File Management ===")
            print(f"Moving processed file to: {DIR_COMPLETED / file_path.name}")
            shutil.move(str(intermediate_output_file), DIR_COMPLETED / file_path.name)
            print(f"Moving original file to: {DIR_ORIGINAL / file_path.name}")
            shutil.move(str(file_path), DIR_ORIGINAL / file_path.name)

            runtime = datetime.now() - start_time
            print(f"\nTotal processing time: {str(runtime).split('.')[0]}")

        except Exception as e:
            print(f"\nAn error occurred while processing '{file_path.name}': {e}", file=sys.stderr)
            if intermediate_output_file.exists():
                intermediate_output_file.unlink()

        finally:
            print("\n=== Cleanup ===")
            print("Cleaning up temporary files...")
            if temp_dir is not None and temp_dir.exists():
                shutil.rmtree(temp_dir)
            print("Temporary directory removed.")

            sys.stdout = original_stdout
            sys.stderr = original_stderr
            log_file.close()


if __name__ == "__main__":
    main()
