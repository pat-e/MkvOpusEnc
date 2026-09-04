#!/usr/bin/env python3
"""
Processes or downmixes an MKV file's audio tracks sequentially using a specific toolchain.
This script is cross-platform and optimized for correctness and clean output.

This script intelligently handles audio streams in an MKV file one by one.
- AAC/Opus audio is remuxed.
- Multi-channel audio (DTS, AC3, etc.) can be re-encoded or optionally downmixed to stereo.
- All other streams and metadata (title, language, delay) are preserved.

Audio normalization uses ffmpeg loudnorm two-pass linear (constant gain, true-peak aware).
Downmix is Nightmode Dialogue (Collier / Harrelson) with pan '<' so the mix cannot clip.
No asoftclip, no sox_ng.
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
# loudnorm max. If target LRA < measured LRA, it silently switches to dynamic (compresses).
LOUDNESS_LRA = 20.0


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
    """Two-pass ffmpeg loudnorm, linear (constant gain + true-peak). No asoftclip."""
    print(f"   [Norm] Pass 1 — measuring integrated loudness and true peak...")
    print(
        f"   [Norm] Targets: I={norm_i} LUFS, TP={norm_tp} dBTP, "
        f"LRA={LOUDNESS_LRA} LU (linear; not a compressor)"
    )
    result = subprocess.run(
        [
            "ffmpeg", "-hide_banner", "-v", "info", "-i", str(input_path),
            "-af", (
                f"loudnorm=I={norm_i}:LRA={LOUDNESS_LRA}:tp={norm_tp}"
                f":print_format=json"
            ),
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
        print(f"   [Norm] Fallback — copying without gain adjustment.")
        run_cmd(["ffmpeg", "-v", "quiet", "-y", "-i", str(input_path), "-c:a", "flac", str(output_path)])
        return

    measured_tp = _finite_float(stats.get("input_tp"), -99.0)
    measured_lra = _finite_float(stats.get("input_lra"), 0.0)
    measured_thresh = _finite_float(stats.get("input_thresh"), -70.0)
    offset = _finite_float(stats.get("target_offset"), 0.0)
    gain_db = norm_i - measured_i
    print(
        f"   [Norm] Measured I={measured_i:.2f} LUFS, TP={measured_tp:.2f} dBTP, "
        f"LRA={measured_lra:.2f} LU → {gain_db:+.2f} dB (offset {offset:+.2f})"
    )
    if measured_lra > LOUDNESS_LRA:
        print(
            f"   [Norm] WARNING: source LRA {measured_lra:.2f} > {LOUDNESS_LRA}; "
            "loudnorm may use dynamic mode."
        )
    print(f"   [Norm] Pass 2 — loudnorm linear=true (true-peak aware, not hard clip)...")
    loudnorm_apply = (
        f"loudnorm=I={norm_i}:LRA={LOUDNESS_LRA}:tp={norm_tp}"
        f":measured_I={measured_i:.2f}"
        f":measured_LRA={measured_lra:.2f}"
        f":measured_TP={measured_tp:.2f}"
        f":measured_thresh={measured_thresh:.2f}"
        f":offset={offset:.2f}"
        f":linear=true"
        f":print_format=summary"
    )
    run_cmd([
        "ffmpeg", "-hide_banner", "-v", "error", "-stats", "-y",
        "-i", str(input_path),
        "-af", f"{loudnorm_apply},aformat=sample_fmts=s32",
        "-c:a", "flac", "-sample_fmt", "s32",
        str(output_path),
    ])


# ---------------------------------------------------------------------------
# Audio track conversion
# ---------------------------------------------------------------------------

def downmix_filters(ch):
    """Nightmode Dialogue (Collier / Harrelson). pan '<' renormalizes so the mix cannot clip."""
    if ch == 6:
        return [
            "pan=stereo|FL<FC+0.30*FL+0.30*SL|FR<FC+0.30*FR+0.30*SR",
            "pan=stereo|FL<FC+0.30*FL+0.30*BL|FR<FC+0.30*FR+0.30*BR",
            "aformat=ch_layouts=5.1,pan=stereo|FL<FC+0.30*FL+0.30*BL|FR<FC+0.30*FR+0.30*BR",
            "pan=stereo|c0<c2+0.30*c0+0.30*c4|c1<c2+0.30*c1+0.30*c5",
        ]
    if ch == 8:
        return [
            "pan=stereo|FL<FC+0.30*FL+0.30*SL+0.30*BL|FR<FC+0.30*FR+0.30*SR+0.30*BR",
            "pan=stereo|c0<c2+0.30*c0+0.30*c4+0.30*c6|c1<c2+0.30*c1+0.30*c5+0.30*c7",
        ]
    return []


def convert_audio_track(stream_index, channels, temp_dir, source_file,
                        should_downmix, bitrate_info, norm_i, norm_tp):
    """Extracts, Nightmode-downmixes if requested, loudnorm-linear, encodes Opus."""
    temp_extracted  = temp_dir / f"track_{stream_index}_extracted.flac"
    temp_normalized = temp_dir / f"track_{stream_index}_normalized.flac"
    final_opus      = temp_dir / f"track_{stream_index}_final.opus"

    print(" - Extracting to FLAC...")
    base_args = [
        "ffmpeg", "-hide_banner", "-v", "error", "-stats", "-y",
        "-drc_scale", "0",
        "-i", str(source_file),
        "-map", f"0:{stream_index}",
        "-map_metadata", "-1",
    ]

    final_channels = channels
    attempts = []
    if should_downmix and channels >= 6:
        attempts.extend(downmix_filters(channels))
        attempts.append(None)
        final_channels = 2
        print(f" (Nightmode Dialogue downmix {channels}ch → stereo, pan '<')")
    else:
        attempts.append("keep")
        print(f" (Preserving {channels}-channel layout)")

    last_error = None
    extracted = False
    for n, filt in enumerate(attempts, start=1):
        ffmpeg_args = list(base_args)
        if filt == "keep":
            pass
        elif filt is None:
            ffmpeg_args += ["-ac", "2"]
            print("   - Downmix fallback: -ac 2")
        else:
            ffmpeg_args += ["-af", filt]
            print(f"   - Downmix filter (try {n}): {filt}")
        ffmpeg_args += ["-c:a", "flac", str(temp_extracted)]
        try:
            run_cmd(ffmpeg_args)
            extracted = True
            break
        except subprocess.CalledProcessError as e:
            last_error = e
            print(f"   - Downmix try {n} failed, trying next option...")
    if not extracted:
        raise last_error

    print(" - Normalizing with ffmpeg loudnorm 2-pass linear...")
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
        help="Nightmode Dialogue downmix of 5.1/7.1 to stereo (pan '<', no mix clip).",
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
