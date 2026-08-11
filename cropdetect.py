#!/usr/bin/env python3

import argparse
import subprocess
import sys
import os
import re
from collections import Counter
import shutil
import multiprocessing
import json

# ANSI color codes
COLOR_GREEN = "\033[92m"
COLOR_RED = "\033[91m"
COLOR_YELLOW = "\033[93m"
COLOR_RESET = "\033[0m"

def check_prerequisites():
    """Checks if required tools are available."""
    print("--- Prerequisite Check ---")
    all_found = True
    for tool in ['ffmpeg', 'ffprobe']:
        if not shutil.which(tool):
            print(f"Error: '{tool}' command not found. Is it installed and in your PATH?")
            all_found = False
    if not all_found:
        sys.exit(1)
    print("All required tools found.")

def analyze_segment(task_args):
    """Function to be run by each worker process. Analyzes one video segment."""
    seek_time, input_file, width, height = task_args
    
    ffmpeg_args = [
        'ffmpeg', '-hide_banner',
        '-ss', str(seek_time),
        '-i', input_file, '-t', '1', '-vf', 'cropdetect',
        '-f', 'null', '-'
    ]

    result = subprocess.run(ffmpeg_args, capture_output=True, text=True, encoding='utf-8')
    
    if result.returncode != 0:
        return [] # Return empty list on error

    crop_detections = re.findall(r'crop=(\d+):(\d+):(\d+):(\d+)', result.stderr)
    
    significant_crops = []
    for w_str, h_str, x_str, y_str in crop_detections:
        w, h, x, y = map(int, [w_str, h_str, x_str, y_str])
        
        # Return the crop string along with the timestamp it was found at
        significant_crops.append((f"crop={w}:{h}:{x}:{y}", seek_time))
        
    return significant_crops

def get_frame_luma(input_file, seek_time):
    """Analyzes a single frame at a given timestamp to get its average luma."""
    ffmpeg_args = [
        'ffmpeg', '-hide_banner',
        '-ss', str(seek_time),
        '-i', input_file,
        '-t', '1',
        '-vf', 'signalstats',
        '-f', 'null', '-'
    ]
    result = subprocess.run(ffmpeg_args, capture_output=True, text=True, encoding='utf-8')
    
    if result.returncode != 0:
        return None # Error during analysis

    # Find the average luma (YAVG) for the frame
    match = re.search(r'YAVG:([0-9.]+)', result.stderr)
    if match:
        return float(match.group(1))
    
    return None

def check_luma_for_group(task_args):
    """Worker function to check the luma for a single group."""
    group_key, sample_ts, input_file, luma_threshold = task_args
    luma = get_frame_luma(input_file, sample_ts)
    is_bright = luma is not None and luma >= luma_threshold
    return (group_key, is_bright)

KNOWN_ASPECT_RATIOS = [
    {"name": "HDTV (16:9)", "ratio": 16/9},
    {"name": "Widescreen (Scope)", "ratio": 2.39},
    {"name": "Widescreen (Flat)", "ratio": 1.85},
    {"name": "IMAX Digital (1.90:1)", "ratio": 1.90},
    {"name": "Fullscreen (4:3)", "ratio": 4/3},
    {"name": "IMAX 70mm (1.43:1)", "ratio": 1.43},
]

def snap_to_known_ar(w, h, x, y, video_w, video_h, tolerance=0.03):
    """Snaps a crop rectangle to the nearest standard aspect ratio if it's close enough."""
    if h == 0: return f"crop={w}:{h}:{x}:{y}", None
    detected_ratio = w / h
    
    best_match = None
    smallest_diff = float('inf')

    for ar in KNOWN_ASPECT_RATIOS:
        diff = abs(detected_ratio - ar['ratio'])
        if diff < smallest_diff:
            smallest_diff = diff
            best_match = ar

    # If the best match is not within the tolerance, return the original
    if not best_match or (smallest_diff / best_match['ratio']) >= tolerance:
        return f"crop={w}:{h}:{x}:{y}", None

    # Match found, now snap the dimensions.
    # Heuristic: if width is close to full video width, it's letterboxed.
    if abs(w - video_w) < 16:
        new_h = round(video_w / best_match['ratio'])
        
        # Round height up to the nearest multiple of 8 for cleaner dimensions and less aggressive cropping.
        if new_h % 8 != 0:
            new_h = new_h + (8 - (new_h % 8))

        new_y = round((video_h - new_h) / 2)
        # Ensure y offset is an even number for compatibility.
        if new_y % 2 != 0:
            new_y -= 1
            
        return f"crop={video_w}:{new_h}:0:{new_y}", best_match['name']
    
    # Heuristic: if height is close to full video height, it's pillarboxed.
    if abs(h - video_h) < 16:
        new_w = round(video_h * best_match['ratio'])

        # Round width up to the nearest multiple of 8.
        if new_w % 8 != 0:
            new_w = new_w + (8 - (new_w % 8))

        new_x = round((video_w - new_w) / 2)
        # Ensure x offset is an even number.
        if new_x % 2 != 0:
            new_x -= 1

        return f"crop={new_w}:{video_h}:{new_x}:0", best_match['name']

    # If not clearly letterboxed or pillarboxed, don't snap.
    return f"crop={w}:{h}:{x}:{y}", None

def cluster_crop_values(crop_counts, tolerance=8):
    """Groups similar crop values into clusters based on the top-left corner."""
    clusters = []
    temp_counts = crop_counts.copy()

    while temp_counts:
        # Get the most frequent remaining crop as the new cluster center
        center_str, _ = temp_counts.most_common(1)[0]
        
        try:
            _, values = center_str.split('=')
            cw, ch, cx, cy = map(int, values.split(':'))
        except (ValueError, IndexError):
            del temp_counts[center_str] # Skip malformed strings
            continue

        cluster_total_count = 0
        crops_to_remove = []

        # Find all crops "close" to the center
        for crop_str, count in temp_counts.items():
            try:
                _, values = crop_str.split('=')
                w, h, x, y = map(int, values.split(':'))
                if abs(x - cx) <= tolerance and abs(y - cy) <= tolerance:
                    cluster_total_count += count
                    crops_to_remove.append(crop_str)
            except (ValueError, IndexError):
                continue
        
        if cluster_total_count > 0:
            clusters.append({'center': center_str, 'count': cluster_total_count})

        # Remove the clustered crops from the temporary counter
        for crop_str in crops_to_remove:
            del temp_counts[crop_str]
            
    clusters.sort(key=lambda c: c['count'], reverse=True)
    return clusters

# Main execution block
if __name__ == '__main__':
    os.system('') # Enable ANSI escape codes on Windows
    # Calculate a sensible default for parallel jobs: half of available cores, but at least 1.
    default_jobs = max(1, os.cpu_count() // 2)

    # 1. Set up the argument parser
    parser = argparse.ArgumentParser(
        description="Analyzes a video file in parallel to detect black bars and suggests a crop filter.",
        epilog="The script will output the most frequent crop value found across all samples."
    )
    parser.add_argument('input_file', help='The path to the input video file.')
    parser.add_argument('-j', '--jobs', type=int, default=default_jobs, 
                        help=f'Number of parallel ffmpeg processes to run. Defaults to {default_jobs} (half of available CPU cores).')
    parser.add_argument('-i', '--interval', type=int, default=30,
                        help='Seconds between samples for crop detection. Smaller values are more accurate but slower. Default: 30.')
    args = parser.parse_args()

    check_prerequisites()

    cleaned_input = args.input_file.replace('`', '')
    if not os.path.exists(cleaned_input):
        print(f"Error: Input file not found at '{cleaned_input}'")
        sys.exit(1)

    # 2. Get video properties (duration and resolution) using ffprobe
    try:
        ffprobe_duration_args = ['ffprobe', '-v', 'error', '-show_entries', 'format=duration', '-of', 'default=noprint_wrappers=1:nokey=1', cleaned_input]
        duration_str = subprocess.check_output(ffprobe_duration_args, text=True, encoding='utf-8').strip()
        duration = float(duration_str)

        # Use JSON output to reliably find the main video stream, ignoring attached pictures
        ffprobe_json_args = [
            'ffprobe', '-v', 'error',
            '-select_streams', 'v',
            '-show_entries', 'stream=width,height,disposition',
            '-of', 'json', cleaned_input
        ]
        json_output_str = subprocess.check_output(ffprobe_json_args, text=True, encoding='utf-8')
        video_info = json.loads(json_output_str)

        main_video_stream = None
        if 'streams' in video_info and video_info['streams']:
            # Find the first video stream that is not an attached picture
            for stream in video_info['streams']:
                if not ('disposition' in stream and stream['disposition'].get('attached_pic') == 1):
                    main_video_stream = stream
                    break # Found it, stop looking

        if not main_video_stream or 'width' not in main_video_stream or 'height' not in main_video_stream:
            print("\nError: Could not find a valid video stream or its resolution.")
            print("This can happen with complex files. Please inspect with 'ffprobe -i \"your_video_file.mkv\"'")
            sys.exit(1)

        width = main_video_stream['width']
        height = main_video_stream['height']

    except (subprocess.CalledProcessError, ValueError, json.JSONDecodeError) as e:
        print(f"Error: Could not get video properties. {e}")
        sys.exit(1)

    # Identify the original aspect ratio of the container
    original_ar_label = None
    if height > 0:
        original_ar_val = width / height
        for ar in KNOWN_ASPECT_RATIOS:
            if abs(original_ar_val - ar['ratio']) < 0.02:
                original_ar_label = ar['name']
                break

    print(f"\nVideo properties: {width}x{height}, {duration:.2f}s. Analyzing with up to {args.jobs} parallel jobs...")

    # 3. Generate seek points
    interval_seconds = args.interval
    seek_points = list(range(interval_seconds, int(duration), interval_seconds))
    if not seek_points and duration > 0:
        seek_points = [duration / 2]
        print(f"Info: Video is shorter than {interval_seconds} seconds, taking one sample from the middle.")

    # 4. Run analysis in parallel
    tasks = [(sp, cleaned_input, width, height) for sp in seek_points]
    all_results = []

    print("\n--- Starting Analysis ---")
    with multiprocessing.Pool(processes=args.jobs) as pool:
        results_list = []
        total_tasks = len(tasks)
        for i, result in enumerate(pool.imap_unordered(analyze_segment, tasks), 1):
            results_list.append(result)
            # Print a simple, updating progress line
            progress_message = f"Analyzing Segments: {i}/{total_tasks} completed..."
            sys.stdout.write(f"\r{progress_message}")
            sys.stdout.flush()

    # Move to the next line after the progress indicator is done
    print()

    # Flatten the list of lists into a single list of crop values with their timestamps
    all_crop_values = [item for sublist in results_list for item in sublist]

    # 5. Find and display the most common crop value
    print("\n--- Final Verdict ---")
    if not all_crop_values:
        print("Could not find any significant crop values. The video might not have black bars.")
    else:
        # Group timestamps by crop value to detect crops only in credits
        crop_to_timestamps = {}
        for crop_str, ts in all_crop_values:
            if crop_str not in crop_to_timestamps:
                crop_to_timestamps[crop_str] = []
            crop_to_timestamps[crop_str].append(ts)

        start_credit_threshold = duration * 0.05  # First 5%
        end_credit_threshold = duration * 0.95    # Last 5%
        
        ignorable_crops = set()
        for crop_str, timestamps in crop_to_timestamps.items():
            # If all occurrences are in the first 5% OR all occurrences are in the last 5%
            if max(timestamps) < start_credit_threshold or min(timestamps) > end_credit_threshold:
                ignorable_crops.add(crop_str)

        # Create a new list containing only the crop strings of non-ignorable crops
        filtered_crop_strings = [
            crop_str for crop_str, ts in all_crop_values 
            if crop_str not in ignorable_crops
        ]

        if ignorable_crops:
            print("\n--- Credits/Logo Detection ---")
            print(f"Ignoring {len(ignorable_crops)} crop value(s) that appear only in the first/last 5% of the video.")
        
        if not filtered_crop_strings:
            print("\nCould not find any significant crop values outside of the start/end credits.")
            print("The video might not have black bars in the main content.")
        else:
            # Create a list of non-ignorable crops with their timestamps
            filtered_crops = [
                (crop_str, ts) for crop_str, ts in all_crop_values 
                if crop_str not in ignorable_crops
            ]

            # Group all detections by their snapped aspect ratio
            ar_groups = {}
            for crop_str, ts in filtered_crops:
                try:
                    _, values = crop_str.split('=')
                    w, h, x, y = map(int, values.split(':'))
                    
                    snapped_crop, ar_label = snap_to_known_ar(w, h, x, y, width, height)
                    
                    key = snapped_crop
                    if key not in ar_groups:
                        ar_groups[key] = {'key': key, 'count': 0, 'ar_label': ar_label, 'timestamps': []}
                    
                    ar_groups[key]['count'] += 1
                    ar_groups[key]['timestamps'].append(ts)

                except (ValueError, IndexError):
                    continue # Skip malformed crop strings

            # --- Luma Verification Pass ---
            # For any group that is an "Unidentified AR", check if it's just a dark scene.
            LUMA_THRESHOLD = 15.0
            validated_groups = []
            dark_scene_detections = 0
            
            # Separate known ARs from unidentified ones that need checking
            groups_to_check = []
            for group in ar_groups.values():
                if group['ar_label'] is None:
                    groups_to_check.append(group)
                else:
                    validated_groups.append(group) # Known ARs are always valid

            if groups_to_check:
                print("\n--- Luma Verification ---")
                
                luma_tasks = [
                    (g['key'], g['timestamps'][0], cleaned_input, LUMA_THRESHOLD) 
                    for g in groups_to_check
                ]
                
                # Create a map for quick lookup after parallel processing
                group_map = {g['key']: g for g in groups_to_check}

                with multiprocessing.Pool(processes=args.jobs) as pool:
                    total_luma_tasks = len(luma_tasks)
                    for i, (group_key, is_bright) in enumerate(pool.imap_unordered(check_luma_for_group, luma_tasks), 1):
                        group = group_map[group_key]
                        if is_bright:
                            validated_groups.append(group)
                        else:
                            dark_scene_detections += group['count']
                        
                        progress_message = f"Verifying scenes: {i}/{total_luma_tasks} completed..."
                        sys.stdout.write(f"\r{progress_message}")
                        sys.stdout.flush()
                
                print() # Move to the next line after the progress indicator is done

                if dark_scene_detections > 0:
                    print(f"Ignoring {dark_scene_detections} detections that occurred in very dark scenes.")

            # Convert the validated groups to a sorted list
            sorted_groups = sorted(validated_groups, key=lambda g: g['count'], reverse=True)
            
            if not sorted_groups:
                print("\nNo significant crop values found after filtering for credits and dark scenes.")
            else:
                total_detections = sum(g['count'] for g in sorted_groups)
                significant_crop_threshold = 5  # in percent

                dominant_group = sorted_groups[0]
                no_crop_string = f"crop={width}:{height}:0:0"
                is_dominant_no_crop = (
                    dominant_group['key'] == no_crop_string and
                    (dominant_group['count'] / total_detections) * 100 > 95.0
                )

                # Special case for videos that are overwhelmingly "no crop"
                if is_dominant_no_crop:
                    print(f"\n{COLOR_GREEN}Analysis complete.{COLOR_RESET}")
                    print(f"The video is overwhelmingly '{dominant_group['ar_label']}' and does not require cropping.")
                    print("Minor aspect ratio variations were detected but are considered insignificant due to their low frequency.")
                    print(f"{COLOR_GREEN}Recommendation: No crop needed.{COLOR_RESET}")
                else:
                    significant_groups = [
                        g for g in sorted_groups
                        if (g['count'] / total_detections) * 100 >= significant_crop_threshold
                    ]

                    if len(significant_groups) == 1:
                        # This is the dominant_group from before
                        # Re-introduce the safety check for dramatic but infrequent AR changes
                        dramatically_different_found = False
                        try:
                            # Get the ratio of the dominant group
                            _, dom_values = dominant_group['key'].split('=')
                            dom_w, dom_h, _, _ = map(int, dom_values.split(':'))
                            dominant_ratio = dom_w / dom_h

                            # Check against all other detected groups, even non-significant ones
                            for other_group in sorted_groups:
                                if other_group['key'] == dominant_group['key']:
                                    continue
                                _, other_values = other_group['key'].split('=')
                                other_w, other_h, _, _ = map(int, other_values.split(':'))
                                other_ratio = other_w / other_h
                                if abs(dominant_ratio - other_ratio) / dominant_ratio > 0.10:
                                    dramatically_different_found = True
                                    break
                        except (ValueError, IndexError, ZeroDivisionError):
                            dramatically_different_found = False

                        if dramatically_different_found:
                            print(f"\n{COLOR_RED}--- WARNING: Potentially Mixed Aspect Ratios Detected! ---{COLOR_RESET}")
                            percentage = (dominant_group['count'] / total_detections) * 100
                            label = f"'{dominant_group['ar_label']}'" if dominant_group['ar_label'] else "Unidentified AR"
                            print(f"The dominant aspect ratio is {label} ({dominant_group['key']}), found in {percentage:.1f}% of samples.")
                            print("However, other significantly different aspect ratios were also detected, although less frequently.")
                            print("\nRecommendation: Manually check the video before applying a single crop.")
                            print("You can review the next most common detections below:")
                            for group in sorted_groups[1:4]:
                                if group['count'] > 0:
                                    percentage = (group['count'] / total_detections) * 100
                                    label = f"'{group['ar_label']}'" if group['ar_label'] else "Unidentified AR"
                                    print(f"  - {label} ({group['key']}) was detected {group['count']} time(s) ({percentage:.1f}%).")
                        else:
                            print(f"\n{COLOR_GREEN}Analysis complete.{COLOR_RESET}")
                            dominant_group_ar_label = dominant_group['ar_label']
                            if dominant_group_ar_label:
                                print(f"The video consistently uses the '{dominant_group_ar_label}' aspect ratio.")
                            
                            if original_ar_label and dominant_group_ar_label and original_ar_label != dominant_group_ar_label:
                                if "4:3" in original_ar_label and ("Widescreen" in dominant_group_ar_label or "IMAX" in dominant_group_ar_label):
                                    print(f"\n{COLOR_YELLOW}--- Sanity Check Warning ---{COLOR_RESET}")
                                    print(f"This video appears to be in a '{original_ar_label}' container, but the suggested crop")
                                    print(f"changes it to a '{dominant_group_ar_label}' aspect ratio.")
                                    print("This can be correct for letterboxed content, but please verify it is not cropping the actual picture.")

                            print(f"{COLOR_GREEN}Recommended crop filter: -vf {dominant_group['key']}{COLOR_RESET}")

                    elif len(significant_groups) > 1:
                        print(f"\n{COLOR_RED}--- WARNING: Mixed Aspect Ratios Detected! ---{COLOR_RESET}")
                        print(f"The analysis found multiple dominant aspect ratios (each present in >{significant_crop_threshold}% of samples):")
                        for group in significant_groups:
                            percentage = (group['count'] / total_detections) * 100
                            label = f"'{group['ar_label']}'" if group['ar_label'] else "Unidentified AR"
                            print(f"  - {label} ({group['key']}) was detected {group['count']} time(s) ({percentage:.1f}%).")
                        print("\nThis likely indicates intentional artistic choices (e.g., IMAX scenes, flashbacks).")
                        print("Recommendation: DO NOT CROP this video.")
                    
                    else: # No single group is significant enough
                        print(f"\n{COLOR_RED}--- WARNING: Highly Variable Aspect Ratio Detected! ---{COLOR_RESET}")
                        print(f"No single aspect ratio was present for more than {significant_crop_threshold}% of the analyzed frames.")
                        print("The most common detected aspect ratios were:")
                        for group in sorted_groups[:3]:
                            percentage = (group['count'] / total_detections) * 100
                            label = f"'{group['ar_label']}'" if group['ar_label'] else "Unidentified AR"
                            print(f"  - {label} ({group['key']}) was detected {group['count']} time(s) ({percentage:.1f}%).")
                        print("\nApplying a single crop filter would likely damage parts of the video.")
                        print("Recommendation: DO NOT CROP this video. Let the player handle letterboxing/pillarboxing.")

