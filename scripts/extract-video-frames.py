#!/usr/bin/env python3
"""Extract one frame per second from a selected tutorial-video interval."""

from __future__ import annotations

import argparse
import math
from pathlib import Path

import cv2


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("video", type=Path)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--start", type=float, default=0.0, help="Start time in seconds")
    parser.add_argument("--end", type=float, help="End time in seconds")
    args = parser.parse_args()

    capture = cv2.VideoCapture(str(args.video))
    if not capture.isOpened():
        raise SystemExit(f"Cannot open video: {args.video}")

    fps = capture.get(cv2.CAP_PROP_FPS)
    frame_count = int(capture.get(cv2.CAP_PROP_FRAME_COUNT))
    duration = frame_count / fps if fps else 0.0
    start = max(0, int(math.floor(args.start)))
    requested_end = duration if args.end is None else min(duration, args.end)
    end = max(start, int(math.ceil(requested_end)))

    args.output_dir.mkdir(parents=True, exist_ok=True)
    written = 0
    for second in range(start, end + 1):
        capture.set(cv2.CAP_PROP_POS_MSEC, second * 1000)
        ok, frame = capture.read()
        if not ok:
            continue
        target = args.output_dir / f"frame-{second:04d}s.jpg"
        if not cv2.imwrite(str(target), frame):
            raise SystemExit(f"Cannot write frame: {target}")
        written += 1
    capture.release()

    print(f"video_duration={duration:.2f}s")
    print(f"source_fps={fps:.3f}")
    print(f"interval={start}s..{end}s")
    print(f"frames_written={written}")


if __name__ == "__main__":
    main()

