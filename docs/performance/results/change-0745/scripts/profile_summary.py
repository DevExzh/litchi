#!/usr/bin/env python3
"""Summarize a `perf record --call-graph dwarf` profile of the 0745 probe.

Only samples whose call chain contains the probe's `timed_owner` frame are
counted, so untimed setup and the per-sample oracle digests are excluded.
For those samples the script reports the leaf-symbol distribution and the
share attributed to a few named inclusive frames (first match wins from the
leaf upwards).

Usage: profile_summary.py PERF_DATA [LABEL]
"""

import collections
import json
import subprocess
import sys

INCLUSIVE = [
    # SHA-256 block compression reached from any caller.
    ("sha256_any", ("digest<sha2::Sha256",)),
    # Whole-artifact digest helpers inside the PPT slide-order owner.
    ("artifact_hash", ("artifact_hash",)),
    ("to_durable", ("to_durable",)),
    ("to_deterministic_json", ("to_deterministic_json",)),
    ("apply_durable", ("apply_durable",)),
]


def samples(path):
    result = subprocess.run(
        ["perf", "script", "-i", path, "-F", "comm,ip,sym"],
        check=True,
        capture_output=True,
        text=True,
    )
    frames = None
    for line in result.stdout.splitlines():
        if not line.strip():
            if frames is not None:
                yield frames
            frames = None
            continue
        if not line.startswith(("\t", " ")):
            # A sample header (the command name) starts a new call chain.
            if frames is not None:
                yield frames
            frames = []
            continue
        parts = line.strip().split(" ", 1)
        frames.append(parts[1] if len(parts) == 2 else parts[0])
    if frames is not None:
        yield frames


def main():
    path = sys.argv[1]
    label = sys.argv[2] if len(sys.argv) > 2 else path
    total = 0
    timed = 0
    leaves = collections.Counter()
    inclusive = collections.Counter()
    for frames in samples(path):
        total += 1
        if not any("timed_owner" in frame for frame in frames):
            continue
        timed += 1
        leaves[frames[0].split("+")[0]] += 1
        for name, needles in INCLUSIVE:
            if any(any(needle in frame for needle in needles) for frame in frames):
                inclusive[name] += 1
    frame_names = collections.Counter()
    for frames in samples(path):
        if any("timed_owner" in frame for frame in frames):
            names = {frame.split(" (inlined)")[0].split("<")[0] for frame in frames}
            frame_names.update(names)
    report = {
        "label": label,
        "samples_total": total,
        "samples_in_timed_owner": timed,
        "inclusive_share_of_timed": {
            name: round(inclusive[name] / timed, 4) if timed else None
            for name, _ in INCLUSIVE
        },
        "top_inclusive_frames_share_of_timed": [
            [name[:80], round(count / timed, 4)]
            for name, count in frame_names.most_common(45)
        ],
        "top_leaf_share_of_timed": [
            [symbol[:120], round(count / timed, 4)]
            for symbol, count in leaves.most_common(12)
        ],
    }
    print(json.dumps(report, indent=1))


if __name__ == "__main__":
    main()
