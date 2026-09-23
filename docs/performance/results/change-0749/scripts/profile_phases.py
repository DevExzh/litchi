#!/usr/bin/env python3
"""Phase x leaf attribution of a frame-pointer `perf record` of the 0749 probe.

Only samples whose call chain contains the probe's `timed_owner` frame are
counted, so untimed setup and oracle digests are excluded. Each timed sample is
assigned to the innermost named phase on its stack (searching from the leaf
towards the root), and its leaf symbol is tallied inside that phase.

Phases (innermost match wins):
- validate/readback: `ReusePlan::validate` -> `open_stream` (step B, before),
  or the 0749 streaming comparison (`stream_equals` / `compare_stream`);
- validate/reparse: `ReusePlan::validate` -> `open_with_limits` (step A);
- validate/other: remaining samples under `validate`;
- plan: `plan_sector_layout` / `plan_reuse`;
- emit: `ReusePlan::emit`;
- write_to/other: remaining samples under `write_to`;
- outside write_to: timed-owner samples outside `OleWriter::write_to`.

Usage: profile_phases.py PERF_DATA [LABEL]
"""

import collections
import json
import subprocess
import sys


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
            if frames is not None:
                yield frames
            frames = []
            continue
        parts = line.strip().split(" ", 1)
        frames.append(parts[1] if len(parts) == 2 else parts[0])
    if frames is not None:
        yield frames


def name_of(frame):
    return frame.split(" (inlined)")[0].split("+0x")[0]


def phase(frames):
    names = [name_of(frame) for frame in frames]

    def has(*needles):
        return any(any(needle in name for needle in needles) for name in names)

    if not has("write_to"):
        return "outside write_to"
    if has("validate"):
        if has("open_stream", "stream_equals", "compare_stream", "compare_run"):
            return "validate/readback"
        if has("open_with_limits"):
            return "validate/reparse"
        return "validate/other"
    if has("plan_sector_layout", "plan_reuse"):
        return "plan"
    if has("emit"):
        return "emit"
    return "write_to/other"


def main():
    path = sys.argv[1]
    label = sys.argv[2] if len(sys.argv) > 2 else path
    total = 0
    timed = 0
    phases = collections.Counter()
    leaves = collections.defaultdict(collections.Counter)
    for frames in samples(path):
        total += 1
        if not any("timed_owner" in frame for frame in frames):
            continue
        timed += 1
        bucket = phase(frames)
        phases[bucket] += 1
        leaves[bucket][name_of(frames[0])[:100]] += 1
    report = {
        "label": label,
        "samples_total": total,
        "samples_in_timed_owner": timed,
        "phase_share_of_timed": {
            name: round(count / timed, 4) for name, count in phases.most_common()
        },
        "leaf_share_of_timed_by_phase": {
            name: [
                [symbol, round(count / timed, 4)]
                for symbol, count in leaves[name].most_common(8)
            ]
            for name, _ in phases.most_common()
        },
    }
    print(json.dumps(report, indent=1))


if __name__ == "__main__":
    main()
