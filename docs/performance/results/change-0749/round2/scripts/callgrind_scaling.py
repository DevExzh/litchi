#!/usr/bin/env python3
"""Change 0749 round two: exact instruction scaling of the Reuse validation.

Runs the probe's `cfb-write-reuse --edit grow` under Callgrind, collecting
only inside `timed_owner`, for the generated v3 files with 1,000 and 3,000
mini streams of 2,000 bytes. Each run performs two writes; the listed
inclusive instruction counts are divided by two (per write). The functions
reported:

- `write_to`: the whole Reuse write;
- `validate`: `ReusePlan::validate`;
- `reparse`: `OleFile::open_with_limits` over the planned view, and within
  it `validate_stream_allocations` (A5), both unchanged by this change;
- `readback`: `OleFile::open_stream` (A), `OleFile::stream_equals` (B) or
  `StreamComparer::stream_equals` (C).

Usage: callgrind_scaling.py OUT_DIR > scaling.json
"""

import json
import os
import re
import subprocess
import sys

ROOT = "/home/zhuhe/code/litchi-worktrees/scratch/0749"
RUNS = [
    ("A", "v3-1000x2000"), ("A", "v3-3000x2000"),
    ("B", "v3-1000x2000"),
    ("C", "v3-1000x2000"), ("C", "v3-3000x2000"),
]
WRITES = 2
PATTERNS = {
    "write_to": r"OleWriter::write_to \[",
    "validate": r"ReusePlan::validate \[",
    "reparse": r"OleFile<R>::open_with_limits \[",
    "reparse_stream_allocations": r"OleFile<R>::validate_stream_allocations \[",
    "readback_open_stream": r"OleFile<R>::open_stream \[",
    "readback_stream_equals_b": r"OleFile<R>::stream_equals \[",
    "readback_stream_equals_c": r"StreamComparer<R,C>::stream_equals \[",
}


def inclusive(listing, pattern):
    for line in listing.splitlines():
        if re.search(pattern, line):
            number = line.strip().split(" ", 1)[0].replace(",", "")
            if number.isdigit():
                return int(number)
    return None


def main():
    out_dir = sys.argv[1]
    os.makedirs(out_dir, exist_ok=True)
    rows = []
    for arm, name in RUNS:
        stem = f"{out_dir}/{arm}-{name}"
        command = [
            "taskset", "-c", "20", "valgrind", "--tool=callgrind",
            "--toggle-collect=probe_0749::timed_owner",
            f"--callgrind-out-file={stem}.callgrind",
            f"{ROOT}/bin/{arm}/probe", "--mode", "cfb-write-reuse", "--edit", "grow",
            "--input", f"{ROOT}/gen/{name}.cfb",
            "--warmups", "0", "--samples", str(WRITES), "--oracle", "first",
        ]
        completed = subprocess.run(command, capture_output=True, text=True, timeout=590)
        if completed.returncode != 0:
            raise SystemExit(f"{command} failed: {completed.stderr[-2000:]}")
        listing = subprocess.run(
            ["callgrind_annotate", "--inclusive=yes", "--threshold=100", f"{stem}.callgrind"],
            capture_output=True, text=True, check=True,
        ).stdout
        with open(f"{stem}.inclusive.txt", "w") as handle:
            handle.write(listing)
        values = {key: inclusive(listing, pattern) for key, pattern in PATTERNS.items()}
        per_write = {key: (value // WRITES if value is not None else None) for key, value in values.items()}
        rows.append({"arm": arm, "input": name, "writes": WRITES, "inclusive_ir_per_write": per_write})
        os.remove(f"{stem}.callgrind")
    json.dump({"schema": "0749-callgrind-scaling-v1", "rows": rows}, sys.stdout, indent=1)
    print()


if __name__ == "__main__":
    main()
