#!/usr/bin/env python3
"""Summarize the positional reads one source-backed XLS open performs.

The capture is one perf-harness child with one warmup and one measured sample,
so it contains two opens; per-open figures are halved. The isolation pair at 1
and 11 samples differences to the same per-open counts independently.
"""

from __future__ import annotations

import argparse
import collections
import json
import pathlib
import re

PREAD = re.compile(r"pread64\(\s*\d+,\s*.*?,\s*(\d+),\s*(\d+)\)\s*=\s*(\d+)")


def summarize(path: pathlib.Path, opens: int) -> dict:
    rows = []
    for line in path.read_text(errors="replace").splitlines():
        match = PREAD.search(line)
        if match:
            requested, offset = int(match.group(1)), int(match.group(2))
            rows.append((offset, requested))
    sizes = collections.Counter(requested for _offset, requested in rows)
    sequential = sum(
        1
        for index in range(1, len(rows))
        if rows[index][0] == rows[index - 1][0] + rows[index - 1][1]
    )
    return {
        "capture": path.name,
        "opens_in_capture": opens,
        "pread64_calls": len(rows),
        "pread64_calls_per_open": len(rows) / opens,
        "bytes_requested": sum(requested for _offset, requested in rows),
        "size_histogram": {str(size): count for size, count in sorted(sizes.items())},
        "four_byte_reads": sizes.get(4, 0),
        "four_byte_share_percent": sizes.get(4, 0) / len(rows) * 100.0 if rows else 0.0,
        "reads_continuing_the_previous_read": sequential,
    }


def counted(path: pathlib.Path, name: str) -> int:
    for line in path.read_text().splitlines():
        fields = line.split()
        if len(fields) >= 4 and fields[-1] == name:
            return int(fields[3])
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--traces", default="docs/performance/results/change-0564/traces")
    parser.add_argument("--output", default="docs/performance/results/change-0564/open-reads.json")
    args = parser.parse_args()
    root = pathlib.Path(args.traces)

    detail = summarize(root / "xls_file_source_open.pread.txt", opens=2)
    low = root / "xls_file_source_open.samples-1.strace.txt"
    high = root / "xls_file_source_open.samples-11.strace.txt"
    isolation = {
        "statx_per_open": (counted(high, "statx") - counted(low, "statx")) / 10,
        "pread64_per_open": (counted(high, "pread64") - counted(low, "pread64")) / 10,
        "method": "one warmup plus 1 and 11 samples; the difference divided by ten isolates one open",
    }
    result = {"schema_version": 1, "summary_kind": "litchi-xls-open-read-shape",
              "detail": detail, "isolation": isolation}
    pathlib.Path(args.output).write_text(json.dumps(result, indent=2) + "\n")
    print(f"per open: {isolation['pread64_per_open']:.0f} pread64, {isolation['statx_per_open']:.0f} statx")
    print(f"four-byte reads: {detail['four_byte_reads']} of {detail['pread64_calls']} "
          f"({detail['four_byte_share_percent']:.1f}%)")
    print(f"reads continuing the previous read: {detail['reads_continuing_the_previous_read']} "
          f"of {detail['pread64_calls'] - 1}")
    print("size histogram:", detail["size_histogram"])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
