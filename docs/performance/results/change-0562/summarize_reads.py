#!/usr/bin/env python3
"""Summarize repeated positional reads from retained strace captures.

Each capture is one child of the perf harness with `--warmup 0 --samples 1`,
traced with `strace -f -e trace=pread64`. The summary counts how many calls
re-read a byte range the same child already read, which is a property of the
retained trace and not a timing measurement.
"""

from __future__ import annotations

import argparse
import collections
import gzip
import json
import pathlib
import re

PREAD = re.compile(r"pread64\(\s*\d+,\s*.*?,\s*(\d+),\s*(\d+)\)\s*=\s*(\d+)")


def read_text(path: pathlib.Path) -> str:
    if path.suffix == ".gz":
        with gzip.open(path, "rt", errors="replace") as handle:
            return handle.read()
    return path.read_text(errors="replace")


def summarize(path: pathlib.Path) -> dict:
    seen: collections.Counter = collections.Counter()
    order: list[tuple[int, int]] = []
    sizes: collections.Counter = collections.Counter()
    total_bytes = 0
    for line in read_text(path).splitlines():
        match = PREAD.search(line)
        if not match:
            continue
        requested, offset, returned = (int(match.group(i)) for i in (1, 2, 3))
        seen[(offset, requested)] += 1
        order.append((offset, requested))
        total_bytes += returned
        bucket = 1
        while bucket < max(returned, 1):
            bucket *= 2
        sizes[bucket] += 1
    calls = len(order)
    if calls == 0:
        return {"capture": path.name, "pread64_calls": 0}
    immediate = sum(1 for index in range(1, calls) if order[index] == order[index - 1])
    return {
        "capture": path.name,
        "pread64_calls": calls,
        "distinct_ranges": len(seen),
        "repeat_calls": calls - len(seen),
        "repeat_call_percent": (calls - len(seen)) / calls * 100.0,
        "immediate_duplicate_calls": immediate,
        "immediate_duplicate_percent": immediate / calls * 100.0,
        "bytes_returned": total_bytes,
        "mean_read_bytes": total_bytes / calls,
        "repeat_count_distribution": dict(sorted(collections.Counter(seen.values()).items())),
        "size_histogram_upper_bound_to_calls": dict(sorted(sizes.items())),
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--traces", default="docs/performance/results/change-0561/traces")
    parser.add_argument("--output", default="docs/performance/results/change-0561/read-repetition.json")
    args = parser.parse_args()
    traces = sorted(
        list(pathlib.Path(args.traces).glob("*.pread.txt"))
        + list(pathlib.Path(args.traces).glob("*.pread.txt.gz")),
        key=lambda path: path.name,
    )
    rows = [summarize(path) for path in traces]
    pathlib.Path(args.output).write_text(
        json.dumps({"schema_version": 1, "summary_kind": "litchi-opc-read-repetition", "rows": rows}, indent=2) + "\n"
    )
    for row in rows:
        if not row["pread64_calls"]:
            print(f"{row['capture']}: no positional reads")
            continue
        print(
            f"{row['capture']:44s} calls={row['pread64_calls']:6d} distinct={row['distinct_ranges']:6d} "
            f"repeats={row['repeat_call_percent']:5.1f}% immediate={row['immediate_duplicate_percent']:5.1f}% "
            f"mean={row['mean_read_bytes']:6.1f}B"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
