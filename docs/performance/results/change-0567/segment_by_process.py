#!/usr/bin/env python3
"""Attribute ZIP archive constructions to processes in the change-0562 traces.

Change 0561 segmented each traced child's positional reads at every 22-byte
end-of-central-directory read and counted ten constructions, then inferred that
one library-level open builds the ZIP index about three times.

Counting the same reads *per process* shows five processes at two each: four
per-sample children and the parent.  Two per process is the lifecycle change
0561 itself describes - one untimed preparing open and one post-timer oracle
open - so one library-level open builds the index once.

Reads only traces already retained in the repository.  Deterministic; stdlib only.
"""

from __future__ import annotations

import argparse
import collections
import gzip
import json
import pathlib
import re

PREAD = re.compile(r"^(\d+)\s+pread64\(\s*(\d+),.*?,\s*(\d+),\s*(\d+)\)\s*=\s*(\d+)")
EOCD_BYTES = 22


def summarize(path: pathlib.Path) -> dict:
    per_process: collections.Counter[str] = collections.Counter()
    total = 0
    opener = gzip.open if path.suffix == ".gz" else open
    with opener(path, "rt", errors="replace") as handle:
        lines = handle.read().splitlines()
    for line in lines:
        match = PREAD.match(line)
        if not match:
            continue
        pid, _fd, requested, _offset, returned = match.groups()
        if int(requested) == EOCD_BYTES and int(returned) == EOCD_BYTES:
            per_process[pid] += 1
            total += 1
    counts = sorted(per_process.values())
    return {
        "capture": path.name,
        "eocd_reads_total": total,
        "processes_issuing_them": len(per_process),
        "per_process": dict(sorted(per_process.items())),
        "distinct_per_process_counts": sorted(set(counts)),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--traces",
                        default="docs/performance/results/change-0562/traces")
    parser.add_argument("--output", default="-")
    args = parser.parse_args()
    root = pathlib.Path(args.traces)
    rows = [summarize(p) for p in sorted(root.iterdir())
            if p.is_file() and "pread" in p.name]
    payload = {
        "note": ("Counts 22-byte end-of-central-directory reads per process in the traces "
                 "change 0562 retained. One such read marks one IndexedArchive construction."),
        "captures": rows,
    }
    text = json.dumps(payload, indent=2, sort_keys=True) + "\n"
    if args.output == "-":
        print(text, end="")
    else:
        pathlib.Path(args.output).write_text(text)


if __name__ == "__main__":
    main()
