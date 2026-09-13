#!/usr/bin/env python3
"""Attribute repeated positional reads to one archive instance or to re-opens.

Each traced child opens the package several times. Segmenting every `(pid, fd)`
read stream at its own 22-byte end-of-central-directory read — the unambiguous
marker for one archive construction — separates repeats that happen *inside* one
archive instance from repeats that happen because the package was opened again.

The distinction decides which optimizations can help: a per-entry memo can only
remove repeats inside one instance.
"""

from __future__ import annotations

import argparse
import collections
import gzip
import json
import pathlib
import re

PREAD = re.compile(r"(\d+)\s+pread64\(\s*(\d+),\s*.*?,\s*(\d+),\s*(\d+)\)\s*=\s*(\d+)")
EOCD_BYTES = 22
LOCAL_HEADER_BYTES = 30
DESCRIPTOR_BYTES = 16


def reads(path: pathlib.Path) -> list[tuple[int, int, int, int]]:
    opener = gzip.open if path.suffix == ".gz" else open
    mode = "rt" if path.suffix == ".gz" else "r"
    out = []
    with opener(path, mode, errors="replace") as handle:
        for line in handle:
            match = PREAD.search(line)
            if match:
                pid, descriptor, requested, offset = (int(match.group(i)) for i in (1, 2, 3, 4))
                out.append((pid, descriptor, offset, requested))
    return out


def summarize(path: pathlib.Path) -> dict:
    rows = reads(path)
    segments: dict[int, list[tuple[int, int]]] = collections.defaultdict(list)
    current: dict[tuple[int, int], int] = {}
    constructions = 0
    for pid, descriptor, offset, requested in rows:
        key = (pid, descriptor)
        if requested == EOCD_BYTES:
            constructions += 1
            current[key] = constructions
        segments[current.get(key, 0)].append((offset, requested))

    counts = {LOCAL_HEADER_BYTES: 0, DESCRIPTOR_BYTES: 0}
    repeats = {LOCAL_HEADER_BYTES: 0, DESCRIPTOR_BYTES: 0}
    for reads_in_segment in segments.values():
        for (_offset, requested), seen in collections.Counter(reads_in_segment).items():
            if requested in counts:
                counts[requested] += seen
                repeats[requested] += seen - 1
    return {
        "capture": path.name,
        "pread64_calls": len(rows),
        "archive_constructions": constructions,
        "local_header_reads": counts[LOCAL_HEADER_BYTES],
        "local_header_repeats_within_one_archive": repeats[LOCAL_HEADER_BYTES],
        "descriptor_reads": counts[DESCRIPTOR_BYTES],
        "descriptor_repeats_within_one_archive": repeats[DESCRIPTOR_BYTES],
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--traces", default="docs/performance/results/change-0562/traces")
    parser.add_argument("--output", default="docs/performance/results/change-0561/read-segmentation.json")
    args = parser.parse_args()
    root = pathlib.Path(args.traces)
    paths = sorted(
        [path for path in root.iterdir() if ".pread.txt" in path.name],
        key=lambda path: path.name,
    )
    rows = [summarize(path) for path in paths]
    pathlib.Path(args.output).write_text(
        json.dumps({"schema_version": 1, "summary_kind": "litchi-opc-read-segmentation", "rows": rows}, indent=2)
        + "\n"
    )
    for row in rows:
        print(
            f"{row['capture']:44s} calls={row['pread64_calls']:6d} archives={row['archive_constructions']:3d} "
            f"headers={row['local_header_reads']:5d} (repeat-in-archive {row['local_header_repeats_within_one_archive']:3d}) "
            f"descriptors={row['descriptor_reads']:5d} (repeat-in-archive {row['descriptor_repeats_within_one_archive']:3d})"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
