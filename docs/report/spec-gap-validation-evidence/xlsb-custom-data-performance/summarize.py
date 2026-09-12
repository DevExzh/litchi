#!/usr/bin/env python3
"""Render a candidate-only report from retained raw JSONL receipts."""

from __future__ import annotations

import argparse
import json
import statistics
from pathlib import Path


LANES = (
    "read",
    "snapshotclone",
    "noop",
    "editpayload",
    "rename-many-to-one",
    "mixedinsertremove",
    "patchinverse",
    "publicapply",
)
LEVELS = ("small", "medium", "large")


def rows(result: Path) -> list[dict]:
    values = []
    for path in sorted((result / "raw").glob("*.jsonl")):
        for line in path.read_text().splitlines():
            if line.strip():
                values.append(json.loads(line))
    return values


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--results", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    values = rows(args.results)
    if len(values) != 120:
        raise SystemExit(f"expected 120 raw samples, found {len(values)}")
    first = values[0]
    lines = [
        "# XLSB Custom Data lifecycle performance",
        "",
        "This is candidate-only characterization of the public XLSB Custom Data",
        "lifecycle at source commit `16102fe751d7c5492042330f1bd1f49c304495f0`.",
        "The corpus is deterministic authored OPC+BIFF12 source based on the",
        "committed lifecycle test. It is valid synthetic evidence and carries no",
        "native-producer, baseline, speedup, Office-acceptance, or general",
        "throughput claim.",
        "",
        "The retained harness is an optimized Cargo `--release` binary with",
        "`CARGO_INCREMENTAL=0` and no compiler override flags. Elapsed values",
        "include the process-local allocator observer's atomic accounting overhead.",
        "",
        "The 24 lane/class groups each contain five measured samples after one",
        "warm-up in a fresh process. Medians, minima, and maxima below are",
        "descriptive observations for those five samples; they are not",
        "tail-latency or uncertainty certification.",
        "Transaction lanes build caller replacement data inside the timed",
        "transaction scope. Patch/inverse and public-apply lanes reuse equal",
        "prepared inputs outside timing; each row records this boundary.",
        "",
        "## Matrix observations",
        "",
        "| lane | small median ns | medium median ns | large median ns |",
        "| --- | ---: | ---: | ---: |",
    ]
    for lane in LANES:
        cells = [lane]
        for level in LEVELS:
            sample_values = [
                row["timing_ns"]
                for row in values
                if row["lane"] == lane and row["size_class"] == level
            ]
            cells.append(str(int(statistics.median(sample_values))))
        lines.append("| " + " | ".join(cells) + " |")

    lines.extend(
        [
            "",
            "The size classes are small (2 storages, 16 references, 1 KiB per",
            "storage), medium (4, 128, 8 KiB), and large (16, 512, 32 KiB).",
            "Each fixture leaves one storage unreferenced for the mixed",
            "insert/remove lane; the first storage receives the many-to-one fan-in.",
            "",
            "## Accounting and correctness boundary",
            "",
            "Receipts retain direct allocation calls and bytes, realloc-old and",
            "realloc-new bytes, requested bytes, operation and through-drop release",
            "bytes, live bytes, and logical peak live bytes. The balance equation is",
            "checked for every sample. Logical peak is an aggregate allocator",
            "observation and excludes transient overlap that an allocator may hide",
            "during reallocation; it is not RSS or a leak proof. `/usr/bin/time -v`",
            "retains process maximum RSS per lane/class process.",
            "",
            "`copies_bytes_observed` counts only explicit source/candidate/opaque",
            "byte copies made by post-timer assertions. Internal parser, archive,",
            "and allocator copies are not instrumented, and validation copies are",
            "excluded from elapsed timing.",
            "",
            "No-op package bytes are compared exactly. Changed lanes preserve the",
            "unrelated opaque member and apply the public inverse back to the exact",
            "source bytes. Semantic storage IDs, cardinality, payload changes, and",
            "connection-reference counts are checked from the public API.",
            "Exact expected payload bytes and lengths are checked for every",
            "untouched, edited, inserted, renamed, and restored storage. Each",
            "sample also emits a candidate package SHA-256.",
            "For read/snapshotclone, this hash identifies the unchanged source; no",
            "new candidate package is produced by those lanes.",
            "The complete per-ID reference-count map, including zero-reference",
            "and inserted/removed IDs, is checked. Patch/inverse receipts retain",
            "a hash of the prepared forward candidate validated outside timing.",
            "",
            "## Replay inputs",
            "",
            f"Source SHA-256 for the first retained fixture row: `{first['fixture']['source_sha256']}`.",
            "The receipt directory retains the clean source manifest, harness",
            "Cargo.lock, binary hash, toolchain, commands, fixture hash",
            "table, raw JSONL samples, `/usr/bin/time` sidecars, and stderr logs.",
            "`replay.sh` reruns the same 24 groups from the retained clean source",
            "checkout when available; if that disposable checkout is absent, it",
            "reconstructs the exact pinned commit before rebuilding, then uses a",
            "fresh result directory.",
        ]
    )
    args.output.write_text("\n".join(lines) + "\n")


if __name__ == "__main__":
    main()
