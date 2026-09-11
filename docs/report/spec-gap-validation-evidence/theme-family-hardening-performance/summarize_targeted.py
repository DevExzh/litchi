#!/usr/bin/env python3
"""Aggregate the targeted current-source theme-family hardening profile."""

from __future__ import annotations

import argparse
import json
import math
from pathlib import Path


LANES = (
    "native_read",
    "native_replace",
    "native_remove",
    "native_add",
    "unknown_32",
    "unknown_1000",
    "duplicate",
    "limit_replace",
    "limit_add",
)


def quantile(values: list[int], percent: int) -> int:
    return sorted(values)[max(0, math.ceil(len(values) * percent / 100) - 1)]


def rss(path: Path) -> int:
    for line in path.read_text().splitlines():
        if line.lstrip().startswith("Maximum resident set size (kbytes):"):
            return int(line.split(":", 1)[1].strip())
    raise ValueError(f"RSS missing from {path}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--results", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    rows: list[dict[str, object]] = []
    for lane in LANES:
        values = []
        rss_values = []
        for path in sorted(args.results.glob(f"{lane}-p*.json")):
            value = json.loads(path.read_text())
            values.extend(value["samples"])
            rss_values.append(rss(path.with_suffix(".time.txt")))
        elapsed = [int(sample["elapsed_ns"]) for sample in values]
        allocated = [int(sample["requested_alloc_bytes"]) for sample in values]
        peak = [int(sample["peak_live_delta"]) for sample in values]
        rows.append(
            {
                "lane": lane,
                "processes": len(rss_values),
                "samples": len(values),
                "p50": quantile(elapsed, 50),
                "p95": quantile(elapsed, 95),
                "p99": quantile(elapsed, 99),
                "alloc50": quantile(allocated, 50),
                "alloc95": quantile(allocated, 95),
                "peak50": quantile(peak, 50),
                "peak95": quantile(peak, 95),
                "rss_min": min(rss_values),
                "rss_max": max(rss_values),
            }
        )
    lines = [
        "# Theme-family hardening profile",
        "",
        "This report contains absolute current-source measurements from three fresh processes and twenty measured samples per lane. The timer covers the named shared DrawingML operation with setup and fixture construction outside the timed region. Allocation bytes and incremental peak live bytes come from a process-local counting allocator; RSS is whole-process `/usr/bin/time -v` RSS.",
        "",
        "| lane | fresh processes | samples | p50 / p95 / p99 ns | requested alloc p50 / p95 B | peak-live delta p50 / p95 B | RSS KiB |",
        "|---|---:|---:|---:|---:|---:|---:|",
    ]
    for row in rows:
        lines.append(
            "| {lane} | {processes} | {samples} | {p50} / {p95} / {p99} | {alloc50} / {alloc95} | {peak50} / {peak95} | {rss_min}–{rss_max} |".format(**row)
        )
    lines.extend(
        [
            "",
            "`native_read`, `native_replace`, `native_remove`, and `native_add` are valid native Theme workflows. `unknown_32` and `unknown_1000` contain an admitted unknown-URI extension with respectively 32 and 1,000 direct family-shaped opaque descendants plus 200 active root namespace declarations; both must read successfully with no typed owner. `duplicate` contains two supported family owners and must reject. `limit_replace` and `limit_add` use a caller output cap of one byte and must reject before producing output.",
            "",
            "These are scoped absolute observations. The run does not provide a before/after comparison, an asymptotic proof, or a whole-library performance claim. The unknown-owner lanes are bounded synthetic stress points, and the 200-declaration root remains below the implementation's active namespace ceiling.",
            "",
            "Raw per-process JSON, `/usr/bin/time -v` receipts, source/build manifests, exact commands, and source hashes are retained in this directory.",
        ]
    )
    args.output.write_text("\n".join(lines) + "\n")


if __name__ == "__main__":
    main()
