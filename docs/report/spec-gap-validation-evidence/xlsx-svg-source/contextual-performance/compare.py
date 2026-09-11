#!/usr/bin/env python3
"""Make a side-by-side diagnostic index from two sealed receipt bundles."""

from __future__ import annotations

import argparse
import json
import re
import statistics
from pathlib import Path

from summarize import LANES


def side_values(results: Path, lane: str) -> dict[str, int]:
    receipts = [json.loads(path.read_text()) for path in sorted(results.glob(f"{lane}-p*.json"))]
    samples = [sample for receipt in receipts for sample in receipt["samples"]]
    rss = []
    for path in sorted(results.glob(f"{lane}-p*.json")):
        text = path.with_suffix(".time.txt").read_text()
        match = re.search(r"Maximum resident set size \(kbytes\):\s*(\d+)", text)
        if match is None:
            raise SystemExit(f"RSS missing for {path}")
        rss.append(int(match.group(1)))
    raw = int(receipts[0]["samples"][0]["retained_raw_source_bytes"])
    context = int(receipts[0]["samples"][0]["context_distinct_count"])
    return {
        "input": int(receipts[0]["input_bytes"]),
        "raw": raw,
        "context": context,
        "elapsed": int(statistics.median(int(sample["elapsed_ns"]) for sample in samples)),
        "alloc": int(statistics.median(int(sample["requested_alloc_bytes"]) for sample in samples)),
        "peak": int(statistics.median(int(sample["peak_live_delta"]) for sample in samples)),
        "rss": int(statistics.median(rss)),
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--before", type=Path, required=True)
    parser.add_argument("--after", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    lines = [
        "# Frozen source-projection comparison index",
        "",
        "This is a mechanical index of the exploratory receipt bundles. It places medians side by side for review; it does not calculate a speedup, establish a threshold, or make a lifecycle/native acceptance claim.",
        "",
        "The before source is `8b5838c59`; the after source is `fc44c4e6c`. Refer to each side's `build-provenance.txt`, `binary.sha256`, manifests, raw JSON, and `/usr/bin/time -v` files for the evidence chain.",
        "",
        "| lane | input bytes | retained raw before/after | context storage nodes before/after | median elapsed ns before/after | median requested alloc bytes before/after | median peak live bytes before/after | median RSS KiB before/after |",
        "|---|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for lane in LANES:
        before = side_values(args.before, lane)
        after = side_values(args.after, lane)
        if before["input"] != after["input"]:
            raise SystemExit(f"fixture input changed for {lane}")
        lines.append(
            f"| {lane} | {before['input']} | {before['raw']}/{after['raw']} | "
            f"{before['context']}/{after['context']} | {before['elapsed']}/{after['elapsed']} | "
            f"{before['alloc']}/{after['alloc']} | {before['peak']}/{after['peak']} | "
            f"{before['rss']}/{after['rss']} |"
        )
    args.output.write_text("\n".join(lines) + "\n")


if __name__ == "__main__":
    main()
