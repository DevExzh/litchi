#!/usr/bin/env python3
"""Summarize raw source-projection receipts without making a performance claim."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from statistics import median


LANES = [
    "contextual_read_p16_n0", "contextual_read_p16_n32", "contextual_read_p16_n128",
    "contextual_read_p32_n0", "contextual_read_p32_n32", "contextual_read_p32_n128",
    "contextual_read_p128_n0", "contextual_read_p128_n32", "contextual_read_p128_n128",
    "standalone_export_p16_n0", "standalone_export_p16_n32", "standalone_export_p16_n128",
    "standalone_export_p32_n0", "standalone_export_p32_n32", "standalone_export_p32_n128",
    "standalone_export_p128_n0", "standalone_export_p128_n32", "standalone_export_p128_n128",
    "scalar_reference_edit_p16_n0", "scalar_reference_edit_p16_n32", "scalar_reference_edit_p16_n128",
    "scalar_reference_edit_p32_n0", "scalar_reference_edit_p32_n32", "scalar_reference_edit_p32_n128",
    "scalar_reference_edit_p128_n0", "scalar_reference_edit_p128_n32", "scalar_reference_edit_p128_n128",
    "small_cap_refusal_p16_n0", "small_cap_refusal_p16_n32", "small_cap_refusal_p16_n128",
    "small_cap_refusal_p32_n0", "small_cap_refusal_p32_n32", "small_cap_refusal_p32_n128",
    "small_cap_refusal_p128_n0", "small_cap_refusal_p128_n32", "small_cap_refusal_p128_n128",
    "contextual_read_original_149433", "standalone_export_original_149433",
    "small_cap_refusal_original_149433",
]


def percentile(values: list[int], fraction: float) -> int:
    values = sorted(values)
    if not values:
        raise ValueError("empty sample set")
    index = min(len(values) - 1, int(round((len(values) - 1) * fraction)))
    return values[index]


def time_seconds(text: str) -> float:
    match = re.search(r"Elapsed \(wall clock\) time \(h:mm:ss or m:ss\):\s*(\S+)", text)
    if not match:
        raise ValueError("elapsed wall time missing")
    fields = match.group(1).split(":")
    if len(fields) == 3:
        hours, minutes, seconds = fields
        return int(hours) * 3600 + int(minutes) * 60 + float(seconds)
    if len(fields) == 2:
        minutes, seconds = fields
        return int(minutes) * 60 + float(seconds)
    return float(fields[0])


def rss_kib(text: str) -> int:
    match = re.search(r"Maximum resident set size \(kbytes\):\s*(\d+)", text)
    if not match:
        raise ValueError("maximum RSS missing")
    return int(match.group(1))


def load_lane(results: Path, lane: str) -> tuple[list[dict[str, object]], list[int], list[int]]:
    receipts: list[dict[str, object]] = []
    rss: list[int] = []
    for path in sorted(results.glob(f"{lane}-p*.json")):
        receipts.append(json.loads(path.read_text()))
        timing = path.with_suffix(".time.txt")
        text = timing.read_text()
        rss.append(rss_kib(text))
    if len(receipts) != 3:
        raise ValueError(f"expected three receipts for {lane}, got {len(receipts)}")
    elapsed: list[int] = []
    requested: list[int] = []
    peak: list[int] = []
    for receipt in receipts:
        for sample in receipt["samples"]:
            elapsed.append(int(sample["elapsed_ns"]))
            requested.append(int(sample["requested_alloc_bytes"]))
            peak.append(int(sample["peak_live_delta"]))
    return receipts, rss, [
        percentile(elapsed, 0.0), percentile(elapsed, 0.5), percentile(elapsed, 1.0),
        percentile(requested, 0.0), percentile(requested, 0.5), percentile(requested, 1.0),
        percentile(peak, 0.0), percentile(peak, 0.5), percentile(peak, 1.0),
    ]


def rows(results: Path) -> list[dict[str, object]]:
    output: list[dict[str, object]] = []
    for lane in LANES:
        receipts, rss, values = load_lane(results, lane)
        first = receipts[0]
        output.append(
            {
                "lane": lane,
                "processes": len(receipts),
                "samples": int(first["sample_count"]),
                "input_bytes": int(first["input_bytes"]),
                "elapsed": tuple(values[0:3]),
                "alloc": tuple(values[3:6]),
                "peak": tuple(values[6:9]),
                "rss": tuple(sorted(rss)),
            }
        )
    return output


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--results", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    values = rows(args.results)
    lines = [
        "# XLSX SVG source contextual profile",
        "",
        "This report is mechanically recomputed from the raw JSON receipts and `/usr/bin/time -v` files. It is exploratory source-bound evidence; it makes no XLSX lifecycle, native acceptance, or optimization-factor claim.",
        "",
        "Elapsed, allocation, and peak-live columns are `min/median/max` across all measured samples (nanoseconds or bytes). RSS is `min/median/max` across the three fresh processes (KiB).",
        "",
        "| lane | processes | samples | input bytes | elapsed ns min/median/max | requested alloc bytes min/median/max | peak live bytes min/median/max | RSS KiB min/median/max |",
        "|---|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for row in values:
        lines.append(
            "| {lane} | {processes} | {samples} | {input_bytes} | {elapsed} | {alloc} | {peak} | {rss} |".format(
                lane=row["lane"],
                processes=row["processes"],
                samples=row["samples"],
                input_bytes=row["input_bytes"],
                elapsed="/".join(str(value) for value in row["elapsed"]),
                alloc="/".join(str(value) for value in row["alloc"]),
                peak="/".join(str(value) for value in row["peak"]),
                rss="/".join(str(value) for value in row["rss"]),
            )
        )
    args.output.write_text("\n".join(lines) + "\n")


if __name__ == "__main__":
    main()
