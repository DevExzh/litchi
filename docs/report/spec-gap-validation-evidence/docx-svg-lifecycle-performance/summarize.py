#!/usr/bin/env python3
"""Recompute DOCX SVG lifecycle profile rows from raw receipts."""

from __future__ import annotations

import argparse
import json
import math
from pathlib import Path


LANES = (
    "native_svg_capture",
    "native_floating_capture",
    "lazy_inventory_1",
    "lazy_inventory_64",
    "single_attach_1",
    "single_attach_16",
    "single_attach_64",
    "single_detach_1",
    "single_detach_16",
    "single_detach_64",
    "batch_attach_1",
    "batch_attach_16",
    "batch_attach_64",
    "batch_detach_1",
    "batch_detach_16",
    "batch_detach_64",
    "shared_svg_cleanup",
    "exact_inverse_single_1",
    "exact_inverse_batch_64",
    "large_unchanged_media_managed_cap",
    "noop_detach_64",
)


def quantile(values: list[int], percent: int) -> int:
    return sorted(values)[max(0, math.ceil(len(values) * percent / 100) - 1)]


def rss(path: Path) -> int:
    marker = "Maximum resident set size (kbytes):"
    for line in path.read_text().splitlines():
        if line.lstrip().startswith(marker):
            return int(line.split(":", 1)[1].strip())
    raise ValueError(f"RSS is missing from {path}")


def rows(results: Path, mode: str) -> list[dict[str, object]]:
    output: list[dict[str, object]] = []
    expected_processes = 1 if mode == "smoke" else 3
    for lane in LANES:
        paths = [
            results / f"{mode}-{lane}-p{process}.json"
            for process in range(1, expected_processes + 1)
            if (results / f"{mode}-{lane}-p{process}.json").is_file()
        ]
        if len(paths) != expected_processes:
            raise ValueError(f"process count mismatch for {lane}")
        payloads = [json.loads(path.read_text()) for path in paths]
        samples = [sample for payload in payloads for sample in payload["samples"]]
        input_bytes = {int(payload["input_bytes"]) for payload in payloads}
        elapsed = [int(sample["elapsed_ns"]) for sample in samples]
        allocated = [int(sample["requested_alloc_bytes"]) for sample in samples]
        peak = [int(sample["peak_live_delta"]) for sample in samples]
        phase_names = (
            "capture_ns",
            "stage_ns",
            "commit_ns",
            "publish_ns",
            "reopen_ns",
            "inverse_reopen_ns",
            "inverse_ns",
            "payload_ns",
            "validation_ns",
            "readback_ns",
        )
        phase_p50 = {
            name: quantile([int(sample["phases"][name]) for sample in samples], 50)
            for name in phase_names
        }
        dominant = max(phase_p50, key=phase_p50.get)
        rss_values = [rss(path.with_suffix(".time.txt")) for path in paths]
        output.append(
            {
                "lane": lane,
                "processes": len(paths),
                "samples": len(samples),
                "input_bytes": next(iter(input_bytes)),
                "elapsed": (quantile(elapsed, 50), quantile(elapsed, 95), quantile(elapsed, 99)),
                "alloc": (quantile(allocated, 50), quantile(allocated, 95)),
                "peak": (quantile(peak, 50), quantile(peak, 95)),
                "rss": (min(rss_values), max(rss_values)),
                "phase_p50": phase_p50,
                "dominant_phase": dominant,
            }
        )
    return output


def write_report(path: Path, mode: str, table: list[dict[str, object]]) -> None:
    title = "smoke correctness profile" if mode == "smoke" else "profile"
    lines = [
        f"# DOCX source-backed SVG lifecycle {title}",
        "",
        "This report is recomputed from raw receipts. It is an absolute,",
        "scenario-scoped observation and makes no before/after speedup claim.",
        "Allocation bytes, peak live bytes, and process RSS are distinct metrics.",
        "",
        "| lane | processes | samples | input bytes | elapsed ns p50/p95/p99 | requested alloc bytes p50/p95 | peak live delta p50/p95 | RSS KiB min-max | dominant phase (p50 ns) |",
        "|---|---:|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for row in table:
        phase = row["phase_p50"][row["dominant_phase"]]
        lines.append(
            f"| {row['lane']} | {row['processes']} | {row['samples']} | {row['input_bytes']} | "
            f"{row['elapsed'][0]}/{row['elapsed'][1]}/{row['elapsed'][2]} | "
            f"{row['alloc'][0]}/{row['alloc'][1]} | "
            f"{row['peak'][0]}/{row['peak'][1]} | "
            f"{row['rss'][0]}–{row['rss'][1]} | {row['dominant_phase']} ({phase}) |"
        )
    path.write_text("\n".join(lines) + "\n")


def write_dominant(path: Path, mode: str, table: list[dict[str, object]]) -> None:
    lines = [
        "# DOCX SVG lifecycle dominant phase observations",
        "",
        "These are p50 phase maxima within each named lane. They identify where",
        "the scaffold spends time; they are not a causal attribution or speedup",
        "claim. Fixture setup and caller-owned payload construction are outside",
        "the timed operation.",
        "",
        f"mode={mode}",
        "",
        "| lane | dominant phase | p50 ns | capture | stage | commit | publish | reopen | inverse reopen | inverse | payload | validation | readback |",
        "|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for row in table:
        phases = row["phase_p50"]
        lines.append(
            f"| {row['lane']} | {row['dominant_phase']} | {phases[row['dominant_phase']]} | "
            f"{phases['capture_ns']} | {phases['stage_ns']} | {phases['commit_ns']} | "
            f"{phases['publish_ns']} | {phases['reopen_ns']} | {phases['inverse_reopen_ns']} | "
            f"{phases['inverse_ns']} | {phases['payload_ns']} | {phases['validation_ns']} | "
            f"{phases['readback_ns']} |"
        )
    path.write_text("\n".join(lines) + "\n")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--results", type=Path, required=True)
    parser.add_argument("--mode", choices=("smoke", "full"), required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--dominant-output", type=Path, required=True)
    args = parser.parse_args()
    table = rows(args.results, args.mode)
    write_report(args.output, args.mode, table)
    write_dominant(args.dominant_output, args.mode, table)


if __name__ == "__main__":
    main()
