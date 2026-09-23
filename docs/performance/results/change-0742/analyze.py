#!/usr/bin/env python3
"""Summarize the change 0742 ABBA matrix from its retained raw reports.

For every lane (``native``, ``alloc``) and case it reports, per process, the
p50/p95/mean of the harness's elapsed samples; per arm, the median of the
process p50s and their spread; and the paired after/before ratio of the two
processes that share a round and a position pair ((s0 before, s1 after) and
(s3 before, s2 after)) with a percentile bootstrap interval of the median
paired ratio (10,000 resamples of the pairs, seed 742). Owned cross-copy
cases also report the medians of the plan/commit/publication phase vectors,
and the allocator lane reports allocator counters. Any case whose median
paired ratio exceeds 1.05 is flagged as a regression.

Usage: analyze.py RAW_DIR [--json OUT] [--markdown OUT]
"""

from __future__ import annotations

import argparse
import json
import random
import re
import statistics
import sys
from pathlib import Path

REPORT = re.compile(r"r(?P<round>\d+)-s(?P<slot>\d+)-(?P<arm>before|after)\.json$")
PAIRS = ((0, 1), (3, 2))
PHASES = ("plan_ns", "commit_ns", "publication_ns", "open_ns")
ALLOCATION_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "allocated_bytes",
    "region_peak_live_bytes",
)


def ms(value: float) -> float:
    return value / 1e6


def phase_vectors(result: dict) -> dict[str, list[int]]:
    source = result.get("source") or {}
    for key in ("pptx_cross_copy", "pptx_source_backed_cross_copy_lifecycle"):
        summary = source.get(key)
        if summary:
            return {
                phase: summary[phase]
                for phase in PHASES
                if isinstance(summary.get(phase), list) and summary[phase]
            }
    return {}


def process_row(path: Path) -> dict:
    report = json.loads(path.read_text(encoding="utf-8"))
    (result,) = report["results"]
    elapsed = result["elapsed_ns"]
    row = {
        "report": path.name,
        "binary_sha256": report["binary_identity"]["binary_sha256"],
        "cpu_affinity": report["environment"].get("cpu_affinity"),
        "samples": elapsed["samples"],
        "p50_ms": ms(elapsed["p50"]),
        "p95_ms": ms(elapsed["p95"]),
        "mean_ms": ms(elapsed["mean"]),
        "output_sha256": result.get("output_sha256"),
        "phases_median_ms": {
            phase: ms(statistics.median(values))
            for phase, values in phase_vectors(result).items()
        },
    }
    allocation = (result.get("operation_metrics") or {}).get("allocation") or {}
    if allocation.get("status") == "measured":
        row["allocation_median"] = {
            field: statistics.median(allocation[field]["values"])
            for field in ALLOCATION_FIELDS
        }
    return row


def bootstrap_median(values: list[float], seed: int = 742, draws: int = 10_000) -> list[float]:
    generator = random.Random(seed)
    medians = sorted(
        statistics.median(generator.choices(values, k=len(values))) for _ in range(draws)
    )
    return [medians[int(0.025 * draws)], medians[int(0.975 * draws) - 1]]


def summarize_case(case_dir: Path) -> dict:
    processes: dict[tuple[int, int], dict] = {}
    for path in sorted(case_dir.glob("*.json")):
        match = REPORT.search(path.name)
        if not match:
            continue
        row = process_row(path)
        row["arm"] = match["arm"]
        processes[(int(match["round"]), int(match["slot"]))] = row
    arms = {
        arm: [row for row in processes.values() if row["arm"] == arm]
        for arm in ("before", "after")
    }
    summary: dict = {"processes": {f"r{r}-s{s}": row for (r, s), row in sorted(processes.items())}}
    for arm, rows in arms.items():
        p50s = [row["p50_ms"] for row in rows]
        summary[arm] = {
            "processes": len(rows),
            "median_process_p50_ms": statistics.median(p50s),
            "min_process_p50_ms": min(p50s),
            "max_process_p50_ms": max(p50s),
            "median_process_p95_ms": statistics.median(row["p95_ms"] for row in rows),
            "median_process_mean_ms": statistics.median(row["mean_ms"] for row in rows),
            "binary_sha256": sorted({row["binary_sha256"] for row in rows}),
            "output_sha256": sorted({str(row["output_sha256"]) for row in rows}),
        }
        phases = {phase for row in rows for phase in row["phases_median_ms"]}
        if phases:
            summary[arm]["phase_median_of_process_medians_ms"] = {
                phase: statistics.median(row["phases_median_ms"][phase] for row in rows)
                for phase in sorted(phases)
            }
        if all("allocation_median" in row for row in rows):
            summary[arm]["allocation_median_of_processes"] = {
                field: statistics.median(row["allocation_median"][field] for row in rows)
                for field in ALLOCATION_FIELDS
            }
    ratios = []
    rounds = sorted({round_index for round_index, _ in processes})
    for round_index in rounds:
        for before_slot, after_slot in PAIRS:
            before = processes.get((round_index, before_slot))
            after = processes.get((round_index, after_slot))
            if before and after:
                ratios.append(after["p50_ms"] / before["p50_ms"])
    summary["paired_ratios_after_over_before"] = ratios
    summary["median_paired_ratio"] = statistics.median(ratios)
    summary["bootstrap95_median_paired_ratio"] = bootstrap_median(ratios)
    summary["ratio_of_arm_medians"] = (
        summary["after"]["median_process_p50_ms"] / summary["before"]["median_process_p50_ms"]
    )
    summary["regression_flag_over_5pct"] = summary["median_paired_ratio"] > 1.05
    return summary


def markdown(analysis: dict) -> str:
    lines = []
    for lane, cases in analysis.items():
        lines.append(f"### {lane}\n")
        lines.append(
            "| case | before median p50 ms [min–max] | after median p50 ms [min–max] | "
            "median paired ratio [bootstrap 95%] | flag |"
        )
        lines.append("|---|---:|---:|---:|---|")
        for case, summary in cases.items():
            before, after = summary["before"], summary["after"]
            low, high = summary["bootstrap95_median_paired_ratio"]
            lines.append(
                f"| `{case}` | {before['median_process_p50_ms']:.3f} "
                f"[{before['min_process_p50_ms']:.3f}–{before['max_process_p50_ms']:.3f}] | "
                f"{after['median_process_p50_ms']:.3f} "
                f"[{after['min_process_p50_ms']:.3f}–{after['max_process_p50_ms']:.3f}] | "
                f"{summary['median_paired_ratio']:.4f} [{low:.4f}, {high:.4f}] | "
                f"{'REGRESSION >5%' if summary['regression_flag_over_5pct'] else '-'} |"
            )
        lines.append("")
        for case, summary in cases.items():
            phases_before = summary["before"].get("phase_median_of_process_medians_ms")
            phases_after = summary["after"].get("phase_median_of_process_medians_ms")
            if phases_before and phases_after:
                lines.append(f"`{case}` phases (median of process medians, ms):")
                for phase in phases_before:
                    lines.append(
                        f"- {phase}: {phases_before[phase]:.3f} -> {phases_after.get(phase, float('nan')):.3f}"
                    )
                lines.append("")
            alloc_before = summary["before"].get("allocation_median_of_processes")
            alloc_after = summary["after"].get("allocation_median_of_processes")
            if alloc_before and alloc_after:
                lines.append(f"`{case}` allocator counters (median of process medians):")
                for field in ALLOCATION_FIELDS:
                    lines.append(
                        f"- {field}: {alloc_before[field]:,.0f} -> {alloc_after[field]:,.0f}"
                    )
                lines.append("")
    return "\n".join(lines).rstrip("\n")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("raw", type=Path)
    parser.add_argument("--json", type=Path)
    parser.add_argument("--markdown", type=Path)
    args = parser.parse_args()
    analysis: dict = {}
    for lane_dir in sorted(path for path in args.raw.iterdir() if path.is_dir()):
        analysis[lane_dir.name] = {
            case_dir.name: summarize_case(case_dir)
            for case_dir in sorted(path for path in lane_dir.iterdir() if path.is_dir())
        }
    text = json.dumps(analysis, indent=2, sort_keys=True)
    if args.json:
        args.json.write_text(text + "\n", encoding="utf-8")
    rendered = markdown(analysis)
    if args.markdown:
        args.markdown.write_text(rendered + "\n", encoding="utf-8")
    print(rendered)
    return 0


if __name__ == "__main__":
    sys.exit(main())
