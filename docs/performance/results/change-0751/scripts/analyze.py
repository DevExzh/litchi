#!/usr/bin/env python3
"""Summarize the change 0751 ABBA matrix from its retained raw reports.

Adapted from change 0742's analyzer. For every lane and case it reports, per
process, the p50/p95/mean of the harness's elapsed samples, the phase medians
and the whole-child ``perf stat`` cycles and instructions when present; per
arm, the median of the process p50s with its min-max, and the medians of the
p95s, means, phases, cycles and instructions; and the paired after/before
ratio of the two processes that share a round and a position pair
((s0 before, s1 after) and (s3 before, s2 after)) with a percentile bootstrap
interval of the median paired ratio (10,000 resamples of the pairs, seed
751). Phase, cycle and instruction ratios are paired the same way. The
allocator lane reports allocator counters. Any case whose median paired
ratio exceeds 1.05 is flagged as a regression.

Usage: analyze.py RAW_DIR [--json OUT] [--markdown OUT]
"""

from __future__ import annotations

import argparse
import gzip
import json
import random
import re
import statistics
import sys
from pathlib import Path

REPORT = re.compile(r"r(?P<round>\d+)-s(?P<slot>\d+)-(?P<arm>before|after)\.json(?:\.gz)?$")
PAIRS = ((0, 1), (3, 2))
PHASES = ("plan_ns", "commit_ns", "publication_ns", "open_ns")
ALLOCATION_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "allocated_bytes",
    "region_peak_live_bytes",
)
COUNTERS = ("cycles", "instructions")


def ms(value: float) -> float:
    return value / 1e6


def read_text(path: Path) -> str:
    """Read a raw file, gzipped (as retained in the packet) or not."""
    if path.suffix == ".gz":
        return gzip.decompress(path.read_bytes()).decode("utf-8")
    return path.read_text(encoding="utf-8")


def sibling(path: Path, suffix: str) -> Path:
    """The file beside a report with `suffix` in place of `.json`, gzipped
    when the report is."""
    name = path.name.removesuffix(".gz").removesuffix(".json") + suffix
    return path.with_name(name + (".gz" if path.name.endswith(".gz") else ""))


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


def perf_counters(path: Path) -> dict[str, float]:
    counters: dict[str, float] = {}
    if not path.exists():
        return counters
    for line in read_text(path).splitlines():
        fields = line.split(",")
        if len(fields) < 3 or not fields[0].strip():
            continue
        event = fields[2].split(":")[0]
        if event in COUNTERS:
            try:
                counters[event] = float(fields[0])
            except ValueError:
                continue
    return counters


def process_rows(path: Path) -> dict[str, dict]:
    """One row per result of a report, keyed by corpus name when a case runs
    more than one corpus (the semantic controls run three)."""
    report = json.loads(read_text(path))
    results = report["results"]
    rows = {}
    for result in results:
        key = "" if len(results) == 1 else (result.get("corpus") or {}).get("name", "")
        rows[key] = process_row(path, report, result)
    return rows


def process_row(path: Path, report: dict, result: dict) -> dict:
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
        "perf": perf_counters(sibling(path, ".perf.csv")),
    }
    allocation = (result.get("operation_metrics") or {}).get("allocation") or {}
    if allocation.get("status") == "measured":
        row["allocation_median"] = {
            field: statistics.median(allocation[field]["values"])
            for field in ALLOCATION_FIELDS
        }
    return row


def bootstrap_median(values: list[float], seed: int = 751, draws: int = 10_000) -> list[float]:
    generator = random.Random(seed)
    medians = sorted(
        statistics.median(generator.choices(values, k=len(values))) for _ in range(draws)
    )
    return [medians[int(0.025 * draws)], medians[int(0.975 * draws) - 1]]


def paired(processes: dict, value) -> list[float]:
    ratios = []
    rounds = sorted({round_index for round_index, _ in processes})
    for round_index in rounds:
        for before_slot, after_slot in PAIRS:
            before = processes.get((round_index, before_slot))
            after = processes.get((round_index, after_slot))
            if before and after:
                numerator, denominator = value(after), value(before)
                if numerator is not None and denominator:
                    ratios.append(numerator / denominator)
    return ratios


def summarize_case_dir(case_dir: Path) -> dict[str, dict]:
    by_corpus: dict[str, dict[tuple[int, int], dict]] = {}
    for path in sorted(case_dir.glob("*.json*")):
        match = REPORT.search(path.name)
        if not match:
            continue
        for corpus, row in process_rows(path).items():
            row["arm"] = match["arm"]
            by_corpus.setdefault(corpus, {})[(int(match["round"]), int(match["slot"]))] = row
    return {
        (case_dir.name if not corpus else f"{case_dir.name} [{corpus}]"): summarize_case(processes)
        for corpus, processes in sorted(by_corpus.items())
    }


def summarize_case(processes: dict[tuple[int, int], dict]) -> dict:
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
        for counter in COUNTERS:
            values = [row["perf"][counter] for row in rows if counter in row["perf"]]
            if len(values) == len(rows) and values:
                summary[arm][f"median_process_{counter}"] = statistics.median(values)
        if all("allocation_median" in row for row in rows):
            summary[arm]["allocation_median_of_processes"] = {
                field: statistics.median(row["allocation_median"][field] for row in rows)
                for field in ALLOCATION_FIELDS
            }
    ratios = paired(processes, lambda row: row["p50_ms"])
    summary["paired_ratios_after_over_before"] = ratios
    summary["median_paired_ratio"] = statistics.median(ratios)
    summary["bootstrap95_median_paired_ratio"] = bootstrap_median(ratios)
    summary["ratio_of_arm_medians"] = (
        summary["after"]["median_process_p50_ms"] / summary["before"]["median_process_p50_ms"]
    )
    summary["regression_flag_over_5pct"] = summary["median_paired_ratio"] > 1.05
    phase_ratios = {}
    for phase in ("plan_ns", "commit_ns", "publication_ns"):
        values = paired(processes, lambda row, phase=phase: row["phases_median_ms"].get(phase))
        if values:
            phase_ratios[phase] = {
                "median": statistics.median(values),
                "bootstrap95": bootstrap_median(values),
            }
    if phase_ratios:
        summary["phase_paired_ratios"] = phase_ratios
    counter_ratios = {}
    for counter in COUNTERS:
        values = paired(processes, lambda row, counter=counter: row["perf"].get(counter))
        if values:
            counter_ratios[counter] = {
                "median": statistics.median(values),
                "bootstrap95": bootstrap_median(values),
            }
    if counter_ratios:
        summary["perf_paired_ratios"] = counter_ratios
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
        lines.append("| case | p95 before → after (ms) | mean before → after (ms) |")
        lines.append("|---|---:|---:|")
        for case, summary in cases.items():
            before, after = summary["before"], summary["after"]
            lines.append(
                f"| `{case}` | {before['median_process_p95_ms']:.3f} → "
                f"{after['median_process_p95_ms']:.3f} | {before['median_process_mean_ms']:.3f} → "
                f"{after['median_process_mean_ms']:.3f} |"
            )
        lines.append("")
        if any("perf_paired_ratios" in summary for summary in cases.values()):
            lines.append(
                "| case | cycles before → after (whole child) | ratio [95%] | "
                "instructions before → after | ratio [95%] |"
            )
            lines.append("|---|---:|---:|---:|---:|")
            for case, summary in cases.items():
                ratios = summary.get("perf_paired_ratios")
                if not ratios:
                    continue
                before, after = summary["before"], summary["after"]
                cycles, instructions = ratios["cycles"], ratios["instructions"]
                lines.append(
                    f"| `{case}` | {before['median_process_cycles']:,.0f} → "
                    f"{after['median_process_cycles']:,.0f} | {cycles['median']:.4f} "
                    f"[{cycles['bootstrap95'][0]:.4f}, {cycles['bootstrap95'][1]:.4f}] | "
                    f"{before['median_process_instructions']:,.0f} → "
                    f"{after['median_process_instructions']:,.0f} | {instructions['median']:.4f} "
                    f"[{instructions['bootstrap95'][0]:.4f}, {instructions['bootstrap95'][1]:.4f}] |"
                )
            lines.append("")
        for case, summary in cases.items():
            phases_before = summary["before"].get("phase_median_of_process_medians_ms")
            phases_after = summary["after"].get("phase_median_of_process_medians_ms")
            if phases_before and phases_after:
                lines.append(f"`{case}` phases (median of process medians, ms; median paired ratio):")
                for phase in phases_before:
                    ratio = summary.get("phase_paired_ratios", {}).get(phase)
                    ratio_text = (
                        f" ({ratio['median']:.4f} [{ratio['bootstrap95'][0]:.4f}, "
                        f"{ratio['bootstrap95'][1]:.4f}])"
                        if ratio
                        else ""
                    )
                    lines.append(
                        f"- {phase}: {phases_before[phase]:.3f} -> "
                        f"{phases_after.get(phase, float('nan')):.3f}{ratio_text}"
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
    lines.append("## Per-process rows\n")
    for lane, cases in analysis.items():
        for case, summary in cases.items():
            lines.append(f"`{lane}` `{case}`:\n")
            lines.append("| process | arm | p50 ms | p95 ms | mean ms |")
            lines.append("|---|---|---:|---:|---:|")
            for name, row in summary["processes"].items():
                lines.append(
                    f"| {name} | {row['arm']} | {row['p50_ms']:.3f} | "
                    f"{row['p95_ms']:.3f} | {row['mean_ms']:.3f} |"
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
        cases: dict = {}
        for case_dir in sorted(path for path in lane_dir.iterdir() if path.is_dir()):
            cases.update(summarize_case_dir(case_dir))
        analysis[lane_dir.name] = cases
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
