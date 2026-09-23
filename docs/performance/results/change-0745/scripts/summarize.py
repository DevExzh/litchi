#!/usr/bin/env python3
"""Render change 0745 packet tables (Markdown) from the analysis outputs.

Usage: summarize.py PACKET_DIR > summary.md
"""

import collections
import json
import statistics
import sys


def fmt_change(entry):
    return "{:+.1f}% [{:+.1f}, {:+.1f}] ({:+.0f}..{:+.0f})".format(
        entry["median_change_pct"],
        entry["bootstrap95_pct"][0],
        entry["bootstrap95_pct"][1],
        entry["min_change_pct"],
        entry["max_change_pct"],
    )


def main():
    packet = sys.argv[1]
    analysis = json.load(open(f"{packet}/matrix/analysis.json"))
    counters = json.load(open(f"{packet}/counters/counters.json"))
    allocation = json.load(open(f"{packet}/allocation/allocation.json"))
    out = []
    out.append("## Timing: medians over 18 paired heap layouts (54 processes per case)\n")
    out.append("Process p50 in µs is the median over the 18 processes of each arm. Changes are the median of 18 paired per-layout changes, the bootstrap 95% interval of that median, and the min..max over layouts.\n")
    out.append("| Case | A p50 | B p50 | C p50 | B vs A p50 | C vs B p50 | C vs A p50 |")
    out.append("|---|---:|---:|---:|---|---|---|")
    for case, entry in analysis["cases"].items():
        medians = entry["arm_medians"]
        comparisons = entry["comparisons"]
        out.append(
            "| {} | {:.0f} | {:.0f} | {:.0f} | {} | {} | {} |".format(
                case,
                medians["A"]["p50_ns"] / 1000,
                medians["B"]["p50_ns"] / 1000,
                medians["C"]["p50_ns"] / 1000,
                fmt_change(comparisons["B/A"]["p50_ns"]),
                fmt_change(comparisons["C/B"]["p50_ns"]),
                fmt_change(comparisons["C/A"]["p50_ns"]),
            )
        )
    out.append("\n### Mean and p95 (median paired change over layouts)\n")
    out.append("| Case | B vs A mean | B vs A p95 | C vs B mean | C vs B p95 | C vs A mean | C vs A p95 |")
    out.append("|---|---:|---:|---:|---:|---:|---:|")
    for case, entry in analysis["cases"].items():
        comparisons = entry["comparisons"]
        cells = []
        for pair in ("B/A", "C/B", "C/A"):
            for stat in ("mean_ns", "p95_ns"):
                cells.append("{:+.1f}%".format(comparisons[pair][stat]["median_change_pct"]))
        out.append("| {} | {} |".format(case, " | ".join(cells)))
    out.append("\n### Every paired change above +5% (regression flags)\n")
    grouped = collections.defaultdict(list)
    for flag in analysis["flags"]:
        grouped[(flag["case"], flag["comparison"], flag["stat"])].append(flag)
    out.append("| Case | Comparison | Statistic | Flagged layouts (of 18) | Largest | Median-level flag |")
    out.append("|---|---|---|---:|---:|---|")
    for (case, pair, stat), flags in sorted(grouped.items()):
        rounds = [flag for flag in flags if flag["round"] != "median"]
        median_flag = [flag for flag in flags if flag["round"] == "median"]
        out.append(
            "| {} | {} | {} | {} | {:+.1f}% | {} |".format(
                case,
                pair,
                stat.replace("_ns", ""),
                len(rounds),
                max((flag["change_pct"] for flag in rounds), default=0.0),
                "{:+.1f}%".format(median_flag[0]["change_pct"]) if median_flag else "no",
            )
        )
    out.append("\n## Per-owner hardware counters (perf stat, 120 minus 20 owners, three layouts)\n")
    out.append("Median per-owner value over the three layouts; instruction counts are user mode. Page faults are listed per layout.\n")
    out.append("| Case | Instructions A (M) | B − A (M) | C − B (M) | Page faults A | Page faults B | Page faults C |")
    out.append("|---|---:|---:|---:|---|---|---|")
    for case, events in counters["summary"].items():
        instructions = events["instructions:u"]
        faults = events["page-faults"]
        med = {arm: statistics.median(instructions[arm]) for arm in "ABC"}
        out.append(
            "| {} | {:.2f} | {:+.2f} | {:+.2f} | {} | {} | {} |".format(
                case,
                med["A"] / 1e6,
                (med["B"] - med["A"]) / 1e6,
                (med["C"] - med["B"]) / 1e6,
                ", ".join(f"{value:.0f}" for value in faults["A"]),
                ", ".join(f"{value:.0f}" for value in faults["B"]),
                ", ".join(f"{value:.0f}" for value in faults["C"]),
            )
        )
    out.append("\n## Allocation per owner (counting allocator, deterministic across owners)\n")
    out.append("| Case | Allocated bytes A | B − A | C − B | Calls A | B − A | C − B | Peak live A | B − A | C − B | Retained A | B − A | C − B |")
    out.append("|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|")
    for case, arms in allocation.items():
        row = [case]
        for key in ("allocated_bytes", "allocation_calls", "peak_live_bytes", "retained_bytes"):
            row += [
                f"{arms['A'][key]:,}",
                f"{arms['B'][key] - arms['A'][key]:+,}",
                f"{arms['C'][key] - arms['B'][key]:+,}",
            ]
        out.append("| " + " | ".join(row) + " |")
    print("\n".join(out))


if __name__ == "__main__":
    main()
