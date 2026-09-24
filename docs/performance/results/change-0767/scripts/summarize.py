#!/usr/bin/env python3
"""Change 0767: summary tables from the lanes' outputs.

Usage: summarize.py PACKET_DIR > summary.md

Reads PACKET_DIR/matrix/analysis.json, counters/counters.json,
scaling/callgrind-scaling.json, allocation/allocation.json and
verdicts/verdicts-summary.json (each optional) and prints Markdown tables:
every timing row, every flag above +5%, per-owner instructions and cycles,
the per-stream scaling curves, the Callgrind scaling and the allocation
deltas.
"""

import json
import os
import statistics
import sys

COUNTS = (1000, 3000, 10000)


def load(path):
    if not os.path.exists(path):
        return None
    with open(path) as handle:
        return json.load(handle)


def us(ns):
    return f"{ns / 1000:,.1f}"


def timing(analysis):
    print("## Timing (median of per-round paired changes, cand vs base)\n")
    print("| Case | Rounds | Base p50 µs | Cand p50 µs | Δ p50 [95% CI] | Δ mean | Δ p95 |")
    print("|---|---:|---:|---:|---:|---:|---:|")
    for name, case in analysis["cases"].items():
        arms = case["arms"]
        change = case["median_paired_change_pct"]
        interval = case["bootstrap95_median_paired_change_pct"]["p50"]
        print(f"| `{name}` | {case['paired_rounds']} | {us(arms['base']['median_process_p50_ns'])} | "
              f"{us(arms['cand']['median_process_p50_ns'])} | {change['p50']:+.2f}% "
              f"[{interval[0]:+.2f}, {interval[1]:+.2f}] | {change['mean']:+.2f}% | {change['p95']:+.2f}% |")
    print(f"\nEvery probe output identical across arms: "
          f"{analysis['all_probe_outputs_identical_across_arms']}.\n")
    print("## Flags (a round's paired p50, mean or p95 above +5%)\n")
    flags = analysis["flags_above_plus_5_pct"]
    print(f"{len(flags)} flags.\n")
    print("| Case | Round | Metric | Change | Base µs | Cand µs |")
    print("|---|---:|---|---:|---:|---:|")
    for flag in flags:
        print(f"| `{flag['case']}` | {flag['round']} | {flag['metric']} | {flag['change_pct']:+.2f}% | "
              f"{us(flag['base_ns'])} | {us(flag['cand_ns'])} |")
    print()


def counters(data):
    print("## Per-owner counters (median of three layouts)\n")
    print("| Case | instructions:u base | cand | Δ | cycles:u base | cand | Δ | page faults base / cand |")
    print("|---|---:|---:|---:|---:|---:|---:|---:|")
    for name, events in data["summary"].items():
        ins, cyc, faults = events["instructions:u"], events["cycles:u"], events["page-faults"]
        print(f"| `{name}` | {statistics.median(ins['base']):,.0f} | {statistics.median(ins['cand']):,.0f} | "
              f"{ins.get('median_change_pct', 0):+.2f}% | {statistics.median(cyc['base']):,.0f} | "
              f"{statistics.median(cyc['cand']):,.0f} | {cyc.get('median_change_pct', 0):+.2f}% | "
              f"{statistics.median(faults['base']):.1f} / {statistics.median(faults['cand']):.1f} |")
    print()


def scaling(analysis, data):
    """Per-stream time and instructions for open and read-all, by scale."""
    print("## Scaling: work and time per stream\n")
    print("Read-all owners open a fresh reader untimed but inside the counted loop, so the "
          "read-all instructions are the read-all owner's minus the open owner's.\n")
    print("| Input | Streams | Open ns/stream base → cand | Open instr/stream base → cand | "
          "Read-all ns/stream base → cand | Read-all instr/stream base → cand |")
    print("|---|---:|---:|---:|---:|---:|")
    for version in ("v3", "v4"):
        for kind in ("mini", "regular"):
            for count in COUNTS:
                name = f"{version}-{kind}-{count}"
                cells = []
                for lane in ("open", "readall"):
                    case = analysis["cases"].get(f"{lane}-{name}") if analysis else None
                    if case:
                        base = case["arms"]["base"]["median_process_p50_ns"] / count
                        cand = case["arms"]["cand"]["median_process_p50_ns"] / count
                        cells.append(f"{base:,.0f} → {cand:,.0f}")
                    else:
                        cells.append("")
                    if data:
                        def per(arm, lane=lane):
                            owner = statistics.median(data["summary"][f"{lane}-{name}"]["instructions:u"][arm])
                            if lane == "readall":
                                owner -= statistics.median(data["summary"][f"open-{name}"]["instructions:u"][arm])
                            return owner / count
                        cells.append(f"{per('base'):,.0f} → {per('cand'):,.0f}")
                    else:
                        cells.append("")
                print(f"| {version} {kind} | {count:,} | {cells[0]} | {cells[1]} | {cells[2]} | {cells[3]} |")
    print()


def callgrind(data):
    print("## Callgrind: one timed owner, by scale\n")
    print("| Arm | Mode | Input | Streams | Owner Ir | Ir/stream | A5 (`validate_stream_allocations`) Ir | "
          "`collect_exact` Ir | `open_stream` Ir | `find_entry` Ir | `memset` Ir |")
    print("|---|---|---|---:|---:|---:|---:|---:|---:|---:|---:|")
    for row in data["rows"]:
        print(f"| {row['arm']} | {row['mode']} | {row['input']} | {row['streams']:,} | {row['owner_ir']:,} | "
              f"{row['owner_ir_per_stream']:,.0f} | {row['validate_stream_allocations_ir']:,} | "
              f"{row['collect_exact_ir']:,} | {row['open_stream_ir']:,} | {row['find_entry_ir']:,} | "
              f"{row['memset_ir']:,} |")
    print("\n| Arm | Mode | Input | Streams ×| Owner Ir × | A5 Ir × |")
    print("|---|---|---|---:|---:|---:|")
    for ratio in data["ratios"]:
        a5 = ratio["validate_stream_allocations_ir_ratio"]
        print(f"| {ratio['arm']} | {ratio['mode']} | {ratio['input']} | {ratio['streams_ratio']:.2f} | "
              f"{ratio['owner_ir_ratio']:.3f} | {a5 if a5 is not None else '—'} |")
    print()


def allocation(data):
    print("## Allocation per owner (first of five identical owners)\n")
    print("| Case | Bytes base → cand | Calls base → cand | Peak live base → cand | Retained base → cand | "
          "Owners identical |")
    print("|---|---:|---:|---:|---:|---|")
    for name, arms in data.items():
        base, cand = arms["base"]["first"], arms["cand"]["first"]
        identical = arms["base"]["owners_identical"] and arms["cand"]["owners_identical"]
        print(f"| `{name}` | {base['allocated_bytes']:,} → {cand['allocated_bytes']:,} | "
              f"{base['allocation_calls']:,} → {cand['allocation_calls']:,} | "
              f"{base['peak_live_bytes']:,} → {cand['peak_live_bytes']:,} | "
              f"{base['retained_bytes']:,} → {cand['retained_bytes']:,} | {identical} |")
    print()


def main():
    packet = sys.argv[1]
    analysis = load(f"{packet}/matrix/analysis.json")
    data = load(f"{packet}/counters/counters.json")
    print("# Change 0767 summary\n")
    print("Generated by `scripts/summarize.py` from the packet's lane outputs.\n")
    if analysis:
        timing(analysis)
    if analysis or data:
        scaling(analysis, data)
    if data:
        counters(data)
    scaling_data = load(f"{packet}/scaling/callgrind-scaling.json")
    if scaling_data:
        callgrind(scaling_data)
    allocation_data = load(f"{packet}/allocation/allocation.json")
    if allocation_data:
        allocation(allocation_data)
    verdicts = load(f"{packet}/verdicts/verdicts-summary.json")
    if verdicts:
        print("## Cross-build verdicts\n")
        print("```json")
        print(json.dumps(verdicts, indent=1))
        print("```")


if __name__ == "__main__":
    main()
