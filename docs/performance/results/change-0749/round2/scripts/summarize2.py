#!/usr/bin/env python3
"""Render summary-r2.md for the change 0749 review round: every timing,
counter, scaling, corpus and allocation table, and every flag.

Usage: summarize2.py PACKET > PACKET/summary-r2.md
"""

import json
import os
import statistics
import sys


def load(packet, *parts):
    with open(os.path.join(packet, *parts)) as handle:
        return json.load(handle)


def us(value):
    return f"{value / 1000:,.1f}"


def change(comparison):
    if not comparison:
        return "—"
    median = comparison["median_paired_change_pct"]["p50"]
    low, high = comparison["bootstrap95_median_paired_change_pct"]["p50"]
    return f"{median:+.2f}% [{low:+.2f}, {high:+.2f}]"


def main():
    packet = sys.argv[1]
    out = ["# 0749 review round: summary tables\n"]
    analysis = load(packet, "round2", "matrix", "analysis.json")
    out.append(
        "Arms: A = base `ab29ac6291`; B = the first 0749 version `776ad05175` (quadratic mini-stream "
        "readback); C = the fix `f2a57ac936`. Timings are the median of per-process p50s (µs) and the median "
        "per-round paired change of p50 with its 95% percentile-bootstrap interval (10,000 resamples, seed 749); "
        "12 rounds, each of the six A/B/C orders twice, argv[0] +8 bytes per round, core 20. Harness selectors "
        "run A and C only. Every probe output digest is identical across all processes of all arms: "
        f"`{analysis['all_probe_outputs_identical_across_arms']}`.\n")
    out.append("## Timing\n")
    out.append("| Case | A p50 | B p50 | C p50 | C vs A | B vs A | C vs B | C vs A mean | C vs A p95 |")
    out.append("|---|---:|---:|---:|---:|---:|---:|---:|---:|")
    for name, case in analysis["cases"].items():
        arms = case["arms"]
        cells = [us(arms[arm]["median_process_p50_ns"]) if arm in arms else "—" for arm in "ABC"]
        comparisons = case["comparisons"]
        ca = comparisons.get("C_vs_A")
        mean = f"{ca['median_paired_change_pct']['mean']:+.2f}%" if ca else "—"
        p95 = f"{ca['median_paired_change_pct']['p95']:+.2f}%" if ca else "—"
        out.append(f"| `{name}` | {' | '.join(cells)} | {change(ca)} | {change(comparisons.get('B_vs_A'))} | "
                   f"{change(comparisons.get('C_vs_B'))} | {mean} | {p95} |")
    out.append("\n## Every per-round C-vs-A change above +5%\n")
    out.append("| Case | Round | Metric | A µs | C µs | Change |")
    out.append("|---|---:|---|---:|---:|---:|")
    flags = analysis["flags_c_vs_a_above_plus_5_pct"]
    for flag in sorted(flags, key=lambda f: (f["case"], f["round"], f["metric"])):
        out.append(f"| `{flag['case']}` | {flag['round']} | {flag['metric']} | {us(flag['base_ns'])} | "
                   f"{us(flag['cand_ns'])} | {flag['change_pct']:+.2f}% |")
    out.append(f"\n{len(flags)} flags.\n")

    for title, file in (("Per-owner counters, probe cases (median of three argv[0] layouts)", "counters-probe.json"),
                        ("Per-owner counters, harness selectors", "counters-harness.json"),
                        ("`doc_semantic_one_edit_save/large` under fixed glibc malloc thresholds", "counters-glibc-tuned.json")):
        data = load(packet, "round2", "counters", file)
        out.append(f"## {title}\n")
        out.append("| Case | instructions:u A | B | C | C vs A | C vs B | cycles A | B | C | C vs A | page faults A / C (per layout) |")
        out.append("|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---|")
        for name, events in data["summary"].items():
            def median(event, arm):
                values = events[event].get(arm)
                return f"{statistics.median(values):,.0f}" if values else "—"
            iu, cy = events["instructions:u"], events["cycles"]
            ca_i = iu.get("median_change_C_vs_A_pct")
            cb_i = iu.get("median_change_C_vs_B_pct")
            ca_c = cy.get("median_change_C_vs_A_pct")
            faults = f"{events['page-faults']['A']} / {events['page-faults']['C']}"
            out.append(f"| `{name}` | {median('instructions:u', 'A')} | {median('instructions:u', 'B')} | "
                       f"{median('instructions:u', 'C')} | {ca_i:+.2f}% | "
                       f"{(f'{cb_i:+.2f}%' if cb_i is not None else '—')} | {median('cycles', 'A')} | "
                       f"{median('cycles', 'B')} | {median('cycles', 'C')} | {ca_c:+.2f}% | {faults} |")
        out.append("")

    scaling = load(packet, "round2", "scaling", "callgrind-scaling.json")
    out.append("## Callgrind instruction scaling per write (generated v3 files, 2,000-byte mini streams)\n")
    out.append("| Arm | Streams | `write_to` | `validate` | reparse (`open_with_limits`) | of which A5 (`validate_stream_allocations`) | readback |")
    out.append("|---|---:|---:|---:|---:|---:|---:|")
    for row in scaling["rows"]:
        values = row["inclusive_ir_per_write"]
        readback = (values["readback_open_stream"] or values["readback_stream_equals_b"]
                    or values["readback_stream_equals_c"])
        streams = row["input"].split("-")[1].split("x")[0]
        out.append(f"| {row['arm']} | {int(streams):,} | {values['write_to']:,} | {values['validate']:,} | "
                   f"{values['reparse']:,} | {values['reparse_stream_allocations']:,} | {readback:,} |")
    out.append("")

    corpus = load(packet, "round2", "corpus", "corpus-instructions.json")
    summary = corpus["summary"]
    out.append("## Per-owner instructions over the OLE2 corpus (both `sector_layout_corpus` edits)\n")
    out.append(f"{summary['fixtures_examined']} fixtures, {summary['pairs_measured']} fixture/edit pairs, "
               f"outputs identical across arms: `{summary['all_outputs_identical_across_arms']}`.\n")
    out.append("| Group | Pairs | median B vs A | max B vs A | B > +5% | median C vs A | max C vs A | C > +1% | C > +5% |")
    out.append("|---|---:|---:|---:|---:|---:|---:|---:|---:|")
    for label, key in (("All reused pairs", "reused_pairs"), ("All-mini-stream fixtures", "reused_all_mini_stream_pairs")):
        group = summary[key]
        out.append(f"| {label} | {group['pairs']} | {group['median_B_vs_A_pct']:+.2f}% | {group['max_B_vs_A_pct']:+.2f}% | "
                   f"{group['B_above_plus_5_pct']} | {group['median_C_vs_A_pct']:+.2f}% | {group['max_C_vs_A_pct']:+.2f}% | "
                   f"{group['C_above_plus_1_pct']} | {group['C_above_plus_5_pct']} |")
    out.append("\nPairs with C above +1% against A:\n")
    for pair in summary["c_above_plus_1_pct"]:
        out.append(f"- `{pair['fixture']}` ({pair['edit']}): C vs A {pair['C_vs_A_pct']:+.2f}%, "
                   f"B vs A {pair['B_vs_A_pct']:+.2f}%, C vs B {pair['C_vs_B_pct']:+.2f}%.")
    out.append("")

    allocation = load(packet, "round2", "allocation", "allocation.json")["summary"]
    out.append("## Allocation per owner (counting allocator)\n")
    out.append("| Case | Bytes A | Bytes B | Bytes C | Calls A | Calls B | Calls C | Peak live A | Peak live C |")
    out.append("|---|---:|---:|---:|---:|---:|---:|---:|---:|")
    for name, arms in allocation.items():
        value = lambda arm, field: arms[arm]["per_owner"][field]
        out.append(f"| `{name}` | {value('A', 'allocated_bytes'):,} | {value('B', 'allocated_bytes'):,} | "
                   f"{value('C', 'allocated_bytes'):,} | {value('A', 'allocation_calls'):,} | "
                   f"{value('B', 'allocation_calls'):,} | {value('C', 'allocation_calls'):,} | "
                   f"{value('A', 'peak_live_bytes'):,} | {value('C', 'peak_live_bytes'):,} |")
    identical = all(arms[arm]["owners_identical"] for arms in allocation.values() for arm in "ABC")
    same = all(len({arms[arm]["output_sha256"] for arm in "ABC"}) == 1 for arms in allocation.values())
    out.append(f"\nEvery owner identical within its process: `{identical}`. Outputs identical across arms: `{same}`.\n")
    print("\n".join(out))


if __name__ == "__main__":
    main()
