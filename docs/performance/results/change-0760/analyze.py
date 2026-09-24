#!/usr/bin/env python3
"""Summarize the change-0760 ABBA timing runs (adapted from change 0743).

Per process: p50, p95 and mean of the harness's timed samples. Per arm: the
median of the process p50s. Pairs are formed inside each ABBA block (position
0 before with position 1 after, position 3 before with position 2 after), so
every pair ran back to back; the paired ratio is after / before. The bootstrap
resamples paired processes (never samples within a process), 10,000 times with
a fixed seed, and reports the percentile 95% interval of the median ratio.
Every p50, p95 or mean pair ratio above 1.05 is listed as a regression flag.
"""

import argparse
import json
import os
import random
import statistics

PHASES = [
    "opened_presentation_ns",
    "snapshot_edit_ns",
    "set_shape_text_ns",
    "transaction_commit_ns",
    "apply_commit_ns",
    "publication_ns",
    "total_ns",
]


def percentile(values, fraction):
    ordered = sorted(values)
    rank = max(0, min(len(ordered) - 1, int(round(fraction * (len(ordered) - 1)))))
    return ordered[rank]


def bootstrap_median(ratios, seed, resamples=10_000):
    rng = random.Random(seed)
    medians = []
    for _ in range(resamples):
        sample = [rng.choice(ratios) for _ in ratios]
        medians.append(statistics.median(sample))
    medians.sort()
    return medians[int(0.025 * resamples)], medians[int(0.975 * resamples) - 1]


def load(run_dir):
    log = []
    for name in sorted(os.listdir(run_dir)):
        if name.startswith("run-log") and name.endswith(".json"):
            with open(os.path.join(run_dir, name), encoding="utf-8") as handle:
                log.extend(json.load(handle))
    processes = []
    for record in log:
        with open(os.path.join(run_dir, record["json"]), encoding="utf-8") as handle:
            report = json.load(handle)
        for result in report["results"]:
            samples = result["elapsed_ns"]["samples"]
            entry = {
                "round": record["round"],
                "position": record["position"],
                "arm": record["arm"],
                "case": result["case"],
                "corpus": result["corpus"]["name"],
                "samples": len(samples),
                "p50": result["elapsed_ns"]["p50"],
                "p95": result["elapsed_ns"]["p95"],
                "mean": result["elapsed_ns"]["mean"],
                "output_sha256": result.get("output_sha256"),
                "load": record["load_average_before"],
            }
            phases = (result.get("source") or {}).get("pptx_opened_transaction_phases")
            if phases:
                entry["phases"] = {
                    name: statistics.median(phases[name]) for name in PHASES if phases.get(name)
                }
            processes.append(entry)
    return processes


def pairs_for(entries):
    by_slot = {(e["round"], e["position"]): e for e in entries}
    pairs = []
    for round_index in sorted({e["round"] for e in entries}):
        for before_position, after_position in ((0, 1), (3, 2)):
            before = by_slot.get((round_index, before_position))
            after = by_slot.get((round_index, after_position))
            if before and after:
                pairs.append((before, after))
    return pairs


def summarize(processes, seed):
    keys = sorted({(p["case"], p["corpus"]) for p in processes})
    summary = []
    flags = []
    for case, corpus in keys:
        entries = [p for p in processes if p["case"] == case and p["corpus"] == corpus]
        arms = {arm: [e for e in entries if e["arm"] == arm] for arm in ("before", "after")}
        pairs = pairs_for(entries)
        row = {
            "case": case,
            "corpus": corpus,
            "processes": {arm: len(arms[arm]) for arm in arms},
            "samples_per_process": sorted({e["samples"] for e in entries}),
            "per_process": {
                arm: [
                    {k: e[k] for k in ("round", "position", "p50", "p95", "mean", "output_sha256")}
                    for e in arms[arm]
                ]
                for arm in arms
            },
            "median_p50_ns": {arm: statistics.median(e["p50"] for e in arms[arm]) for arm in arms},
            "median_p95_ns": {arm: statistics.median(e["p95"] for e in arms[arm]) for arm in arms},
            "median_mean_ns": {arm: statistics.median(e["mean"] for e in arms[arm]) for arm in arms},
        }
        for metric in ("p50", "p95", "mean"):
            ratios = [after[metric] / before[metric] for before, after in pairs]
            low, high = bootstrap_median(ratios, seed)
            row[f"paired_{metric}_ratio"] = {
                "median": statistics.median(ratios),
                "min": min(ratios),
                "max": max(ratios),
                "bootstrap95": [low, high],
                "pairs": len(ratios),
            }
            for (before, after), ratio in zip(pairs, ratios):
                if ratio > 1.05:
                    flags.append(
                        {
                            "case": case,
                            "corpus": corpus,
                            "metric": metric,
                            "round": before["round"],
                            "before_position": before["position"],
                            "after_position": after["position"],
                            "ratio": ratio,
                        }
                    )
        if any("phases" in e for e in entries):
            row["phases_median_ns"] = {
                arm: {
                    name: statistics.median(e["phases"][name] for e in arms[arm])
                    for name in PHASES
                }
                for arm in arms
            }
        summary.append(row)
    return summary, flags


def fmt_ms(ns):
    return f"{ns / 1e6:,.3f}"


def pct(ratio):
    return f"{(ratio - 1) * 100:+.2f}%"


def tables(summary, flags):
    lines = [
        "| case | corpus | before p50 ms | after p50 ms | paired p50 change | 95% CI | paired mean change | before p95 ms | after p95 ms |",
        "| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |",
    ]
    for row in summary:
        if row["case"] == "pptx_semantic_opened_transaction_phases":
            continue
        p50 = row["paired_p50_ratio"]
        mean = row["paired_mean_ratio"]
        lines.append(
            f"| `{row['case']}` | {row['corpus']} | {fmt_ms(row['median_p50_ns']['before'])} | "
            f"{fmt_ms(row['median_p50_ns']['after'])} | {pct(p50['median'])} | "
            f"[{pct(p50['bootstrap95'][0])}, {pct(p50['bootstrap95'][1])}] | {pct(mean['median'])} | "
            f"{fmt_ms(row['median_p95_ns']['before'])} | {fmt_ms(row['median_p95_ns']['after'])} |"
        )
    for row in summary:
        if "phases_median_ns" not in row:
            continue
        lines += [
            "",
            f"Phases, `{row['case']}` on {row['corpus']} (median of per-process medians, ms):",
            "",
            "| phase | before | after | change |",
            "| --- | ---: | ---: | ---: |",
        ]
        for name in PHASES:
            before = row["phases_median_ns"]["before"][name]
            after = row["phases_median_ns"]["after"][name]
            change = pct(after / before) if before else "n/a"
            lines.append(f"| {name} | {fmt_ms(before)} | {fmt_ms(after)} | {change} |")
    lines += ["", f"Regression flags (pair ratio > 1.05): {len(flags)}", ""]
    for flag in flags:
        lines.append(
            f"- `{flag['case']}` {flag['corpus']} {flag['metric']} round {flag['round']} "
            f"(positions {flag['before_position']}/{flag['after_position']}): {pct(flag['ratio'])}"
        )
    return "\n".join(lines) + "\n"


def perf_counters(run_dir):
    """Whole-process user instructions and cycles per process, by group."""
    log = []
    for name in sorted(os.listdir(run_dir)):
        if name.startswith("run-log") and name.endswith(".json"):
            with open(os.path.join(run_dir, name), encoding="utf-8") as handle:
                log.extend(json.load(handle))
    rows = []
    for record in log:
        path = os.path.join(run_dir, record["json"] + ".perf.csv")
        if not os.path.exists(path):
            continue
        values = {}
        with open(path, encoding="utf-8") as handle:
            for line in handle:
                fields = line.strip().split(",")
                if len(fields) > 2 and fields[0].replace(".", "").isdigit():
                    values[fields[2]] = float(fields[0])
        rows.append({
            "round": record["round"], "group": record["group"], "position": record["position"],
            "arm": record["arm"], "instructions": values.get("instructions:u"),
            "cycles": values.get("cycles:u"),
        })
    summary = []
    for group in sorted({row["group"] for row in rows}):
        entries = [row for row in rows if row["group"] == group]
        by_slot = {(row["round"], row["position"]): row for row in entries}
        pairs = []
        for round_index in sorted({row["round"] for row in entries}):
            for before_position, after_position in ((0, 1), (3, 2)):
                before = by_slot.get((round_index, before_position))
                after = by_slot.get((round_index, after_position))
                if before and after:
                    pairs.append((before, after))
        item = {"group": group}
        for metric in ("instructions", "cycles"):
            item[metric] = {
                arm: statistics.median(row[metric] for row in entries if row["arm"] == arm)
                for arm in ("before", "after")
            }
            ratios = [after[metric] / before[metric] for before, after in pairs]
            item[metric + "_paired_ratio"] = {
                "median": statistics.median(ratios), "min": min(ratios), "max": max(ratios),
                "pairs": len(ratios),
            }
        summary.append(item)
    return rows, summary


def perf_table(summary):
    lines = [
        "",
        "Whole-process user instructions and cycles (perf stat; includes corpus construction and verification):",
        "",
        "| group | before instructions | after instructions | paired change | before cycles | after cycles | paired change |",
        "| --- | ---: | ---: | ---: | ---: | ---: | ---: |",
    ]
    for item in summary:
        lines.append(
            f"| {item['group']} | {item['instructions']['before'] / 1e9:,.3f} G | "
            f"{item['instructions']['after'] / 1e9:,.3f} G | {pct(item['instructions_paired_ratio']['median'])} | "
            f"{item['cycles']['before'] / 1e9:,.3f} G | {item['cycles']['after'] / 1e9:,.3f} G | "
            f"{pct(item['cycles_paired_ratio']['median'])} |"
        )
    return "\n".join(lines) + "\n"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--runs", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--seed", type=int, default=760)
    args = parser.parse_args()
    processes = load(args.runs)
    summary, flags = summarize(processes, args.seed)
    perf_rows, perf_summary = perf_counters(args.runs)
    with open(os.path.join(args.out, "analysis.json"), "w", encoding="utf-8") as handle:
        json.dump({"seed": args.seed, "summary": summary, "flags": flags,
                   "perf_processes": perf_rows, "perf_summary": perf_summary}, handle, indent=1)
    rendered = tables(summary, flags) + perf_table(perf_summary)
    with open(os.path.join(args.out, "tables.md"), "w", encoding="utf-8") as handle:
        handle.write(rendered)
    print(rendered)


if __name__ == "__main__":
    main()
