#!/usr/bin/env python3
"""ABBA before/after runs for record 0764.

A is the before binary (base 1d1044e3ac plus the harness commits), B the
after binary. Each round runs four processes in the order A, B, B, A, each
pinned to one core under `perf stat -e instructions:u,cycles:u`. Every process
writes its own JSON report, kept next to the summary.

A report holds one or more series: the xml_attribute_bounds binary reports
one; litchi-perf-baseline reports one per corpus shape. For each series the
summary gives per-process p50/p95/mean, the median of process p50s per arm,
the paired after/before ratios of process p50s (A1-B1 and A2-B2 of every
round), and a percentile bootstrap interval (10,000 resamples, fixed seed) of
the median paired ratio. Instruction and cycle counts are per process and
cover the whole process, input construction included.

Usage:
  abba.py --kind bounds|harness --a BIN --b BIN --case NAME
          [--rounds 4] [--samples 15] [--warmup 3] [--core 16] --out DIR
"""

import argparse
import json
import os
import random
import statistics
import subprocess
from pathlib import Path


def perf_counters(path: Path) -> dict:
    counters = {}
    for line in path.read_text().splitlines():
        fields = line.split(",")
        if len(fields) < 3 or not fields[0].strip().isdigit():
            continue
        counters[fields[2].split(":")[0]] = int(fields[0])
    return counters


def series_from_report(kind: str, report: dict) -> dict:
    """Series name -> retained samples in nanoseconds."""
    if kind == "bounds":
        return {"default": [int(value) for value in report["elapsed_ns"]]}
    series = {}
    for result in report["results"]:
        corpus = result.get("corpus", {})
        name = corpus.get("shape") or corpus.get("name") or result["case"]
        series[name] = [int(value) for value in result["elapsed_ns"]["samples"]]
    return series


def percentile(values: list, fraction: float) -> float:
    ordered = sorted(values)
    index = min(len(ordered) - 1, max(0, round(fraction * (len(ordered) - 1))))
    return ordered[index]


def summarize(name: str, processes: list, rounds: int) -> dict:
    def arm(letter):
        return [process for process in processes if process["arm"] == letter]

    pairs = []
    for round_index in range(1, rounds + 1):
        legs = {process["leg"]: process for process in processes if process["round"] == round_index}
        pairs.append(legs["B1"]["p50_ns"] / legs["A1"]["p50_ns"])
        pairs.append(legs["B2"]["p50_ns"] / legs["A2"]["p50_ns"])
    generator = random.Random(764)
    medians = [
        statistics.median(generator.choice(pairs) for _ in pairs) for _ in range(10_000)
    ]
    return {
        "series": name,
        "processes": processes,
        "median_p50_ns": {
            "before": statistics.median(process["p50_ns"] for process in arm("A")),
            "after": statistics.median(process["p50_ns"] for process in arm("B")),
        },
        "paired_ratios_after_over_before": pairs,
        "paired_ratio_median": statistics.median(pairs),
        "paired_ratio_bootstrap_95": [percentile(medians, 0.025), percentile(medians, 0.975)],
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--kind", choices=["bounds", "harness"], required=True)
    parser.add_argument("--a", type=Path, required=True)
    parser.add_argument("--b", type=Path, required=True)
    parser.add_argument("--case", required=True)
    parser.add_argument("--rounds", type=int, default=4)
    parser.add_argument("--samples", type=int, default=15)
    parser.add_argument("--warmup", type=int, default=3)
    parser.add_argument("--core", type=int, default=16)
    parser.add_argument("--out", type=Path, required=True)
    args = parser.parse_args()

    args.out.mkdir(parents=True, exist_ok=True)
    per_series = {}
    counters_by_process = []
    for round_index in range(1, args.rounds + 1):
        for leg, binary in (("A1", args.a), ("B1", args.b), ("B2", args.b), ("A2", args.a)):
            stem = f"{args.case}-r{round_index}-{leg}"
            report_path = args.out / f"{stem}.json"
            perf_path = args.out / f"{stem}.perf.csv"
            if args.kind == "bounds":
                command = [str(binary), "adversarial", "--case", args.case]
            else:
                command = [str(binary), "--case", args.case]
            command += [
                "--samples", str(args.samples),
                "--warmup", str(args.warmup),
                "--json", str(report_path),
            ]
            full = [
                "taskset", "-c", str(args.core),
                "perf", "stat", "-x", ",", "-e", "instructions:u,cycles:u",
                "-o", str(perf_path), "--",
            ] + command
            subprocess.run(full, check=True, env=dict(os.environ),
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            report = json.loads(report_path.read_text())
            counters = perf_counters(perf_path)
            counters_by_process.append({
                "round": round_index, "leg": leg, "arm": leg[0],
                "instructions": counters.get("instructions"),
                "cycles": counters.get("cycles"),
            })
            for name, samples in series_from_report(args.kind, report).items():
                per_series.setdefault(name, []).append({
                    "round": round_index,
                    "leg": leg,
                    "arm": leg[0],
                    "report": report_path.name,
                    "samples": len(samples),
                    "p50_ns": statistics.median(samples),
                    "p95_ns": percentile(samples, 0.95),
                    "mean_ns": statistics.fmean(samples),
                    "outcomes": report.get("outcomes"),
                })

    def median_counter(letter, key):
        values = [process[key] for process in counters_by_process
                  if process["arm"] == letter and process[key] is not None]
        return statistics.median(values) if values else None

    summary = {
        "case": args.case,
        "kind": args.kind,
        "binaries": {"before": str(args.a), "after": str(args.b)},
        "core": args.core,
        "rounds": args.rounds,
        "samples_per_process": args.samples,
        "warmup_per_process": args.warmup,
        "process_counters": counters_by_process,
        "median_process_instructions": {
            "before": median_counter("A", "instructions"),
            "after": median_counter("B", "instructions"),
        },
        "median_process_cycles": {
            "before": median_counter("A", "cycles"),
            "after": median_counter("B", "cycles"),
        },
        "series": [summarize(name, processes, args.rounds) for name, processes in per_series.items()],
    }
    (args.out / f"{args.case}-summary.json").write_text(json.dumps(summary, indent=1))
    for series in summary["series"]:
        print(json.dumps({
            "case": args.case,
            "series": series["series"],
            "median_p50_ns": series["median_p50_ns"],
            "ratio": round(series["paired_ratio_median"], 4),
            "ci95": [round(value, 4) for value in series["paired_ratio_bootstrap_95"]],
        }))
    print(json.dumps({"case": args.case, "instructions": summary["median_process_instructions"]}))


if __name__ == "__main__":
    main()
