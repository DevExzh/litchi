#!/usr/bin/env python3
"""Run one A1 B1 B2 A2 (plus A/A) timing sequence for two harness binaries.

Each leg is one pinned harness invocation with identical flags. The legs are
ordered before, after, after, before so that drift during the window shows up
as a difference between the two same-leg runs rather than as a result. The
A/A pair at the end reports this host's floor in the same window.

Usage: paired-timing.py <before-binary> <after-binary> <cpu> <case> <samples>
                        <output-json>
"""

import json
import statistics
import subprocess
import sys


def run(binary, cpu, case, samples):
    process = subprocess.run(
        [
            "taskset", "-c", str(cpu), binary,
            "--case", case, "--warmup", "3", "--samples", str(samples),
        ],
        capture_output=True,
        text=True,
        check=True,
    )
    report = json.loads(process.stdout)
    legs = {}
    for result in report["results"]:
        shape = result["corpus"]["shape"]
        legs[shape] = {
            "samples": result["elapsed_ns"]["samples"],
            "output_sha256": result["output_sha256"],
        }
    return legs


def summarize(samples):
    ordered = sorted(samples)
    return {
        "n": len(samples),
        "p50": statistics.median(ordered),
        "mean": statistics.fmean(ordered),
        "p95": ordered[min(len(ordered) - 1, int(round(0.95 * (len(ordered) - 1))))],
        "p99": ordered[min(len(ordered) - 1, int(round(0.99 * (len(ordered) - 1))))],
        "min": ordered[0],
        "max": ordered[-1],
    }


def main():
    before, after, cpu, case, samples, output = sys.argv[1:7]
    samples = int(samples)
    plan = [
        ("A1", before), ("B1", after), ("B2", after), ("A2", before),
        ("A3", before), ("A4", before),
    ]
    legs = {}
    for label, binary in plan:
        legs[label] = run(binary, cpu, case, samples)
        print(f"{label} done", file=sys.stderr)

    report = {"case": case, "cpu": int(cpu), "samples_per_leg": samples, "shapes": {}}
    shapes = sorted(legs["A1"])
    for shape in shapes:
        entry = {"legs": {}, "digests": {}}
        for label in legs:
            entry["legs"][label] = summarize(legs[label][shape]["samples"])
            entry["digests"][label] = legs[label][shape]["output_sha256"]
        before_p50 = statistics.median(legs["A1"][shape]["samples"] + legs["A2"][shape]["samples"])
        after_p50 = statistics.median(legs["B1"][shape]["samples"] + legs["B2"][shape]["samples"])
        floor_a = statistics.median(legs["A3"][shape]["samples"])
        floor_b = statistics.median(legs["A4"][shape]["samples"])
        entry["pooled"] = {
            "before_p50": before_p50,
            "after_p50": after_p50,
            "after_vs_before_pct": 100.0 * (after_p50 - before_p50) / before_p50,
            "before_vs_after_pct": 100.0 * (before_p50 - after_p50) / after_p50,
            "aa_floor_pct": 100.0 * abs(floor_b - floor_a) / floor_a,
            "aa_p50": [floor_a, floor_b],
        }
        report["shapes"][shape] = entry
    with open(output, "w") as handle:
        json.dump(report, handle, indent=2)
        handle.write("\n")
    print(json.dumps(report["shapes"], indent=2)[:4000])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
