#!/usr/bin/env python3
"""Summarize the change 0664 baseline runs.

Reads every `<group>-<repeat>.json` report in a directory and prints, per
selector, p50/mean/p95/p99 for each repeat, the median of the three repeat p50s,
and the widest p50 spread across the repeats -- the in-window A/A floor this
program reports when a dedicated A/A pair is not available for a selector.

Also prints the marker/control and marker/generated ratios that are the point of
the pair, and the allocation counters from the allocator runs.

Usage: summarize.py <dir> [--json OUT]
"""

from __future__ import annotations

import argparse
import json
import pathlib
import statistics
import sys


def load(directory: pathlib.Path) -> dict:
    timing: dict[str, dict[str, dict]] = {}
    alloc: dict[str, dict] = {}
    for path in sorted(directory.glob("*.json")):
        name = path.stem
        try:
            report = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            continue
        results = report.get("results")
        if not results:
            continue
        if name.startswith("alloc-"):
            for row in results:
                metrics = (row.get("operation_metrics") or {}).get("allocation")
                if metrics:
                    alloc[row["case"]] = metrics
            continue
        if name.startswith("census-"):
            # The one-sample census run is not a timing sample.
            continue
        repeat = name.rsplit("-", 1)[-1]
        for row in results:
            timing.setdefault(row["case"], {})[repeat] = row
    return {"timing": timing, "alloc": alloc}


def vector(metrics: dict, field: str) -> float | None:
    """Median of a metric vector; allocator counters are usually constant."""
    entry = metrics.get(field)
    if not isinstance(entry, dict):
        return None
    values = entry.get("values")
    if values:
        return statistics.median(values)
    return entry.get("p50", entry.get("mean"))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory")
    parser.add_argument("--json")
    arguments = parser.parse_args()
    directory = pathlib.Path(arguments.directory)
    data = load(directory)
    timing, alloc = data["timing"], data["alloc"]

    summary: dict[str, dict] = {}
    summary_alloc: dict[str, dict] = {}
    print(f"{'selector':52} {'p50 median (ms)':>16} {'mean (ms)':>11} "
          f"{'p95 (ms)':>10} {'p99 (ms)':>10} {'p50 spread':>11} {'A/A':>9} {'n':>3}")
    for case in sorted(timing):
        repeats = timing[case]
        ordered = sorted(repeats)
        p50s = [repeats[key]["elapsed_ns"]["p50"] for key in ordered]
        aa = [repeats[key]["elapsed_ns"]["p50"] for key in ordered if key.startswith("A")]
        aa_floor = (max(aa) - min(aa)) / min(aa) if len(aa) == 2 and min(aa) else None
        means = [row["elapsed_ns"]["mean"] for row in repeats.values()]
        p95s = [row["elapsed_ns"]["p95"] for row in repeats.values()]
        p99s = [row["elapsed_ns"]["p99"] for row in repeats.values()]
        median_p50 = statistics.median(p50s)
        spread = (max(p50s) - min(p50s)) / min(p50s) if min(p50s) else 0.0
        summary[case] = {
            "repeats": {k: v["elapsed_ns"]["p50"] for k, v in repeats.items()},
            "p50_median_ns": median_p50,
            "mean_ns": statistics.median(means),
            "p95_ns": statistics.median(p95s),
            "p99_ns": statistics.median(p99s),
            "p50_spread": spread,
            "aa_floor": aa_floor,
            "samples": len(repeats[next(iter(repeats))]["elapsed_ns"]["samples"]),
        }
        print(f"{case:52} {median_p50/1e6:16.4f} {statistics.median(means)/1e6:11.4f} "
              f"{statistics.median(p95s)/1e6:10.4f} {statistics.median(p99s)/1e6:10.4f} "
              f"{spread*100:10.2f}% "
              f"{('-' if aa_floor is None else format(aa_floor*100, '.2f') + '%'):>9} "
              f"{len(repeats):3}")

    print("\nratios (median p50):")
    ratios = {}
    for marker, control, label in [
        ("pptx_marker_ordinary_save_edit", "pptx_marker_control_ordinary_save_edit", "pptx edit marker/control"),
        ("pptx_marker_ordinary_save_lifecycle", "pptx_marker_control_ordinary_save_lifecycle", "pptx lifecycle marker/control"),
        ("pptx_marker_ordinary_save_atomic_publish", "pptx_marker_control_ordinary_save_atomic_publish", "pptx atomic marker/control"),
        ("pptx_marker_ordinary_save_counting_publish", "pptx_marker_control_ordinary_save_counting_publish", "pptx counting marker/control"),
        ("docx_marker_ordinary_save_edit", "docx_marker_control_ordinary_save_edit", "docx edit marker/control"),
        ("docx_marker_ordinary_save_lifecycle", "docx_marker_control_ordinary_save_lifecycle", "docx lifecycle marker/control"),
        ("pptx_marker_eager_full_text", "pptx_marker_control_eager_full_text", "pptx eager text marker/control"),
        ("pptx_marker_source_full_text", "pptx_marker_control_source_full_text", "pptx source text marker/control"),
        ("docx_marker_eager_full_text", "docx_marker_control_eager_full_text", "docx eager text marker/control"),
        ("docx_marker_source_full_text", "docx_marker_control_source_full_text", "docx source text marker/control"),
        ("pptx_marker_ordinary_save_edit", "pptx_ordinary_save_edit", "pptx edit marker/generated"),
        ("pptx_marker_ordinary_save_lifecycle", "pptx_ordinary_save_lifecycle", "pptx lifecycle marker/generated"),
        ("docx_marker_ordinary_save_edit", "docx_ordinary_save_edit", "docx edit marker/generated"),
        ("pptx_marker_control_ordinary_save_edit", "pptx_ordinary_save_edit", "pptx edit control/generated"),
    ]:
        if marker in summary and control in summary and summary[control]["p50_median_ns"]:
            value = summary[marker]["p50_median_ns"] / summary[control]["p50_median_ns"]
            ratios[label] = value
            print(f"  {label:36} {value:8.2f}x")

    if alloc:
        print("\nallocations (median of the retained samples):")
        print(f"{'selector':52} {'calls':>12} {'alloc bytes':>16} {'reallocs':>10} "
              f"{'retained bytes':>15} {'region peak':>14} {'constant':>9}")
        for case in sorted(alloc):
            metrics = alloc[case]
            before = vector(metrics, "live_bytes_before")
            after = vector(metrics, "live_bytes_after")
            retained = (after - before) if (before is not None and after is not None) else None
            # Only the work counters are expected to be constant. The live and
            # peak byte counters are absolute process state and move with what
            # the process already had allocated when the region opened.
            constant = all(
                len(set(metrics[field]["values"])) == 1
                for field in (
                    "allocation_calls",
                    "deallocation_calls",
                    "reallocation_calls",
                    "failed_allocation_calls",
                    "allocated_bytes",
                    "deallocated_bytes",
                )
                if isinstance(metrics.get(field), dict) and metrics[field].get("values")
            )
            print(f"{case:52} "
                  f"{str(vector(metrics,'allocation_calls')):>12} "
                  f"{str(vector(metrics,'allocated_bytes')):>16} "
                  f"{str(vector(metrics,'reallocation_calls')):>10} "
                  f"{str(retained):>15} "
                  f"{str(vector(metrics,'region_peak_live_bytes')):>14} "
                  f"{str(constant):>9}")
            summary_alloc[case] = {
                "allocation_calls": vector(metrics, "allocation_calls"),
                "allocated_bytes": vector(metrics, "allocated_bytes"),
                "deallocated_bytes": vector(metrics, "deallocated_bytes"),
                "reallocation_calls": vector(metrics, "reallocation_calls"),
                "retained_live_bytes": retained,
                "region_peak_live_bytes": vector(metrics, "region_peak_live_bytes"),
                "constant_across_samples": constant,
            }

    if arguments.json:
        pathlib.Path(arguments.json).write_text(
            json.dumps({"timing": summary, "ratios": ratios,
                        "allocation": summary_alloc}, indent=2, sort_keys=True) + "\n",
            encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
