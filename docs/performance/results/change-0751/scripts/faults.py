#!/usr/bin/env python3
"""Regroup lifecycle samples by their minor-fault count (change 0742; reused unchanged by change 0751).

The harness's lifecycle runners report, per measured sample, the operation's
procfs minor-fault delta aligned with the phase vectors (both are ordered by
elapsed time, then sample index). One freshly mapped 33.6 MB buffer costs
about 8,204 first-touch faults on this corpus, so ``round(faults / 8204)``
counts the fresh large mappings a sample touched. This script groups every
sample of the owned media-rich lifecycle (and of the source-backed control)
by that count and reports the median phase times in each group.

Usage: faults.py RAW_DIR [--json OUT]
"""

from __future__ import annotations

import argparse
import collections
import gzip
import json
import statistics
import sys
from pathlib import Path

REGION_FAULTS = 8204
CASES = {
    "pptx_cross_copy_media_rich_lifecycle": (
        "pptx_cross_copy",
        ("lifecycle_ns", "plan_ns", "commit_ns", "publication_ns"),
    ),
    "pptx_source_backed_cross_copy_media_rich_lifecycle": (
        "pptx_source_backed_cross_copy_lifecycle",
        ("lifecycle_ns", "plan_ns", "publication_ns"),
    ),
}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("raw", type=Path)
    parser.add_argument("--json", type=Path)
    args = parser.parse_args()
    output: dict = {}
    for case, (summary_key, phases) in CASES.items():
        groups: dict = collections.defaultdict(lambda: collections.defaultdict(list))
        processes: dict = collections.defaultdict(collections.Counter)
        for path in sorted((args.raw / "native" / case).glob("*.json*")):
            arm = "after" if "-after.json" in path.name else "before"
            raw = path.read_bytes()
            text = gzip.decompress(raw).decode("utf-8") if path.suffix == ".gz" else raw.decode("utf-8")
            (result,) = json.loads(text)["results"]
            metrics = result["operation_metrics"]
            if metrics["alignment"] != "elapsed_ns.samples_by_elapsed_then_sample_index":
                raise SystemExit(f"{path}: unexpected sample alignment")
            summary = result["source"][summary_key]
            faults = metrics["process"]["minor_faults"]["values"]
            for index, value in enumerate(faults):
                regions = round(value / REGION_FAULTS)
                groups[arm][regions].append(
                    {phase: summary[phase][index] / 1e6 for phase in phases}
                )
            processes[arm][round(statistics.median(faults) / REGION_FAULTS)] += 1
        case_output = {}
        for arm in ("before", "after"):
            case_output[arm] = {
                "processes_by_median_fresh_regions": dict(sorted(processes[arm].items())),
                "samples_by_fresh_regions": {
                    str(regions): {
                        "samples": len(rows),
                        **{
                            f"median_{phase[:-3]}_ms": round(
                                statistics.median(row[phase] for row in rows), 3
                            )
                            for phase in phases
                        },
                    }
                    for regions, rows in sorted(groups[arm].items())
                },
            }
        output[case] = case_output
    text = json.dumps(output, indent=2, sort_keys=True)
    if args.json:
        args.json.write_text(text + "\n", encoding="utf-8")
    print(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
