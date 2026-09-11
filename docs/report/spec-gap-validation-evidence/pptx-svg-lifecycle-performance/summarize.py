#!/usr/bin/env python3
"""Recompute the SVG lifecycle profile table from raw process receipts."""

from __future__ import annotations

import argparse
import json
import math
from pathlib import Path


LANES = (
    "capture_raster_small",
    "capture_raster_large",
    "capture_attached_small",
    "capture_attached_large",
    "capture_namespace_heavy",
    "inventory_many_raster_256",
    "inventory_many_raster_1024",
    "inventory_distinct_local_namespace_256",
    "inventory_distinct_local_namespace_1024",
    "attach_end_to_end_small",
    "attach_end_to_end_large",
    "detach_end_to_end_small",
    "detach_end_to_end_large",
    "noop_detach_end_to_end_small",
    "noop_detach_end_to_end_large",
    "clone_raster_small",
    "clone_raster_large",
    "clone_attached_small",
    "clone_attached_large",
    "limit_small",
    "limit_large",
    "malformed_small",
    "malformed_large",
    "namespace_limit_refusal",
)


def quantile(values: list[int], percent: int) -> int:
    return sorted(values)[max(0, math.ceil(len(values) * percent / 100) - 1)]


def rss(path: Path) -> int:
    marker = "Maximum resident set size (kbytes):"
    for line in path.read_text().splitlines():
        if line.lstrip().startswith(marker):
            return int(line.split(":", 1)[1].strip())
    raise ValueError(f"RSS missing from {path}")


def rows(results: Path) -> list[dict[str, object]]:
    output = []
    for lane in LANES:
        paths = sorted(results.glob(f"{lane}-p*.json"))
        samples = []
        rss_values = []
        input_bytes = set()
        for path in paths:
            payload = json.loads(path.read_text())
            input_bytes.add(int(payload["input_bytes"]))
            rss_values.append(rss(path.with_suffix(".time.txt")))
            samples.extend(payload["samples"])
        elapsed = [int(sample["elapsed_ns"]) for sample in samples]
        allocated = [int(sample["requested_alloc_bytes"]) for sample in samples]
        peak = [int(sample["peak_live_delta"]) for sample in samples]
        output.append(
            {
                "lane": lane,
                "processes": len(paths),
                "samples": len(samples),
                "input_bytes": next(iter(input_bytes)),
                "elapsed": (quantile(elapsed, 50), quantile(elapsed, 95), quantile(elapsed, 99)),
                "alloc": (quantile(allocated, 50), quantile(allocated, 95)),
                "peak": (quantile(peak, 50), quantile(peak, 95)),
                "rss": (min(rss_values), max(rss_values)),
            }
        )
    return output


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--results", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    table = rows(args.results)
    lines = [
        "# Source-backed PPTX SVG lifecycle profile",
        "",
        "This is an absolute profile of synthetic bounded inputs. It is not a",
        "before/after speedup claim and makes no native Office or rendering claim.",
        "The timer includes the operation named by each lane; end-to-end lanes",
        "include capture, commit, publication, and reopen semantic checks.",
        "",
        "| lane | processes | samples | input bytes | elapsed ns p50/p95/p99 | requested alloc bytes p50/p95 | peak live delta p50/p95 | RSS KiB min-max |",
        "|---|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for row in table:
        lines.append(
            f"| {row['lane']} | {row['processes']} | {row['samples']} | {row['input_bytes']} | "
            f"{row['elapsed'][0]}/{row['elapsed'][1]}/{row['elapsed'][2]} | "
            f"{row['alloc'][0]}/{row['alloc'][1]} | "
            f"{row['peak'][0]}/{row['peak'][1]} | "
            f"{row['rss'][0]}–{row['rss'][1]} |"
        )
    args.output.write_text("\n".join(lines) + "\n")


if __name__ == "__main__":
    main()
