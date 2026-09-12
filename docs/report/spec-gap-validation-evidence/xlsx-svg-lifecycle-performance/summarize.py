#!/usr/bin/env python3
"""Recompute the XLSX profile table from raw process receipts."""

from __future__ import annotations

import argparse
import json
import math
from pathlib import Path


LANES = (
    "capture_native_fixture",
    "capture_raster_two_cell_small",
    "capture_raster_two_cell_large",
    "capture_raster_one_cell_small",
    "capture_raster_one_cell_large",
    "capture_raster_absolute_small",
    "capture_raster_absolute_large",
    "capture_attached_two_cell_small",
    "capture_attached_two_cell_large",
    "capture_attached_one_cell_small",
    "capture_attached_one_cell_large",
    "capture_attached_absolute_small",
    "capture_attached_absolute_large",
    "clone_raster_small",
    "clone_raster_large",
    "clone_attached_small",
    "clone_attached_large",
    "clone_captured_owner_small",
    "clone_captured_owner_large",
    "inventory_shared_256",
    "inventory_shared_1024",
    "inventory_distinct_256",
    "inventory_distinct_1024",
    "namespace_heavy",
    "namespace_limit_refusal",
    "attach_end_to_end_two_cell_small",
    "attach_end_to_end_two_cell_large",
    "attach_end_to_end_one_cell_small",
    "attach_end_to_end_one_cell_large",
    "attach_end_to_end_absolute_small",
    "attach_end_to_end_absolute_large",
    "strict_attach_end_to_end_two_cell_small",
    "inverse_attach_detach_two_cell_small",
    "inverse_attach_detach_two_cell_large",
    "inverse_attach_detach_one_cell_small",
    "inverse_attach_detach_one_cell_large",
    "inverse_attach_detach_absolute_small",
    "inverse_attach_detach_absolute_large",
    "detach_end_to_end_shared_first_two_cell",
    "detach_end_to_end_shared_first_one_cell",
    "detach_end_to_end_shared_first_absolute",
    "detach_end_to_end_shared_final_two_cell",
    "detach_end_to_end_shared_final_one_cell",
    "detach_end_to_end_shared_final_absolute",
    "strict_detach_end_to_end_shared_final_two_cell",
    "incoming_edge_shared_final_two_cell",
    "detach_end_to_end_distinct_two_cell_small",
    "detach_end_to_end_distinct_two_cell_large",
    "detach_end_to_end_distinct_one_cell_small",
    "detach_end_to_end_distinct_one_cell_large",
    "detach_end_to_end_distinct_absolute_small",
    "detach_end_to_end_distinct_absolute_large",
    "same_picture_attach_detach_two_cell",
    "same_picture_attach_detach_one_cell",
    "same_picture_attach_detach_absolute",
    "multisheet_attach_detach",
    "noop_detach_two_cell",
    "noop_detach_one_cell",
    "noop_detach_absolute",
    "limit_small",
    "limit_large",
    "mixed_caps_rejection",
    "malformed_duplicate_owner",
    "malformed_mce_owner",
    "malformed_linked_owner",
    "malformed_unknown_uri",
    "multi_picture_same_drawing_16",
    "multi_picture_same_drawing_64",
    "multi_picture_same_drawing_256",
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
        if not paths:
            raise ValueError(f"no receipts found for {lane}")
        samples = []
        rss_values = []
        input_bytes = set()
        for path in paths:
            payload = json.loads(path.read_text())
            input_bytes.add(int(payload["input_bytes"]))
            rss_values.append(rss(path.with_suffix(".time.txt")))
            samples.extend(payload["samples"])
        if len(input_bytes) != 1:
            raise ValueError(f"input size changed across processes for {lane}")
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
        "# XLSX ordinary worksheet SVG lifecycle profile",
        "",
        "This is an absolute profile of deterministic bounded synthetic inputs.",
        "It is not a before/after speedup claim and makes no native Office or",
        "rendering claim. End-to-end rows include only the named commit,",
        "publication, reopen, and semantic checks described in requirements.md.",
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
