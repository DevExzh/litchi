#!/usr/bin/env python3
"""Aggregate raw current-source ink-action edit profile receipts."""

from __future__ import annotations

import argparse
import json
import math
from pathlib import Path


LANES = (
    "draft_small_8",
    "draft_scaled_128",
    "draft_near_1024",
    "draft_opaque_64",
    "scalar_edit_small_8",
    "scalar_edit_scaled_128",
    "scalar_edit_near_1024",
    "scalar_batch_scaled_128",
    "scalar_batch_near_1024",
    "scalar_coalesce_scaled_128",
    "scalar_coalesce_near_1024",
    "no_op_small_8",
    "no_op_scaled_128",
    "no_op_near_1024",
    "add_small_8",
    "add_scaled_128",
    "add_near_1024",
    "insert_batch_scaled_128",
    "insert_batch_near_1024",
    "remove_small_8",
    "remove_scaled_128",
    "remove_near_1024",
    "remove_batch_scaled_128",
    "remove_batch_near_1024",
    "clear_batch_scaled_128",
    "clear_batch_near_1024",
    "move_small_8",
    "move_scaled_128",
    "move_near_1024",
    "move_batch_scaled_128",
    "move_batch_near_1024",
    "cap_refusal_small_8",
    "cap_refusal_scaled_128",
    "cap_refusal_near_1024",
)


def quantile(values: list[int], percent: int) -> int:
    return sorted(values)[max(0, math.ceil(len(values) * percent / 100) - 1)]


def rss(path: Path) -> int:
    for line in path.read_text().splitlines():
        if line.lstrip().startswith("Maximum resident set size (kbytes):"):
            return int(line.split(":", 1)[1].strip())
    raise ValueError(f"RSS missing from {path}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--results", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()

    rows: list[dict[str, object]] = []
    for lane in LANES:
        values = []
        rss_values = []
        input_bytes = set()
        action_counts = set()
        result_action_counts = set()
        operation_counts = set()
        for path in sorted(args.results.glob(f"{lane}-p*.json")):
            value = json.loads(path.read_text())
            values.extend(value["samples"])
            input_bytes.add(int(value["input_bytes"]))
            action_counts.add(int(value["action_count"]))
            result_action_counts.add(int(value["result_action_count"]))
            operation_counts.add(int(value["operation_count"]))
            rss_values.append(rss(path.with_suffix(".time.txt")))
        if (
            len(input_bytes) != 1
            or len(action_counts) != 1
            or len(result_action_counts) != 1
            or len(operation_counts) != 1
        ):
            raise ValueError(f"fixture metadata changed across {lane}")
        elapsed = [int(sample["elapsed_ns"]) for sample in values]
        allocated = [int(sample["requested_alloc_bytes"]) for sample in values]
        peak = [int(sample["peak_live_delta"]) for sample in values]
        rows.append(
            {
                "lane": lane,
                "processes": len(rss_values),
                "samples": len(values),
                "action_count": next(iter(action_counts)),
                "result_action_count": next(iter(result_action_counts)),
                "operation_count": next(iter(operation_counts)),
                "input_bytes": next(iter(input_bytes)),
                "p50": quantile(elapsed, 50),
                "p95": quantile(elapsed, 95),
                "p99": quantile(elapsed, 99),
                "alloc50": quantile(allocated, 50),
                "alloc95": quantile(allocated, 95),
                "peak50": quantile(peak, 50),
                "peak95": quantile(peak, 95),
                "rss_min": min(rss_values),
                "rss_max": max(rss_values),
            }
        )

    lines = [
        "# Ink-action edit bounded allocator/runtime profile",
        "",
        "This report contains absolute current-source measurements from three fresh processes and twenty measured samples per lane across 34 bounded lanes. The timer covers the named detached draft or source-backed edit workflow, including its bounded validation and output allocation; fixture construction and post-timer semantic, source, patch, and inverse checks stay outside the timed interval. Requested allocation bytes and incremental peak live bytes come from a process-local counting allocator; RSS is whole-process `/usr/bin/time -v` RSS.",
        "",
        "| lane | fresh processes | samples | input actions | result actions | queued operations | input bytes | p50 / p95 / p99 ns | requested alloc p50 / p95 B | peak-live delta p50 / p95 B | RSS KiB |",
        "|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for row in rows:
        lines.append(
            "| {lane} | {processes} | {samples} | {action_count} | {result_action_count} | {operation_count} | {input_bytes} | {p50} / {p95} / {p99} | {alloc50} / {alloc95} | {peak50} / {peak95} | {rss_min}–{rss_max} |".format(
                **row
            )
        )
    lines.extend(
        [
            "",
            "The small, scaled, and near-limit source-backed lanes use 8, 128, and 1,024 direct actions. Draft creation adds a 64-action lane with complete namespace-bearing `inkml:definitions` and `inkml:trace` opaque payloads. Scalar replacement, distinct scalar batches, repeated writes to one scalar, root insertion batches, structural add/remove/clear/move, and exact no-op edits retain source comments, identifiers, namespaces, and opaque bytes; the no-op lanes additionally require source allocation sharing. The repeated-write lane reports only the final scalar value and measured cost; the public API exposes no internal coalescing diagnostic, so the report makes no coalescing claim. Caller-cap lanes attempt a bounded property expansion and must refuse at the configured complete-output limit without changing the source snapshot. Batch operation counts and resulting action counts are retained in each raw receipt so the report distinguishes one edit in a large source from many queued edits.",
            "",
            "Every successful edit is checked after timing by applying its source-checked patch and inverse. The report records only absolute observations for this detached API and fixture matrix. It makes no before/after speedup, native-application, asymptotic, or host-placement claim. It makes no package-wide performance claim.",
            "",
            "Raw per-process JSON, `/usr/bin/time -v` receipts, source/build manifests, exact commands, and source hashes are retained in this directory.",
        ]
    )
    args.output.write_text("\n".join(lines) + "\n")


if __name__ == "__main__":
    main()
