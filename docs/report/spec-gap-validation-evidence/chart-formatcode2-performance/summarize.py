#!/usr/bin/env python3
"""Aggregate the raw current-source formatcode2 profile receipts."""

from __future__ import annotations

import argparse
import json
import math
from pathlib import Path


LANES = (
    "element_small_read",
    "element_small_noop",
    "element_small_scalar_edit",
    "element_small_clone",
    "element_small_read_shared",
    "element_small_write_to",
    "element_small_malformed",
    "element_near_limit_read",
    "element_near_limit_noop",
    "element_near_limit_scalar_edit",
    "element_near_limit_clone",
    "element_near_limit_read_shared",
    "element_near_limit_write_to",
    "element_near_limit_malformed",
    "attribute_small_read",
    "attribute_small_noop",
    "attribute_small_scalar_edit",
    "attribute_small_clone",
    "attribute_small_read_shared",
    "attribute_small_write_to",
    "attribute_small_malformed",
    "attribute_near_limit_read",
    "attribute_near_limit_noop",
    "attribute_near_limit_scalar_edit",
    "attribute_near_limit_clone",
    "attribute_near_limit_read_shared",
    "attribute_near_limit_write_to",
    "attribute_near_limit_malformed",
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
        for path in sorted(args.results.glob(f"{lane}-p*.json")):
            value = json.loads(path.read_text())
            values.extend(value["samples"])
            input_bytes.add(int(value["input_bytes"]))
            rss_values.append(rss(path.with_suffix(".time.txt")))
        if len(input_bytes) != 1:
            raise ValueError(f"input size changed across {lane}")
        elapsed = [int(sample["elapsed_ns"]) for sample in values]
        allocated = [int(sample["requested_alloc_bytes"]) for sample in values]
        peak = [int(sample["peak_live_delta"]) for sample in values]
        rows.append(
            {
                "lane": lane,
                "processes": len(rss_values),
                "samples": len(values),
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
        "# `formatcode2` bounded allocator/runtime profile",
        "",
        "This report contains absolute current-source measurements from three fresh processes and twenty measured samples per lane across 28 bounded lanes. The timer covers the named `formatcode2` owner operation with fixture construction and post-timer semantic checks outside the timed interval. `write_to` timing includes the sink's nonallocating byte-count, FNV-1a checksum, and write-call counting work; it makes no comparison claim against the Vec-returning lane. Requested allocation bytes and incremental peak live bytes come from a process-local counting allocator; RSS is whole-process `/usr/bin/time -v` RSS.",
        "",
        "| lane | fresh processes | samples | input bytes | p50 / p95 / p99 ns | requested alloc p50 / p95 B | peak-live delta p50 / p95 B | RSS KiB |",
        "|---|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for row in rows:
        lines.append(
            "| {lane} | {processes} | {samples} | {input_bytes} | {p50} / {p95} / {p99} | {alloc50} / {alloc95} | {peak50} / {peak95} | {rss_min}–{rss_max} |".format(
                **row
            )
        )
    lines.extend(
        [
            "",
            "`element_*` uses a short source-preserving element or a valid source of `MAX_XML_BYTES - 1` bytes. `attribute_*` uses a complete host start tag containing the qualified chart attribute; its near-limit form fills bounded ordinary attributes while keeping every value below the per-attribute ceiling. Each near-limit decoded value is close to `MAX_VALUE_BYTES`; its malformed lane appends one non-whitespace byte and remains within the source ceiling. Each small malformed lane exercises an unpaired ST_Xstring surrogate escape. `clone` is the prepared typed owner's `Clone`, `read_shared` parses an existing `Arc<[u8]>` prepared outside the timer, `no-op` serializes the unchanged prepared owner into an owned `Vec`, `write_to` serializes the same unchanged owner into a counting/hash sink, and `scalar_edit` changes the decoded value and serializes it.",
            "",
            "These are scoped absolute observations. They do not establish a before/after speedup, an asymptotic result, a host-placement guarantee, or a whole-library performance claim. The shared attribute owner is measured only for its complete start-tag contract; the package host still owns placement and parent grammar.",
            "",
            "Raw per-process JSON, `/usr/bin/time -v` receipts, source/build manifests, exact commands, and source hashes are retained in this directory.",
        ]
    )
    args.output.write_text("\n".join(lines) + "\n")


if __name__ == "__main__":
    main()
