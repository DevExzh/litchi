#!/usr/bin/env python3
"""Derive descriptive phase medians and retained-live ranges."""

from __future__ import annotations

import argparse
import statistics
from pathlib import Path

from phase_verify import PHASES, PICTURE_COUNTS, verify_receipt, verify_results


def value_range(values: list[int | float]) -> str:
    return f"[{min(values)},{max(values)}]"


def metric_values(payload: dict, phase: str, field: str) -> list[int]:
    return [sample["phases"][phase][field] for sample in payload["samples"]]


def process_metric(payload: dict, phase: str, field: str) -> int | float:
    return statistics.median(metric_values(payload, phase, field))


def largest_phase(payloads: list[dict], field: str) -> list[tuple[str, int | float]]:
    """Return each process's largest phase median for one descriptive metric."""
    return [
        max(
            ((phase, process_metric(payload, phase, field)) for phase in PHASES),
            key=lambda item: item[1],
        )
        for payload in payloads
    ]


def summarize(results: Path) -> str:
    verify_results(results)
    lines = [
        "# XLSX SVG same-drawing phase decomposition",
        "",
        "This is an exploratory descriptive view of one public attach operation",
        "split into explicit phases. It reports absolute per-process medians and",
        "ranges. It does not compare against the sealed baseline and makes no",
        "speedup, regression, causal, or scaling claim.",
        "",
        "Each phase resets the process-local counter baseline, while objects needed",
        "by later phases remain live. `requested_alloc_bytes` is phase-local;",
        "`live_after_bytes` and `retained_live_bytes_after` expose the retained",
        "boundary. Phase values must not be summed or subtracted across unlike",
        "live sets.",
        "",
        "The six phase clocks are `open` (Workbook::from_bytes), `stages`",
        "(Workbook::edit plus all public attach calls), `commit`, `firstsave`,",
        "`reopen_secondsave`, and `validation` (the complete attach semantic",
        "predicate, including graph, picture, opaque, and reopen-byte checks).",
    ]
    for pictures in PICTURE_COUNTS:
        lines.extend(
            [
                "",
                f"## `multi_picture_same_drawing_{pictures}`",
                "",
                "| phase | elapsed median p1/p2/p3 (ns) | elapsed process range (ns) | requested allocation median p1/p2/p3 (bytes) | live delta median p1/p2/p3 (bytes) | retained live after median p1/p2/p3 (bytes) | RSS p1/p2/p3 (KiB) |",
                "|---|---:|---:|---:|---:|---:|---:|",
            ]
        )
        payloads = []
        rss_values = []
        for process in (1, 2, 3):
            path = results / f"phase_{pictures}-p{process}.json"
            payload, rss = verify_receipt(path, pictures, process)
            payloads.append(payload)
            rss_values.append(rss)
        for phase in PHASES:
            elapsed = [process_metric(payload, phase, "elapsed_ns") for payload in payloads]
            requested = [process_metric(payload, phase, "requested_alloc_bytes") for payload in payloads]
            live_delta = [process_metric(payload, phase, "live_delta_bytes") for payload in payloads]
            retained = [process_metric(payload, phase, "retained_live_bytes_after") for payload in payloads]
            all_elapsed = [value for payload in payloads for value in metric_values(payload, phase, "elapsed_ns")]
            lines.append(
                f"| `{phase}` | {' / '.join(map(str, elapsed))} | {value_range(all_elapsed)} | "
                f"{' / '.join(map(str, requested))} | {' / '.join(map(str, live_delta))} | "
                f"{' / '.join(map(str, retained))} | {' / '.join(map(str, rss_values))} |"
            )
        lines.append("")
        largest_elapsed = largest_phase(payloads, "elapsed_ns")
        largest_requested = largest_phase(payloads, "requested_alloc_bytes")
        lines.append(
            "Largest observed phase medians by fresh process (descriptive only): "
            + "; ".join(f"p{index}: `{phase}`={value} ns" for index, (phase, value) in enumerate(largest_elapsed, 1))
            + "."
        )
        lines.append(
            "Largest phase-local requested-allocation medians by fresh process "
            "(descriptive only): "
            + "; ".join(f"p{index}: `{phase}`={value} bytes" for index, (phase, value) in enumerate(largest_requested, 1))
            + "."
        )
        lines.append(f"RSS process range: {value_range(rss_values)} KiB; this is whole-process `/usr/bin/time -v` RSS.")
    return "\n".join(lines) + "\n"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--results", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    args.output.write_text(summarize(args.results.resolve()))


if __name__ == "__main__":
    main()
