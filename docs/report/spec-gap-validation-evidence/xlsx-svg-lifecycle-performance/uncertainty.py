#!/usr/bin/env python3
"""Summarize process-level uncertainty without making causal claims."""

from __future__ import annotations

import argparse
import json
import re
import statistics
from pathlib import Path
from typing import Any

from summarize import LANES, rss
from verify import REFUSAL_LANES


SCHEMA = "xlsx-svg-lifecycle-uncertainty-v1"
PROCESS_COUNT = 3
MIN_WARMUP = 2
MIN_SAMPLES = 20
METRICS = (
    ("elapsed_ns", "elapsed_ns"),
    ("requested_alloc_bytes", "requested_alloc_bytes"),
    ("peak_live_delta", "peak_live_delta"),
)
SHA256 = re.compile(r"^[0-9a-f]{64}$")


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def number(value: Any, field: str, path: Path, *, nonnegative: bool = False) -> int:
    require(type(value) is int, f"{field} must be an integer: {path}")
    if nonnegative:
        require(value >= 0, f"{field} must be nonnegative: {path}")
    return value


def value_range(values: list[int | float]) -> dict[str, int | float | list[int | float]]:
    low = min(values)
    high = max(values)
    return {
        "min": low,
        "max": high,
        "range": [low, high],
        "span": high - low,
    }


def metric_summary(values: list[int]) -> dict[str, int | float | list[int | float]]:
    return {
        "median": statistics.median(values),
        **value_range(values),
    }


def process_paths(results: Path, lane: str) -> list[Path]:
    paths = sorted(results.glob(f"{lane}-p*.json"), key=lambda path: path.name)
    expected = [results / f"{lane}-p{process}.json" for process in range(1, PROCESS_COUNT + 1)]
    require(paths == expected, f"expected exactly p1..p{PROCESS_COUNT} receipts for {lane}")
    return paths


def process_summary(path: Path, lane: str, process: int) -> dict[str, object]:
    payload = json.loads(path.read_text())
    require(payload.get("schema") == "xlsx-svg-lifecycle-profile-v1", f"schema mismatch: {path}")
    require(payload.get("lane") == lane, f"lane mismatch: {path}")
    warmup = number(payload.get("warmup"), "warmup", path)
    require(warmup >= MIN_WARMUP, f"warmup below {MIN_WARMUP}: {path}")
    samples = payload.get("samples")
    require(isinstance(samples, list), f"samples missing: {path}")
    sample_count = number(payload.get("sample_count"), "sample_count", path)
    require(sample_count == len(samples), f"sample count does not match samples: {path}")
    require(sample_count >= MIN_SAMPLES, f"sample count below {MIN_SAMPLES}: {path}")
    expected_success = lane not in REFUSAL_LANES
    require(
        payload.get("expected_success") is expected_success,
        f"expected status does not match lane: {path}",
    )
    expected_error = REFUSAL_LANES.get(lane)
    for sample in samples:
        require(isinstance(sample, dict), f"sample malformed: {path}")
        require(
            sample.get("expected_success") is expected_success,
            f"sample expected status mismatch: {path}",
        )
        require(
            sample.get("actual_success") is expected_success,
            f"sample actual status mismatch: {path}",
        )
        require(sample.get("semantic_ok") is True, f"semantic gate failed: {path}")
        require(sample.get("output_exact") is True, f"output gate failed: {path}")
        require(sample.get("alloc_balance_ok") is True, f"allocator balance failed: {path}")
        require(sample.get("alloc_invalid") is False, f"allocator underflow failed: {path}")
        require(sample.get("alloc_failed") == 0, f"allocation failure failed: {path}")
        if expected_success:
            require(sample.get("error") is None, f"unexpected profile error: {path}")
        else:
            error = sample.get("error")
            require(isinstance(error, dict), f"refusal reason missing: {path}")
            require(
                error.get("class") == expected_error,
                f"refusal class mismatch: {path}",
            )

    metrics = {
        name: metric_summary(
            [number(sample.get(field), field, path, nonnegative=True) for sample in samples]
        )
        for name, field in METRICS
    }
    input_bytes = number(payload.get("input_bytes"), "input_bytes", path, nonnegative=True)
    input_sha256 = payload.get("input_sha256")
    require(
        isinstance(input_sha256, str) and SHA256.fullmatch(input_sha256) is not None,
        f"input_sha256 must be a lowercase SHA-256 digest: {path}",
    )
    rss_kib = rss(path.with_suffix(".time.txt"))
    require(type(rss_kib) is int and rss_kib >= 0, f"rss must be nonnegative: {path}")
    return {
        "process": process,
        "receipt": path.name,
        "warmup": warmup,
        "sample_count": sample_count,
        "metrics": metrics,
        "rss_kib": rss_kib,
        "input": {
            "bytes": input_bytes,
            "sha256": input_sha256,
        },
        "expected_success": expected_success,
    }


def between_processes(processes: list[dict[str, object]], metric: str) -> dict[str, object]:
    medians = [process["metrics"][metric]["median"] for process in processes]  # type: ignore[index]
    return {
        "process_medians": medians,
        "median": statistics.median(medians),
        **value_range(medians),
    }


def summarize_results(results: Path) -> dict[str, object]:
    known_lanes = set(LANES)
    for path in sorted(results.glob("*.json")):
        receipt = re.fullmatch(r"(.+)-p[0-9]+\.json", path.name)
        if receipt is not None:
            require(receipt[1] in known_lanes, f"unknown lane receipt: {path}")
    lanes: list[dict[str, object]] = []
    for lane in LANES:
        paths = process_paths(results, lane)
        processes = [process_summary(path, lane, index) for index, path in enumerate(paths, 1)]
        identities = {
            (process["input"]["bytes"], process["input"]["sha256"])  # type: ignore[index]
            for process in processes
        }
        require(len(identities) == 1, f"input identity changed across processes: {lane}")
        sample_counts = {process["sample_count"] for process in processes}
        require(len(sample_counts) == 1, f"sample count changed across processes: {lane}")
        expected_statuses = {process["expected_success"] for process in processes}
        require(len(expected_statuses) == 1, f"expected status changed across processes: {lane}")
        medians = {metric: between_processes(processes, metric) for metric, _ in METRICS}
        rss_values = [process["rss_kib"] for process in processes]
        lanes.append(
            {
                "lane": lane,
                "process_count": len(processes),
                "sample_count_per_process": next(iter(sample_counts)),
                "input": processes[0]["input"],
                "expected_success": processes[0]["expected_success"],
                "processes": processes,
                "between_process_medians": medians,
                "rss_kib": {
                    "process_values": rss_values,
                    "median": statistics.median(rss_values),
                    **value_range(rss_values),
                },
            }
        )
    return {
        "schema": SCHEMA,
        "basis": "descriptive per-process medians and within-process min/max ranges",
        "process_count": PROCESS_COUNT,
        "minimum_warmup": MIN_WARMUP,
        "minimum_samples_per_process": MIN_SAMPLES,
        "limitations": [
            "The process count is n=3; between-process ranges are descriptive and are not confidence intervals.",
            "Within-process ranges describe the retained samples and do not model scheduler, cache, or host variation.",
            "This summary does not estimate a speedup, regression, causal effect, or algorithmic scaling law.",
        ],
        "lanes": lanes,
    }


def format_value(value: int | float) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def metric_markdown(summary: dict[str, object], metric: str) -> str:
    processes = summary["processes"]
    medians = "; ".join(
        f"p{process['process']}={format_value(process['metrics'][metric]['median'])}"  # type: ignore[index]
        for process in processes  # type: ignore[union-attr]
    )
    ranges = "; ".join(
        f"p{process['process']}=[{format_value(process['metrics'][metric]['min'])},{format_value(process['metrics'][metric]['max'])}]"  # type: ignore[index]
        for process in processes  # type: ignore[union-attr]
    )
    between = summary["between_process_medians"][metric]  # type: ignore[index]
    return (
        f"- `{metric}`: per-process medians {medians}; within-process sample ranges "
        f"{ranges}; between-process median range "
        f"[{format_value(between['min'])},{format_value(between['max'])}]."  # type: ignore[index]
    )


def to_markdown(summary: dict[str, object]) -> str:
    lines = [
        "# XLSX ordinary worksheet SVG lifecycle uncertainty summary",
        "",
        "This is a descriptive summary of retained raw receipts. It reports each "
        "fresh process median, each process's sample min/max range, and the "
        "range of process medians. It does not estimate a speedup, regression, "
        "causal effect, confidence interval, or scaling law.",
        "",
        f"- Processes per lane: n={summary['process_count']}",
        f"- Minimum warmups per process: {summary['minimum_warmup']}",
        f"- Minimum measured samples per process: {summary['minimum_samples_per_process']}",
        "- Raw receipts are unchanged; this file is a derived view.",
        "",
        "## Limitations",
        "",
    ]
    lines.extend(f"- {limitation}" for limitation in summary["limitations"])
    for lane in summary["lanes"]:  # type: ignore[union-attr]
        lines.extend(
            [
                "",
                f"## `{lane['lane']}`",  # type: ignore[index]
                "",
                f"- Input bytes: {lane['input']['bytes']}; SHA-256: `{lane['input']['sha256']}`.",  # type: ignore[index]
                f"- Shape: n={lane['process_count']}; samples/process={lane['sample_count_per_process']}.",  # type: ignore[index]
            ]
        )
        for metric, _ in METRICS:
            lines.append(metric_markdown(lane, metric))  # type: ignore[arg-type]
        rss_values = "; ".join(
            f"p{process['process']}={process['rss_kib']}"  # type: ignore[index]
            for process in lane["processes"]  # type: ignore[index]
        )
        rss_summary = lane["rss_kib"]  # type: ignore[index]
        lines.append(
            f"- `rss_kib`: per-process values {rss_values}; range "
            f"[{rss_summary['min']},{rss_summary['max']}]."  # type: ignore[index]
        )
    return "\n".join(lines) + "\n"


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--results", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True, help="Markdown summary path")
    parser.add_argument("--json-output", type=Path, help="Optional machine-readable summary path")
    args = parser.parse_args()
    summary = summarize_results(args.results.resolve())
    args.output.write_text(to_markdown(summary))
    if args.json_output is not None:
        args.json_output.write_text(json.dumps(summary, indent=2) + "\n")


if __name__ == "__main__":
    main()
