#!/usr/bin/env python3
"""Summarize the retained 0502 ODG before/after reports deterministically."""

from __future__ import annotations

import argparse
import json
import math
import random
from pathlib import Path


CORPORA = ("plain-small", "plain-large", "metadata-small", "metadata-large")
REPEATS = ("r1", "r2")
BOOTSTRAP_SAMPLES = 10_000


def probe_percentile(values: list[float], quantile: float) -> float:
    ordered = sorted(values)
    # Rust's f64::round used by the probe rounds halfway cases away from zero;
    # Python's round uses ties-to-even.
    position = math.floor((len(ordered) - 1) * quantile + 0.5)
    return ordered[position]


def validate_report(report: dict, path: Path, corpus: str) -> list[int]:
    if report.get("schema") != "litchi.odg.open-probe.v1":
        raise ValueError(f"{path}: unexpected probe schema")
    for field in ("corpus", "input_bytes", "input_sha256", "warmups", "samples"):
        if field not in report:
            raise ValueError(f"{path}: missing {field}")
    if report["corpus"] != corpus:
        raise ValueError(f"{path}: corpus does not match {corpus}")
    if not isinstance(report["input_bytes"], int) or report["input_bytes"] <= 0:
        raise ValueError(f"{path}: input_bytes must be positive")
    if not isinstance(report["input_sha256"], str) or not report["input_sha256"]:
        raise ValueError(f"{path}: input_sha256 must be non-empty")
    if not isinstance(report["warmups"], int) or report["warmups"] < 0:
        raise ValueError(f"{path}: warmups must be non-negative")
    if not isinstance(report["samples"], int) or report["samples"] <= 0:
        raise ValueError(f"{path}: samples must be positive")
    elapsed = report.get("elapsed_ns")
    if not isinstance(elapsed, list) or len(elapsed) != report["samples"]:
        raise ValueError(f"{path}: samples does not match elapsed_ns length")
    if any(
        isinstance(value, bool) or not isinstance(value, int) or value <= 0
        for value in elapsed
    ):
        raise ValueError(f"{path}: elapsed_ns must contain positive integers")
    statistics = report.get("statistics")
    if not isinstance(statistics, dict):
        raise ValueError(f"{path}: missing statistics")
    for field in ("min_ns", "mean_ns", "p50_ns", "p95_ns", "p99_ns", "max_ns"):
        value = statistics.get(field)
        if not isinstance(value, (int, float)) or not math.isfinite(float(value)) or value <= 0:
            raise ValueError(f"{path}: statistics.{field} must be finite and positive")
    expected_p50 = probe_percentile([float(value) for value in elapsed], 0.50)
    if statistics["p50_ns"] != expected_p50:
        raise ValueError(f"{path}: statistics.p50_ns does not match probe percentile")
    return elapsed


def validate_pair(before: dict, after: dict, before_path: Path, after_path: Path, corpus: str) -> None:
    before_elapsed = validate_report(before, before_path, corpus)
    after_elapsed = validate_report(after, after_path, corpus)
    for field in ("input_bytes", "input_sha256", "warmups", "samples", "rich_metadata", "pages", "shapes_per_page"):
        if before.get(field) != after.get(field):
            raise ValueError(f"{before_path} / {after_path}: {field} differs")
    if len(before_elapsed) != len(after_elapsed):
        raise ValueError(f"{before_path} / {after_path}: sample counts differ")


def load_report(phase_root: Path, corpus: str, repeat: str) -> dict:
    path = phase_root / f"{corpus}-{repeat}.json"
    with path.open(encoding="utf-8") as stream:
        return json.load(stream)


def bootstrap_ratio_ci(
    before: list[float], after: list[float], seed: int
) -> tuple[float, float]:
    rng = random.Random(seed)
    ratios: list[float] = []
    for _ in range(BOOTSTRAP_SAMPLES):
        before_sample = [before[rng.randrange(len(before))] for _ in before]
        after_sample = [after[rng.randrange(len(after))] for _ in after]
        ratios.append(
            probe_percentile(after_sample, 0.50)
            / probe_percentile(before_sample, 0.50)
            - 1.0
        )
    return (
        probe_percentile(ratios, 0.025),
        probe_percentile(ratios, 0.975),
    )


def summarize(root: Path) -> dict:
    rows = []
    for corpus_index, corpus in enumerate(CORPORA):
        for repeat_index, repeat in enumerate(REPEATS):
            before_report = load_report(root / "before", corpus, repeat)
            after_report = load_report(root / "after", corpus, repeat)
            validate_pair(
                before_report,
                after_report,
                root / "before" / f"{corpus}-{repeat}.json",
                root / "after" / f"{corpus}-{repeat}.json",
                corpus,
            )
            # Keep the sample arrays in the reports as nanoseconds and use
            # floating point only for the ratio calculation.
            before = [float(value) for value in before_report["elapsed_ns"]]
            after = [float(value) for value in after_report["elapsed_ns"]]
            ratio_low, ratio_high = bootstrap_ratio_ci(
                before, after, 50_200 + corpus_index * 10 + repeat_index
            )
            observed_ratio = (
                after_report["statistics"]["p50_ns"]
                / before_report["statistics"]["p50_ns"]
                - 1.0
            )
            rows.append(
                {
                    "corpus": corpus,
                    "repeat": repeat,
                    "input_bytes": before_report["input_bytes"],
                    "input_sha256": before_report["input_sha256"],
                    "before_binary_sha256": before_report["binary_sha256"],
                    "after_binary_sha256": after_report["binary_sha256"],
                    "before_p50_ns": before_report["statistics"]["p50_ns"],
                    "after_p50_ns": after_report["statistics"]["p50_ns"],
                    "p50_change_ratio": observed_ratio,
                    "bootstrap_95pct_ratio_ci": [ratio_low, ratio_high],
                    "samples": len(before),
                    "warmups": before_report["warmups"],
                }
            )
    return {
        "schema": "litchi.odg.open-probe-summary.v1",
        "corpora": list(CORPORA),
        "repeats": list(REPEATS),
        "bootstrap": {
            "samples": BOOTSTRAP_SAMPLES,
            "seed_base": 50200,
            "metric": "ratio of independently resampled within-run medians minus one",
            "interpretation": "within-run sampling uncertainty only; not machine-to-machine uncertainty",
        },
        "rows": rows,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    default_root = Path(__file__).resolve().parent
    parser.add_argument("--root", type=Path, default=default_root)
    parser.add_argument("--output", type=Path, default=default_root / "summary.json")
    args = parser.parse_args()
    result = summarize(args.root)
    with args.output.open("w", encoding="utf-8") as stream:
        json.dump(result, stream, indent=2, sort_keys=True)
        stream.write("\n")


if __name__ == "__main__":
    main()
