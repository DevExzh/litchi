#!/usr/bin/env python3
"""Verify and summarize raw XLSB binary-index profile reports."""

from __future__ import annotations

import json
import statistics
import sys
from collections import defaultdict
from pathlib import Path


CASES = {
    "indexed_cold",
    "materialize_cold",
    "indexed_warm",
    "materialize_warm",
}
IDENTITY_KEYS = (
    "input_bytes",
    "input_digest",
    "cell_count",
    "dimensions",
    "targets",
)


def nearest_rank(values: list[int], percentile: int) -> int:
    ordered = sorted(values)
    rank = ((len(ordered) * percentile) + 99) // 100
    return ordered[max(0, min(rank - 1, len(ordered) - 1))]


def source_observation_stable(report: dict[str, object]) -> bool:
    samples = report["samples"]
    assert isinstance(samples, list)
    for first, second in zip(samples, samples[1:]):
        assert isinstance(first, dict) and isinstance(second, dict)
        if any(
            first[key] != second[key]
            for key in (
                "reads_operation",
                "cache_operation",
                "retained_entries",
                "retained_bytes",
            )
        ):
            return False
    return True


def check_allocation(allocation: dict[str, object], where: str) -> None:
    if allocation.get("invalid") is not False:
        raise SystemExit(f"allocator live accounting invalid in {where}")
    if allocation.get("failed") != 0:
        raise SystemExit(f"allocator failure count is nonzero in {where}")
    if allocation["peak_after"] < allocation["peak_before"]:
        raise SystemExit(f"interval peak regressed below baseline in {where}")
    if allocation["peak_before"] != allocation["live_before"]:
        raise SystemExit(f"interval peak was not reset to the live baseline in {where}")
    if allocation["peak_after"] < allocation["live_after"]:
        raise SystemExit(f"interval peak is below live bytes in {where}")
    if (
        allocation["live_before"]
        + allocation["allocated_bytes"]
        - allocation["deallocated_bytes"]
        != allocation["live_after"]
    ):
        raise SystemExit(f"allocator byte interval does not balance in {where}")


def check_source_intervals(sample: dict[str, object], where: str) -> None:
    for field in ("calls", "requested_bytes", "returned_bytes"):
        if (
            sample["reads_before"][field] + sample["reads_operation"][field]
            != sample["reads_total"][field]
        ):
            raise SystemExit(f"read interval does not balance in {where}: {field}")
    for field in (
        "hits",
        "cold_loads",
        "waiter_joins",
        "successful_loads",
        "failed_loads",
        "evictions",
        "bypasses",
        "oversized_bypasses",
        "allocation_bypasses",
        "budget_reservation_failures",
    ):
        if (
            sample["cache_before"][field] + sample["cache_operation"][field]
            != sample["cache_total"][field]
        ):
            raise SystemExit(f"cache interval does not balance in {where}: {field}")


def verify_report(path: Path, warmup: int, samples: int) -> dict[str, object]:
    report = json.loads(path.read_text(encoding="utf-8"))
    if report.get("schema") != "xlsb-binary-index-profile-v1":
        raise SystemExit(f"unexpected schema in {path.name}")
    case = report.get("case")
    if case not in CASES:
        raise SystemExit(f"unexpected case {case!r} in {path.name}")
    if report.get("warmup") != warmup or report.get("sample_count") != samples:
        raise SystemExit(f"sample contract failed in {path.name}")
    raw_samples = report.get("samples")
    if not isinstance(raw_samples, list) or len(raw_samples) != samples:
        raise SystemExit(f"raw sample cardinality failed in {path.name}")
    elapsed = []
    digests = []
    for index, sample in enumerate(raw_samples):
        if not isinstance(sample, dict):
            raise SystemExit(f"sample {index} is not an object in {path.name}")
        if sample.get("semantic_ok") is not True:
            raise SystemExit(f"semantic gate failed in {path.name}")
        elapsed.append(int(sample["elapsed_ns"]))
        digests.append(sample["digest"])
        check_allocation(sample["allocation"], f"{path.name} sample {index}")
        check_source_intervals(sample, f"{path.name} sample {index}")
    if report.get("semantic_ok") is not True or report.get("digest_stable") is not True:
        raise SystemExit(f"semantic/digest report gate failed in {path.name}")
    if len(set(digests)) != 1:
        raise SystemExit(f"raw digest changed in {path.name}")
    if report.get("source_observation_stable") is not True or not source_observation_stable(report):
        raise SystemExit(f"source observation changed in {path.name}")
    expected = {
        "p50_ns": nearest_rank(elapsed, 50),
        "p95_ns": nearest_rank(elapsed, 95),
        "p99_ns": nearest_rank(elapsed, 99),
    }
    for key, value in expected.items():
        if report.get(key) != value:
            raise SystemExit(f"{key} is not reproducible from raw samples in {path.name}")
    expected_mean = statistics.fmean(elapsed)
    if abs(float(report["mean_ns"]) - expected_mean) > 0.0005:
        raise SystemExit(f"mean_ns is not reproducible from raw samples in {path.name}")
    if case.endswith("_cold"):
        if any(
            report.get(key) is not None
            for key in ("setup_reads", "setup_cache", "setup_allocation")
        ):
            raise SystemExit(f"cold report has warm setup fields in {path.name}")
        if (
            report.get("setup_retained_entries") is not None
            or report.get("setup_retained_bytes") is not None
        ):
            raise SystemExit(f"cold report has setup retained gauges in {path.name}")
    else:
        if any(
            report.get(key) is None
            for key in ("setup_reads", "setup_cache", "setup_allocation")
        ):
            raise SystemExit(f"warm report is missing setup fields in {path.name}")
        if (
            report.get("setup_retained_entries") is None
            or report.get("setup_retained_bytes") is None
        ):
            raise SystemExit(f"warm report is missing setup retained gauges in {path.name}")
        check_allocation(report["setup_allocation"], f"{path.name} setup")
    return report


def main() -> None:
    run = (
        Path(sys.argv[1]).resolve()
        if len(sys.argv) > 1
        else Path(__file__).resolve().parent / "raw"
    )
    warmup = int(sys.argv[2]) if len(sys.argv) > 2 else 3
    samples = int(sys.argv[3]) if len(sys.argv) > 3 else 30
    paths = sorted(
        path for path in run.glob("*.json") if path.name != "matrix-summary.json"
    )
    if not paths:
        raise SystemExit(f"no raw JSON reports in {run}")

    reports = [(path, verify_report(path, warmup, samples)) for path in paths]
    groups: dict[tuple[str, str], list[dict[str, object]]] = defaultdict(list)
    fixture_identity: dict[str, tuple[object, ...]] = {}
    for path, report in reports:
        fixture = report["fixture"]
        key = (fixture, report["case"])
        groups[key].append(report)
        identity = tuple(report[name] for name in IDENTITY_KEYS)
        prior = fixture_identity.setdefault(fixture, identity)
        if identity != prior:
            raise SystemExit(f"fixture identity changed across reports in {path.name}")

    for fixture in fixture_identity:
        cases = {case for (name, case) in groups if name == fixture}
        if cases != CASES:
            raise SystemExit(
                f"fixture {fixture!r} does not have exactly the four profile cases"
            )

    summary_groups = {}
    for fixture, case in sorted(groups):
        group = groups[(fixture, case)]
        summary_groups[f"{fixture}/{case}"] = {
            "fixture": fixture,
            "case": case,
            "process_count": len(group),
            "timing_scopes": sorted({report["timing_scope"] for report in group}),
            "process_p50_ns": [report["p50_ns"] for report in group],
            "process_p95_ns": [report["p95_ns"] for report in group],
            "process_p99_ns": [report["p99_ns"] for report in group],
            "process_mean_ns": [report["mean_ns"] for report in group],
            "median_process_p50_ns": statistics.median(
                report["p50_ns"] for report in group
            ),
            "median_process_p95_ns": statistics.median(
                report["p95_ns"] for report in group
            ),
            "median_process_p99_ns": statistics.median(
                report["p99_ns"] for report in group
            ),
            "setup_allocation": [report["setup_allocation"] for report in group],
            "setup_retained_entries": [
                report["setup_retained_entries"] for report in group
            ],
            "setup_retained_bytes": [
                report["setup_retained_bytes"] for report in group
            ],
            "sample_source_observations": [
                {
                    "reads_operation": report["samples"][0]["reads_operation"],
                    "cache_operation": report["samples"][0]["cache_operation"],
                    "retained_entries": report["samples"][0]["retained_entries"],
                    "retained_bytes": report["samples"][0]["retained_bytes"],
                }
                for report in group
            ],
        }

    output = run.parent / "matrix-summary.json"
    summary = {
        "schema": "xlsb-binary-index-profile-matrix-v1",
        "raw_run_directory": str(run),
        "contract": {
            "fixtures": len(fixture_identity),
            "cases_per_fixture": len(CASES),
            "process_reports": len(reports),
            "warmup_per_report": warmup,
            "samples_per_report": samples,
        },
        "comparison_scope": "No equivalent-work speedup inferred; indexed and materialized API validation scopes differ.",
        "groups": summary_groups,
    }
    output.write_text(
        json.dumps(summary, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )
    print(f"verified_reports={len(reports)}")
    print(f"verified_fixtures={len(fixture_identity)}")
    print(f"summary={output}")


if __name__ == "__main__":
    main()
