#!/usr/bin/env python3
"""Verify raw XLSB Theme profile reports and summarize process repeats."""

from __future__ import annotations

import json
import statistics
import sys
from collections import defaultdict
from pathlib import Path


CASES = {
    "source_cold",
    "eager_cold",
    "source_warm",
    "eager_warm",
    "noop",
    "change_inverse",
}
IDENTITY_KEYS = (
    "input_bytes",
    "theme_bytes",
    "input_digest",
    "theme_digest",
    "expected_semantic_digest",
    "expected_metadata_digest",
)
MANDATORY_REPORT_GATES = (
    "semantic_ok",
    "digest_stable",
    "preservation_ok",
    "inverse_ok",
    "change_ok",
    "source_observation_stable",
)
MANDATORY_SAMPLE_GATES = (
    "preservation_ok",
    "inverse_ok",
    "change_ok",
)


def nearest_rank(values: list[int], percentile: int) -> int:
    ordered = sorted(values)
    rank = ((len(ordered) * percentile) + 99) // 100
    return ordered[max(0, min(rank - 1, len(ordered) - 1))]


def check_allocator(allocation: dict[str, object], where: str) -> None:
    if allocation.get("invalid") is not False:
        raise SystemExit(f"allocator accounting invalid in {where}")
    if allocation.get("failed") != 0:
        raise SystemExit(f"allocator failure count nonzero in {where}")
    if allocation["peak_before"] != allocation["live_before"]:
        raise SystemExit(f"interval peak baseline mismatch in {where}")
    if allocation["peak_after"] < allocation["peak_before"]:
        raise SystemExit(f"interval peak regressed in {where}")
    if allocation["peak_after"] < allocation["live_after"]:
        raise SystemExit(f"interval peak below live bytes in {where}")
    if (
        allocation["live_before"]
        + allocation["allocated_bytes"]
        - allocation["deallocated_bytes"]
        != allocation["live_after"]
    ):
        raise SystemExit(f"allocator interval does not balance in {where}")


def check_reads(sample: dict[str, object], where: str) -> None:
    before = sample.get("reads_before")
    operation = sample.get("reads")
    total = sample.get("reads_total")
    if not all(isinstance(value, dict) for value in (before, operation, total)):
        raise SystemExit(f"missing read interval fields in {where}")
    for field in ("calls", "requested_bytes", "returned_bytes"):
        if before[field] + operation[field] != total[field]:
            raise SystemExit(f"read interval does not balance in {where}: {field}")
        if operation[field] < 0:
            raise SystemExit(f"negative read counter in {where}: {field}")
    if operation["returned_bytes"] > operation["requested_bytes"]:
        raise SystemExit(f"read returned bytes exceed requested bytes in {where}")


def sample_signature(sample: dict[str, object]) -> tuple[object, ...]:
    return (
        sample["semantic_digest"],
        sample["metadata_digest"],
        sample["source_digest"],
        sample["preservation_ok"],
        sample["inverse_ok"],
        sample["change_ok"],
        json.dumps(sample["reads_before"], sort_keys=True),
        json.dumps(sample["reads"], sort_keys=True),
        json.dumps(sample["reads_total"], sort_keys=True),
    )


def verify_report(path: Path, warmup: int, samples: int) -> dict[str, object]:
    report = json.loads(path.read_text(encoding="utf-8"))
    if report.get("schema") != "xlsb-theme-profile-v1":
        raise SystemExit(f"unexpected schema in {path.name}")
    if report.get("case") not in CASES:
        raise SystemExit(f"unexpected case in {path.name}: {report.get('case')!r}")
    if report.get("warmup") != warmup or report.get("sample_count") != samples:
        raise SystemExit(f"sample contract failed in {path.name}")
    for gate in MANDATORY_REPORT_GATES:
        if report.get(gate) is not True:
            raise SystemExit(f"report gate {gate} failed in {path.name}")
    raw = report.get("samples")
    if not isinstance(raw, list) or len(raw) != samples:
        raise SystemExit(f"raw sample cardinality failed in {path.name}")
    elapsed: list[int] = []
    signatures: list[tuple[object, ...]] = []
    for index, sample in enumerate(raw):
        if not isinstance(sample, dict):
            raise SystemExit(f"sample {index} is not an object in {path.name}")
        for gate in MANDATORY_SAMPLE_GATES:
            if sample.get(gate) is not True:
                raise SystemExit(f"sample gate {gate} failed in {path.name} #{index}")
        elapsed.append(int(sample["elapsed_ns"]))
        signatures.append(sample_signature(sample))
        check_reads(sample, f"{path.name} sample {index}")
        check_allocator(sample["allocation"], f"{path.name} sample {index}")
    if len(set(signatures)) != 1:
        raise SystemExit(f"semantic/source observations changed in {path.name}")
    expected = {
        "p50_ns": nearest_rank(elapsed, 50),
        "p95_ns": nearest_rank(elapsed, 95),
        "p99_ns": nearest_rank(elapsed, 99),
    }
    for key, value in expected.items():
        if report.get(key) != value:
            raise SystemExit(f"{key} is not reproducible in {path.name}")
    mean = sum(elapsed) // len(elapsed)
    if report["mean_ns"] != mean:
        raise SystemExit(f"mean_ns is not reproducible in {path.name}")
    if report["case"].endswith("_warm"):
        if not isinstance(report.get("setup_reads"), dict):
            raise SystemExit(f"warm report is missing setup reads in {path.name}")
        if not isinstance(report.get("setup_allocation"), dict):
            raise SystemExit(f"warm report is missing setup allocation in {path.name}")
        check_allocator(report["setup_allocation"], f"{path.name} setup")
    elif "setup_reads" in report or "setup_allocation" in report:
        raise SystemExit(f"cold/transaction report has warm setup fields in {path.name}")
    return report


def main() -> None:
    run = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else Path(__file__).resolve().parent / "raw"
    warmup = int(sys.argv[2]) if len(sys.argv) > 2 else 3
    samples = int(sys.argv[3]) if len(sys.argv) > 3 else 30
    expected_processes = int(sys.argv[4]) if len(sys.argv) > 4 else None
    paths = sorted(
        path
        for path in run.glob("*.json")
        if path.name not in {"matrix-summary.json"}
    )
    if not paths:
        raise SystemExit(f"no raw reports in {run}")

    reports = [(path, verify_report(path, warmup, samples)) for path in paths]
    groups: dict[tuple[str, str], list[dict[str, object]]] = defaultdict(list)
    identities: dict[str, tuple[object, ...]] = {}
    for path, report in reports:
        fixture = report["fixture"]
        key = (fixture, report["case"])
        groups[key].append(report)
        identity = tuple(report[name] for name in IDENTITY_KEYS)
        if fixture in identities and identities[fixture] != identity:
            raise SystemExit(f"fixture identity changed in {path.name}")
        identities.setdefault(fixture, identity)

    for fixture in identities:
        cases = {case for name, case in groups if name == fixture}
        if cases != CASES:
            raise SystemExit(f"fixture {fixture!r} does not have exactly six cases")

    summary_groups: dict[str, object] = {}
    for (fixture, case), group in sorted(groups.items()):
        if expected_processes is not None and len(group) != expected_processes:
            raise SystemExit(
                f"{fixture}/{case} has {len(group)} reports; expected {expected_processes}"
            )
        signatures = [
            [sample_signature(sample) for sample in report["samples"]]
            for report in group
        ]
        if len({json.dumps(signature, sort_keys=True) for signature in signatures}) != 1:
            raise SystemExit(f"cross-process output/source observations differ for {fixture}/{case}")
        if case.endswith("_warm"):
            setup_reads = {
                json.dumps(report["setup_reads"], sort_keys=True) for report in group
            }
            if len(setup_reads) != 1:
                raise SystemExit(f"cross-process setup read observations differ for {fixture}/{case}")
        summary_groups[f"{fixture}/{case}"] = {
            "fixture": fixture,
            "case": case,
            "process_count": len(group),
            "timing_scopes": sorted({report["timing_scope"] for report in group}),
            "process_p50_ns": [report["p50_ns"] for report in group],
            "process_p95_ns": [report["p95_ns"] for report in group],
            "process_p99_ns": [report["p99_ns"] for report in group],
            "process_mean_ns": [report["mean_ns"] for report in group],
            "median_process_p50_ns": statistics.median(report["p50_ns"] for report in group),
            "median_process_p95_ns": statistics.median(report["p95_ns"] for report in group),
            "median_process_p99_ns": statistics.median(report["p99_ns"] for report in group),
            "source_observation": group[0]["samples"][0]["reads"],
            "setup_reads": [report.get("setup_reads") for report in group],
            "setup_allocation": [report.get("setup_allocation") for report in group],
        }

    output = run.parent / "matrix-summary.json"
    summary = {
        "schema": "xlsb-theme-profile-matrix-v1",
        "raw_run_directory": str(run),
        "contract": {
            "fixtures": len(identities),
            "cases_per_fixture": len(CASES),
            "process_reports": len(reports),
            "warmup_per_report": warmup,
            "samples_per_report": samples,
        },
        "comparison_scope": "No equivalent-work speedup inferred; source-backed, eager, and transaction API scopes differ.",
        "groups": summary_groups,
    }
    output.write_text(json.dumps(summary, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(f"verified_reports={len(reports)}")
    print(f"verified_fixtures={len(identities)}")
    print(f"summary={output}")


if __name__ == "__main__":
    main()
