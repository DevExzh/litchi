#!/usr/bin/env python3
"""Verify and summarize the raw XLSB projection matrix without comparing lanes."""

from __future__ import annotations

import hashlib
import json
import statistics
import sys
from collections import defaultdict
from pathlib import Path


ROOT = Path(__file__).resolve().parents[5]
RUN = Path(sys.argv[1]).resolve() if len(sys.argv) == 2 else Path(
    __file__).resolve().parent / "runs/final-release"
EXPECTED_BACKENDS = {
    "owned": {
        "open_identify",
        "worksheet_catalog",
        "selected_worksheet_cell",
        "full_stored_cell_scan",
        "full_text",
        "noop_transaction_commit_save",
        "edit_one_existing_scalar_save",
        "edit_ceil_one_percent_existing_cells_save",
    },
    "owned_direct": {
        "open_identify",
        "worksheet_catalog",
        "selected_worksheet_cell",
        "full_stored_cell_scan",
        "noop_transaction_commit_save",
        "edit_one_existing_scalar_save",
        "edit_ceil_one_percent_existing_cells_save",
    },
    "owned_without_drawings": {
        "open_identify",
        "worksheet_catalog",
        "selected_worksheet_cell",
        "full_stored_cell_scan",
        "noop_transaction_commit_save",
        "edit_one_existing_scalar_save",
        "edit_ceil_one_percent_existing_cells_save",
    },
    "source_backed": {
        "open_identify",
        "worksheet_catalog",
        "selected_worksheet_cell",
        "full_stored_cell_scan",
    },
}
EXPECTED_GROUPS = {
    (backend, case)
    for backend, cases in EXPECTED_BACKENDS.items()
    for case in cases
}
GATE_KEYS = (
    "representative_output_reopen_ok",
    "semantic_readback_ok",
    "exact_noop_patch",
    "output_matches_across_samples",
    "unchanged_parts_ok",
    "malformed_input_refused",
    "tight_limits_refused",
    "tight_cell_limits_refused",
    "fixture_cell_count_within_dimensions",
)


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def median3(values: list[int]) -> int:
    return int(statistics.median(values))


def nearest_rank(samples: list[int], percent: int) -> int:
    ordered = sorted(samples)
    index = ((len(ordered) * percent) + 99) // 100 - 1
    return ordered[max(0, min(index, len(ordered) - 1))]


def check_source_intervals(observation: dict[str, object], path: Path) -> None:
    for total, open_value, operation in (
        ("source_read_calls", "open_source_read_calls", "operation_source_read_calls"),
        (
            "source_read_requested_bytes",
            "open_source_read_requested_bytes",
            "operation_source_read_requested_bytes",
        ),
        ("source_read_bytes", "open_source_read_bytes", "operation_source_read_bytes"),
        ("part_materializations", "open_part_materializations", "operation_part_materializations"),
        ("part_cache_hits", "open_part_cache_hits", "operation_part_cache_hits"),
    ):
        if observation[total] != observation[open_value] + observation[operation]:
            raise SystemExit(
                f"source counter interval does not sum in {path.name}: {total}"
            )


def main() -> None:
    paths = sorted(RUN.glob("*.json"))
    if len(paths) != 78:
        raise SystemExit(f"expected 78 raw JSON reports, found {len(paths)}")
    if (RUN / "completed.txt").read_text() != (
        "supported_backend_cases=26\nprocesses_per_case=3\ncompleted_invocations=78\n"
    ):
        raise SystemExit("completed.txt does not certify the expected matrix")

    groups: dict[tuple[str, str], list[dict[str, object]]] = defaultdict(list)
    binary_ids: set[str] = set()
    fixture_ids: set[tuple[str, int]] = set()
    for path in paths:
        report = json.loads(path.read_text())
        if report.get("schema") != "xlsb-crud-v2":
            raise SystemExit(f"unexpected schema in {path.name}")
        binary = report["binary_identity"]
        binary_ids.add(binary["binary_sha256"])
        corpus = report["corpus"]
        fixture_ids.add((corpus["source_sha256"], corpus["input_bytes"]))
        case_report = report["cases"]
        if len(case_report) != 1:
            raise SystemExit(f"{path.name} should contain one case")
        case = case_report[0]
        key = (case["backend"], case["case"])
        if key not in EXPECTED_GROUPS:
            raise SystemExit(f"unexpected backend/case {key} in {path.name}")
        stats = case["statistics"]
        if stats["warmup"] != 3 or stats["samples"] != 30 or len(stats["samples_ns"]) != 30:
            raise SystemExit(f"wrong sample contract in {path.name}")
        recomputed = {
            "p50_ns": nearest_rank(stats["samples_ns"], 50),
            "p95_ns": nearest_rank(stats["samples_ns"], 95),
            "p99_ns": nearest_rank(stats["samples_ns"], 99),
            "mean_ns": sum(stats["samples_ns"]) / len(stats["samples_ns"]),
        }
        for name, expected in recomputed.items():
            if stats[name] != expected:
                raise SystemExit(f"{name} is not reproducible from raw samples in {path.name}")
        required_true = {
            "semantic_readback_ok",
            "malformed_input_refused",
            "tight_limits_refused",
            "fixture_cell_count_within_dimensions",
        }
        if case["backend"] != "source_backed":
            required_true.add("tight_cell_limits_refused")
        if case["case"] in {
            "noop_transaction_commit_save",
            "edit_one_existing_scalar_save",
            "edit_ceil_one_percent_existing_cells_save",
        }:
            required_true.update(
                {
                    "representative_output_reopen_ok",
                    "output_matches_across_samples",
                    "unchanged_parts_ok",
                }
            )
        if case["case"] == "noop_transaction_commit_save":
            required_true.add("exact_noop_patch")
        for gate in required_true:
            if case["gates"].get(gate) is not True:
                raise SystemExit(f"required gate {gate} is not true in {path.name}")
        # Source-backed mode is read-only and deliberately does not run the
        # eager cell-limit refusal probe; false here is the expected scope
        # marker. Other optional gates are None for projection-only cases.
        if case["backend"] == "source_backed":
            if case["gates"].get("tight_cell_limits_refused") is not False:
                raise SystemExit(f"source cell-limit scope marker changed in {path.name}")
        if case["backend"] == "source_backed":
            if case["source_observation"] is None or case["source_observation_stable"] is not True:
                raise SystemExit(f"source counter stability failed in {path.name}")
            check_source_intervals(case["source_observation"], path)
        groups[key].append(case)

    if set(groups) != EXPECTED_GROUPS:
        missing = sorted(EXPECTED_GROUPS - set(groups))
        extra = sorted(set(groups) - EXPECTED_GROUPS)
        raise SystemExit(f"matrix groups mismatch; missing={missing}, extra={extra}")
    if any(len(reports) != 3 for reports in groups.values()):
        raise SystemExit("every backend/case group must have three processes")
    if len(binary_ids) != 1 or len(fixture_ids) != 1:
        raise SystemExit("binary or fixture identity changed across raw reports")
    if any(path.stat().st_size != 0 for path in RUN.glob("*.log")):
        # Logs are retained for exact reproduction; nonempty logs are allowed,
        # but record whether they contain a diagnostic-looking failure marker.
        for path in RUN.glob("*.log"):
            text = path.read_text(errors="replace").lower()
            if "panicked at" in text or "error:" in text:
                raise SystemExit(f"failure marker in {path.name}")

    summary_groups: dict[str, object] = {}
    for key in sorted(groups):
        backend, case = key
        reports = groups[key]
        raw_output_ids = {
            (report["output_bytes"], report["output_sha256"])
            for report in reports
            if report["output_sha256"] is not None
        }
        if len(raw_output_ids) > 1:
            raise SystemExit(f"saved output identity changed across processes for {key}")
        if backend == "source_backed":
            source_ids = {
                json.dumps(report["source_observation"], sort_keys=True)
                for report in reports
            }
            if len(source_ids) != 1:
                raise SystemExit(f"source observations changed across processes for {key}")
        source_observations = [
            report["source_observation"]
            for report in reports
            if report["source_observation"] is not None
        ]
        summary_groups[f"{backend}/{case}"] = {
            "backend": backend,
            "case": case,
            "timing_scopes": sorted({report["timing_scope"] for report in reports}),
            "process_count": len(reports),
            "process_p50_ns": [report["statistics"]["p50_ns"] for report in reports],
            "process_p95_ns": [report["statistics"]["p95_ns"] for report in reports],
            "process_p99_ns": [report["statistics"]["p99_ns"] for report in reports],
            "process_mean_ns": [report["statistics"]["mean_ns"] for report in reports],
            "median_process_p50_ns": median3(
                [report["statistics"]["p50_ns"] for report in reports]
            ),
            "median_process_p95_ns": median3(
                [report["statistics"]["p95_ns"] for report in reports]
            ),
            "median_process_p99_ns": median3(
                [report["statistics"]["p99_ns"] for report in reports]
            ),
            "output_sha256": sorted(
                {
                    report["output_sha256"]
                    for report in reports
                    if report["output_sha256"] is not None
                }
            ),
            "source_observations": source_observations,
        }

    result = {
        "schema": "litchi-xlsb-drawing-projection-matrix-v1",
        "raw_run_directory": str(RUN),
        "binary_sha256": next(iter(binary_ids)),
        "fixture": {
            "sha256": next(iter(fixture_ids))[0],
            "bytes": next(iter(fixture_ids))[1],
        },
        "contract": {
            "supported_backend_cases": 26,
            "processes_per_backend_case": 3,
            "warmup_per_process": 3,
            "samples_per_process": 30,
            "raw_json_reports": 78,
        },
        "cross_lane_comparison": "No equivalent-work speedup inferred; backend API/validation scopes differ.",
        "expected_unavailable_gates": {
            "source_backed/tight_cell_limits_refused": (
                "source-backed is read-only and does not run the eager cell-limit probe"
            ),
        },
        "groups": summary_groups,
    }
    output = RUN.parent.parent / "matrix-summary.json"
    output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    print(f"verified_groups={len(summary_groups)}")
    print("verified_reports=78")
    print(f"binary_sha256={next(iter(binary_ids))}")
    print(f"summary={output}")


if __name__ == "__main__":
    main()
