#!/usr/bin/env python3
"""Audit raw phase shares in the retained 0923 candidate reports.

This is an offline audit of an already-frozen packet.  It never starts a
benchmark, changes a source tree, or treats phase shares as an admission gate.
The report vectors are checked before they are summed so a copied or partial
packet cannot silently produce an Amdahl estimate.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


PHASES = (
    "preparation_ns",
    "open_ns",
    "query_ns",
    "edit_ns",
    "commit_ns",
    "publication_ns",
)
ROUTES = ("generic_read_at", "owned_vec", "owned_reader")
CASE = "docx_source_backed_existing_section_layout_edit_save"
MANIFEST_SHA256 = "b3089f7dcd5cc1517d22615a066e16088a50b4edbd4e703f7594bc5e66a568ae"
SUMMARY_SHA256 = "abbb158c87104a9f5cad85165d9a578984c23c0035722d491c1a9bb3319cc54f"
RAW_SAMPLE_SUMMARY_SHA256 = "edac04901aa00de5d726aef7229daca96ce58fb9441b13f093d50f26a9e43b7b"
EXPECTED_OUTPUT_SHA256 = "5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec"
SAMPLES_PER_REPORT = 200
REPORTS_PER_ROUTE = 8


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def read_json(path: Path) -> Any:
    with path.open(encoding="utf-8") as stream:
        return json.load(stream)


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def git_revision(repo: Path) -> str | None:
    try:
        return subprocess.check_output(
            ["git", "-C", str(repo), "rev-parse", "HEAD"],
            text=True,
            stderr=subprocess.DEVNULL,
        ).strip()
    except (OSError, subprocess.CalledProcessError):
        return None


def amdahl_bound(phase_shares: dict[str, float], eliminated: list[str]) -> dict[str, Any]:
    fraction = sum(phase_shares[name] for name in eliminated)
    remaining = 1.0 - fraction
    return {
        "eliminated_phases": eliminated,
        "eliminated_fraction": fraction,
        "remaining_fraction": remaining,
        "maximum_speedup_if_eliminated": (1.0 / remaining) if remaining > 0 else None,
        "interpretation": "descriptive bound from summed raw samples, not a timing gate",
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--repo",
        type=Path,
        default=Path("/home/zhuhe/code/litchi"),
        help="checkout containing the retained 0923 packet",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("/tmp/litchi-performance-0924/phase-share-v2.json"),
    )
    args = parser.parse_args()
    repo = args.repo.resolve()
    manifest_path = repo / "docs/performance/results/change-0923/capture-final/packet/comparison-manifest-v2.json"
    summary_path = repo / "docs/performance/results/change-0923/comparison-result-summary-v1.json"
    raw_sample_summary_path = repo / "docs/performance/results/change-0923/root-raw-sample-summary.json"
    report_root = repo / "docs/performance/results/change-0923/capture-final/packet/children"

    require(manifest_path.is_file(), f"missing manifest: {manifest_path}")
    require(summary_path.is_file(), f"missing comparison summary: {summary_path}")
    require(raw_sample_summary_path.is_file(), f"missing raw sample summary: {raw_sample_summary_path}")
    manifest_sha = sha256_file(manifest_path)
    summary_sha = sha256_file(summary_path)
    raw_sample_summary_sha = sha256_file(raw_sample_summary_path)
    require(manifest_sha == MANIFEST_SHA256, f"manifest hash changed: {manifest_sha}")
    require(summary_sha == SUMMARY_SHA256, f"summary hash changed: {summary_sha}")
    require(
        raw_sample_summary_sha == RAW_SAMPLE_SUMMARY_SHA256,
        f"raw sample summary hash changed: {raw_sample_summary_sha}",
    )

    manifest = read_json(manifest_path)
    summary = read_json(summary_path)
    raw_sample_summary = read_json(raw_sample_summary_path)
    require(manifest.get("status") == "frozen", "manifest is not frozen")
    require(manifest.get("case") == CASE, "manifest case mismatch")
    native = manifest.get("native", {})
    require(native.get("samples_per_child") == SAMPLES_PER_REPORT, "sample count mismatch")
    require(native.get("warmups_per_child") == 30, "warmup count mismatch")
    schedule = native.get("schedule")
    require(isinstance(schedule, list) and len(schedule) == 48, "schedule must contain 48 children")
    schedule_by_id = {entry.get("id"): entry for entry in schedule}
    require(len(schedule_by_id) == 48 and None not in schedule_by_id, "schedule IDs are not unique")
    require(set(manifest.get("routes", [])) == set(ROUTES), "manifest route set mismatch")
    candidate_schedule = {
        entry["id"]: entry for entry in schedule if entry.get("variant") == "candidate"
    }
    require(len(candidate_schedule) == 24, "candidate schedule must contain 24 children")
    candidate_counts = Counter(entry["route"] for entry in candidate_schedule.values())
    require(candidate_counts == Counter({route: REPORTS_PER_ROUTE for route in ROUTES}), "candidate route counts mismatch")

    raw_report_index: dict[str, dict[str, Any]] = {}
    for entry in raw_sample_summary.get("reports", []):
        report_path = Path(entry.get("path", ""))
        report_id = report_path.parent.name
        require(report_id not in raw_report_index, f"duplicate raw report index entry: {report_id}")
        raw_report_index[report_id] = {
            "bytes": entry.get("bytes"),
            "sha256": entry.get("sha256"),
        }
    require(len(raw_report_index) == 48, "raw report index must contain 48 reports")
    require(set(candidate_schedule).issubset(raw_report_index), "candidate reports missing from raw report index")

    expected_binary = manifest["binaries"]["candidate"]["binary"]["sha256"]
    expected_binary_bytes = manifest["binaries"]["candidate"]["binary"]["bytes"]
    report_paths = sorted(report_root.glob("*-candidate/report.json"))
    require(len(report_paths) == 24, f"expected 24 candidate reports, found {len(report_paths)}")
    require({path.parent.name for path in report_paths} == set(candidate_schedule), "report IDs differ from schedule")

    route_data: dict[str, dict[str, Any]] = {
        route: {
            "report_count": 0,
            "sample_count": 0,
            "phase_sums_ns": {phase: 0 for phase in PHASES},
            "phase_sum_mismatches": 0,
            "invalid_sample_orders": 0,
            "source_identities": set(),
            "output_identities": set(),
            "report_files": [],
        }
        for route in ROUTES
    }
    all_phase_sums = {phase: 0 for phase in PHASES}
    all_report_hashes: list[dict[str, Any]] = []
    source_identities: set[tuple[int, str]] = set()
    output_identities: set[tuple[int, str]] = set()

    for report_path in report_paths:
        report_id = report_path.parent.name
        schedule_entry = candidate_schedule[report_id]
        report = read_json(report_path)
        expected_report = raw_report_index[report_id]
        actual_report_sha = sha256_file(report_path)
        require(report_path.stat().st_size == expected_report["bytes"], f"{report_id}: report size mismatch")
        require(actual_report_sha == expected_report["sha256"], f"{report_id}: report hash mismatch")
        identity = report.get("binary_identity", {})
        require(identity.get("binary_sha256") == expected_binary, f"{report_id}: binary hash mismatch")
        require(identity.get("binary_bytes") == expected_binary_bytes, f"{report_id}: binary size mismatch")
        results = report.get("results")
        require(isinstance(results, list) and len(results) == 1, f"{report_id}: result cardinality mismatch")
        result = results[0]
        require(result.get("case") == CASE, f"{report_id}: case mismatch")
        require(result.get("output_sha256") == EXPECTED_OUTPUT_SHA256, f"{report_id}: output digest mismatch")
        detail = result.get("source", {}).get("docx_section_layout", {})
        require(
            detail.get("expected_output_sha256") == EXPECTED_OUTPUT_SHA256,
            f"{report_id}: expected output digest mismatch",
        )
        route = detail.get("ingress_route")
        require(route == schedule_entry.get("route") and route in ROUTES, f"{report_id}: route mismatch")
        source_bytes = detail.get("source_bytes")
        source_sha = detail.get("source_sha256")
        require(isinstance(source_bytes, int) and isinstance(source_sha, str), f"{report_id}: source identity missing")
        source_identities.add((source_bytes, source_sha))
        route_data[route]["source_identities"].add((source_bytes, source_sha))

        output_sha = detail.get("output_sha256")
        if isinstance(output_sha, list):
            require(len(output_sha) == SAMPLES_PER_REPORT, f"{report_id}: output vector length mismatch")
            require(all(value == EXPECTED_OUTPUT_SHA256 for value in output_sha), f"{report_id}: output vector digest mismatch")
            output_values = set(output_sha)
        else:
            output_values = {output_sha}
        output_identities.update((source_bytes, value) for value in output_values if isinstance(value, str))
        route_data[route]["output_identities"].update(
            (source_bytes, value) for value in output_values if isinstance(value, str)
        )

        elapsed = result.get("elapsed_ns", {}).get("samples")
        elapsed_order = result.get("elapsed_ns", {}).get("sample_order")
        require(isinstance(elapsed, list) and len(elapsed) == SAMPLES_PER_REPORT, f"{report_id}: elapsed vector length mismatch")
        require(
            isinstance(elapsed_order, list) and len(elapsed_order) == SAMPLES_PER_REPORT,
            f"{report_id}: elapsed sample order missing",
        )
        phase_vectors = {phase: detail.get(phase) for phase in PHASES}
        for phase, values in phase_vectors.items():
            require(isinstance(values, list) and len(values) == SAMPLES_PER_REPORT, f"{report_id}: {phase} length mismatch")
            require(all(isinstance(value, int) and value >= 0 for value in values), f"{report_id}: invalid {phase} value")
        require(sorted(elapsed_order) == list(range(SAMPLES_PER_REPORT)), f"{report_id}: elapsed order is not a permutation")
        detail_order = detail.get("sample_order")
        require(detail_order == elapsed_order, f"{report_id}: phase and elapsed sample orders differ")
        require(sorted(detail_order) == list(range(SAMPLES_PER_REPORT)), f"{report_id}: phase order is not a permutation")

        mismatch_count = 0
        for index in range(SAMPLES_PER_REPORT):
            phase_total = sum(phase_vectors[phase][index] for phase in PHASES)
            if phase_total != elapsed[index]:
                mismatch_count += 1
        route_data[route]["phase_sum_mismatches"] += mismatch_count
        route_data[route]["invalid_sample_orders"] += int(sorted(detail_order) != list(range(SAMPLES_PER_REPORT)))
        route_data[route]["report_count"] += 1
        route_data[route]["sample_count"] += SAMPLES_PER_REPORT
        for phase in PHASES:
            phase_sum = sum(phase_vectors[phase])
            route_data[route]["phase_sums_ns"][phase] += phase_sum
            all_phase_sums[phase] += phase_sum
        all_report_hashes.append(
            {
                "id": report_id,
                "route": route,
                "variant": schedule_entry.get("variant"),
                "bytes": report_path.stat().st_size,
                "sha256": actual_report_sha,
                "expected_bytes": expected_report["bytes"],
                "expected_sha256": expected_report["sha256"],
            }
        )

    require(source_identities == {(2_101_427, "183204a832a3518bf1d4474e9b7ff347ad7d0dc95ad2e9b1dbdde405837a29e8")}, "source identity drift")
    for route in ROUTES:
        require(route_data[route]["report_count"] == REPORTS_PER_ROUTE, f"{route}: report count mismatch")
        require(route_data[route]["sample_count"] == REPORTS_PER_ROUTE * SAMPLES_PER_REPORT, f"{route}: sample count mismatch")
        require(route_data[route]["phase_sum_mismatches"] == 0, f"{route}: phase sum mismatch")
        require(route_data[route]["invalid_sample_orders"] == 0, f"{route}: invalid sample order")

    def finalize_route(data: dict[str, Any]) -> dict[str, Any]:
        phase_sums = data["phase_sums_ns"]
        total = sum(phase_sums.values())
        shares = {phase: phase_sums[phase] / total for phase in PHASES}
        return {
            "report_count": data["report_count"],
            "sample_count": data["sample_count"],
            "phase_sums_ns": phase_sums,
            "total_elapsed_ns": total,
            "phase_shares": shares,
            "query_plus_publication_share": shares["query_ns"] + shares["publication_ns"],
            "amdahl_bounds": {
                "open_only": amdahl_bound(shares, ["open_ns"]),
                "publication_only": amdahl_bound(shares, ["publication_ns"]),
                "query_plus_publication": amdahl_bound(shares, ["query_ns", "publication_ns"]),
            },
            "phase_sum_mismatches": data["phase_sum_mismatches"],
            "invalid_sample_orders": data["invalid_sample_orders"],
            "source_identities": [
                {"bytes": size, "sha256": digest}
                for size, digest in sorted(data["source_identities"])
            ],
            "output_identities": [
                {"source_bytes": size, "sha256": digest}
                for size, digest in sorted(data["output_identities"])
            ],
            "report_files": data["report_files"],
        }

    routes_out: dict[str, Any] = {}
    for route in ROUTES:
        route_out = finalize_route(route_data[route])
        route_summary = summary.get("routes", {}).get(route, {})
        timing_gate = route_summary.get("timing_gate", {})
        route_out["paired_median_timing_gate"] = {
            "median_ratio": timing_gate.get("median_ratio"),
            "lower_95": timing_gate.get("lower_95"),
            "eligible": timing_gate.get("eligible"),
            "note": "paired p50 ratio from retained summary; separate from summed raw phase shares",
        }
        routes_out[route] = route_out

    total_elapsed = sum(all_phase_sums.values())
    all_shares = {phase: all_phase_sums[phase] / total_elapsed for phase in PHASES}
    for route in ROUTES:
        route_data[route]["report_files"] = [entry for entry in all_report_hashes if entry["route"] == route]
    for route in ROUTES:
        routes_out[route]["report_files"] = route_data[route]["report_files"]

    output = {
        "schema": "litchi.performance.0924.raw-phase-share-audit.v2",
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "purpose": "offline phase-share and descriptive Amdahl-bound audit of retained 0923 candidate raw samples",
        "capture_performed": False,
        "source_mutation_performed": False,
        "repo_git_revision_at_audit": git_revision(repo),
        "inputs": {
            "manifest": {"path": str(manifest_path), "bytes": manifest_path.stat().st_size, "sha256": manifest_sha},
            "comparison_summary": {"path": str(summary_path), "bytes": summary_path.stat().st_size, "sha256": summary_sha},
            "raw_sample_summary": {
                "path": str(raw_sample_summary_path),
                "bytes": raw_sample_summary_path.stat().st_size,
                "sha256": raw_sample_summary_sha,
                "report_index_entries": len(raw_report_index),
            },
            "case": CASE,
            "routes": list(ROUTES),
            "candidate_reports": 24,
            "reports_per_route": REPORTS_PER_ROUTE,
            "samples_per_report": SAMPLES_PER_REPORT,
            "source_identity": {"bytes": 2_101_427, "sha256": "183204a832a3518bf1d4474e9b7ff347ad7d0dc95ad2e9b1dbdde405837a29e8"},
            "candidate_binary": {"bytes": expected_binary_bytes, "sha256": expected_binary},
        },
        "validation": {
            "manifest_status": manifest["status"],
            "schedule_children": len(schedule),
            "candidate_schedule_children": len(candidate_schedule),
            "candidate_route_counts": dict(sorted(candidate_counts.items())),
            "report_hashes": sorted(all_report_hashes, key=lambda item: item["id"]),
            "all_candidate_report_hashes_match_root_raw_sample_summary": True,
            "all_candidate_report_sizes_match_root_raw_sample_summary": True,
            "all_candidate_reports_have_exact_binary_identity": True,
            "all_candidate_reports_have_exact_case_and_route": True,
            "all_phase_vectors_have_200_samples": True,
            "all_sample_orders_are_permutations": True,
            "all_phase_sums_equal_elapsed_samples": True,
        },
        "all_routes_sum_of_raw_samples": {
            "sample_count": sum(routes_out[route]["sample_count"] for route in ROUTES),
            "phase_sums_ns": all_phase_sums,
            "total_elapsed_ns": total_elapsed,
            "phase_shares": all_shares,
            "amdahl_bounds": {
                "open_only": amdahl_bound(all_shares, ["open_ns"]),
                "publication_only": amdahl_bound(all_shares, ["publication_ns"]),
                "query_plus_publication": amdahl_bound(all_shares, ["query_ns", "publication_ns"]),
            },
        },
        "routes": routes_out,
        "interpretation": {
            "raw_share_definition": "sum of each phase's 1600 raw samples divided by the sum of all six phases' 1600 raw samples for that route",
            "amdahl_definition": "counterfactual descriptive ceiling if the listed phase(s) consumed zero time while all other raw samples stayed fixed",
            "median_distinction": "paired_median_timing_gate values are p50 baseline/candidate ratios and are not phase shares or pooled raw-sample quantiles",
            "reader_ingress_limit": "owned_reader open is measurable but bounded by its open-only share; query and publication dominate the retained route's summed time",
            "no_adoption_claim": True,
        },
        "regeneration": {
            "command": "python3 -B /tmp/litchi-performance-0924/phase-share-v2.py --repo /home/zhuhe/code/litchi --output /tmp/litchi-performance-0924/phase-share-v2.json",
            "script_path": "/tmp/litchi-performance-0924/phase-share-v2.py",
        },
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open("w", encoding="utf-8") as stream:
        json.dump(output, stream, indent=2, sort_keys=True)
        stream.write("\n")
    print(json.dumps({"output": str(args.output), "bytes": args.output.stat().st_size, "sha256": sha256_file(args.output)}, sort_keys=True))
    for route, value in routes_out.items():
        print(json.dumps({"route": route, "phase_shares": value["phase_shares"], "amdahl_open": value["amdahl_bounds"]["open_only"]["maximum_speedup_if_eliminated"]}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
