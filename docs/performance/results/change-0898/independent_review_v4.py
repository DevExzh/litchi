#!/usr/bin/env python3
"""Independently replay the sealed 0898 native capture.

This file only reopens a frozen plan, the completed ``capture-v1`` tree, and
the reducer/verifier receipts.  It never launches a child, invokes Cargo, or
changes a source/build tree.  Statistics are recomputed from every retained
raw vector and paired child artifact; the retained reducer is used only for a
cross-check after the independent values have been produced.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import math
from pathlib import Path
import random
import re
import statistics
import stat
from typing import Any


PACKET = Path(__file__).resolve().parent
CAPTURE_DEFAULT = PACKET / "capture-v1"
PLAN_DEFAULT = PACKET / "plan.json"
REDUCTION_DEFAULT = PACKET / "reduction.json"
VERIFICATION_DEFAULT = PACKET / "verification.json"
OUTPUT_DEFAULT = PACKET / "native-review.json"

SCHEMA = "litchi.mce-codec-emitted-name-output-span-native-review-0898.v1"
PLAN_SCHEMA = "litchi.xlsx-mce-codec-emitted-name-output-span-experiment-0898.v1"
EXPERIMENT_SCHEMA = "litchi.xlsx-mce-codec-emitted-name-output-span-experiment-receipt-0898.v1"
COMMAND_SCHEMA = "litchi.xlsx-mce-codec-emitted-name-output-span-child-command-0898.v1"
RECEIPT_SCHEMA = "litchi.xlsx-mce-codec-emitted-name-output-span-child-receipt-0898.v1"
VERIFICATION_SCHEMA = "litchi.xlsx-mce-codec-emitted-name-output-span-verification-0898.v1"
REDUCTION_SCHEMA = "litchi.xlsx-mce-codec-emitted-name-output-span-reduction-0898.v1"

BASELINE_COMMIT = "506aa2edae6666d9c76ae6827c2f5941a8e47896"
EVALUATION_HEAD = "8d517a7861163f04663dae6481e6123b6a314b3c"
GUARD_CONTRACT = "post-endpoint-per-worksheet-tolerance-v1"
PROFILE = "release"
CPU_AFFINITY = "12"
CHILD_ENV = {"PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC"}
CASES = (
    "xlsx_producer_medium_source_full_cell_scan",
    "xlsx_full_cell_scan",
    "docx_producer_source_selected_paragraph",
    "pptx_producer_source_selected_slide",
)
PRIMARY_CASE = CASES[0]
NATIVE_BLOCKS = 8
NATIVE_WARMUP = 3
NATIVE_SAMPLES = 50
ALLOCATOR_BLOCKS = 4
ALLOCATOR_WARMUP = 3
ALLOCATOR_SAMPLES = 15
TOTAL_NATIVE_CHILDREN = NATIVE_BLOCKS * len(CASES) * 2
TOTAL_ALLOCATOR_CHILDREN = ALLOCATOR_BLOCKS * 2
TOTAL_CHILDREN = TOTAL_NATIVE_CHILDREN + TOTAL_ALLOCATOR_CHILDREN
TOTAL_SAMPLES = TOTAL_NATIVE_CHILDREN * NATIVE_SAMPLES + TOTAL_ALLOCATOR_CHILDREN * ALLOCATOR_SAMPLES
BOOTSTRAP_SEED = 898
BOOTSTRAP_DRAWS = 10_000
BOOTSTRAP_LOW_INDEX = 249
BOOTSTRAP_HIGH_INDEX = 9_749
PAIR_RATIO_THRESHOLD = 1.05
P50_SPREAD_THRESHOLD = 1.20
PRIMARY_ARCHIVE_SHA256 = "55e0901c093c27d80bcd4e5fef888f00798654cdef8c0d79d130a2ce7cbe0ae8"
PRIMARY_ARCHIVE_BYTES = 22_739
PRIMARY_DIGEST_SHA256 = "8d29b40718aa9d5b08c746f841195bba17ab8e01608b12e6093daf160c00764b"
PRIMARY_CENSUS_SHA256 = "552c35c8b8f958de28792f00c1e9e831ae884471aa28b5b2a7a1479bd06c5207"
BUILD_CONFIGURATION = {
    "toolchain": "1.95.0",
    "profile": "release",
    "opt_level": 3,
    "debug": 1,
    "lto": "fat",
    "codegen_units": 1,
    "jobs": 2,
}
REQUIRED_BINDINGS = {
    "source_manifest",
    "source_identity",
    "candidate_freeze",
    "baseline_build",
    "candidate_build",
    "baseline_quality",
    "candidate_quality",
    "guard_reference",
}
SOURCE_OVERLAY = {
    "crates/litchi-ooxml-common/src/mce/codec.rs",
    "crates/litchi-ooxml-common/src/mce/tests.rs",
}
ALLOCATION_METRICS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "allocated_bytes",
    "region_peak_live_bytes",
)
PRIMARY_SCAN_DYNAMIC = {
    "sample_indices",
    "observed_cell_counts",
    "observed_numeric_counts",
    "observed_text_counts",
    "observed_digest_sha256",
    "source_version_stable",
    "forced_worksheet_counts",
    "forced_worksheet_tolerance_clean",
    "warmup_forced_worksheet_counts",
    "warmup_forced_worksheet_tolerance_clean",
    "forced_worksheet_guard_iterations",
    "timing_scope",
}


class ReviewError(Exception):
    """A custody, guard, or arithmetic replay failure."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ReviewError(message)


def read(path: Path) -> Any:
    with Path(path).open(encoding="utf-8") as handle:
        return json.load(handle)


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    path = Path(path)
    require(path.is_file() and not path.is_symlink(), f"expected regular non-symlink file: {path}")
    path = path.resolve()
    info = path.stat()
    require(stat.S_ISREG(info.st_mode), f"expected regular file: {path}")
    return {
        "path": str(path),
        "bytes": info.st_size,
        "sha256": sha(path),
        "mode": stat.S_IMODE(info.st_mode),
        "executable": bool(info.st_mode & stat.S_IXUSR),
    }


def same_descriptor(record: Any, actual: dict[str, Any], label: str, *, path: bool = True) -> None:
    require(isinstance(record, dict), f"{label}: descriptor missing")
    fields = ("path", "bytes", "sha256", "mode", "executable") if path else ("bytes", "sha256", "mode", "executable")
    for field in fields:
        if field in record:
            require(record[field] == actual[field], f"{label}: {field} mismatch")


def path_from_record(record: Any, base: Path, label: str) -> Path:
    require(isinstance(record, dict) and isinstance(record.get("path"), str), f"{label}: path missing")
    path = Path(record["path"])
    if not path.is_absolute():
        path = base / path
    return path.resolve()


def check_descriptor(record: Any, path: Path, label: str) -> dict[str, Any]:
    actual = descriptor(path)
    same_descriptor(record, actual, label, path=False)
    return actual


def load_object(path: Path, label: str) -> dict[str, Any]:
    value = read(path)
    require(isinstance(value, dict), f"{label}: JSON root is not an object")
    return value


def floor_midpoint(left: int, right: int) -> int:
    return left // 2 + right // 2 + ((left % 2 + right % 2) // 2)


def nearest_rank(values: list[int], percentile: int) -> int:
    rank = (len(values) * percentile + 99) // 100
    return values[max(0, min(len(values) - 1, rank - 1))]


def raw_stats(values: Any, label: str) -> dict[str, int]:
    require(
        type(values) is list and values and all(type(value) is int and value > 0 for value in values),
        f"{label}: elapsed vector is invalid",
    )
    require(values == sorted(values), f"{label}: raw elapsed vector is not sorted")
    return {
        "min": values[0],
        "p50": floor_midpoint(values[(len(values) - 1) // 2], values[len(values) // 2]),
        "p95": nearest_rank(values, 95),
        "p99": nearest_rank(values, 99),
        "max": values[-1],
    }


def bootstrap(ratios: list[float]) -> dict[str, Any]:
    require(ratios and all(math.isfinite(value) and value > 0 for value in ratios), "cannot bootstrap empty/invalid ratios")
    rng = random.Random(BOOTSTRAP_SEED)
    medians = []
    for _ in range(BOOTSTRAP_DRAWS):
        draw = [ratios[rng.randrange(len(ratios))] for _ in ratios]
        medians.append(statistics.median(draw))
    ordered = sorted(medians)
    return {
        "seed": BOOTSTRAP_SEED,
        "draws": BOOTSTRAP_DRAWS,
        "low_index": BOOTSTRAP_LOW_INDEX,
        "high_index": BOOTSTRAP_HIGH_INDEX,
        "median_ratio": statistics.median(ratios),
        "ci95_low": ordered[BOOTSTRAP_LOW_INDEX],
        "ci95_high": ordered[BOOTSTRAP_HIGH_INDEX],
        "ratios": ratios,
    }


def expected_rows() -> list[dict[str, Any]]:
    rows = []
    ordinal = 0
    for block in range(NATIVE_BLOCKS):
        for case_index, case in enumerate(CASES):
            order = ("before", "after") if (block + case_index) % 2 == 0 else ("after", "before")
            pair = f"native-b{block:02d}-c{case_index:02d}"
            for role in order:
                rows.append({
                    "ordinal": ordinal,
                    "lane": "native",
                    "block": block,
                    "case_index": case_index,
                    "case": case,
                    "role": role,
                    "pair": pair,
                    "warmup": NATIVE_WARMUP,
                    "samples": NATIVE_SAMPLES,
                })
                ordinal += 1
    for block in range(ALLOCATOR_BLOCKS):
        order = ("before", "after") if block % 2 == 0 else ("after", "before")
        pair = f"allocator-b{block:02d}-c00"
        for role in order:
            rows.append({
                "ordinal": ordinal,
                "lane": "allocator",
                "block": block,
                "case_index": 0,
                "case": PRIMARY_CASE,
                "role": role,
                "pair": pair,
                "warmup": ALLOCATOR_WARMUP,
                "samples": ALLOCATOR_SAMPLES,
            })
            ordinal += 1
    require(len(rows) == TOTAL_CHILDREN, "schedule population mismatch")
    return rows


def expected_corpora() -> dict[str, dict[str, Any]]:
    root = PACKET.parent / "change-0887" / "baseline-qualification"
    result = {}
    for case in CASES:
        folder = root / case
        report_path = folder / "report.json"
        report = load_object(report_path, f"sealed corpus {case}")
        rows = report.get("results")
        require(type(rows) is list and len(rows) == 1 and rows[0].get("case") == case, f"sealed corpus result mismatch: {case}")
        producer_path = folder / "producer.json"
        producer = read(producer_path) if producer_path.is_file() else None
        if producer is not None:
            require(producer.get("schema") == "litchi.perf-baseline.producer-shape-evidence.v1", f"producer schema mismatch: {case}")
        if case in (PRIMARY_CASE, CASES[2], CASES[3]):
            require(producer is not None, f"required producer sidecar missing: {case}")
        result[case] = {
            "corpus": rows[0].get("corpus"),
            "report": report,
            "producer": producer,
        }
    return result


def validate_source_census(plan: dict[str, Any], bindings: dict[str, Any], candidate_source: dict[str, Any]) -> dict[str, Any]:
    source_identity = bindings["source_identity"]
    source_path = path_from_record(source_identity, PACKET, "source identity binding")
    source = load_object(source_path, "source identity")
    require(source.get("baseline_commit", source.get("baseline_revision")) == BASELINE_COMMIT, "source census baseline mismatch")
    require(source.get("evaluation_head", source.get("current_revision")) == EVALUATION_HEAD, "source census evaluation head mismatch")
    if "source_selection_unchanged" in source:
        require(source["source_selection_unchanged"] is True, "source selection changed")
    require(source.get("file_count", source.get("verified_files", 10_777)) >= 10_771, "source census is incomplete")
    if isinstance(source.get("files"), dict):
        require(len(source["files"]) >= 10_771, "source file map is incomplete")
    source_after = None
    source_after_path = None
    if "source_after_build" in bindings:
        after_path = path_from_record(bindings["source_after_build"], PACKET, "source after-build binding")
        source_after_path = after_path
        source_after = load_object(after_path, "source after-build census")
        require(source_after.get("status") == "pass", "source after-build census did not pass")
        if isinstance(source_after.get("candidate_source"), dict):
            same_descriptor(source_after["candidate_source"], descriptor(PACKET / "candidate-source.json"), "source after-build candidate source")
        if isinstance(source_after.get("files"), dict):
            require(len(source_after["files"]) >= 10_771, "source after-build file map is incomplete")
    overlay = candidate_source.get("overlay") or candidate_source.get("files")
    require(isinstance(overlay, dict) and set(overlay) == SOURCE_OVERLAY, "candidate source overlay scope changed")
    overlay_hashes = {}
    for relative, record in overlay.items():
        overlay_path = path_from_record(record, PACKET, f"candidate overlay {relative}")
        actual = descriptor(overlay_path)
        same_descriptor(record, actual, f"candidate overlay {relative}")
        overlay_hashes[relative] = actual
    return {
        "source": descriptor(source_path),
        "source_after_build": descriptor(source_after_path) if source_after_path is not None else None,
        "source_file_count": source.get("file_count", source.get("verified_files")),
        "source_map_count": len(source.get("files", {})) if isinstance(source.get("files"), dict) else None,
        "candidate_source": descriptor(PACKET / "candidate-source.json"),
        "overlay": overlay_hashes,
    }


def validate_bindings(plan: dict[str, Any]) -> dict[str, Any]:
    bindings = plan.get("bindings")
    require(isinstance(bindings, dict) and REQUIRED_BINDINGS <= set(bindings), "plan evidence bindings are incomplete")
    binding_descriptors = {}
    for name, record in bindings.items():
        path = path_from_record(record, PACKET, f"binding {name}")
        actual = descriptor(path)
        same_descriptor(record, actual, f"binding {name}")
        binding_descriptors[name] = actual
    candidate_source_path = path_from_record(bindings["candidate_freeze"], PACKET, "candidate freeze binding")
    candidate_source = load_object(candidate_source_path, "candidate source receipt")
    require(candidate_source.get("status") == "pass", "candidate source receipt does not pass")
    require(candidate_source.get("production_adopted", candidate_source.get("production_adoption")) is False, "candidate source is marked adopted")
    require(candidate_source.get("guard_contract") == GUARD_CONTRACT, "candidate source guard contract mismatch")
    source_binding = candidate_source.get("source")
    require(isinstance(source_binding, dict), "candidate source source binding missing")
    source_identity_value = load_object(path_from_record(bindings["source_identity"], PACKET, "source identity binding"), "source identity")
    allowed_source_bindings = [binding_descriptors["source_manifest"]]
    if isinstance(source_identity_value.get("source"), dict):
        source_record_path = path_from_record(source_identity_value["source"], PACKET, "source census manifest")
        allowed_source_bindings.append(descriptor(source_record_path))
    require(any(
        all(source_binding.get(field) == allowed.get(field) for field in ("path", "bytes", "sha256", "mode", "executable"))
        for allowed in allowed_source_bindings
    ), "candidate source source binding differs from the retained source census")
    quality_path = path_from_record(bindings["candidate_quality"], PACKET, "candidate quality binding")
    quality = load_object(quality_path, "candidate quality")
    require(quality.get("status") == "pass" and quality.get("gate_count") == 10, "candidate quality receipt is incomplete")
    if isinstance(quality.get("source"), dict):
        same_descriptor(quality["source"], binding_descriptors["candidate_freeze"], "candidate quality source")
    candidate_build_path = path_from_record(bindings["candidate_build"], PACKET, "candidate build binding")
    candidate_build = load_object(candidate_build_path, "candidate build")
    require(candidate_build.get("status") == "pass", "candidate build receipt does not pass")
    require(candidate_build.get("guard_contract") in (None, GUARD_CONTRACT), "candidate build guard contract mismatch")
    if isinstance(candidate_build.get("source"), dict):
        same_descriptor(candidate_build["source"], binding_descriptors["candidate_freeze"], "candidate build source")
    if isinstance(candidate_build.get("quality"), dict):
        same_descriptor(candidate_build["quality"], binding_descriptors["candidate_quality"], "candidate build quality")
    source_census = validate_source_census(plan, bindings, candidate_source)
    guard_path = path_from_record(bindings["guard_reference"], PACKET, "guard reference binding")
    guard = load_object(guard_path, "guard reference")
    guard_scan = guard.get("source", {}).get("producer_full_cell_scan")
    require(isinstance(guard_scan, dict), "guard reference scan is missing")
    review_bindings = {}
    for name in ("source_review", "candidate_review", "corpus_review", "corpus_comparison"):
        if name not in bindings:
            continue
        review_path = path_from_record(bindings[name], PACKET, f"{name} binding")
        review = load_object(review_path, name)
        require(review.get("status") in {"pass", "approved", "approved_for_measurement", "frozen"}, f"{name} receipt status")
        review_bindings[name] = descriptor(review_path)
    return {
        "descriptors": binding_descriptors,
        "candidate_source": candidate_source,
        "candidate_build": candidate_build,
        "quality": quality,
        "source_census": source_census,
        "guard_reference": descriptor(guard_path),
        "guard_scan": guard_scan,
        "review_bindings": review_bindings,
    }


def validate_plan(plan_path: Path, plan: dict[str, Any]) -> dict[str, Any]:
    require(plan.get("schema") == PLAN_SCHEMA, f"plan schema mismatch: {plan.get('schema')!r}")
    require(plan.get("status") == "frozen", "plan is not frozen")
    require(plan.get("base_commit") == BASELINE_COMMIT, "plan baseline commit mismatch")
    require(plan.get("evaluation_head") == EVALUATION_HEAD, "plan evaluation head mismatch")
    require(plan.get("profile") == PROFILE and str(plan.get("cpu_affinity")) == CPU_AFFINITY, "plan profile/CPU mismatch")
    require(plan.get("cache") == "warm" and plan.get("cases") == list(CASES) and plan.get("xlsx_shape") == "medium", "plan workload scope mismatch")
    require(plan.get("native_blocks") == NATIVE_BLOCKS and plan.get("native_warmup") == NATIVE_WARMUP and plan.get("native_samples") == NATIVE_SAMPLES, "plan native protocol mismatch")
    require(plan.get("allocator_blocks") == ALLOCATOR_BLOCKS and plan.get("allocator_warmup") == ALLOCATOR_WARMUP and plan.get("allocator_samples") == ALLOCATOR_SAMPLES, "plan allocator protocol mismatch")
    require(plan.get("total_children") == TOTAL_CHILDREN and plan.get("total_samples") == TOTAL_SAMPLES, "plan population mismatch")
    require(plan.get("bootstrap_seed") == BOOTSTRAP_SEED and plan.get("bootstrap_draws") == BOOTSTRAP_DRAWS and plan.get("bootstrap_ci_indices") == [BOOTSTRAP_LOW_INDEX, BOOTSTRAP_HIGH_INDEX], "plan bootstrap mismatch")
    require(plan.get("fresh_child") is True and plan.get("fail_fast") is True and plan.get("guard_contract") == GUARD_CONTRACT, "plan guard/fail-fast contract mismatch")
    require(plan.get("build_configuration") == BUILD_CONFIGURATION, "plan build configuration mismatch")
    require(plan.get("rows") == expected_rows(), "plan schedule mismatch")
    require(plan.get("candidate", {}).get("hypothesis") == "mce-codec-emitted-name-output-span-0898", "candidate hypothesis mismatch")
    require(set(plan.get("candidate", {}).get("paths", [])) == SOURCE_OVERLAY, "candidate source inventory mismatch")
    require(plan.get("candidate", {}).get("production_adoption") is False and plan.get("production_adoption") is False, "plan production adoption is enabled")
    require(plan.get("primary_gate") == {"median_ratio_at_most": 0.98, "ci95_upper_below": 1.0}, "primary gate changed")
    require(plan.get("review_flags") == {"native_p50_spread_above": P50_SPREAD_THRESHOLD, "pair_ratio_above": PAIR_RATIO_THRESHOLD, "rss_causal_claim": False}, "review thresholds changed")
    binding_review = validate_bindings(plan)
    binaries = plan.get("binaries")
    require(isinstance(binaries, dict) and set(binaries) == {"before", "after"}, "plan binary roles are incomplete")
    binary_map = {}
    for role in ("before", "after"):
        require(set(binaries[role]) == {"native", "allocator"}, f"{role} binary lanes are incomplete")
        for mode in ("native", "allocator"):
            item = binaries[role][mode]
            require(isinstance(item, dict), f"{role}/{mode} binary descriptor missing")
            expected_name = "litchi-perf-baseline" if mode == "native" else "litchi-perf-baseline-alloc"
            expected_features = [] if mode == "native" else ["allocator-metrics"]
            require(item.get("role") == mode and item.get("name") == expected_name and item.get("features") == expected_features, f"{role}/{mode} binary metadata mismatch")
            path = path_from_record(item, PACKET, f"{role}/{mode} binary")
            actual = descriptor(path)
            same_descriptor(item, actual, f"{role}/{mode} binary")
            require(actual["executable"] is True, f"{role}/{mode} binary is not executable")
            for nested in ("build_receipt", "depfile"):
                nested_path = path_from_record(item.get(nested), PACKET, f"{role}/{mode} {nested}")
                nested_actual = descriptor(nested_path)
                same_descriptor(item[nested], nested_actual, f"{role}/{mode} {nested}")
            require(isinstance(item.get("elf_build_id"), str) and item["elf_build_id"], f"{role}/{mode} ELF build ID missing")
            binary_map[(role, mode)] = item
    return {
        "descriptor": descriptor(plan_path),
        "bindings": binding_review,
        "binaries": binary_map,
        "plan_sha256": sha(plan_path),
    }


def validate_guard(result: dict[str, Any], sample_order: list[int], warmup: int, samples: int, expected_scan: dict[str, Any], label: str) -> None:
    source = result.get("source")
    require(isinstance(source, dict), f"{label}: source evidence missing")
    scan = source.get("producer_full_cell_scan")
    require(isinstance(scan, dict), f"{label}: producer scan missing")
    operation = result.get("operation_metrics")
    require(isinstance(operation, dict), f"{label}: operation metrics missing")
    require(operation.get("sample_count") == samples and operation.get("sample_indices") == sample_order, f"{label}: operation sample alignment")
    require(scan.get("guard_contract") == GUARD_CONTRACT, f"{label}: guard contract")
    require(scan.get("sample_indices") == sample_order, f"{label}: guard sample alignment")
    require(scan.get("archive_sha256") == PRIMARY_ARCHIVE_SHA256 and scan.get("archive_bytes") == PRIMARY_ARCHIVE_BYTES, f"{label}: archive identity")
    require(scan.get("oracle_digest_sha256") == PRIMARY_DIGEST_SHA256 and scan.get("marker_refusal_census_sha256") == PRIMARY_CENSUS_SHA256, f"{label}: guard census/digest")
    require(scan.get("expected_cell_count") == 4096 and scan.get("expected_numeric_count") == 3686 and scan.get("expected_text_count") == 410, f"{label}: expected cell census")
    require(scan.get("observed_cell_counts") == [4096] * samples, f"{label}: observed cells")
    require(scan.get("observed_numeric_counts") == [3686] * samples, f"{label}: observed numeric")
    require(scan.get("observed_text_counts") == [410] * samples, f"{label}: observed text")
    require(scan.get("observed_digest_sha256") == [PRIMARY_DIGEST_SHA256] * samples, f"{label}: observed digest")
    require(scan.get("source_version_stable") == [True] * samples, f"{label}: source freshness")
    require(scan.get("forced_worksheet_counts") == [4] * samples and scan.get("forced_worksheet_tolerance_clean") == [True] * samples, f"{label}: worksheet guard")
    require(scan.get("warmup_forced_worksheet_counts") == [4] * warmup and scan.get("warmup_forced_worksheet_tolerance_clean") == [True] * warmup, f"{label}: warmup guard")
    require(scan.get("forced_worksheet_guard_iterations") == warmup + samples, f"{label}: guard iteration count")
    require(scan.get("runtime_route_observed") is False and scan.get("physical_cold_eligible") is False, f"{label}: route/cold scope")
    timing_scope = scan.get("timing_scope")
    require(isinstance(timing_scope, str), f"{label}: timing scope missing")
    for phrase in ("process/allocation endpoint snapshots", "post-timing per-worksheet tolerance guard", "aggregate tolerance report", "post-guard source-freshness fence"):
        require(phrase in timing_scope, f"{label}: timing scope omits {phrase}")
    for field, value in expected_scan.items():
        if field not in PRIMARY_SCAN_DYNAMIC:
            require(scan.get(field) == value, f"{label}: guard static field {field}")


def validate_operation(result: dict[str, Any], mode: str, case: str, sample_order: list[int], samples: int, label: str) -> dict[str, Any] | None:
    operation = result.get("operation_metrics")
    if mode == "native" and case != PRIMARY_CASE:
        require(operation in (None, {}), f"{label}: control unexpectedly has operation metrics")
        return None
    require(isinstance(operation, dict), f"{label}: operation metrics missing")
    require(operation.get("sample_count") == samples and operation.get("sample_indices") == sample_order, f"{label}: operation alignment")
    process = operation.get("process")
    require(isinstance(process, dict) and process.get("status") == "measured", f"{label}: process metrics")
    peak = process.get("peak_rss_bytes")
    require(isinstance(peak, dict) and peak.get("status") == "measured" and peak.get("scope") == "process_lifetime_high_water_after_not_operation_peak", f"{label}: process peak scope")
    peak_values = peak.get("values")
    require(type(peak_values) is list and len(peak_values) == samples and all(type(value) is int and value > 0 for value in peak_values), f"{label}: process peak vector")
    if mode != "allocator":
        return None
    allocation = operation.get("allocation")
    require(isinstance(allocation, dict) and allocation.get("status") == "measured", f"{label}: allocation metrics")
    for metric in ALLOCATION_METRICS:
        entry = allocation.get(metric)
        require(isinstance(entry, dict) and entry.get("status") == "measured" and entry.get("scope") == "operation_global_system_allocator", f"{label}: allocator {metric} scope")
        values = entry.get("values")
        require(type(values) is list and len(values) == samples and all(type(value) is int and value >= 0 for value in values), f"{label}: allocator {metric} vector")
        if metric in {"allocation_calls", "deallocation_calls", "allocated_bytes", "region_peak_live_bytes"}:
            require(any(value > 0 for value in values), f"{label}: allocator {metric} is empty")
    failed = allocation.get("failed_allocation_calls")
    require(isinstance(failed, dict) and failed.get("status") == "measured" and failed.get("scope") == "operation_global_system_allocator", f"{label}: failed allocation metric")
    failed_values = failed.get("values")
    require(type(failed_values) is list and len(failed_values) == samples and failed_values == [0] * samples, f"{label}: failed allocation calls")
    return allocation


def validate_report(report: dict[str, Any], producer: dict[str, Any] | None, row: dict[str, Any], binary: dict[str, Any], expected: dict[str, Any], guard_scan: dict[str, Any]) -> tuple[dict[str, int], int, dict[str, Any] | None]:
    label = f"child {row['ordinal']}"
    mode = "native" if row["lane"] == "native" else "allocator"
    require(report.get("schema_version") == 1, f"{label}: report schema")
    tool = report.get("tool", {})
    require(tool.get("name") == "litchi-perf-baseline" and tool.get("version") == "0.1.0", f"{label}: tool identity")
    require(tool.get("profile") == PROFILE and tool.get("target_os") == "linux" and tool.get("target_arch") == "x86_64", f"{label}: target identity")
    require(tool.get("binary") == ("litchi-perf-baseline" if mode == "native" else "litchi-perf-baseline-alloc"), f"{label}: binary name")
    require(tool.get("instrumentation") == ("none" if mode == "native" else "system_allocator_operation_scoped"), f"{label}: instrumentation")
    if mode == "allocator":
        require(tool.get("allocator_counter_revision") == "serialized_region_peak_v3", f"{label}: allocator counter revision")
    config = report.get("configuration", {})
    require(config.get("samples_per_case") == row["samples"] and config.get("warmup_iterations_per_case") == row["warmup"], f"{label}: sample protocol")
    require(config.get("filesystem_cache_states") == ["warm"] and config.get("filesystem_fresh_child_per_sample") is True and config.get("filesystem_process_isolated") is True, f"{label}: filesystem protocol")
    require(config.get("xlsx_shapes") == ["medium"] and config.get("cases") == [row["case"]], f"{label}: case protocol")
    require(report.get("environment", {}).get("cpu_affinity") == CPU_AFFINITY, f"{label}: CPU affinity")
    identity = report.get("binary_identity", {})
    require(identity.get("path") == binary["path"] and identity.get("binary_sha256") == binary["sha256"] and identity.get("binary_bytes") == binary["bytes"] and identity.get("mode_bits") == binary["mode"] and identity.get("executable") is True and identity.get("profile") == PROFILE, f"{label}: binary custody")
    results = report.get("results")
    require(type(results) is list and len(results) == 1 and results[0].get("case") == row["case"], f"{label}: result case")
    result = results[0]
    elapsed = result.get("elapsed_ns", {})
    raw = elapsed.get("samples")
    stats = raw_stats(raw, f"{label} elapsed")
    require(len(raw) == row["samples"] and sorted(elapsed.get("sample_order", [])) == list(range(row["samples"])), f"{label}: sample order")
    for key, value in stats.items():
        require(elapsed.get(key) == value, f"{label}: elapsed {key}")
    require(result.get("corpus") == expected["corpus"], f"{label}: sealed corpus mismatch")
    if expected["producer"] is None:
        require(producer is None, f"{label}: unexpected producer sidecar")
    else:
        require(producer == expected["producer"], f"{label}: producer sidecar mismatch")
    allocation = validate_operation(result, mode, row["case"], elapsed["sample_order"], row["samples"], label)
    if row["case"] == PRIMARY_CASE:
        require(producer is not None, f"{label}: primary producer sidecar missing")
        validate_guard(result, elapsed["sample_order"], row["warmup"], row["samples"], guard_scan, label)
    return stats, 0, allocation


def pair_groups(children: dict[int, dict[str, Any]], lane: str, case: str) -> list[dict[str, Any]]:
    selected = [item for item in children.values() if item["row"]["lane"] == lane and item["row"]["case"] == case]
    expected_count = NATIVE_BLOCKS if lane == "native" else ALLOCATOR_BLOCKS
    groups: dict[str, dict[str, Any]] = {}
    for item in selected:
        groups.setdefault(item["row"]["pair"], {})[item["row"]["role"]] = item
    require(len(groups) == expected_count, f"{lane}/{case}: pair count")
    pairs = []
    for pair in sorted(groups):
        require(set(groups[pair]) == {"before", "after"}, f"{pair}: pair roles")
        pairs.append({"pair": pair, "before": groups[pair]["before"], "after": groups[pair]["after"]})
    return pairs


def native_summary(pairs: list[dict[str, Any]], metric: str) -> dict[str, Any]:
    values = []
    for pair in pairs:
        before = pair["before"]["stats"][metric]
        after = pair["after"]["stats"][metric]
        values.append({"pair": pair["pair"], "before": before, "after": after, "ratio": after / before})
    summary = bootstrap([item["ratio"] for item in values])
    summary["pairs"] = values
    return summary


def rss_summary(pairs: list[dict[str, Any]]) -> dict[str, Any]:
    values = []
    for pair in pairs:
        before = pair["before"]["peak_rss_kib"]
        after = pair["after"]["peak_rss_kib"]
        values.append({"pair": pair["pair"], "before": before, "after": after, "ratio": after / before})
    summary = bootstrap([item["ratio"] for item in values])
    summary["pairs"] = values
    return summary


def allocation_values(item: dict[str, Any], metric: str) -> list[int]:
    allocation = item["allocation"]
    require(isinstance(allocation, dict), f"{item['row']['ordinal']}: allocator data missing")
    values = allocation[metric]["values"]
    require(type(values) is list, f"{item['row']['ordinal']}: allocator {metric} vector missing")
    return values


def allocator_summary(pairs: list[dict[str, Any]], metric: str) -> dict[str, Any]:
    values = []
    for pair in pairs:
        before_vector = allocation_values(pair["before"], metric)
        after_vector = allocation_values(pair["after"], metric)
        before = statistics.median(before_vector)
        after = statistics.median(after_vector)
        ratio = after / before if before else None
        values.append({
            "pair": pair["pair"],
            "before": before,
            "after": after,
            "delta": after - before,
            "ratio": ratio,
            "before_vector_sha256": hashlib.sha256(json.dumps(before_vector, separators=(",", ":")).encode()).hexdigest(),
            "after_vector_sha256": hashlib.sha256(json.dumps(after_vector, separators=(",", ":")).encode()).hexdigest(),
        })
    ratios = [item["ratio"] for item in values if item["ratio"] is not None]
    result: dict[str, Any] = {"pairs": values, "ratio_defined_for_all_pairs": len(ratios) == len(values)}
    if ratios:
        result.update(bootstrap(ratios))
    else:
        result.update({"seed": BOOTSTRAP_SEED, "draws": BOOTSTRAP_DRAWS, "low_index": BOOTSTRAP_LOW_INDEX, "high_index": BOOTSTRAP_HIGH_INDEX, "median_ratio": None, "ci95_low": None, "ci95_high": None, "ratios": []})
    return result


def failed_summary(pairs: list[dict[str, Any]]) -> dict[str, Any]:
    children = []
    for pair in pairs:
        before = allocation_values(pair["before"], "failed_allocation_calls")
        after = allocation_values(pair["after"], "failed_allocation_calls")
        children.append({"pair": pair["pair"], "before": before, "after": after, "all_zero": before == [0] * len(before) and after == [0] * len(after)})
    return {"all_zero": all(item["all_zero"] for item in children), "children": children}


def spread_flags(children: dict[int, dict[str, Any]]) -> list[dict[str, Any]]:
    result = []
    for case in CASES:
        for role in ("before", "after"):
            values = [item["stats"]["p50"] for item in children.values() if item["row"]["lane"] == "native" and item["row"]["case"] == case and item["row"]["role"] == role]
            require(len(values) == NATIVE_BLOCKS, f"{case}/{role}: spread population")
            ratio = max(values) / min(values)
            result.append({"case": case, "role": role, "values": values, "min": min(values), "max": max(values), "max_over_min": ratio, "threshold": P50_SPREAD_THRESHOLD, "flag": ratio > P50_SPREAD_THRESHOLD})
    return result


def ratio_flags(native: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    primary = []
    controls = []
    for case, summary in native.items():
        destination = primary if case == PRIMARY_CASE else controls
        reason = "primary pair ratio above 1.05" if case == PRIMARY_CASE else "control pair ratio above 1.05"
        for metric, point in summary.items():
            over = [pair for pair in point["pairs"] if pair["ratio"] > PAIR_RATIO_THRESHOLD]
            if over:
                destination.append({"case": case, "metric": metric, "reason": reason, "pairs_over_threshold": over, "summary": point})
    return primary, controls


def compare_reduction(reduction_path: Path, plan_sha: str, experiment_sha: str, native: dict[str, Any], allocator: dict[str, Any], gate: dict[str, Any], primary_flags: list[dict[str, Any]], control_flags: list[dict[str, Any]], spreads: list[dict[str, Any]]) -> dict[str, Any]:
    reduction = load_object(reduction_path, "reduction")
    require(reduction.get("schema") == REDUCTION_SCHEMA and reduction.get("status") == "pass", "reduction receipt schema/status")
    require(reduction.get("plan_sha256") == plan_sha and reduction.get("experiment_sha256") == experiment_sha, "reduction plan/capture binding")
    for case in CASES:
        for ours, theirs in (("p50_ns", "p50_ns"), ("p95_ns", "p95_ns"), ("p99_ns", "p99_ns"), ("peak_rss_kib", "peak_rss_kib")):
            expected = reduction["native"][case][theirs]
            actual = native[case][ours]
            for field in ("seed", "draws", "low_index", "high_index", "median_ratio", "ci95_low", "ci95_high", "ratios", "pairs"):
                require(actual.get(field) == expected.get(field), f"reduction mismatch {case}/{ours}/{field}")
    for metric in ("allocation_calls", "allocated_bytes", "region_peak_live_bytes"):
        expected = reduction["allocator_primary"][metric]
        actual = allocator[metric]
        for field in ("seed", "draws", "low_index", "high_index", "median_ratio", "ci95_low", "ci95_high", "ratios"):
            require(actual.get(field) == expected.get(field), f"reduction mismatch allocator/{metric}/{field}")
        for left, right in zip(actual["pairs"], expected["pairs"]):
            for field in ("pair", "before", "after", "ratio"):
                require(left[field] == right[field], f"reduction mismatch allocator/{metric}/pair/{field}")
    require(reduction.get("primary_native_p50_gate") == gate, "reduction primary gate mismatch")

    def flag_signature(value: dict[str, Any]) -> dict[str, Any]:
        return {key: value[key] for key in ("case", "metric", "reason", "summary")}

    require([flag_signature(item) for item in reduction.get("control_review_flags", [])] == [flag_signature(item) for item in control_flags], "reduction control flags mismatch")
    require([flag_signature(item) for item in reduction.get("primary_review_flags", [])] == [flag_signature(item) for item in primary_flags], "reduction primary flags mismatch")

    def spread_signature(value: dict[str, Any]) -> dict[str, Any]:
        return {key: value[key] for key in ("case", "role", "values", "max_over_min", "flag")}

    require([spread_signature(item) for item in reduction.get("native_p50_spread_flags", [])] == [spread_signature(item) for item in spreads if item["flag"]], "reduction spread flags mismatch")
    return {"descriptor": descriptor(reduction_path), "schema": reduction["schema"], "status": reduction["status"]}


def replay(plan_path: Path, capture_root: Path, reduction_path: Path, verification_path: Path) -> dict[str, Any]:
    plan = load_object(plan_path, "plan")
    plan_check = validate_plan(plan_path, plan)
    plan_sha = plan_check["plan_sha256"]
    capture_root = capture_root.resolve()
    require(capture_root.is_dir() and not capture_root.is_symlink(), f"capture root missing or symlinked: {capture_root}")
    experiment_path = capture_root / "experiment.json"
    experiment = load_object(experiment_path, "experiment receipt")
    experiment_identity = descriptor(experiment_path)
    require(experiment.get("schema") == EXPERIMENT_SCHEMA and experiment.get("status") == "pass", "experiment receipt schema/status")
    require(experiment.get("plan_sha256") == plan_sha and experiment.get("children") == TOTAL_CHILDREN and experiment.get("native_children") == TOTAL_NATIVE_CHILDREN and experiment.get("allocator_children") == TOTAL_ALLOCATOR_CHILDREN and experiment.get("no_retries_or_replacements") is True, "experiment receipt population/custody")
    rows = expected_rows()
    expected_receipts = [{"path": f"runs/{row['ordinal']:03d}-{row['lane']}-{row['pair']}-{row['role']}/receipt.json", "row": row} for row in rows]
    require(experiment.get("rows") == expected_receipts, "experiment receipt schedule mismatch")
    expected = expected_corpora()
    guard_scan = plan_check["bindings"]["guard_scan"]
    children: dict[int, dict[str, Any]] = {}
    receipt_rows = []
    artifact_counts = {"receipt": 0, "command": 0, "report": 0, "peak": 0, "stdout": 0, "stderr": 0, "producer": 0, "semantic": 0, "binary": 0}
    previous_finished = None
    for row in rows:
        folder = capture_root / "runs" / f"{row['ordinal']:03d}-{row['lane']}-{row['pair']}-{row['role']}"
        receipt_path = folder / "receipt.json"
        receipt = load_object(receipt_path, f"child {row['ordinal']} receipt")
        require(receipt.get("schema") == RECEIPT_SCHEMA and receipt.get("row") == row and receipt.get("exit_code") == 0 and receipt.get("plan_sha256") == plan_sha, f"child {row['ordinal']}: receipt identity")
        mode = "native" if row["lane"] == "native" else "allocator"
        binary = plan_check["binaries"][(row["role"], mode)]
        receipt_binary = receipt.get("binary")
        require(isinstance(receipt_binary, dict), f"child {row['ordinal']}: binary binding wrapper")
        require(receipt_binary.get("declared") == binary, f"child {row['ordinal']}: binary declared binding")
        same_descriptor(receipt_binary.get("identity"), binary, f"child {row['ordinal']}: binary identity")
        binary_path = path_from_record(binary, PACKET, f"child {row['ordinal']} binary")
        binary_actual = descriptor(binary_path)
        same_descriptor(binary, binary_actual, f"child {row['ordinal']} binary")
        artifact_counts["binary"] += 1
        command_path = folder / "command.json"
        command = load_object(command_path, f"child {row['ordinal']} command")
        require(command.get("schema") == COMMAND_SCHEMA and command.get("row") == row and command.get("plan_sha256") == plan_sha, f"child {row['ordinal']}: command identity")
        require(command.get("binary") == binary and command.get("environment") == CHILD_ENV and command.get("cwd") == str(PACKET), f"child {row['ordinal']}: command custody")
        report_path = folder / "report.json"
        producer_path = folder / "producer.json"
        peak_path = folder / "peak-rss-kib.txt"
        stdout_path = folder / "stdout.txt"
        stderr_path = folder / "stderr.txt"
        expected_argv = ["/usr/bin/time", "-f", "%M", "-o", str(peak_path), "/usr/bin/taskset", "-c", CPU_AFFINITY, binary["path"], "--case", row["case"], "--xlsx-shape", "medium", "--warmup", str(row["warmup"]), "--samples", str(row["samples"]), "--filesystem-cache", "warm", "--json", str(report_path), "--producer-evidence", str(producer_path)]
        require(command.get("argv") == expected_argv and receipt.get("argv") == expected_argv and receipt.get("cwd") == command.get("cwd") and receipt.get("environment") == command.get("environment"), f"child {row['ordinal']}: command argv")
        require(receipt.get("command", {}).get("sha256") == sha(command_path), f"child {row['ordinal']}: command descriptor")
        command_actual = descriptor(command_path)
        artifact_records = receipt.get("artifacts", {})
        for name, path, key in (("report.json", report_path, "report"), ("peak-rss-kib.txt", peak_path, "peak"), ("stdout.txt", stdout_path, "stdout"), ("stderr.txt", stderr_path, "stderr")):
            check_descriptor(artifact_records.get(name), path, f"child {row['ordinal']} {name}")
            artifact_counts[key] += 1
        producer = None
        producer_actual = None
        if producer_path.is_file():
            producer = load_object(producer_path, f"child {row['ordinal']} producer")
            check_descriptor(artifact_records.get("producer.json"), producer_path, f"child {row['ordinal']} producer")
            producer_actual = descriptor(producer_path)
            artifact_counts["producer"] += 1
        else:
            require("producer.json" not in artifact_records, f"child {row['ordinal']}: missing producer artifact custody")
        semantic_path = folder / "semantic.json"
        semantic = load_object(semantic_path, f"child {row['ordinal']} semantic")
        require(semantic.get("status") == "pass" and semantic.get("report_sha256") == sha(report_path) and semantic.get("producer_sha256") == (sha(producer_path) if producer_path.is_file() else None), f"child {row['ordinal']}: semantic custody")
        semantic_actual = descriptor(semantic_path)
        artifact_counts["semantic"] += 1
        started = _datetime.datetime.fromisoformat(receipt["started_utc"])
        finished = _datetime.datetime.fromisoformat(receipt["finished_utc"])
        require(finished >= started, f"child {row['ordinal']}: invalid interval")
        if previous_finished is not None:
            require(started >= previous_finished, f"child {row['ordinal']}: children are not sequential")
        previous_finished = finished
        peak_text = peak_path.read_text(encoding="utf-8").strip()
        require(re.fullmatch(r"[0-9]+", peak_text) is not None and int(peak_text) > 0, f"child {row['ordinal']}: peak RSS value")
        report = load_object(report_path, f"child {row['ordinal']} report")
        stats, _, allocation = validate_report(report, producer, row, binary, expected[row["case"]], guard_scan)
        process_peak = report.get("results", [{}])[0].get("operation_metrics", {}).get("process", {}).get("peak_rss_bytes", {}).get("values")
        peak_rss = max(process_peak) if isinstance(process_peak, list) and process_peak else int(peak_text)
        children[row["ordinal"]] = {"row": row, "stats": stats, "report": report, "allocation": allocation, "peak_rss_kib": int(peak_text), "process_peak_rss_bytes": process_peak}
        receipt_rows.append({"ordinal": row["ordinal"], "row": row, "receipt": descriptor(receipt_path), "command": command_actual, "artifacts": {name: descriptor(path) for name, path in (("report.json", report_path), ("peak-rss-kib.txt", peak_path), ("stdout.txt", stdout_path), ("stderr.txt", stderr_path))} | ({"producer.json": producer_actual} if producer_actual else {}), "semantic": semantic_actual})
        artifact_counts["receipt"] += 1
        artifact_counts["command"] += 1
        _ = peak_rss
    require(len(children) == TOTAL_CHILDREN and previous_finished is not None, "child population incomplete")
    native = {}
    for case in CASES:
        pairs = pair_groups(children, "native", case)
        native[case] = {"p50_ns": native_summary(pairs, "p50"), "p95_ns": native_summary(pairs, "p95"), "p99_ns": native_summary(pairs, "p99"), "peak_rss_kib": rss_summary(pairs)}
    allocator_pairs = pair_groups(children, "allocator", PRIMARY_CASE)
    allocator = {metric: allocator_summary(allocator_pairs, metric) for metric in ALLOCATION_METRICS}
    allocator["failed_allocation_calls"] = failed_summary(allocator_pairs)
    primary_flags, control_flags = ratio_flags(native)
    allocator_flags = []
    for metric, summary in allocator.items():
        if metric == "failed_allocation_calls":
            continue
        over = [pair for pair in summary["pairs"] if pair.get("ratio") is not None and pair["ratio"] > PAIR_RATIO_THRESHOLD]
        if over:
            allocator_flags.append({"metric": metric, "reason": "allocator pair ratio above 1.05", "pairs_over_threshold": over, "summary": summary})
    spreads = spread_flags(children)
    primary_gate = {"median_ratio_at_most": 0.98, "ci95_upper_below": 1.0, "median_ratio": native[PRIMARY_CASE]["p50_ns"]["median_ratio"], "ci95_high": native[PRIMARY_CASE]["p50_ns"]["ci95_high"], "pass": native[PRIMARY_CASE]["p50_ns"]["median_ratio"] <= 0.98 and native[PRIMARY_CASE]["p50_ns"]["ci95_high"] < 1.0}
    verification = load_object(verification_path, "verification")
    verification_identity = descriptor(verification_path)
    require(verification.get("schema") == VERIFICATION_SCHEMA and verification.get("status") == "pass", "verification receipt schema/status")
    require(verification.get("plan", {}).get("sha256") == plan_sha and verification.get("experiment", {}).get("sha256") == sha(experiment_path), "verification plan/capture binding")
    reduction = compare_reduction(reduction_path, plan_sha, sha(experiment_path), native, allocator, primary_gate, primary_flags, control_flags, spreads)
    helper_identity = descriptor(Path(__file__))
    source_census = plan_check["bindings"]["source_census"]
    return {
        "schema": SCHEMA,
        "status": "pass",
        "reviewer": "independent_review.py",
        "scope": {"independent_raw_replay": True, "timed_workload_rerun": False, "cargo_rerun": False, "profiler_rerun": False, "source_mutation": False, "performance_claim": False, "production_adoption": False},
        "bindings": {"helper": helper_identity, "plan": plan_check["descriptor"], "capture": experiment_identity, "reduction": reduction["descriptor"], "verification": verification_identity, "guard_reference": plan_check["bindings"]["guard_reference"], "source_census": source_census},
        "protocol_replay": {"base_commit": BASELINE_COMMIT, "evaluation_head": EVALUATION_HEAD, "cases": list(CASES), "primary_case": PRIMARY_CASE, "cpu_affinity": CPU_AFFINITY, "environment": CHILD_ENV, "filesystem_cache": "warm", "fresh_child_per_scheduled_row": True, "fail_fast": True, "native_children": TOTAL_NATIVE_CHILDREN, "allocator_children": TOTAL_ALLOCATOR_CHILDREN, "native_pairs_per_case": NATIVE_BLOCKS, "allocator_pairs": ALLOCATOR_BLOCKS, "native_samples": NATIVE_SAMPLES, "allocator_samples": ALLOCATOR_SAMPLES, "warmup": NATIVE_WARMUP, "bootstrap": {"seed": BOOTSTRAP_SEED, "draws": BOOTSTRAP_DRAWS, "ci95_indices": [BOOTSTRAP_LOW_INDEX, BOOTSTRAP_HIGH_INDEX], "unit": "whole paired child ratios"}, "pooling": "none; each ratio uses one child aggregate statistic and pairs are summarized by median"},
        "raw_capture_replay": {"capture_status": "pass", "experiment_schema": EXPERIMENT_SCHEMA, "experiment_sha256": sha(experiment_path), "plan_sha256": plan_sha, "expected_children": TOTAL_CHILDREN, "observed_children": len(children), "native_children": TOTAL_NATIVE_CHILDREN, "allocator_children": TOTAL_ALLOCATOR_CHILDREN, "raw_sample_total": TOTAL_SAMPLES, "raw_sample_counts": {"native": TOTAL_NATIVE_CHILDREN * NATIVE_SAMPLES, "allocator": TOTAL_ALLOCATOR_CHILDREN * ALLOCATOR_SAMPLES}, "artifact_hashes_checked": artifact_counts, "no_retries_or_replacements": True, "sequential_intervals_checked": TOTAL_CHILDREN, "receipt_rows": receipt_rows},
        "guard_census_custody": {"guard_contract": GUARD_CONTRACT, "primary_guard_children": sum(row["case"] == PRIMARY_CASE for row in rows), "guard_reference": plan_check["bindings"]["guard_reference"], "source_census": source_census, "candidate_overlay_files": sorted(SOURCE_OVERLAY), "binary_identities_checked": artifact_counts["binary"], "all_receipts_and_artifacts_bound": True},
        "native": native,
        "allocator_primary": allocator,
        "allocator_review": {"status": "descriptive_review_only", "adoption": "not_evaluated", "metrics": {metric: allocator[metric] for metric in ALLOCATION_METRICS}, "failed_allocation_calls": allocator["failed_allocation_calls"], "reason": "Allocator counters are separate observations and are not the native adoption gate."},
        "primary_native_p50_gate": primary_gate,
        "primary_review_flags": primary_flags,
        "control_review_flags": control_flags,
        "allocator_ratio_flags": allocator_flags,
        "native_p50_spread_flags": [item for item in spreads if item["flag"]],
        "native_p50_spread_review": spreads,
        "issues": [],
        "performance_claim": None,
        "production_adoption": False,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, default=PLAN_DEFAULT)
    parser.add_argument("--experiment-root", type=Path, default=CAPTURE_DEFAULT)
    parser.add_argument("--reduction", type=Path, default=REDUCTION_DEFAULT)
    parser.add_argument("--verification", type=Path, default=VERIFICATION_DEFAULT)
    parser.add_argument("--output", type=Path, default=OUTPUT_DEFAULT)
    args = parser.parse_args()
    output = args.output.resolve()
    require(not output.exists() and not output.is_symlink(), f"review output already exists: {output}")
    result = replay(args.plan.resolve(), args.experiment_root.resolve(), args.reduction.resolve(), args.verification.resolve())
    with output.open("x", encoding="utf-8") as handle:
        json.dump(result, handle, indent=2, sort_keys=True)
        handle.write("\n")
    print(json.dumps({"status": result["status"], "output": str(output), "sha256": sha(output), "bytes": output.stat().st_size, "primary_gate": result["primary_native_p50_gate"]}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
