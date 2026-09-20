#!/usr/bin/env python3
"""Validate the separate 0705 XLSX allocator lane.

This analyzer never builds or runs a benchmark.  It checks the artifacts
planned by ``allocations.py`` after capture, binds every report to its build,
source manifest, command and receipt, and compares allocator reports with the
matching uninstrumented native reports only for deterministic identity.  The
allocator samples are diagnostics; this file makes no latency claim.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import statistics
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
TARGET = ROOT.parent / "litchi-target-0705"
CASE = "xlsx_source_backed_cell_values_one_percent_edit_save"
SHAPES = ("medium", "dense-sparse")
REPEATS = (1, 2)
ALLOC_SAMPLES = 3
ALLOC_WARMUPS = 2
NATIVE_SAMPLES = 100
NATIVE_WARMUPS = 20
CPU = 12

PHASES = ("plan", "staging", "commit_core", "commit", "publication")
PHASE_KEYS = {
    "plan": "plan_allocation_metrics",
    "staging": "staging_allocation_metrics",
    "commit_core": "commit_core_allocation_metrics",
    "commit": "commit_allocation_metrics",
    "publication": "publication_allocation_metrics",
}
TIME_KEYS = ("open_ns", "plan_ns", "commit_ns", "publication_ns", "reopen_ns")
SAMPLE_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
    "region_peak_live_bytes",
)
COUNTER_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
)
ALLOCATOR_SCOPE = "operation_global_system_allocator"
ALLOCATOR_INSTRUMENTATION = "system_allocator_operation_scoped"
COUNTER_REVISION = "serialized_region_peak_v3"
GENERATOR = "litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1"


class EvidenceError(Exception):
    """A present artifact is malformed or violates the frozen contract."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read {path.name}: {error}") from error


def sha(path: Path) -> str:
    try:
        return hashlib.sha256(path.read_bytes()).hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def hash_value(value: Any, label: str) -> str:
    require(
        isinstance(value, str)
        and len(value) == 64
        and all(character in "0123456789abcdef" for character in value),
        f"{label} is not a lowercase SHA-256 digest",
    )
    return value


def nonnegative_integer(value: Any, label: str) -> int:
    require(isinstance(value, int) and not isinstance(value, bool), f"{label} is not an integer")
    require(value >= 0, f"{label} is negative")
    return value


def finite_nonnegative(value: Any, label: str) -> float:
    require(isinstance(value, (int, float)) and not isinstance(value, bool), f"{label} is not numeric")
    result = float(value)
    require(math.isfinite(result) and result >= 0, f"{label} is not finite and nonnegative")
    return result


def object_value(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is not an object")
    return value


def list_value(value: Any, label: str) -> list[Any]:
    require(isinstance(value, list), f"{label} is not an array")
    return value


def expected_alloc_name(repeat: int, shape: str) -> str:
    return f"alloc-r{repeat}-{shape}"


def expected_native_name(repeat: int, shape: str) -> str:
    return f"native-r{repeat}-{shape}"


def expected_alloc_command(name: str, shape: str, binary: Path) -> list[str]:
    return [
        "taskset",
        "-c",
        str(CPU),
        str(binary),
        "--warmup",
        str(ALLOC_WARMUPS),
        "--samples",
        str(ALLOC_SAMPLES),
        "--case",
        CASE,
        "--xlsx-cell-crud-shape",
        shape,
        "--json",
        str(HERE / f"{name}.json"),
    ]


def expected_native_command(name: str, shape: str, binary: Path) -> list[str]:
    return [
        "taskset",
        "-c",
        str(CPU),
        str(binary),
        "--warmup",
        str(NATIVE_WARMUPS),
        "--samples",
        str(NATIVE_SAMPLES),
        "--case",
        CASE,
        "--xlsx-cell-crud-shape",
        shape,
        "--json",
        str(HERE / f"{name}.json"),
    ]


def expected_build_command(allocator: bool) -> list[str]:
    binary_name = "litchi-perf-baseline-alloc" if allocator else "litchi-perf-baseline"
    command = [
        "cargo",
        "build",
        "--release",
        "--locked",
        "--manifest-path",
        "tools/perf-baseline/Cargo.toml",
    ]
    if allocator:
        command += ["--features", "allocator-metrics"]
    command += ["--bin", binary_name, "--target-dir", str(TARGET), "-j", "2"]
    return command


def plan_identity() -> dict[str, Any]:
    path = HERE / "plan.json"
    require(path.is_file() and not path.is_symlink(), "plan.json is missing")
    plan = object_value(read_json(path), "plan")
    expected = {
        "revision": "eac6f43db52f4092c50ecc3e1058685024b3bfaf",
        "case": CASE,
        "shapes": list(SHAPES),
        "native_repeats": 2,
        "warmup": NATIVE_WARMUPS,
        "samples": NATIVE_SAMPLES,
        "cpu": CPU,
        "performance_claim": "none; current-head diagnostic, not a historical speed comparison",
        "timing": "sum of open, plan, staged sets plus commit, publication including returned snapshot drop; excludes sink setup, other handle drops, reopen and oracles",
        "provider": "instrumented in-memory ReadAt; fresh editor/cache per iteration, warm process",
    }
    for key, value in expected.items():
        require(plan.get(key) == value, f"plan.{key} differs from the frozen 0705 plan")
    return {"path": str(path), "sha256": sha(path), **plan}


def source_manifest() -> dict[str, Any]:
    path = HERE / "source-manifest.json"
    require(path.is_file() and not path.is_symlink(), "source-manifest.json is missing")
    manifest = object_value(read_json(path), "source manifest")
    require(bool(manifest), "source manifest is empty")
    for name, digest in manifest.items():
        require(isinstance(name, str) and name and not Path(name).is_absolute(), "source manifest path is invalid")
        hash_value(digest, f"source manifest {name}")
    return {"path": str(path), "sha256": sha(path), "entries": len(manifest)}


def validate_binary_file(path: Path, expected_digest: str, label: str) -> dict[str, Any]:
    if not path.exists():
        raise EvidenceError(f"{label} binary is missing: {path}")
    require(not path.is_symlink(), f"{label} binary is a symlink")
    require(path.is_file(), f"{label} binary is not a regular file")
    actual = sha(path)
    require(actual == expected_digest, f"{label} binary digest differs from its receipt")
    return {"path": str(path), "sha256": actual, "bytes": path.stat().st_size}


def cleanup_binary_witness(path: Path, expected_digest: str, label: str) -> dict[str, Any] | None:
    """Accept a post-capture removal only with an exact cleanup witness.

    0705's capture script retains the build receipt and report identities, and
    the coordinator may remove the owned target before the final audit.  Older
    batches used several cleanup receipt shapes, so recognize their common
    path/digest forms without making cleanup itself part of the output identity.
    """

    candidates = (
        HERE / "cleanup.json",
        HERE / "allocation-cleanup.json",
        HERE / "cleanup-alloc.json",
    )
    wanted = str(path)
    for cleanup_path in candidates:
        if not cleanup_path.is_file() or cleanup_path.is_symlink():
            continue
        cleanup = object_value(read_json(cleanup_path), f"{label} cleanup witness")
        removed_paths: set[str] = set()
        digest_matches = False
        digest_bytes: int | None = None
        root_absent = cleanup.get("owned_paths_absent") is True

        def visit(value: Any) -> None:
            nonlocal digest_matches, digest_bytes, root_absent
            if isinstance(value, dict):
                if value.get("owned_paths_absent") is True:
                    root_absent = True
                path_value = value.get("path")
                if isinstance(path_value, str):
                    removed = value.get("removed") is True or value.get("exists_after") is False
                    if removed:
                        removed_paths.add(path_value)
                    digest_value = value.get("binary_sha256") or value.get("sha256")
                    if path_value == wanted and digest_value == expected_digest:
                        digest_matches = True
                        for key in ("bytes", "binary_bytes", "size"):
                            if isinstance(value.get(key), int) and value[key] >= 0:
                                digest_bytes = value[key]
                                break
                for key in ("retained_binary_sha256_before_removal", "binary_sha256"):
                    mapping = value.get(key)
                    if isinstance(mapping, dict) and mapping.get(wanted) == expected_digest:
                        digest_matches = True
                for child in value.values():
                    visit(child)
            elif isinstance(value, list):
                for child in value:
                    visit(child)

        visit(cleanup)
        if digest_matches and (wanted in removed_paths or root_absent):
            return {"path": wanted, "sha256": expected_digest, "bytes": digest_bytes}
    return None


def validate_build_receipt(
    path: Path, *, allocator: bool, manifest: dict[str, Any], plan: dict[str, Any]
) -> dict[str, Any]:
    label = "allocator" if allocator else "native"
    require(path.is_file() and not path.is_symlink(), f"{label} build receipt is missing")
    receipt = object_value(read_json(path), f"{label} build receipt")
    require(receipt.get("command") == expected_build_command(allocator), f"{label} build command differs from allocations/build.py")
    require(receipt.get("exit_code") == 0, f"{label} build did not succeed")
    seconds = finite_nonnegative(receipt.get("seconds"), f"{label} build seconds")
    digest = hash_value(receipt.get("binary_sha256"), f"{label} build binary_sha256")
    require(receipt.get("source_manifest_sha256") == manifest["sha256"], f"{label} build source manifest hash differs")
    binary_name = "litchi-perf-baseline-alloc" if allocator else "litchi-perf-baseline"
    expected_binary = TARGET / "release" / binary_name
    require(Path(receipt.get("binary", "")) == expected_binary, f"{label} build binary path differs")
    require(not expected_binary.is_symlink(), f"{label} binary is a symlink")
    if expected_binary.exists():
        binary = validate_binary_file(expected_binary, digest, label)
    else:
        witness = cleanup_binary_witness(expected_binary, digest, label)
        require(witness is not None, f"{label} binary is missing without an exact cleanup witness")
        binary = witness
    if not allocator and "revision" in receipt:
        require(receipt["revision"] == plan["revision"], "native build revision differs from plan")
    if allocator:
        require("revision" not in receipt or receipt["revision"] == plan["revision"], "allocator build revision differs from plan")
    return {
        "path": str(path),
        "sha256": sha(path),
        "binary": binary,
        "binary_sha256": digest,
        "seconds": seconds,
        "source_manifest_sha256": receipt["source_manifest_sha256"],
        "command": receipt["command"],
    }


def validate_capture_receipt(
    name: str,
    *,
    shape: str,
    repeat: int,
    allocator: bool,
    binary: Path,
    binary_sha256: str,
    manifest_sha256: str,
    plan_sha256: str,
) -> dict[str, Any]:
    receipt_path = HERE / f"{name}.receipt.json"
    report_path = HERE / f"{name}.json"
    stdout_path = HERE / f"{name}.stdout"
    stderr_path = HERE / f"{name}.stderr"
    require(receipt_path.is_file() and not receipt_path.is_symlink(), f"{name} receipt is missing")
    receipt = object_value(read_json(receipt_path), f"{name} receipt")
    expected_command = expected_alloc_command(name, shape, binary) if allocator else expected_native_command(name, shape, binary)
    require(receipt.get("command") == expected_command, f"{name} command differs from planned command")
    require(receipt.get("exit_code") == 0, f"{name} child did not succeed")
    require(receipt.get("binary_sha256") == binary_sha256, f"{name} binary hash differs from build receipt")
    require(receipt.get("source_manifest_sha256") == manifest_sha256, f"{name} source manifest hash differs")
    require(receipt.get("script_sha256") == sha(HERE / ("allocations.py" if allocator else "capture.py")), f"{name} acquisition script hash differs")
    if not allocator:
        require(receipt.get("plan_sha256") == plan_sha256, f"{name} plan hash differs")
    artifacts = object_value(receipt.get("artifacts"), f"{name} receipt artifacts")
    expected_artifacts = {report_path.name, stdout_path.name, stderr_path.name}
    require(set(artifacts) == expected_artifacts, f"{name} receipt artifact inventory differs")
    for artifact_name, digest in artifacts.items():
        artifact_path = HERE / artifact_name
        require(artifact_path.is_file() and not artifact_path.is_symlink(), f"{name} artifact is missing: {artifact_name}")
        require(hash_value(digest, f"{name} {artifact_name} hash") == sha(artifact_path), f"{name} artifact hash differs: {artifact_name}")
    return {
        "path": str(receipt_path),
        "sha256": sha(receipt_path),
        "artifacts": artifacts,
        "command": receipt["command"],
        "binary_sha256": binary_sha256,
        "source_manifest_sha256": manifest_sha256,
    }


def percentile(values: list[int], fraction: float) -> float:
    ordered = sorted(values)
    if len(ordered) == 1:
        return float(ordered[0])
    position = (len(ordered) - 1) * fraction
    lower = math.floor(position)
    upper = math.ceil(position)
    if lower == upper:
        return float(ordered[lower])
    weight = position - lower
    return ordered[lower] + (ordered[upper] - ordered[lower]) * weight


def metric_summary(values: list[int]) -> dict[str, Any]:
    require(bool(values), "cannot summarize an empty metric vector")
    return {
        "samples": values,
        "min": min(values),
        "p50": statistics.median(values),
        "p95": percentile(values, 0.95),
        "p99": percentile(values, 0.99),
        "max": max(values),
        "mean": statistics.fmean(values),
    }


def validate_allocator_sample(sample: Any, label: str) -> dict[str, int]:
    value = object_value(sample, label)
    require(value.get("status") == "measured", f"{label} is not measured")
    require(value.get("scope") == ALLOCATOR_SCOPE, f"{label} has the wrong allocator scope")
    parsed: dict[str, int] = {}
    for field in SAMPLE_FIELDS:
        parsed[field] = nonnegative_integer(value.get(field), f"{label}.{field}")
    require(parsed["failed_allocation_calls"] == 0, f"{label} recorded a failed allocation")
    require(parsed["peak_live_bytes_before"] >= parsed["live_bytes_before"], f"{label} peak before is below live before")
    require(parsed["peak_live_bytes_after"] >= parsed["live_bytes_after"], f"{label} peak after is below live after")
    require(parsed["peak_live_bytes_after"] >= parsed["peak_live_bytes_before"], f"{label} peak high water regressed")
    require(parsed["region_peak_live_bytes"] >= parsed["live_bytes_before"], f"{label} region peak is below live before")
    require(parsed["region_peak_live_bytes"] >= parsed["live_bytes_after"], f"{label} region peak is below live after")
    require(parsed["region_peak_live_bytes"] <= parsed["peak_live_bytes_after"], f"{label} region peak exceeds process high water")
    return parsed


def validate_native_unavailable_sample(sample: Any, label: str) -> None:
    value = object_value(sample, label)
    require(value.get("status") == "unavailable", f"{label} unexpectedly contains measured allocator evidence")
    require(value.get("scope") == ALLOCATOR_SCOPE, f"{label} has the wrong allocator scope")
    require(set(value) <= {"status", "scope"}, f"{label} unavailable sample contains numeric evidence")


def validate_phase_reconciliation(
    phase_samples: dict[str, list[dict[str, int]]], label: str
) -> dict[str, bool]:
    staging = phase_samples["staging"]
    core = phase_samples["commit_core"]
    combined = phase_samples["commit"]
    require(len(staging) == len(core) == len(combined), f"{label} split commit vectors are not aligned")
    for index, (left, middle, total) in enumerate(zip(staging, core, combined)):
        for field in COUNTER_FIELDS:
            require(
                left[field] + middle[field] == total[field],
                f"{label} commit counter {field} does not reconcile at sample {index}",
            )
        require(left["live_bytes_after"] == middle["live_bytes_before"], f"{label} split live boundary differs at sample {index}")
        require(total["live_bytes_before"] == left["live_bytes_before"], f"{label} aggregate live start differs at sample {index}")
        require(total["live_bytes_after"] == middle["live_bytes_after"], f"{label} aggregate live end differs at sample {index}")
        require(total["peak_live_bytes_before"] == left["peak_live_bytes_before"], f"{label} aggregate peak start differs at sample {index}")
        require(total["peak_live_bytes_after"] == middle["peak_live_bytes_after"], f"{label} aggregate peak end differs at sample {index}")
        require(total["region_peak_live_bytes"] == max(left["region_peak_live_bytes"], middle["region_peak_live_bytes"]), f"{label} aggregate region peak differs at sample {index}")
    return {
        "counter_sums_match": True,
        "live_boundaries_match": True,
        "peak_boundaries_match": True,
        "aggregate_region_peak_is_segment_max": True,
        "whole_operation_peak_was_not_summed": True,
    }


def validate_phase_times(result: dict[str, Any], source: dict[str, Any], label: str, samples: int) -> dict[str, Any]:
    elapsed = object_value(result.get("elapsed_ns"), f"{label}.elapsed_ns")
    elapsed_values = [nonnegative_integer(item, f"{label}.elapsed_ns.samples[{index}]") for index, item in enumerate(list_value(elapsed.get("samples"), f"{label}.elapsed_ns.samples"))]
    order = [nonnegative_integer(item, f"{label}.elapsed_ns.sample_order[{index}]") for index, item in enumerate(list_value(elapsed.get("sample_order"), f"{label}.elapsed_ns.sample_order"))]
    require(len(elapsed_values) == len(order) == samples, f"{label} elapsed vector cardinality differs from configuration")
    require(sorted(order) == list(range(samples)), f"{label} elapsed sample order is not a permutation")
    phase_values: dict[str, list[int]] = {}
    for key in ("open_ns", "plan_ns", "commit_ns", "publication_ns"):
        vector = [nonnegative_integer(item, f"{label}.source.xlsx_cell_values.{key}[{index}]") for index, item in enumerate(list_value(source.get(key), f"{label}.source.xlsx_cell_values.{key}"))]
        require(len(vector) == samples, f"{label}.{key} cardinality differs from configuration")
        phase_values[key] = vector
    reopen = [nonnegative_integer(item, f"{label}.source.xlsx_cell_values.reopen_ns[{index}]") for index, item in enumerate(list_value(source.get("reopen_ns"), f"{label}.source.xlsx_cell_values.reopen_ns"))]
    require(len(reopen) == samples, f"{label}.reopen_ns cardinality differs from configuration")
    for sorted_index, acquisition_index in enumerate(order):
        total = sum(phase_values[key][acquisition_index] for key in ("open_ns", "plan_ns", "commit_ns", "publication_ns"))
        require(total == elapsed_values[sorted_index], f"{label} timed phase sum differs at acquisition index {acquisition_index}")
    return {"elapsed_ns": elapsed_values, "sample_order": order, "phase_ns": phase_values, "reopen_ns": reopen}


def validate_report(
    path: Path,
    *,
    name: str,
    shape: str,
    repeat: int,
    allocator: bool,
    expected_samples: int,
    expected_warmups: int,
    binary: dict[str, Any],
    plan: dict[str, Any],
) -> dict[str, Any]:
    label = name
    report = object_value(read_json(path), label)
    require(report.get("schema_version") == 1, f"{label} schema version is not 1")
    tool = object_value(report.get("tool"), f"{label}.tool")
    require(tool.get("name") == "litchi-perf-baseline", f"{label} tool name differs")
    require(tool.get("version") == "0.1.0", f"{label} tool version differs")
    require(tool.get("binary") == ("litchi-perf-baseline-alloc" if allocator else "litchi-perf-baseline"), f"{label} binary identity differs")
    require(tool.get("profile") == "release", f"{label} profile differs")
    require(tool.get("target_os") == "linux" and tool.get("target_arch") == "x86_64", f"{label} target differs")
    require(tool.get("instrumentation") == (ALLOCATOR_INSTRUMENTATION if allocator else "none"), f"{label} instrumentation differs")
    if allocator:
        require(tool.get("allocator_counter_revision") == COUNTER_REVISION, f"{label} allocator counter revision differs")
    else:
        require("allocator_counter_revision" not in tool or tool.get("allocator_counter_revision") is None, f"{label} native report carries allocator counter identity")
    binary_identity = object_value(report.get("binary_identity"), f"{label}.binary_identity")
    require(binary_identity.get("binary_sha256") == binary["sha256"], f"{label} report binary hash differs")
    require(binary_identity.get("profile") == "release", f"{label} report binary profile differs")
    require(Path(binary_identity.get("path", "")) == Path(binary["path"]), f"{label} report binary path differs")
    report_binary_bytes = nonnegative_integer(binary_identity.get("binary_bytes"), f"{label}.binary_identity.binary_bytes")
    if binary["bytes"] is None:
        # When the owned target has already been removed, the retained report
        # identity supplies the size that was paired with the receipt digest.
        binary["bytes"] = report_binary_bytes
    require(report_binary_bytes == binary["bytes"], f"{label} report binary size differs")
    environment = object_value(report.get("environment"), f"{label}.environment")
    require(environment.get("git_revision") == plan["revision"], f"{label} git revision differs")
    require(environment.get("cpu_affinity") == str(CPU), f"{label} CPU affinity differs")
    require(environment.get("allocator") == ("CountingSystemAllocator(std::alloc::System)" if allocator else "Rust system allocator"), f"{label} allocator identity differs")
    configuration = object_value(report.get("configuration"), f"{label}.configuration")
    require(configuration.get("cases") == [CASE], f"{label} case configuration differs")
    require(configuration.get("xlsx_cell_crud_shapes") == [shape], f"{label} shape configuration differs")
    require(configuration.get("samples_per_case") == expected_samples, f"{label} sample count differs")
    require(configuration.get("warmup_iterations_per_case") == expected_warmups, f"{label} warmup count differs")
    results = list_value(report.get("results"), f"{label}.results")
    require(len(results) == 1, f"{label} must contain one result")
    result = object_value(results[0], f"{label}.result")
    require(result.get("case") == CASE, f"{label} result case differs")
    corpus = object_value(result.get("corpus"), f"{label}.corpus")
    require(corpus.get("name") == f"xlsx-cell-values-{shape}", f"{label} corpus name differs")
    require(corpus.get("generator") == GENERATOR, f"{label} corpus generator differs")
    require(corpus.get("package_format") == "XLSX/OPC/ZIP", f"{label} corpus package format differs")
    require(corpus.get("shape") == shape, f"{label} corpus shape differs")
    corpus_xlsx = object_value(corpus.get("xlsx"), f"{label}.corpus.xlsx")
    update_count = nonnegative_integer(corpus_xlsx.get("one_percent_update_count"), f"{label}.corpus.xlsx.one_percent_update_count")
    require(update_count > 0, f"{label} corpus has no one-percent updates")
    result_output = hash_value(result.get("output_sha256"), f"{label}.output_sha256")
    source_root = object_value(result.get("source"), f"{label}.source")
    source = object_value(source_root.get("xlsx_cell_values"), f"{label}.source.xlsx_cell_values")
    require(source.get("implementation") == "source-backed", f"{label} implementation differs")
    require(source.get("cache_mode") == "unmanaged-control", f"{label} cache mode differs")
    require(source.get("update_count") == update_count, f"{label} update count differs from corpus")
    require(nonnegative_integer(source.get("selected_worksheet_count"), f"{label}.selected_worksheet_count") > 0, f"{label} has no selected worksheet")
    for key in TIME_KEYS:
        vector = list_value(source.get(key), f"{label}.source.xlsx_cell_values.{key}")
        require(len(vector) == expected_samples, f"{label}.{key} cardinality differs")
    for key in ("read_calls", "read_bytes", "ordinary_payload_read_calls", "ordinary_payload_read_bytes", "max_in_flight_reads", "ordinary_payload_materializations"):
        vector = list_value(source_root.get(key), f"{label}.source.{key}")
        require(len(vector) == expected_samples, f"{label}.source.{key} cardinality differs")
    source_outputs = [hash_value(item, f"{label}.source.output_sha256[{index}]") for index, item in enumerate(list_value(source.get("output_sha256"), f"{label}.source.output_sha256"))]
    source_semantics = [hash_value(item, f"{label}.source.semantic_sha256[{index}]") for index, item in enumerate(list_value(source.get("semantic_sha256"), f"{label}.source.semantic_sha256"))]
    require(len(source_outputs) == len(source_semantics) == expected_samples, f"{label} output/semantic vector cardinality differs")
    require(all(item == result_output for item in source_outputs), f"{label} source output digests disagree with result digest")
    require(len(set(source_semantics)) == 1, f"{label} semantic output is not repeated deterministically")
    untouched_hashes = [hash_value(item, f"{label}.source.untouched_member_sha256[{index}]") for index, item in enumerate(list_value(source.get("untouched_member_sha256"), f"{label}.source.untouched_member_sha256"))]
    require(len(untouched_hashes) == expected_samples and len(set(untouched_hashes)) == 1, f"{label} untouched-member identity is not repeated deterministically")
    timing = validate_phase_times(result, source, label, expected_samples)
    phase_samples: dict[str, list[dict[str, int]]] = {}
    phase_metrics: dict[str, Any] = {}
    for phase in PHASES:
        key = PHASE_KEYS[phase]
        raw_vector = source.get(key)
        vector = list_value(raw_vector, f"{label}.source.xlsx_cell_values.{key}")
        require(len(vector) == expected_samples, f"{label}.{key} cardinality differs")
        parsed = [validate_allocator_sample(sample, f"{label}.{key}[{index}]") if allocator else None for index, sample in enumerate(vector)]
        if allocator:
            phase_samples[phase] = [item for item in parsed if item is not None]
            phase_metrics[phase] = {
                "sample_count": expected_samples,
                "metrics": {
                    field: metric_summary([item[field] for item in phase_samples[phase]])
                    for field in SAMPLE_FIELDS
                },
            }
        else:
            for index, sample in enumerate(vector):
                validate_native_unavailable_sample(sample, f"{label}.{key}[{index}]")
    consistency = validate_phase_reconciliation(phase_samples, label) if allocator else None
    sink = result.get("sink")
    return {
        "name": name,
        "shape": shape,
        "repeat": repeat,
        "allocator": allocator,
        "report_sha256": sha(path),
        "tool": tool,
        "binary_identity": {"path": binary_identity["path"], "binary_sha256": binary_identity["binary_sha256"], "binary_bytes": binary_identity["binary_bytes"]},
        "corpus": corpus,
        "output_sha256": result_output,
        "source_output_sha256": source_outputs,
        "semantic_sha256": source_semantics,
        "untouched_member_count": source.get("untouched_member_count"),
        "untouched_member_sha256": untouched_hashes,
        "sink": sink,
        "timing": timing,
        "phase_metrics": phase_metrics,
        "phase_reconciliation": consistency,
    }


def parity_record(allocator: dict[str, Any], native: dict[str, Any]) -> dict[str, Any]:
    require(allocator["shape"] == native["shape"] and allocator["repeat"] == native["repeat"], f"{allocator['name']} native pairing differs")
    corpus_equal = allocator["corpus"] == native["corpus"]
    output_equal = allocator["output_sha256"] == native["output_sha256"]
    # The allocator lane intentionally uses three samples while the native
    # lane uses one hundred.  Both validators already require each digest
    # vector to be constant within its report, so compare the stable identity
    # value rather than the differently sized vectors.
    source_output_equal = allocator["source_output_sha256"][0] == native["source_output_sha256"][0]
    semantic_equal = allocator["semantic_sha256"][0] == native["semantic_sha256"][0]
    untouched_equal = allocator["untouched_member_count"] == native["untouched_member_count"] and allocator["untouched_member_sha256"][0] == native["untouched_member_sha256"][0]
    require(corpus_equal, f"{allocator['name']} corpus differs from native report")
    require(output_equal, f"{allocator['name']} output digest differs from native report")
    require(source_output_equal, f"{allocator['name']} source output digests differ from native report")
    require(semantic_equal, f"{allocator['name']} semantic digests differ from native report")
    require(untouched_equal, f"{allocator['name']} untouched-member identity differs from native report")
    return {
        "native_name": native["name"],
        "corpus_equal": corpus_equal,
        "output_sha256_equal": output_equal,
        "source_output_sha256_equal": source_output_equal,
        "semantic_sha256_equal": semantic_equal,
        "untouched_member_identity_equal": untouched_equal,
        "allocator_timing_compared": False,
    }


def repeat_record(first: dict[str, Any], second: dict[str, Any]) -> dict[str, Any]:
    require(first["shape"] == second["shape"], "allocator repeat pairing has different shapes")
    same_identity = {
        "corpus_equal": first["corpus"] == second["corpus"],
        "output_sha256_equal": first["output_sha256"] == second["output_sha256"],
        "source_output_sha256_equal": first["source_output_sha256"] == second["source_output_sha256"],
        "semantic_sha256_equal": first["semantic_sha256"] == second["semantic_sha256"],
        "untouched_member_identity_equal": first["untouched_member_count"] == second["untouched_member_count"] and first["untouched_member_sha256"] == second["untouched_member_sha256"],
    }
    require(all(same_identity.values()), f"allocator repeats disagree for {first['shape']}")
    phase_shapes_equal = set(first["phase_metrics"]) == set(second["phase_metrics"]) == set(PHASES)
    require(phase_shapes_equal, f"allocator repeat phase inventory differs for {first['shape']}")
    metric_fields_equal = all(
        set(first["phase_metrics"][phase]["metrics"]) == set(second["phase_metrics"][phase]["metrics"]) == set(SAMPLE_FIELDS)
        for phase in PHASES
    )
    require(metric_fields_equal, f"allocator repeat metric inventory differs for {first['shape']}")
    return {
        "shape": first["shape"],
        "repeats": [first["repeat"], second["repeat"]],
        **same_identity,
        "phase_inventory_equal": phase_shapes_equal,
        "metric_inventory_equal": metric_fields_equal,
        "sample_count": {phase: [first["phase_metrics"][phase]["sample_count"], second["phase_metrics"][phase]["sample_count"]] for phase in PHASES},
        "allocator_values_are_reported_per_phase": True,
        "whole_operation_peak_not_derived": True,
    }


def relative_repeat_records(allocator_runs: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Recheck exact relative gauges over both three-sample repeats.

    These are phase-local differences.  In particular, ``peak_above_start``
    uses each region's own live start and never adds peaks from adjacent
    phases into a whole-operation number.
    """

    records: list[dict[str, Any]] = []
    relative_fields = (
        "allocation_calls",
        "reallocation_calls",
        "allocated_bytes",
        "deallocated_bytes",
        "failed_allocation_calls",
        "net_live",
        "peak_above_start",
    )
    for shape in SHAPES:
        runs = sorted((run for run in allocator_runs if run["shape"] == shape), key=lambda run: run["repeat"])
        require(len(runs) == 2, f"relative repeat check is missing a repeat for {shape}")
        for phase in PHASES:
            metrics = runs[0]["phase_metrics"][phase]["metrics"]
            vectors: dict[str, list[int]] = {
                field: list(metrics[field]["samples"])
                for field in relative_fields[:5]
            }
            for run in runs[1:]:
                next_metrics = run["phase_metrics"][phase]["metrics"]
                for field in relative_fields[:5]:
                    vectors[field].extend(next_metrics[field]["samples"])
            live_before: list[int] = []
            live_after: list[int] = []
            region_peak: list[int] = []
            for run in runs:
                phase_metrics = run["phase_metrics"][phase]["metrics"]
                live_before.extend(phase_metrics["live_bytes_before"]["samples"])
                live_after.extend(phase_metrics["live_bytes_after"]["samples"])
                region_peak.extend(phase_metrics["region_peak_live_bytes"]["samples"])
            vectors["net_live"] = [after - before for before, after in zip(live_before, live_after)]
            vectors["peak_above_start"] = [peak - before for before, peak in zip(live_before, region_peak)]
            require(all(len(values) == 6 for values in vectors.values()), f"relative repeat vector cardinality differs for {shape}/{phase}")
            require(all(len(set(values)) == 1 for values in vectors.values()), f"relative repeat samples differ for {shape}/{phase}")
            records.append({
                "shape": shape,
                "phase": phase,
                "samples": 6,
                "metrics": {field: values[0] for field, values in vectors.items()},
                "vectors": vectors,
                "all_six_samples_identical": True,
                "whole_operation_peak_not_derived": True,
            })
    return records


def missing_artifacts() -> list[str]:
    expected = ["source-manifest.json", "build-alloc.json", "build.json"]
    for repeat in REPEATS:
        for shape in SHAPES:
            expected += [
                f"{expected_alloc_name(repeat, shape)}{suffix}"
                for suffix in (".json", ".receipt.json", ".stdout", ".stderr")
            ]
            expected += [
                f"{expected_native_name(repeat, shape)}{suffix}"
                for suffix in (".json", ".receipt.json", ".stdout", ".stderr")
            ]
    return [name for name in expected if not (HERE / name).is_file()]


def analyze() -> dict[str, Any]:
    plan = plan_identity()
    missing = missing_artifacts()
    if missing:
        return {
            "schema_version": 1,
            "status": "pending",
            "scope": "0705 separate allocator diagnostics; no allocator timing claim",
            "plan": plan,
            "missing_artifacts": missing,
            "expected_allocator_runs": [expected_alloc_name(repeat, shape) for repeat in REPEATS for shape in SHAPES],
            "expected_native_runs": [expected_native_name(repeat, shape) for repeat in REPEATS for shape in SHAPES],
            "whole_operation_peak": None,
            "whole_operation_peak_policy": "not derived by summing phase peaks or live-region deltas",
        }
    manifest = source_manifest()
    alloc_build = validate_build_receipt(HERE / "build-alloc.json", allocator=True, manifest=manifest, plan=plan)
    native_build = validate_build_receipt(HERE / "build.json", allocator=False, manifest=manifest, plan=plan)
    require(alloc_build["binary_sha256"] != native_build["binary_sha256"], "allocator and native binaries unexpectedly share a digest")
    alloc_runs: list[dict[str, Any]] = []
    native_runs: list[dict[str, Any]] = []
    for repeat in REPEATS:
        for shape in SHAPES:
            alloc_name = expected_alloc_name(repeat, shape)
            native_name = expected_native_name(repeat, shape)
            alloc_receipt = validate_capture_receipt(alloc_name, shape=shape, repeat=repeat, allocator=True, binary=Path(alloc_build["binary"]["path"]), binary_sha256=alloc_build["binary_sha256"], manifest_sha256=manifest["sha256"], plan_sha256=plan["sha256"])
            native_receipt = validate_capture_receipt(native_name, shape=shape, repeat=repeat, allocator=False, binary=Path(native_build["binary"]["path"]), binary_sha256=native_build["binary_sha256"], manifest_sha256=manifest["sha256"], plan_sha256=plan["sha256"])
            alloc = validate_report(HERE / f"{alloc_name}.json", name=alloc_name, shape=shape, repeat=repeat, allocator=True, expected_samples=ALLOC_SAMPLES, expected_warmups=ALLOC_WARMUPS, binary=alloc_build["binary"], plan=plan)
            native = validate_report(HERE / f"{native_name}.json", name=native_name, shape=shape, repeat=repeat, allocator=False, expected_samples=NATIVE_SAMPLES, expected_warmups=NATIVE_WARMUPS, binary=native_build["binary"], plan=plan)
            alloc["receipt"] = alloc_receipt
            native["receipt"] = native_receipt
            alloc["native_parity"] = parity_record(alloc, native)
            alloc_runs.append(alloc)
            native_runs.append(native)
    repeats = []
    for shape in SHAPES:
        pair = [run for run in alloc_runs if run["shape"] == shape]
        require({run["repeat"] for run in pair} == set(REPEATS), f"allocator repeat set differs for {shape}")
        repeats.append(repeat_record(sorted(pair, key=lambda run: run["repeat"])[0], sorted(pair, key=lambda run: run["repeat"])[1]))
    return {
        "schema_version": 1,
        "status": "pass",
        "scope": "0705 separate allocator diagnostics; allocator-instrumented elapsed samples are not native timing evidence",
        "plan": plan,
        "source_manifest": manifest,
        "acquisition": {
            "allocations_script_sha256": sha(HERE / "allocations.py"),
            "capture_script_sha256": sha(HERE / "capture.py"),
            "allocator_build": alloc_build,
            "native_build": native_build,
            "separate_instrumentation": {
                "allocator_binary": "litchi-perf-baseline-alloc",
                "native_binary": "litchi-perf-baseline",
                "allocator_instrumentation": ALLOCATOR_INSTRUMENTATION,
                "native_instrumentation": "none",
                "binary_digests_distinct": True,
            },
        },
        "allocator_runs": alloc_runs,
        "native_runs": [
            {
                key: value
                for key, value in run.items()
                if key not in {"phase_metrics", "timing", "sink"}
            }
            for run in native_runs
        ],
        "repeat_consistency": repeats,
        "relative_repeat_consistency": relative_repeat_records(alloc_runs),
        "whole_operation_peak": None,
        "whole_operation_peak_policy": "not derived by summing phase peaks or live-region deltas; aggregate commit retains its own observer-ordered region peak",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=HERE / "allocation-analysis.json")
    args = parser.parse_args()
    try:
        result = analyze()
    except EvidenceError as error:
        result = {
            "schema_version": 1,
            "status": "error",
            "scope": "0705 separate allocator diagnostics; no allocator timing claim",
            "error": str(error),
            "whole_operation_peak": None,
            "whole_operation_peak_policy": "not derived by summing phase peaks or live-region deltas",
        }
        args.output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        print(f"0705 allocation evidence error: {error}")
        return 1
    args.output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(f"0705 allocation evidence {result['status']}: {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
