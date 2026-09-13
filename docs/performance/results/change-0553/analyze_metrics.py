#!/usr/bin/env python3
"""Validate the frozen 0553 compact XLSX source-proof capture.

This analyzer consumes the serial preflight, native, and allocator children
from the frozen driver.  The driver intentionally has no ``jobs()`` helper,
so the matrix is reconstructed from ``plan.json``.  Native elapsed samples
are reported with the workflow and its four source-backed phases.  The
allocator executable is a separate instrumentation binary: its phase
allocation vectors and absolute live-byte values are reported, while its
instrumented elapsed samples are never interpreted as latency evidence.
Guard, cap, quality, and adoption decisions remain outside this main-lane
analyzer.

The command is read-only with respect to captured evidence.  ``--output``
uses exclusive-create/identical-replay semantics, so a replay may reuse an
existing identical report but cannot silently replace it.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import importlib.util
import json
import math
import re
import sys
from pathlib import Path
from typing import Any


sys.dont_write_bytecode = True

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]

# The frozen capture driver owns the matrix and retained binary layout.  Both
# imports are deliberately bound to this bundle; no earlier campaign globals
# are consulted for paths or counts.
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))
import run as RUN  # noqa: E402


_binding_path = REPO / "tools" / "validate_perf_corpus_binding.py"
_binding_spec = importlib.util.spec_from_file_location("perf_corpus_binding_for_0553", _binding_path)
if _binding_spec is None or _binding_spec.loader is None:
    raise ImportError(f"cannot load corpus binding validator: {_binding_path}")
_BINDING = importlib.util.module_from_spec(_binding_spec)
_binding_spec.loader.exec_module(_BINDING)


# Reuse the established 0546 numerical helper for report-statistics
# calculation.  Its source-bound validator is tied to an older plan, so only
# the pure arithmetic helper is used here; this analyzer owns the 0553 matrix
# and all evidence bindings below.
_helper_path = REPO / "docs" / "performance" / "results" / "change-0546" / "integration" / "analyze.py"
_helper_spec = importlib.util.spec_from_file_location(
    "xlsx_0546_numerical_helper_for_0553", _helper_path
)
if _helper_spec is None or _helper_spec.loader is None:
    raise ImportError(f"cannot load numerical helper: {_helper_path}")
_HELPER = importlib.util.module_from_spec(_helper_spec)
_helper_spec.loader.exec_module(_HELPER)
_REPORT_STATS = _HELPER.BASE.report_stats
_helper_base_path = Path(_HELPER.HELPER).resolve()


SCHEMA = "xlsx_multisource_edit_metrics_0553_v1"
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
REVISION_RE = re.compile(r"^[0-9a-f]{40}$")
LANES = ("preflight", "native", "alloc")
PHASES = ("open_ns", "plan_ns", "commit_ns", "publication_ns", "reopen_ns")
TIMED_PHASES = ("open_ns", "plan_ns", "commit_ns", "publication_ns")
PHASE_LABELS = {
    "open_ns": "open",
    "plan_ns": "planning",
    "commit_ns": "commit",
    "publication_ns": "publication",
    "reopen_ns": "reopen",
}
ALLOCATION_PHASES = ("plan", "commit", "publication")
ALLOCATION_REPORT_FIELDS = (
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
    "incremental_region_peak_live_bytes",
    "retained_live_delta_bytes",
    "attributable_region_peak_live_bytes",
)
ALLOCATION_SAMPLE_FIELDS = (
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
REPORT_STATS = ("p50", "mean", "p95", "p99")
DRIFT_THRESHOLD_PERCENT = 5.0
PRIMARY_IMPROVEMENT_PERCENT = 3.0
LATENCY_GUARD_PERCENT = 5.0
HOST_SCOPE = "Accessible compiler processes; no host quiescence guarantee"
RECEIPT_KEYS = {
    "command",
    "start_utc",
    "end_utc",
    "seconds",
    "exit_code",
    "execution_stage",
    "execution_manifest_sha256",
    "binary_sha256",
    "source_manifest_sha256",
    "script_sha256",
    "plan_sha256",
    "environment",
    "artifacts",
}
RECEIPT_ENVIRONMENT_KEYS = {
    "RUSTFLAGS",
    "CARGO_ENCODED_RUSTFLAGS",
    "LD_PRELOAD",
    "MALLOC_CONF",
    "GLIBC_TUNABLES",
}
NORMAL_BINARY = "normal"
ALLOC_BINARY = "alloc"
BINARY_NAMES = {
    NORMAL_BINARY: "litchi-perf-baseline",
    ALLOC_BINARY: "litchi-perf-baseline-alloc",
}
ROOT_SOURCE_METRIC_VECTORS = {
    "read_calls",
    "read_bytes",
    "ordinary_payload_read_calls",
    "ordinary_payload_read_bytes",
    "max_in_flight_reads",
    "ordinary_payload_materializations",
}
XLSX_SOURCE_METRIC_VECTORS = {
    "source_read_calls",
    "source_read_bytes",
    "workbook_read_calls",
    "workbook_read_bytes",
    "selected_worksheet_read_calls",
    "selected_worksheet_read_bytes",
    "unselected_worksheet_read_calls",
    "unselected_worksheet_read_bytes",
    "payload_materializations",
    "cache_hits",
    "cache_cold_loads",
    "cache_waiter_joins",
    "cache_successful_loads",
    "cache_failed_loads",
    "cache_evictions",
    "cache_bypasses",
    "cache_oversized_bypasses",
    "cache_allocation_bypasses",
    "cache_in_flight_loads",
    "cache_retained_entries",
    "cache_retained_bytes",
    "cache_budget_memory_used",
    "cache_budget_reserved_bytes",
    "cache_budget_reservation_failures",
    "pre_publication_budget",
    "post_publication_budget",
    "budget_used_after_package_drop",
    "budget_used_after_handles_drop",
    "budget_objects_used_after_handles_drop",
}
XLSX_IDENTITY_EXCLUDED = {
    *XLSX_SOURCE_METRIC_VECTORS,
    "output_sha256",
    "semantic_sha256",
    "untouched_member_sha256",
    "output_budget_refusal",
    "partial_sink_verified",
}


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory capture artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {path}: {error}") from error


def sha256(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def is_hash(value: Any) -> bool:
    return isinstance(value, str) and SHA256_RE.fullmatch(value) is not None


def check_hash(value: Any, label: str) -> str:
    require(is_hash(value), f"{label} is not a lowercase SHA-256")
    return value


def nonnegative_integer(value: Any, label: str) -> None:
    require(
        isinstance(value, int) and not isinstance(value, bool) and value >= 0,
        f"{label} is not a nonnegative integer",
    )


def finite_number(value: Any, label: str) -> None:
    require(
        isinstance(value, (int, float)) and not isinstance(value, bool),
        f"{label} is not numeric",
    )
    require(math.isfinite(float(value)), f"{label} is not finite")


def timestamp(value: Any, label: str) -> _datetime.datetime:
    require(isinstance(value, str), f"{label} timestamp is missing")
    try:
        parsed = _datetime.datetime.fromisoformat(value)
    except ValueError as error:
        raise EvidenceError(f"{label} timestamp is malformed") from error
    require(parsed.tzinfo is not None, f"{label} timestamp has no timezone")
    return parsed


def validate_receipt_common(receipt: dict[str, Any], label: str) -> None:
    require(set(receipt) == RECEIPT_KEYS, f"{label} receipt schema differs")
    start = timestamp(receipt.get("start_utc"), f"{label}.start_utc")
    end = timestamp(receipt.get("end_utc"), f"{label}.end_utc")
    require(start < end, f"{label} timestamps are not increasing")
    finite_number(receipt.get("seconds"), f"{label}.seconds")
    require(receipt["seconds"] > 0.0, f"{label}.seconds is not positive")
    wall_seconds = (end - start).total_seconds()
    require(
        abs(float(receipt["seconds"]) - wall_seconds)
        <= max(0.25, wall_seconds * 0.02 + 0.05),
        f"{label}.seconds does not match its UTC interval",
    )
    environment = receipt.get("environment")
    require(isinstance(environment, dict)
            and set(environment) == RECEIPT_ENVIRONMENT_KEYS,
            f"{label}.environment fields differ")
    require(all(value is None for value in environment.values()),
            f"{label}.environment has uncontrolled values")
    require(isinstance(receipt.get("command"), list)
            and all(isinstance(item, str) and item for item in receipt["command"]),
            f"{label}.command is malformed")
    require(isinstance(receipt.get("exit_code"), int)
            and not isinstance(receipt["exit_code"], bool),
            f"{label}.exit_code is malformed")
    check_hash(receipt.get("execution_manifest_sha256"),
               f"{label}.execution_manifest_sha256")
    check_hash(receipt.get("source_manifest_sha256"),
               f"{label}.source_manifest_sha256")
    check_hash(receipt.get("script_sha256"), f"{label}.script_sha256")
    check_hash(receipt.get("plan_sha256"), f"{label}.plan_sha256")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{label}.artifacts is missing")
    for name, digest in artifacts.items():
        require(isinstance(name, str) and Path(name).name == name,
                f"{label}.artifacts has a non-local name")
        check_hash(digest, f"{label}.artifacts.{name}")


def vector(value: Any, count: int, label: str) -> list[Any]:
    require(isinstance(value, list), f"{label} is not a vector")
    require(len(value) == count, f"{label} has {len(value)} values, expected {count}")
    return value


def integer_vector(value: Any, count: int, label: str) -> list[int]:
    values = vector(value, count, label)
    for index, item in enumerate(values):
        nonnegative_integer(item, f"{label}[{index}]")
    return values  # type: ignore[return-value]


def midpoint(left: int, right: int) -> int:
    # This is the Rust benchmark's overflow-safe midpoint implementation.
    return left // 2 + right // 2 + (left % 2 + right % 2) // 2


def nearest_rank(values: list[int], percentile: int) -> int:
    index = min((percentile * len(values) + 99) // 100 - 1, len(values) - 1)
    return values[index]


def student_t_critical_95(degrees: int) -> float:
    values = (
        12.706,
        4.303,
        3.182,
        2.776,
        2.571,
        2.447,
        2.365,
        2.306,
        2.262,
        2.228,
        2.201,
        2.179,
        2.160,
        2.145,
        2.131,
        2.120,
        2.110,
        2.101,
        2.093,
        2.086,
        2.080,
        2.074,
        2.069,
        2.064,
        2.060,
        2.056,
        2.052,
        2.048,
        2.045,
        2.042,
    )
    if degrees == 0:
        return 0.0
    if degrees <= len(values):
        return values[degrees - 1]
    z = 1.959963984540054
    z2 = z * z
    z3 = z2 * z
    z5 = z3 * z2
    z7 = z5 * z2
    d = float(degrees)
    return (
        z
        + (z3 + z) / (4.0 * d)
        + (5.0 * z5 + 16.0 * z3 + 3.0 * z) / (96.0 * d * d)
        + (3.0 * z7 + 19.0 * z5 + 17.0 * z3 - 15.0 * z)
        / (384.0 * d * d * d)
    )


def stats(values: list[int], unit: str = "ns") -> dict[str, Any]:
    require(values, "statistics vector is empty")
    for index, value in enumerate(values):
        nonnegative_integer(value, f"statistics[{index}]")
    result = dict(_REPORT_STATS(list(values)))
    # The helper's arithmetic is unit agnostic even though its historical
    # envelope labels the values as nanoseconds.  Allocation summaries retain
    # the same deterministic statistics with their actual count/byte units.
    result["unit"] = unit
    return result


def signed_stats(values: list[int], unit: str = "bytes") -> dict[str, Any]:
    """Statistics for derived retained deltas, which may be negative."""

    require(values, "signed statistics vector is empty")
    for index, value in enumerate(values):
        require(isinstance(value, int) and not isinstance(value, bool),
                f"signed statistics[{index}] is not an integer")
    ordered = sorted(values)
    mean = sum(ordered) / len(ordered)
    if len(ordered) > 1:
        squared = sum((value - mean) ** 2 for value in ordered)
        deviation = math.sqrt(squared / (len(ordered) - 1))
        margin = (_HELPER.BASE.student_t_critical_95(len(ordered) - 1)
                  * deviation / math.sqrt(len(ordered)))
    else:
        deviation = 0.0
        margin = 0.0
    return {
        "unit": unit,
        "samples": ordered,
        "p50": (ordered[(len(ordered) - 1) // 2] + ordered[len(ordered) // 2]) // 2,
        "p95": ordered[min(math.ceil(len(ordered) * 0.95) - 1, len(ordered) - 1)],
        "p99": ordered[min(math.ceil(len(ordered) * 0.99) - 1, len(ordered) - 1)],
        "min": ordered[0],
        "max": ordered[-1],
        "mean": mean,
        "standard_deviation": deviation,
        "confidence_interval_95": {
            "method": "two-sided Student's t interval for the mean",
            "lower": mean - margin,
            "upper": mean + margin,
        },
    }


def compare_number(actual: Any, expected: Any, label: str) -> None:
    finite_number(actual, label)
    finite_number(expected, f"expected {label}")
    require(
        math.isclose(float(actual), float(expected), rel_tol=1e-12, abs_tol=1e-9),
        f"{label}: {actual!r} != {expected!r}",
    )


def validate_report_stats(actual: Any, values: list[int], label: str,
                          unit: str = "ns") -> dict[str, Any]:
    require(isinstance(actual, dict), f"{label} is not an object")
    expected = stats(values, unit)
    require(actual.get("unit") == unit, f"{label}.unit differs")
    require(actual.get("samples") == expected["samples"], f"{label}.samples differ")
    for key in (
        "min",
        "p50",
        "p95",
        "p99",
        "max",
        "mean",
        "standard_deviation",
    ):
        require(key in actual, f"{label} is missing {key}")
        compare_number(actual[key], expected[key], f"{label}.{key}")
    interval = actual.get("confidence_interval_95")
    require(isinstance(interval, dict), f"{label}.confidence_interval_95 is not an object")
    require(
        interval.get("method") == expected["confidence_interval_95"]["method"],
        f"{label}.confidence_interval_95 method differs",
    )
    for key in ("lower", "upper"):
        compare_number(
            interval.get(key), expected["confidence_interval_95"][key],
            f"{label}.confidence_interval_95.{key}",
        )
    return expected


def plan_data() -> dict[str, Any]:
    plan = read_json(HERE / "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    revision = plan.get("revision")
    require(isinstance(revision, str) and REVISION_RE.fullmatch(revision) is not None,
            "plan revision is not a lowercase commit hash")
    require(
        plan.get("scope") ==
        "Matched commit-local compact source-cell proof experiment for source-backed XLSX MultiSourceEdit",
        "plan scope differs",
    )
    require(plan.get("priority") ==
            "OLE2/OOXML first; ODF deferred until that goal completes; iWork excluded",
            "plan priority differs")
    require(plan.get("cpu") == 2, "plan CPU differs from pinned CPU")
    require(plan.get("owned_paths") == ["/home/zhuhe/litchi-goal-0553-target"],
            "plan owned target differs")
    require(plan.get("cases") == [
        "xlsx_source_backed_cell_values_one_edit_save",
        "xlsx_source_backed_cell_values_one_percent_edit_save",
        "xlsx_source_backed_managed_cell_values_one_edit_save",
        "xlsx_source_backed_managed_cell_values_one_percent_edit_save",
    ], "plan case matrix differs")
    require(plan.get("shapes") == ["medium", "dense-sparse", "noncompact", "vendor-extension"],
            "plan shape matrix differs")
    require(plan.get("native") == {"repeats": 2, "samples": 200, "warmup": 20},
            "plan native counts differ")
    require(plan.get("alloc") == {"repeats": 2, "samples": 20, "warmup": 3},
            "plan allocator counts differ")
    require(plan.get("profile") == {
        "repeats": 2,
        "samples": 1,
        "warmup": 0,
        "case": "xlsx_source_backed_cell_values_one_percent_edit_save",
        "owner": "litchi_xlsx::cell_values::source::MultiSourceEdit::commit",
        "parent": "litchi_perf_baseline::run_xlsx_cell_values_edit_save",
    }, "plan profile differs")
    require(plan.get("guard") == {
        "shapes": ["medium", "dense-sparse"],
        "cases": ["valid", "late-validator", "late-raw"],
        "repeats": 2,
        "native_samples": 200,
        "native_warmup": 20,
        "alloc_samples": 20,
        "alloc_warmup": 3,
    }, "plan guard matrix differs")
    require(plan.get("cap") == {
        "sizes": [1, 2, 160, 164, 256],
        "repeats": 2,
        "samples": 200,
        "warmup": 20,
    }, "plan cap matrix differs")
    expected_order = [
        "baseline r1",
        "candidate r1",
        "candidate r2",
        "retained baseline r2 under candidate source manifest",
    ]
    require(plan.get("native_order") == expected_order, "plan native order differs")
    require(plan.get("allocator_order") == expected_order, "plan allocator order differs")
    expected_admission = {
        "primary": "Every one-percent case/shape/repeat, including managed, workflow p50 and mean improve at least 3%.",
        "latency_guards": "Every one-cell case/shape/repeat and valid-noop planning/cap control p50 and mean <= 1.05x matched baseline.",
        "refusals": "Exact errors/retry identities unchanged; each invalid guard p50 and mean <= max(1.05x corresponding baseline invalid,2x baseline valid); each invalid attributable peak <= 1.10x baseline valid attributable peak.",
        "workflow_memory": "For each main case/shape/repeat, max planning/commit/publication absolute region peak minus planning live_bytes_before <= 1.05x baseline; process peakRSS <= 1.05x baseline. Use allocation-instrumented comparisons only for allocation metrics.",
        "noop_memory": "Valid planning guard region peak minus live_bytes_before <= 1.05x baseline. Decline if durable metadata cannot fit representative no-op memory envelope.",
        "allocation": "Workflow sum allocated bytes <= 1.05x baseline; allocation/reallocation calls and per-phase retained-after values reported individually. Any >5% adverse/drift row requires explicit review; no hidden geometric mean.",
        "proof_resources": "Explicit checked byte/event caps and fallible reservation before growth required; immutable original-source binding and complete fallback. Concrete candidate cap/source hash frozen before application.",
        "correctness": "Exact scanner/writer differential tests, source fallback/error-order, no-op, managed execution/cancellation, source identity, inverse/clone/re-edit, preservation and quality gates required.",
        "profile": "Exact commit Ir must decrease for every shape/repeat if pilot native/memory gates pass; absence of scan symbol alone is insufficient.",
        "disposition": "All mandatory gates required; reject and restore baseline on failure. No exception for a necessary enabler without measured justification.",
    }
    require(plan.get("admission") == expected_admission, "plan admission text differs")
    require(plan.get("limitations") == [
        "No hardware/cold/provider/scaling claim from this campaign.",
        "Current single-sheet SourceEdit API is outside measured owner.",
        "Profiles conditional on native/memory pilot; all failures retained.",
        "Historical guards use no-op oracle after planning interval; full no-op publication time is not measured.",
    ], "plan limitations differ")
    return plan


def frozen_inputs() -> dict[str, Any]:
    frozen = read_json(HERE / "frozen-inputs.json")
    require(isinstance(frozen, dict), "frozen-inputs.json is not an object")
    require(set(frozen) == {"frozen_utc", "files"},
            "frozen input envelope differs")
    files = frozen.get("files")
    require(isinstance(files, dict), "frozen input file map is missing")
    expected_names = {
        "docs/performance/results/change-0553/run.py",
        "docs/performance/results/change-0553/capture.py",
        "docs/performance/results/change-0553/plan.json",
        "docs/performance/results/change-0553/adr-manifest.json",
    }
    require(set(files) == expected_names, "frozen input file inventory differs")
    for name, expected in files.items():
        require(is_hash(expected), f"frozen input hash for {name} is malformed")
        require(expected == sha256(REPO / name), f"frozen input {name} has changed")
    return frozen


def _cleanup_binary_hash(stage: str, kind: str, expected_path: Path,
                         plan: dict[str, Any]) -> str | None:
    for cleanup_path in (HERE / stage / "cleanup.json", HERE / "cleanup.json"):
        if not cleanup_path.is_file():
            continue
        cleanup = read_json(cleanup_path)
        if not isinstance(cleanup, dict):
            continue
        if cleanup.get("owned_paths_absent") is not True:
            continue
        if cleanup.get("accessible_process_references") != []:
            continue
        if cleanup.get("removed") != plan.get("owned_paths"):
            continue
        by_kind = cleanup.get("binary_sha256_by_kind")
        if not isinstance(by_kind, dict):
            continue
        for key in (f"{stage}/{kind}", str(expected_path), expected_path.name, kind):
            value = by_kind.get(key)
            if is_hash(value):
                return value
    return None


def _expected_build_command(kind: str, plan: dict[str, Any]) -> list[str]:
    target = Path(plan["owned_paths"][0])
    command = [
        "env", "TMPDIR=" + str(target / "tmp"), "CARGO_BUILD_JOBS=2",
        "CARGO_INCREMENTAL=0", "cargo", "build", "--release", "--locked",
        "--manifest-path", "tools/perf-baseline/Cargo.toml", "--bin",
        BINARY_NAMES[kind], "--target-dir", str(target),
    ]
    if kind == ALLOC_BINARY:
        command += ["--features", "allocator-metrics"]
    return command


def binary_metadata(stage: str, kind: str, plan: dict[str, Any]) -> dict[str, Any]:
    require(kind in (NORMAL_BINARY, ALLOC_BINARY), f"unknown binary kind {kind}")
    folder = HERE / stage
    descriptor_path = folder / f"binary-{kind}.json"
    descriptor = read_json(descriptor_path)
    require(isinstance(descriptor, dict), f"{descriptor_path.name} is not an object")
    path = Path(descriptor.get("path", ""))
    expected_path = Path(RUN.SCRATCH_ROOT) / stage / kind
    require(path == expected_path, f"{descriptor_path.name} path differs from run.py")
    digest = check_hash(descriptor.get("sha256"), f"{descriptor_path.name}.sha256")
    nonnegative_integer(descriptor.get("bytes"), f"{descriptor_path.name}.bytes")
    require(descriptor["bytes"] > 0, f"{descriptor_path.name}.bytes is zero")
    require(not path.is_symlink(), f"{descriptor_path.name} points to a symlink")
    if path.exists():
        require(path.is_file(), f"{descriptor_path.name} path is not a regular file")
        require(sha256(path) == digest, f"{descriptor_path.name} binary hash differs")
        require(path.stat().st_size == descriptor["bytes"],
                f"{descriptor_path.name} binary size differs")
    else:
        require(_cleanup_binary_hash(stage, kind, path, plan) == digest,
                f"{descriptor_path.name} has no post-cleanup binary custody")

    manifest = folder / "source-manifest.json"
    require(descriptor.get("source_manifest_sha256") == sha256(manifest),
            f"{descriptor_path.name} source manifest differs")
    build_name = f"build-{kind}"
    build_path = folder / f"{build_name}.receipt.json"
    require(descriptor.get("build_receipt_sha256") == sha256(build_path),
            f"{descriptor_path.name} build receipt hash differs")
    build = read_json(build_path)
    require(isinstance(build, dict), f"{build_path.name} is not an object")
    validate_receipt_common(build, build_path.name)
    require(build.get("exit_code") == 0 and build.get("binary_sha256") is None,
            f"{build_path.name} is not a successful build receipt")
    require(build.get("execution_stage") == stage,
            f"{build_path.name} execution stage differs")
    require(build.get("execution_manifest_sha256") == sha256(manifest),
            f"{build_path.name} execution manifest differs")
    require(build.get("script_sha256") == sha256(HERE / "run.py"),
            f"{build_path.name} script hash differs")
    require(build.get("plan_sha256") == sha256(HERE / "plan.json"),
            f"{build_path.name} plan hash differs")
    require(build.get("source_manifest_sha256") == sha256(manifest),
            f"{build_path.name} source manifest hash differs")
    require(build.get("command") == _expected_build_command(kind, plan),
            f"{build_path.name} command differs from run.py")
    artifacts = build.get("artifacts")
    require(isinstance(artifacts, dict), f"{build_path.name} artifact inventory is missing")
    expected_artifacts = {
        f"{build_name}.host.json",
        f"{build_name}.stdout",
        f"{build_name}.stderr",
    }
    require(set(artifacts) == expected_artifacts,
            f"{build_path.name} artifact inventory differs")
    for name, digest_value in artifacts.items():
        check_hash(digest_value, f"{build_path.name}/{name}")
        artifact = folder / name
        require(artifact.is_file() and not artifact.is_symlink(),
                f"{build_path.name} artifact is missing: {name}")
        require(sha256(artifact) == digest_value,
                f"{build_path.name} artifact hash differs: {name}")
        if name.endswith(".host.json"):
            check_host(artifact)
    return {
        "path": str(path),
        "sha256": digest,
        "bytes": descriptor["bytes"],
        "build_receipt_sha256": descriptor["build_receipt_sha256"],
        "source_manifest_sha256": descriptor["source_manifest_sha256"],
        "binary": BINARY_NAMES[kind],
        "stage": stage,
    }


def jobs_for(plan: dict[str, Any], stage: str, lane: str) -> list[dict[str, Any]]:
    require(stage in ("baseline", "candidate"), "unknown capture stage")
    require(lane in LANES, "unknown capture lane")
    repeats = (1,) if lane == "preflight" else (1, 2)
    jobs: list[dict[str, Any]] = []
    for repeat in repeats:
        for shape in plan["shapes"][::1 if repeat == 1 else -1]:
            for index, case in enumerate(plan["cases"]):
                jobs.append({
                    "name": f"{lane}-r{repeat}-{shape}-c{index}",
                    "lane": lane,
                    "repeat": repeat,
                    "case": case,
                    "shape": shape,
                    "samples": 1 if lane == "preflight" else plan[lane]["samples"],
                    "warmup": 0 if lane == "preflight" else plan[lane]["warmup"],
                    "stage": stage,
                    # Baseline r2 is deliberately retained in baseline/ but
                    # executed after candidate freeze against candidate's
                    # execution manifest.
                    "execution_stage": (
                        "candidate" if stage == "baseline" and repeat == 2 else stage
                    ),
                })
    return jobs


def expected_command(job: dict[str, Any], binary: dict[str, Any]) -> list[str]:
    folder = HERE / job["stage"]
    return [
        "taskset",
        "-c",
        "2",
        "/usr/bin/time",
        "-v",
        binary["path"],
        "--case",
        job["case"],
        "--xlsx-cell-crud-shape",
        job["shape"],
        "--samples",
        str(job["samples"]),
        "--warmup",
        str(job["warmup"]),
        "--json",
        str(folder / (job["name"] + ".json")),
        "--corpus-manifest",
        str(folder / (job["name"] + ".catalog.json")),
    ]


def check_host(path: Path) -> None:
    host = read_json(path)
    require(isinstance(host, dict), f"{path.name} host sidecar is not an object")
    require(host.get("scope") == HOST_SCOPE, f"{path.name} host scope differs")
    processes = host.get("compiler_processes")
    require(isinstance(processes, list), f"{path.name} compiler process list is missing")
    for index, process in enumerate(processes):
        require(isinstance(process, dict), f"{path.name} process {index} is not an object")
        nonnegative_integer(process.get("pid"), f"{path.name} process {index}.pid")
        require(process["pid"] > 0, f"{path.name} process {index}.pid is zero")
        require(process.get("comm") in ("cargo", "rustc"),
                f"{path.name} process {index}.comm is unexpected")
        require(isinstance(process.get("cwd"), str) and process["cwd"],
                f"{path.name} process {index}.cwd is missing")


def validate_time_v(path: Path, label: str) -> dict[str, Any]:
    """Validate the GNU time sidecar and retain peak RSS for memory gates."""

    try:
        text = path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as error:
        raise EvidenceError(f"cannot read {path}: {error}") from error
    values: dict[str, Any] = {}
    patterns = {
        "user_seconds": r"^\s*User time \(seconds\):\s*([0-9]+(?:\.[0-9]+)?)\s*$",
        "system_seconds": r"^\s*System time \(seconds\):\s*([0-9]+(?:\.[0-9]+)?)\s*$",
        "elapsed_seconds": r"^\s*Elapsed .*:\s*(?:(\d+):)?(\d+):(\d{2})\.(\d{2})\s*$",
        "max_rss_kib": r"^\s*Maximum resident set size \(kbytes\):\s*([0-9]+)\s*$",
        "exit_status": r"^\s*Exit status:\s*([0-9]+)\s*$",
    }
    for line in text.splitlines():
        for name, pattern in patterns.items():
            match = re.match(pattern, line)
            if not match:
                continue
            if name == "elapsed_seconds":
                values[name] = (
                    int(match.group(1) or 0) * 3600
                    + int(match.group(2)) * 60
                    + int(match.group(3))
                    + int(match.group(4)) / 100
                )
            elif name in ("user_seconds", "system_seconds"):
                values[name] = float(match.group(1))
            else:
                values[name] = int(match.group(1))
    for name in ("user_seconds", "system_seconds", "elapsed_seconds"):
        finite_number(values.get(name), f"{label}.{name}")
        require(values[name] >= 0, f"{label}.{name} is negative")
    nonnegative_integer(values.get("max_rss_kib"), f"{label}.max_rss_kib")
    require(values.get("exit_status") == 0, f"{label}.Exit status is not zero")
    return values


def allocation_sample(value: Any, label: str, measured: bool) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is not an object")
    require(value.get("scope") == "operation_global_system_allocator",
            f"{label}.scope differs")
    status = value.get("status")
    if not measured:
        require(status == "unavailable", f"{label} must be unavailable")
        require(all(field not in value for field in ALLOCATION_SAMPLE_FIELDS),
                f"{label} contains unavailable numeric fields")
        return {"status": status, "scope": value["scope"]}
    require(status == "measured", f"{label}.status is not measured")
    values: dict[str, int] = {}
    for field in ALLOCATION_SAMPLE_FIELDS:
        nonnegative_integer(value.get(field), f"{label}.{field}")
        values[field] = value[field]
    require(values["failed_allocation_calls"] == 0,
            f"{label} recorded a failed allocation")
    require(values["live_bytes_before"] + values["allocated_bytes"]
            == values["live_bytes_after"] + values["deallocated_bytes"],
            f"{label} live-byte balance does not reconcile")
    require(values["peak_live_bytes_before"] >= values["live_bytes_before"],
            f"{label} pre-operation peak is below live bytes")
    require(values["peak_live_bytes_after"] >= values["peak_live_bytes_before"],
            f"{label} peak moved backwards")
    require(values["peak_live_bytes_after"] >= values["live_bytes_after"],
            f"{label} post-operation peak is below live bytes")
    require(values["region_peak_live_bytes"] >= max(
        values["live_bytes_before"], values["live_bytes_after"]
    ), f"{label} region peak is below live bytes")
    require(values["region_peak_live_bytes"] <= values["peak_live_bytes_after"],
            f"{label} region peak exceeds process peak")
    return {
        "status": status,
        "scope": value["scope"],
        **values,
        "incremental_region_peak_live_bytes": (
            values["region_peak_live_bytes"] - values["live_bytes_before"]
        ),
        "retained_live_delta_bytes": (
            values["live_bytes_after"] - values["live_bytes_before"]
        ),
    }


def _validate_recursive_value(value: Any, label: str) -> None:
    if isinstance(value, bool) or value is None or isinstance(value, str):
        return
    if isinstance(value, int):
        nonnegative_integer(value, label)
        return
    if isinstance(value, float):
        finite_number(value, label)
        require(value >= 0.0, f"{label} is negative")
        return
    if isinstance(value, list):
        for index, item in enumerate(value):
            _validate_recursive_value(item, f"{label}[{index}]")
        return
    if isinstance(value, dict):
        for key, item in value.items():
            require(isinstance(key, str), f"{label} contains a non-string key")
            _validate_recursive_value(item, f"{label}.{key}")
        return
    raise EvidenceError(f"{label} contains an unsupported JSON value")


def canonical_identity_value(value: Any, count: int, label: str) -> Any:
    """Collapse an identity vector only after proving it is constant."""

    if isinstance(value, list):
        values = vector(value, count, label)
        for index, item in enumerate(values):
            _validate_recursive_value(item, f"{label}[{index}]")
        require(all(item == values[0] for item in values),
                f"{label} varies across samples")
        return canonical_identity_value(values[0], count, f"{label}[0]")
    if isinstance(value, dict):
        return {
            key: canonical_identity_value(value[key], count, f"{label}.{key}")
            for key in sorted(value)
        }
    return value


def source_identity(source: dict[str, Any], count: int) -> dict[str, Any]:
    """Return semantic source identity after excluding mutable diagnostics.

    The compact source proof is expected to alter read/cache counters, so
    those vectors are validated for shape and arithmetic but are compared as
    diagnostics rather than as a cross-stage identity.
    """

    require(isinstance(source, dict), "result.source is not an object")
    result: dict[str, Any] = {}
    for key in sorted(source):
        value = source[key]
        if key in ROOT_SOURCE_METRIC_VECTORS:
            vector(value, count, f"source.{key}")
            continue
        if key != "xlsx_cell_values":
            result[key] = canonical_identity_value(value, count, f"source.{key}")
            continue
        require(isinstance(value, dict), "source.xlsx_cell_values is not an object")
        inner: dict[str, Any] = {}
        for inner_key in sorted(value):
            # These are either timing measurements or allocator-instrumented
            # measurements.  They are validated separately and cannot be an
            # exact logical identity between normal and allocator binaries.
            if inner_key in PHASES or inner_key.endswith("_allocation_metrics"):
                vector(value[inner_key], count,
                       f"source.xlsx_cell_values.{inner_key}")
                continue
            if inner_key in XLSX_IDENTITY_EXCLUDED:
                if isinstance(value[inner_key], list):
                    vector(value[inner_key], count,
                           f"source.xlsx_cell_values.{inner_key}")
                continue
            inner[inner_key] = canonical_identity_value(
                value[inner_key], count, f"source.xlsx_cell_values.{inner_key}"
            )
        result[key] = inner
    return result


def validate_sink(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}.sink is not an object")
    for key in ("accepted_bytes", "write_calls", "largest_write"):
        nonnegative_integer(value.get(key), f"{label}.sink.{key}")
    buckets = value.get("write_size_buckets")
    require(isinstance(buckets, dict), f"{label}.sink.write_size_buckets is missing")
    expected = {
        "bytes_0",
        "bytes_1_to_512",
        "bytes_513_to_4096",
        "bytes_4097_to_16384",
        "bytes_16385_to_65536",
        "bytes_over_65536",
    }
    require(set(buckets) == expected, f"{label}.sink bucket fields differ")
    for key, item in buckets.items():
        nonnegative_integer(item, f"{label}.sink.write_size_buckets.{key}")
    require(value["largest_write"] <= 65_536,
            f"{label}.sink largest write exceeds bound")
    require(sum(buckets.values()) == value["write_calls"],
            f"{label}.sink buckets do not reconcile")
    result = {
        "accepted_bytes": value["accepted_bytes"],
        "write_calls": value["write_calls"],
        "largest_write": value["largest_write"],
        "write_size_buckets": {key: buckets[key] for key in sorted(buckets)},
    }
    # These fields are optional in SinkSummary.  XLSX CountingSink currently
    # retains its verification Vec without advertising it as a retention
    # metric; preserve that absence rather than inferring it from
    # accepted_bytes.
    for field in ("retained_output_bytes", "retained_authoring_window_bytes"):
        if field in value:
            nonnegative_integer(value[field], f"{label}.sink.{field}")
            result[field] = value[field]
    return result


def validate_operation_metrics(value: Any, count: int, sink: dict[str, Any],
                               label: str) -> None:
    require(isinstance(value, dict), f"{label}.operation_metrics is not an object")
    require(value.get("sample_count") == count,
            f"{label}.operation_metrics sample count differs")
    require(value.get("sample_indices") == list(range(count)),
            f"{label}.operation_metrics sample indices differ")
    require(value.get("alignment") == "elapsed_ns.samples_by_elapsed_then_sample_index",
            f"{label}.operation_metrics alignment differs")
    require(value.get("latency_claim") == "comparable_timed_operation",
            f"{label}.operation_metrics latency claim differs")
    sink_metrics = value.get("sink")
    require(isinstance(sink_metrics, dict), f"{label}.operation_metrics.sink is missing")
    require(sink_metrics.get("write_status") == "measured",
            f"{label}.operation_metrics sink write status differs")
    for field in ("accepted_bytes", "write_calls", "largest_write"):
        item = sink_metrics.get(field)
        require(isinstance(item, dict), f"{label}.operation_metrics.sink.{field} missing")
        require(item.get("status") == "measured",
                f"{label}.operation_metrics.sink.{field} status differs")
        values = integer_vector(item.get("values"), count,
                                f"{label}.operation_metrics.sink.{field}.values")
        require(all(item_value == sink[field] for item_value in values),
                f"{label}.operation_metrics.sink.{field} disagrees with sink")
    bucket_metrics = sink_metrics.get("write_size_buckets")
    require(isinstance(bucket_metrics, dict),
            f"{label}.operation_metrics sink buckets missing")
    for field, expected in sink["write_size_buckets"].items():
        item = bucket_metrics.get(field)
        require(isinstance(item, dict),
                f"{label}.operation_metrics.sink.write_size_buckets.{field} missing")
        require(item.get("status") == "measured",
                f"{label}.operation_metrics.sink.write_size_buckets.{field} status differs")
        values = integer_vector(item.get("values"), count,
                                f"{label}.operation_metrics.sink.write_size_buckets.{field}.values")
        require(all(item_value == expected for item_value in values),
                f"{label}.operation_metrics bucket {field} disagrees with sink")


def validate_budget_snapshot(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is not an object")
    fields = {
        "input_bytes_used", "input_bytes_limit", "output_bytes_used", "output_bytes_limit",
        "work_used", "work_limit", "objects_used", "objects_limit",
        "catalog_reserved_objects", "cache_reserved_objects",
    }
    require(set(value) == fields, f"{label} fields differ")
    for field, item in value.items():
        if item is None:
            require(field.endswith("_limit") or field.endswith("_objects"),
                    f"{label}.{field} is unexpectedly null")
        else:
            nonnegative_integer(item, f"{label}.{field}")
    return value


def validate_refusal(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is not an object")
    fields = {
        "successful_output_ceiling", "first_output_request_bytes", "one_under_output_limit",
        "accepted_output_bytes", "source_read_calls", "source_read_bytes", "output_bytes_used",
        "typed_output_resource_refusal", "zero_output_verified", "source_identity_preserved",
    }
    require(set(value) == fields, f"{label} fields differ")
    bool_fields = {"typed_output_resource_refusal", "zero_output_verified", "source_identity_preserved"}
    for field, item in value.items():
        if field in bool_fields:
            require(isinstance(item, bool), f"{label}.{field} is not boolean")
        else:
            nonnegative_integer(item, f"{label}.{field}")
    return value


def validate_source(
    source: Any,
    count: int,
    job: dict[str, Any],
    corpus: dict[str, Any],
    measured: bool,
    label: str,
) -> tuple[dict[str, Any], dict[str, Any], dict[str, list[int]], dict[str, list[dict[str, Any]]]]:
    require(isinstance(source, dict), f"{label}.source is not an object")
    for name in ROOT_SOURCE_METRIC_VECTORS:
        integer_vector(source.get(name), count, f"{label}.source.{name}")
    xlsx = source.get("xlsx_cell_values")
    require(isinstance(xlsx, dict), f"{label}.source.xlsx_cell_values is missing")
    managed_case = "managed" in job["case"]
    expected_impl = "managed-source-backed" if managed_case else "source-backed"
    expected_mode = "managed-budget" if managed_case else "unmanaged-control"
    require(xlsx.get("implementation") == expected_impl,
            f"{label}.implementation differs")
    require(xlsx.get("cache_mode") == expected_mode, f"{label}.cache_mode differs")
    require(xlsx.get("cache_budget_managed") is managed_case,
            f"{label}.cache_budget_managed differs")
    one_edit = "one_edit" in job["case"]
    xlsx_manifest = corpus.get("xlsx")
    require(isinstance(xlsx_manifest, dict), f"{label}.corpus.xlsx is missing")
    expected_update_count = 1 if one_edit else xlsx_manifest.get("one_percent_update_count")
    nonnegative_integer(expected_update_count, f"{label}.expected_update_count")
    require(xlsx.get("update_count") == expected_update_count,
            f"{label}.update_count differs")
    expected_touched = 1 if one_edit else xlsx_manifest.get("sheet_count")
    nonnegative_integer(expected_touched, f"{label}.expected_touched_worksheets")
    require(xlsx.get("selected_worksheet_count") == expected_touched,
            f"{label}.selected_worksheet_count differs")
    require(isinstance(xlsx.get("timing_scope"), str)
            and "commit" in xlsx["timing_scope"]
            and "publication" in xlsx["timing_scope"],
            f"{label}.timing_scope does not include commit/publication")

    phase_values: dict[str, list[int]] = {}
    for name in PHASES:
        phase_values[name] = integer_vector(xlsx.get(name), count,
                                            f"{label}.source.xlsx_cell_values.{name}")
    allocation_values: dict[str, list[dict[str, Any]]] = {}
    for phase in ALLOCATION_PHASES:
        field = f"{phase}_allocation_metrics"
        samples = vector(xlsx.get(field), count,
                         f"{label}.source.xlsx_cell_values.{field}")
        allocation_values[phase] = [
            allocation_sample(item, f"{label}.{field}[{index}]", measured)
            for index, item in enumerate(samples)
        ]

    # All remaining source arrays are diagnostics, but their cardinality and
    # scalar types are still part of the evidence contract.
    for name, value in xlsx.items():
        if name in PHASES or name.endswith("_allocation_metrics"):
            continue
        if isinstance(value, list):
            vector(value, count, f"{label}.source.xlsx_cell_values.{name}")
            for index, item in enumerate(value):
                _validate_recursive_value(item,
                                          f"{label}.source.xlsx_cell_values.{name}[{index}]")

    output_values = vector(xlsx.get("output_sha256"), count,
                           f"{label}.source.xlsx_cell_values.output_sha256")
    for index, item in enumerate(output_values):
        check_hash(item, f"{label}.source.xlsx_cell_values.output_sha256[{index}]")
    require(all(item == output_values[0] for item in output_values),
            f"{label} source output digest varies across samples")
    semantic_values = vector(xlsx.get("semantic_sha256"), count,
                             f"{label}.source.xlsx_cell_values.semantic_sha256")
    for index, item in enumerate(semantic_values):
        check_hash(item, f"{label}.source.xlsx_cell_values.semantic_sha256[{index}]")
    require(all(item == semantic_values[0] for item in semantic_values),
            f"{label} semantic digest varies across samples")
    nonnegative_integer(xlsx.get("untouched_member_count"),
                       f"{label}.untouched_member_count")
    require(xlsx["untouched_member_count"] > 0,
            f"{label}.untouched_member_count is zero")
    untouched_values = vector(xlsx.get("untouched_member_sha256"), count,
                              f"{label}.source.xlsx_cell_values.untouched_member_sha256")
    for index, item in enumerate(untouched_values):
        check_hash(item,
                   f"{label}.source.xlsx_cell_values.untouched_member_sha256[{index}]")
    require(all(item == untouched_values[0] for item in untouched_values),
            f"{label} untouched-member digest varies across samples")

    refusal = validate_refusal(xlsx.get("output_budget_refusal"),
                               f"{label}.output_budget_refusal")
    if managed_case:
        for field in ("payload_memory_limit", "publication_planning_memory_headroom",
                      "cache_budget_memory_limit"):
            nonnegative_integer(xlsx.get(field), f"{label}.{field}")
            require(xlsx[field] > 0, f"{label}.{field} is zero")
        require(xlsx["cache_budget_memory_limit"] ==
                xlsx["payload_memory_limit"] + xlsx["publication_planning_memory_headroom"],
                f"{label} managed memory limits do not reconcile")
        for field in ("pre_publication_budget", "post_publication_budget"):
            snapshots = vector(xlsx.get(field), count, f"{label}.{field}")
            for index, snapshot in enumerate(snapshots):
                validate_budget_snapshot(snapshot, f"{label}.{field}[{index}]")
        for field in ("cache_retained_bytes", "cache_budget_memory_used",
                      "cache_budget_reserved_bytes", "cache_budget_reservation_failures",
                      "budget_used_after_package_drop", "budget_used_after_handles_drop",
                      "budget_objects_used_after_handles_drop"):
            integer_vector(xlsx.get(field), count, f"{label}.{field}")
        require(all(item > 0 for item in xlsx["budget_used_after_package_drop"]),
                f"{label} package-drop budget retention is empty")
        require(all(item == 0 for item in xlsx["budget_used_after_handles_drop"]),
                f"{label} handles-drop memory budget is nonzero")
        require(all(item == 0 for item in xlsx["budget_objects_used_after_handles_drop"]),
                f"{label} handles-drop object budget is nonzero")
        for index in range(count):
            require(
                xlsx["cache_retained_bytes"][index]
                == xlsx["cache_budget_memory_used"][index]
                == xlsx["cache_budget_reserved_bytes"][index]
                == xlsx["budget_used_after_package_drop"][index],
                f"{label} managed cache retention does not reconcile at sample {index}",
            )
        require(refusal["typed_output_resource_refusal"]
                and refusal["zero_output_verified"]
                and refusal["source_identity_preserved"],
                f"{label} managed output refusal proof is incomplete")
    else:
        for field in ("payload_memory_limit", "publication_planning_memory_headroom",
                      "cache_budget_memory_limit"):
            require(xlsx.get(field) is None,
                    f"{label}.{field} is non-null for unmanaged control")
        for field in ("cache_budget_memory_used", "cache_budget_reserved_bytes",
                      "cache_budget_reservation_failures", "budget_used_after_package_drop",
                      "budget_used_after_handles_drop", "budget_objects_used_after_handles_drop"):
            values = integer_vector(xlsx.get(field), count, f"{label}.{field}")
            require(all(item == 0 for item in values), f"{label}.{field} is nonzero")
        for field in ("pre_publication_budget", "post_publication_budget"):
            snapshots = vector(xlsx.get(field), count, f"{label}.{field}")
            for index, snapshot in enumerate(snapshots):
                snapshot = validate_budget_snapshot(snapshot, f"{label}.{field}[{index}]")
                require(all(item in (None, 0) for item in snapshot.values()),
                        f"{label}.{field}[{index}] is nonzero for unmanaged control")

    stable = source_identity(source, count)
    logical = {
        "source": stable,
        "source_output_sha256": output_values[0],
        "semantic_sha256": semantic_values[0],
        "untouched_member_count": xlsx["untouched_member_count"],
        "untouched_member_sha256": untouched_values[0],
    }
    observations: dict[str, Any] = {
        "source_read_calls": list(source["read_calls"]),
        "source_read_bytes": list(source["read_bytes"]),
        "ordinary_payload_read_calls": list(source["ordinary_payload_read_calls"]),
        "ordinary_payload_read_bytes": list(source["ordinary_payload_read_bytes"]),
        "max_in_flight_reads": list(source["max_in_flight_reads"]),
        "ordinary_payload_materializations": list(source["ordinary_payload_materializations"]),
        "xlsx": {
            name: list(xlsx[name]) for name in sorted(XLSX_SOURCE_METRIC_VECTORS)
            if isinstance(xlsx.get(name), list)
        },
        "retention_limits": {
            name: xlsx.get(name) for name in (
                "payload_memory_limit",
                "publication_planning_memory_headroom",
                "cache_budget_memory_limit",
            )
        },
        "output_budget_refusal": refusal,
    }
    return logical, observations, phase_values, allocation_values


def build_allocation_report(
    values: dict[str, list[dict[str, Any]]], count: int, label: str
) -> dict[str, Any]:
    """Build raw absolute, retained, and attributable allocation summaries."""

    planning_before = [item["live_bytes_before"] for item in values["plan"]]
    phase_report: dict[str, Any] = {}
    for phase in ALLOCATION_PHASES:
        samples = values[phase]
        fields: dict[str, Any] = {}
        for field in ALLOCATION_REPORT_FIELDS:
            if field == "attributable_region_peak_live_bytes":
                field_values = [
                    item["region_peak_live_bytes"] - planning_before[index]
                    for index, item in enumerate(samples)
                ]
            else:
                field_values = [item[field] for item in samples]
            unit = "count" if field.endswith("_calls") else "bytes"
            fields[field] = (
                signed_stats(field_values, unit)
                if field in ("retained_live_delta_bytes",
                             "attributable_region_peak_live_bytes")
                else stats(field_values, unit)
            )
        phase_report[PHASE_LABELS[f"{phase}_ns"]] = fields

    workflow_fields = (
        "allocation_calls",
        "deallocation_calls",
        "reallocation_calls",
        "failed_allocation_calls",
        "allocated_bytes",
        "deallocated_bytes",
    )
    workflow: dict[str, Any] = {}
    for field in workflow_fields:
        field_values = [
            sum(values[phase][index][field] for phase in ALLOCATION_PHASES)
            for index in range(count)
        ]
        workflow[field] = stats(
            field_values, "count" if field.endswith("_calls") else "bytes"
        )
    retained_workflow = [
        sum(values[phase][index]["retained_live_delta_bytes"]
            for phase in ALLOCATION_PHASES)
        for index in range(count)
    ]
    workflow["retained_live_delta_bytes"] = signed_stats(retained_workflow, "bytes")
    absolute_peaks = [
        max(
            planning_before[index],
            *(values[phase][index]["region_peak_live_bytes"]
              for phase in ALLOCATION_PHASES),
        )
        for index in range(count)
    ]
    attributable_peaks = [
        absolute_peaks[index] - planning_before[index]
        for index in range(count)
    ]
    workflow["absolute_region_peak_live_bytes"] = stats(absolute_peaks, "bytes")
    workflow["attributable_region_peak_live_bytes"] = stats(attributable_peaks, "bytes")
    return {
        "sample_count": count,
        "planning_live_bytes_before": stats(planning_before, "bytes"),
        "absolute_region_peak_live_bytes": stats(absolute_peaks, "bytes"),
        "attributable_region_peak_live_bytes": stats(attributable_peaks, "bytes"),
        "phases": phase_report,
        "workflow": workflow,
        "field_paths": {
            "plan": "results[0].source.xlsx_cell_values.plan_allocation_metrics",
            "commit": "results[0].source.xlsx_cell_values.commit_allocation_metrics",
            "publication": "results[0].source.xlsx_cell_values.publication_allocation_metrics",
        },
    }


def validate_report(path: Path, catalog_path: Path, stderr_path: Path,
                    job: dict[str, Any], kind: str, binary: dict[str, Any],
                    plan: dict[str, Any]) -> dict[str, Any]:
    label = job["name"]
    path = path.resolve()
    catalog_path = catalog_path.resolve()
    stderr_path = stderr_path.resolve()
    report = read_json(path)
    require(isinstance(report, dict), f"{label} report is not an object")
    require(report.get("schema_version") == 1, f"{label} schema version differs")
    tool = report.get("tool")
    require(isinstance(tool, dict), f"{label}.tool is not an object")
    require(tool.get("binary") == binary["binary"], f"{label} tool binary differs")
    require(tool.get("profile") == "release", f"{label} tool profile differs")
    expected_instrumentation = (
        "system_allocator_operation_scoped" if kind == ALLOC_BINARY else "none"
    )
    require(tool.get("instrumentation") == expected_instrumentation,
            f"{label} tool instrumentation differs")
    identity = report.get("binary_identity")
    require(isinstance(identity, dict), f"{label}.binary_identity is missing")
    require(identity.get("path") == binary["path"], f"{label} binary path differs")
    require(identity.get("binary_sha256") == binary["sha256"],
            f"{label} binary hash differs")
    require(identity.get("binary_bytes") == binary["bytes"],
            f"{label} binary size differs")
    require(identity.get("profile") == "release", f"{label} binary profile differs")
    environment = report.get("environment")
    require(isinstance(environment, dict), f"{label}.environment is missing")
    require(environment.get("git_revision") == plan["revision"],
            f"{label} revision differs")
    require(environment.get("cpu_affinity") == str(plan["cpu"]),
            f"{label} CPU affinity differs")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{label}.configuration is missing")
    require(configuration.get("cases") == [job["case"]],
            f"{label} case configuration differs")
    require(configuration.get("xlsx_cell_crud_shapes") == [job["shape"]],
            f"{label} shape configuration differs")
    require(configuration.get("samples_per_case") == job["samples"],
            f"{label} sample configuration differs")
    require(configuration.get("warmup_iterations_per_case") == job["warmup"],
            f"{label} warmup configuration differs")

    parallel = report.get("parallel_metrics")
    require(isinstance(parallel, dict), f"{label}.parallel_metrics is missing")
    require(parallel.get("schema_version") == 1,
            f"{label}.parallel_metrics schema differs")
    require(parallel.get("scope") == "explicit_local_execution_only",
            f"{label}.parallel_metrics scope differs")
    require(parallel.get("claim") == "descriptive",
            f"{label}.parallel_metrics claim differs")
    worker_budget = parallel.get("configured_worker_budget")
    require(isinstance(worker_budget, dict)
            and worker_budget.get("status") == "measured"
            and worker_budget.get("value") == [1],
            f"{label}.parallel_metrics worker budget differs")

    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{label} must contain one result")
    result = results[0]
    require(isinstance(result, dict), f"{label} result is not an object")
    require(result.get("case") == job["case"], f"{label} result case differs")
    corpus = result.get("corpus")
    require(isinstance(corpus, dict), f"{label}.corpus is missing")
    require(corpus.get("generator") ==
            "litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1",
            f"{label} corpus generator differs")
    require(corpus.get("shape") == job["shape"], f"{label} corpus shape differs")
    require(corpus.get("package_format") == "XLSX/OPC/ZIP",
            f"{label} corpus package format differs")
    for field in ("archive_sha256", "target_payload_sha256"):
        check_hash(corpus.get(field), f"{label}.corpus.{field}")
    for field in ("archive_bytes", "archive_member_count", "entry_count",
                  "entry_bytes", "uncompressed_payload_bytes",
                  "target_payload_bytes"):
        nonnegative_integer(corpus.get(field), f"{label}.corpus.{field}")
    corpus_xlsx = corpus.get("xlsx")
    require(isinstance(corpus_xlsx, dict), f"{label}.corpus.xlsx is missing")
    for field in ("sheet_count", "rows_per_sheet", "columns_per_sheet",
                  "one_percent_update_count"):
        nonnegative_integer(corpus_xlsx.get(field), f"{label}.corpus.xlsx.{field}")
        require(corpus_xlsx[field] > 0, f"{label}.corpus.xlsx.{field} is zero")
    members = corpus_xlsx.get("source_members")
    require(isinstance(members, dict), f"{label}.corpus.xlsx.source_members is missing")
    require(set(members) == {"workbook", "worksheets", "shared_strings", "styles"},
            f"{label}.corpus.xlsx.source_members fields differ")
    require(isinstance(members["workbook"], str) and members["workbook"],
            f"{label}.corpus.xlsx workbook member is missing")
    require(isinstance(members["worksheets"], list)
            and len(members["worksheets"]) == corpus_xlsx["sheet_count"],
            f"{label}.corpus.xlsx worksheet member count differs")
    require(all(isinstance(item, str) and item for item in members["worksheets"]),
            f"{label}.corpus.xlsx worksheet member is malformed")
    require(members["shared_strings"] is None
            or isinstance(members["shared_strings"], str),
            f"{label}.corpus.xlsx shared_strings member is malformed")
    require(isinstance(members["styles"], str) and members["styles"],
            f"{label}.corpus.xlsx styles member is missing")

    elapsed = result.get("elapsed_ns")
    require(isinstance(elapsed, dict), f"{label}.elapsed_ns is missing")
    elapsed_values = integer_vector(elapsed.get("samples"), job["samples"],
                                    f"{label}.elapsed_ns.samples")
    sample_order = elapsed.get("sample_order")
    require(isinstance(sample_order, list)
            and sorted(sample_order) == list(range(job["samples"])),
            f"{label}.elapsed_ns.sample_order is not a permutation")
    require(elapsed_values == sorted(elapsed_values),
            f"{label}.elapsed_ns.samples are not sorted")
    validate_report_stats(elapsed, elapsed_values, f"{label}.elapsed_ns")
    sink = validate_sink(result.get("sink"), label)
    output = check_hash(result.get("output_sha256"), f"{label}.output_sha256")
    validate_operation_metrics(result.get("operation_metrics"), job["samples"], sink, label)
    logical, source_observations, phase_values, allocation_values = validate_source(
        result.get("source"), job["samples"], job, corpus,
        kind == ALLOC_BINARY, label
    )
    if "managed" in job["case"]:
        for index, snapshot in enumerate(
                source_observations["xlsx"]["post_publication_budget"]):
            require(snapshot["output_bytes_used"] == sink["accepted_bytes"],
                    f"{label} post-publication output retention differs at sample {index}")
    require(logical["source_output_sha256"] == output,
            f"{label} source output hash differs from result output")
    workflow_values = [
        sum(phase_values[phase][index] for phase in TIMED_PHASES)
        for index in range(job["samples"])
    ]
    require([workflow_values[index] for index in sample_order] == elapsed_values,
            f"{label} phase vectors do not reconcile with elapsed samples")
    time_values = validate_time_v(stderr_path, label)

    catalog = read_json(catalog_path)
    try:
        _BINDING.validate_binding(report, catalog)
    except Exception as error:
        raise EvidenceError(f"{label} corpus catalog binding failed: {error}") from error
    catalog_ref = report.get("corpus_catalog")
    require(isinstance(catalog_ref, dict), f"{label}.corpus_catalog reference is missing")
    catalog_identity = {
        "manifest_version": catalog.get("manifest_version"),
        "catalog_id": catalog.get("catalog_id"),
        "catalog_sha256": check_hash(catalog.get("catalog_sha256"),
                                      f"{label}.catalog_sha256"),
        "content_set_sha256": check_hash(catalog.get("content_set_sha256"),
                                          f"{label}.content_set_sha256"),
        "corpus_ids": sorted(item.get("id") for item in catalog.get("corpora", [])
                              if isinstance(item, dict)),
        "case_bindings": sorted(
            (item.get("case"), item.get("corpus_id"), item.get("legacy_archive_sha256"))
            for item in catalog.get("case_bindings", []) if isinstance(item, dict)
        ),
    }
    require(catalog_identity["manifest_version"] == 2,
            f"{label} catalog manifest version differs")
    require(catalog_identity["catalog_id"] == "litchi-perf-corpus-v2",
            f"{label} catalog ID differs")
    require(catalog_ref.get("catalog_sha256") == catalog_identity["catalog_sha256"],
            f"{label} catalog reference hash differs")
    require(catalog_ref.get("content_set_sha256") == catalog_identity["content_set_sha256"],
            f"{label} catalog reference content hash differs")

    allocation = (build_allocation_report(allocation_values, job["samples"], label)
                  if kind == ALLOC_BINARY else None)
    timing = None
    if kind == NORMAL_BINARY:
        timing = {"workflow": stats(elapsed_values, "ns")}
        for phase in PHASES:
            timing[PHASE_LABELS[phase]] = stats(phase_values[phase], "ns")
    row_identity = {
        "case": job["case"],
        "shape": job["shape"],
        "corpus": corpus,
        "sink": sink,
        "source": logical["source"],
        "output_sha256": output,
        "semantic_sha256": logical["semantic_sha256"],
        "untouched_member_count": logical["untouched_member_count"],
        "untouched_member_sha256": logical["untouched_member_sha256"],
        "catalog": catalog_identity,
    }
    folder = HERE / job["stage"]
    receipt_path = folder / (label + ".receipt.json")
    row = {
        "stage": job["stage"],
        "execution_stage": job["execution_stage"],
        "lane": job["lane"],
        "name": label,
        "repeat": job["repeat"],
        "case": job["case"],
        "shape": job["shape"],
        "warmup": job["warmup"],
        "samples": job["samples"],
        "report": str(path.relative_to(HERE)),
        "receipt": str(receipt_path.relative_to(HERE)),
        "catalog": str(catalog_path.relative_to(HERE)),
        "stderr": str(stderr_path.relative_to(HERE)),
        "report_sha256": sha256(path),
        "receipt_sha256": sha256(receipt_path),
        "catalog_sha256": catalog_identity["catalog_sha256"],
        "binary_sha256": binary["sha256"],
        "identity": row_identity,
        "identity_sha256": hashlib.sha256(
            json.dumps(row_identity, sort_keys=True, separators=(",", ":")).encode()
        ).hexdigest(),
        "source_observations": source_observations,
        "time_v": time_values,
        "commit_scope": (
            "commit_ns includes edit staging and MultiSourceEdit::commit; "
            "open, selector planning, and stream publication are separate phases"
        ),
        "allocation_instrumented_elapsed_excluded": kind == ALLOC_BINARY,
        "allocation": allocation,
    }
    if kind == NORMAL_BINARY:
        row["timing"] = timing
        row["phase_samples"] = {
            PHASE_LABELS[phase]: list(phase_values[phase]) for phase in PHASES
        }
        row["elapsed_sample_order"] = list(sample_order)
        row["elapsed_samples"] = list(elapsed_values)
    else:
        row["allocation_samples"] = {
            PHASE_LABELS[f"{phase}_ns"]: [dict(item) for item in allocation_values[phase]]
            for phase in ALLOCATION_PHASES
        }
        row["instrumented_elapsed_samples"] = list(elapsed_values)
    return row


def check_receipt(job: dict[str, Any], binary: dict[str, Any],
                  plan: dict[str, Any]) -> tuple[dict[str, Any], Path, Path]:
    folder = HERE / job["stage"]
    path = folder / f"{job['name']}.receipt.json"
    receipt = read_json(path)
    require(isinstance(receipt, dict), f"{path.name} is not an object")
    validate_receipt_common(receipt, path.name)
    require(receipt.get("exit_code") == 0, f"{path.name} did not exit successfully")
    require(receipt.get("binary_sha256") == binary["sha256"],
            f"{path.name} binary hash differs")
    require(receipt.get("execution_stage") == job["execution_stage"],
            f"{path.name} execution stage differs")
    execution_manifest = HERE / job["execution_stage"] / "source-manifest.json"
    require(receipt.get("execution_manifest_sha256") == sha256(execution_manifest),
            f"{path.name} execution manifest differs")
    require(receipt.get("source_manifest_sha256") == sha256(folder / "source-manifest.json"),
            f"{path.name} source manifest differs")
    require(receipt.get("script_sha256") == sha256(HERE / "run.py"),
            f"{path.name} script hash differs")
    require(receipt.get("plan_sha256") == sha256(HERE / "plan.json"),
            f"{path.name} plan hash differs")
    require(receipt.get("command") == expected_command(job, binary),
            f"{path.name} command differs from frozen capture")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{path.name} artifact inventory is missing")
    expected_artifacts = {
        f"{job['name']}.json",
        f"{job['name']}.catalog.json",
        f"{job['name']}.stdout",
        f"{job['name']}.stderr",
        f"{job['name']}.host.json",
    }
    require(set(artifacts) == expected_artifacts,
            f"{path.name} artifact inventory differs")
    for name, digest in artifacts.items():
        check_hash(digest, f"{path.name}/{name}")
        artifact = folder / name
        require(artifact.is_file() and not artifact.is_symlink(),
                f"{path.name} artifact missing: {name}")
        require(sha256(artifact) == digest,
                f"{path.name} artifact hash differs: {name}")
    check_host(folder / f"{job['name']}.host.json")
    report_path = folder / f"{job['name']}.json"
    catalog_path = folder / f"{job['name']}.catalog.json"
    return receipt, report_path, catalog_path


def validate_stage(stage: str, plan: dict[str, Any], lane: str,
                   binaries: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    jobs = jobs_for(plan, stage, lane)
    expected_count = {"preflight": 16, "native": 32, "alloc": 32}[lane]
    require(len(jobs) == expected_count,
            f"{stage} {lane} matrix produced {len(jobs)} jobs, expected {expected_count}")
    expected_names = {job["name"] for job in jobs}
    prefix = f"{lane}-"
    actual_names = {
        path.name[: -len(".receipt.json")]
        for path in (HERE / stage).glob(f"{prefix}*.receipt.json")
    }
    require(actual_names == expected_names,
            f"{stage} {lane} receipt matrix differs: {sorted(actual_names ^ expected_names)}")
    kind = ALLOC_BINARY if lane == "alloc" else NORMAL_BINARY
    rows: list[dict[str, Any]] = []
    for job in jobs:
        receipt, report_path, catalog_path = check_receipt(job, binaries[kind], plan)
        stderr_path = HERE / stage / f"{job['name']}.stderr"
        row = validate_report(report_path, catalog_path, stderr_path, job,
                              kind, binaries[kind], plan)
        row["receipt_execution_manifest_sha256"] = receipt["execution_manifest_sha256"]
        row["receipt_source_manifest_sha256"] = receipt["source_manifest_sha256"]
        rows.append(row)
    return rows


def validate_stage_metadata(stage: str, plan: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    frozen = frozen_inputs()
    folder = HERE / stage
    manifest_path = folder / "source-manifest.json"
    require(manifest_path.is_file() and not manifest_path.is_symlink(),
            f"{stage} source manifest is missing")
    manifest = read_json(manifest_path)
    require(isinstance(manifest, dict) and manifest,
            f"{stage} source manifest is empty")
    for relative, digest in manifest.items():
        require(isinstance(relative, str) and relative and not Path(relative).is_absolute(),
                f"{stage} source manifest has an invalid path")
        check_hash(digest, f"{stage} source manifest {relative}")
    host = read_json(HERE / "host.json")
    require(isinstance(host, dict), "host.json is not an object")
    affinity = host.get("cpu_affinity")
    require(isinstance(affinity, list) and plan["cpu"] in affinity,
            f"{stage} host CPU affinity does not include pinned CPU")
    return frozen, {
        "stage": stage,
        "sha256": sha256(manifest_path),
        "entries": len(manifest),
    }, {
        "host_sha256": sha256(HERE / "host.json"),
        "host": host,
    }


def percent_change(first: float, second: float) -> float | None:
    if first == 0.0:
        return 0.0 if second == 0.0 else None
    return (second / first - 1.0) * 100.0


def comparison_record(first: Any, second: Any, *, lane: str, case: str,
                      shape: str, repeat: int, metric: str,
                      policy: str = "increase") -> dict[str, Any]:
    finite_number(first, f"comparison {metric}.baseline")
    finite_number(second, f"comparison {metric}.candidate")
    change = percent_change(float(first), float(second))
    if policy == "increase":
        adverse = change is None or change > DRIFT_THRESHOLD_PERCENT
    elif policy == "informational":
        adverse = None
    else:
        raise EvidenceError(f"unknown comparison policy {policy}")
    return {
        "lane": lane,
        "case": case,
        "shape": shape,
        "repeat": repeat,
        "metric": metric,
        "baseline": first,
        "candidate": second,
        "delta": float(second) - float(first),
        "change_percent": change,
        "threshold_percent": DRIFT_THRESHOLD_PERCENT,
        "adverse": adverse,
    }


def exact_comparison(first: Any, second: Any, *, lane: str, case: str,
                     shape: str, repeat: int, metric: str) -> dict[str, Any]:
    equal = first == second
    return {
        "lane": lane,
        "case": case,
        "shape": shape,
        "repeat": repeat,
        "metric": metric,
        "baseline": first,
        "candidate": second,
        "equal": equal,
        "adverse": not equal,
    }


def source_metric_summaries(row: dict[str, Any]) -> dict[str, dict[str, Any]]:
    """Summarize every numeric source counter retained by validate_source."""

    def unit_for(name: str) -> str:
        return ("bytes" if "bytes" in name or name.startswith("budget_used_after_")
                else "count")

    result: dict[str, dict[str, Any]] = {}
    observations = row["source_observations"]
    for name, values in observations.items():
        if name in ("xlsx", "output_budget_refusal", "retention_limits"):
            continue
        if isinstance(values, list) and all(isinstance(item, int)
                                            and not isinstance(item, bool)
                                            for item in values):
            result[f"source.{name}"] = stats(
                list(values), unit_for(name)
            )
    xlsx = observations.get("xlsx", {})
    require(isinstance(xlsx, dict), "source observations xlsx is not an object")
    for name, values in xlsx.items():
        if not isinstance(values, list):
            continue
        if all(isinstance(item, int) and not isinstance(item, bool)
               for item in values):
            result[f"source.xlsx.{name}"] = stats(
                list(values), unit_for(name)
            )
            continue
        if not all(isinstance(item, dict) for item in values):
            continue
        field_names = sorted({field for item in values for field in item})
        for field in field_names:
            field_values = [item.get(field) for item in values]
            if all(isinstance(item, int) and not isinstance(item, bool)
                   for item in field_values):
                result[f"source.xlsx.{name}.{field}"] = stats(
                    field_values, unit_for(field)
                )
    return result


def source_exact_summaries(row: dict[str, Any]) -> dict[str, Any]:
    observations = row["source_observations"]
    # Refusal/error identity and the configured managed limits are exact
    # correctness controls.  Budget snapshots and retained byte counters are
    # resource measurements; they remain in numeric comparisons so an
    # optimization may change them while still being reported.
    return {
        "source.output_budget_refusal": observations["output_budget_refusal"],
        "source.retention_limits": observations["retention_limits"],
    }


def row_metric_summaries(row: dict[str, Any], lane: str) -> dict[str, Any]:
    result: dict[str, Any] = {}
    if lane in ("preflight", "native"):
        for phase in ("workflow", "open", "planning", "commit",
                      "publication", "reopen"):
            for metric in REPORT_STATS:
                result[f"timing.{phase}.{metric}"] = row["timing"][phase][metric]
        result["time_v.max_rss_bytes"] = row["time_v"]["max_rss_kib"] * 1024
    if lane == "alloc":
        allocation = row["allocation"]
        require(isinstance(allocation, dict), "allocation row has no allocation report")
        for phase, phase_values in allocation["phases"].items():
            for field, field_stats in phase_values.items():
                for metric in REPORT_STATS:
                    result[f"allocation.{phase}.{field}.{metric}"] = field_stats[metric]
        for field, field_stats in allocation["workflow"].items():
            metrics = REPORT_STATS
            if field in ("allocated_bytes", "attributable_region_peak_live_bytes"):
                metrics = REPORT_STATS + ("max",)
            for metric in metrics:
                result[f"allocation.workflow.{field}.{metric}"] = field_stats[metric]
        for name in ("planning_live_bytes_before", "absolute_region_peak_live_bytes",
                     "attributable_region_peak_live_bytes"):
            for metric in REPORT_STATS:
                result[f"allocation.{name}.{metric}"] = allocation[name][metric]
    for name, field_stats in source_metric_summaries(row).items():
        for metric in REPORT_STATS:
            result[f"{name}.{metric}"] = field_stats[metric]
    return result


def _pair_rows(baseline_rows: list[dict[str, Any]],
               candidate_rows: list[dict[str, Any]], lane: str
               ) -> list[tuple[dict[str, Any], dict[str, Any]]]:
    baseline = {(row["case"], row["shape"], row["repeat"]): row
                for row in baseline_rows if row["lane"] == lane}
    candidate = {(row["case"], row["shape"], row["repeat"]): row
                for row in candidate_rows if row["lane"] == lane}
    require(set(baseline) == set(candidate),
            f"{lane} baseline/candidate job keys differ: "
            f"{sorted(set(baseline) ^ set(candidate))}")
    return [(baseline[key], candidate[key]) for key in sorted(baseline)]


def compare_lane(baseline_rows: list[dict[str, Any]],
                 candidate_rows: list[dict[str, Any]], lane: str
                 ) -> dict[str, Any]:
    rows: list[dict[str, Any]] = []
    all_comparisons: list[dict[str, Any]] = []
    adverse: list[dict[str, Any]] = []
    for baseline, candidate in _pair_rows(baseline_rows, candidate_rows, lane):
        key = {name: baseline[name] for name in ("case", "shape", "repeat")}
        identity_equal = baseline["identity"] == candidate["identity"]
        require(identity_equal,
                f"matched identity differs for {lane} {key}")
        comparison: list[dict[str, Any]] = []
        baseline_metrics = row_metric_summaries(baseline, lane)
        candidate_metrics = row_metric_summaries(candidate, lane)
        require(set(baseline_metrics) == set(candidate_metrics),
                f"{lane} metric fields differ for {key}")
        for metric in sorted(baseline_metrics):
            policy = ("informational" if "retained_live_delta_bytes" in metric
                      else "increase")
            record = comparison_record(
                baseline_metrics[metric], candidate_metrics[metric],
                lane=lane, metric=metric, policy=policy, **key
            )
            comparison.append(record)
            all_comparisons.append(record)
            if record["adverse"] is True:
                adverse.append(record)
        baseline_exact = source_exact_summaries(baseline)
        candidate_exact = source_exact_summaries(candidate)
        require(set(baseline_exact) == set(candidate_exact),
                f"{lane} exact source fields differ for {key}")
        exact: list[dict[str, Any]] = []
        for metric in sorted(baseline_exact):
            record = exact_comparison(
                baseline_exact[metric], candidate_exact[metric],
                lane=lane, metric=metric, **key
            )
            exact.append(record)
            if record["adverse"]:
                adverse.append(record)
        rows.append({
            **key,
            "identity_equal": identity_equal,
            "source_exact": exact,
            "comparisons": comparison,
            "adverse": [record for record in comparison if record["adverse"]]
                       + [record for record in exact if record["adverse"]],
        })
    return {
        "rows": rows,
        "comparisons": all_comparisons,
        "adverse": adverse,
        "pair_count": len(rows),
    }


def drift_record(first: Any, second: Any, *, stage: str, lane: str,
                 case: str, shape: str, metric: str) -> dict[str, Any]:
    finite_number(first, f"drift {metric}.first")
    finite_number(second, f"drift {metric}.second")
    change = percent_change(float(first), float(second))
    return {
        "stage": stage,
        "lane": lane,
        "case": case,
        "shape": shape,
        "repeat_first": 1,
        "repeat_second": 2,
        "metric": metric,
        "first": first,
        "second": second,
        "delta": float(second) - float(first),
        "change_percent": change,
        "threshold_percent": DRIFT_THRESHOLD_PERCENT,
        "over_five_percent": change is None or abs(change) > DRIFT_THRESHOLD_PERCENT,
    }


def repeat_drift(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    for stage in ("baseline", "candidate"):
        stage_rows = [row for row in rows if row["stage"] == stage]
        for lane in ("native", "alloc"):
            groups = sorted({(row["case"], row["shape"]) for row in stage_rows
                             if row["lane"] == lane})
            by_key = {(row["case"], row["shape"], row["repeat"]): row
                      for row in stage_rows if row["lane"] == lane}
            for case, shape in groups:
                first = by_key[(case, shape, 1)]
                second = by_key[(case, shape, 2)]
                first_metrics = row_metric_summaries(first, lane)
                second_metrics = row_metric_summaries(second, lane)
                require(set(first_metrics) == set(second_metrics),
                        f"{stage} {lane} repeat metric fields differ for {case}/{shape}")
                for metric in sorted(first_metrics):
                    output.append(drift_record(
                        first_metrics[metric], second_metrics[metric],
                        stage=stage, lane=lane, case=case, shape=shape,
                        metric=metric,
                    ))
    return output


def _comparison_index(comparisons: list[dict[str, Any]]) -> dict[tuple[str, str, str, int, str], dict[str, Any]]:
    return {
        (item["lane"], item["case"], item["shape"], item["repeat"], item["metric"]): item
        for item in comparisons
    }


def gate_check(index: dict[tuple[str, str, str, int, str], dict[str, Any]],
               *, lane: str, case: str, shape: str, repeat: int, metric: str,
               max_change: float | None = None,
               min_improvement: float | None = None) -> dict[str, Any]:
    key = (lane, case, shape, repeat, metric)
    comparison = index.get(key)
    require(comparison is not None, f"missing comparison for gate metric {key}")
    change = comparison["change_percent"]
    if min_improvement is not None:
        passed = change is not None and change <= -min_improvement
        criterion = f"candidate change <= {-min_improvement:.2f}%"
    elif max_change is not None:
        passed = change is not None and change <= max_change
        criterion = f"candidate change <= {max_change:.2f}%"
    else:
        raise EvidenceError("gate has no threshold")
    return {
        "lane": lane,
        "case": case,
        "shape": shape,
        "repeat": repeat,
        "metric": metric,
        "criterion": criterion,
        "pass": passed,
        "baseline": comparison["baseline"],
        "candidate": comparison["candidate"],
        "change_percent": change,
    }


def gate_group(name: str, checks: list[dict[str, Any]], description: str) -> dict[str, Any]:
    return {
        "name": name,
        "description": description,
        "pass": bool(checks) and all(check["pass"] for check in checks),
        "check_count": len(checks),
        "checks": checks,
    }


def main_gates(plan: dict[str, Any], baseline_rows: list[dict[str, Any]],
               candidate_rows: list[dict[str, Any]], comparisons: dict[str, dict[str, Any]],
               identity_equal: bool) -> dict[str, Any]:
    all_comparisons = [
        record for lane_result in comparisons.values()
        for record in lane_result["comparisons"]
    ]
    index = _comparison_index(all_comparisons)
    native_rows = [row for row in baseline_rows if row["lane"] == "native"]
    alloc_rows = [row for row in baseline_rows if row["lane"] == "alloc"]
    primary: list[dict[str, Any]] = []
    latency: list[dict[str, Any]] = []
    for row in native_rows:
        key = {name: row[name] for name in ("case", "shape", "repeat")}
        if "one_percent" in row["case"]:
            for metric in ("timing.workflow.p50", "timing.workflow.mean"):
                primary.append(gate_check(
                    index, lane="native", metric=metric,
                    min_improvement=PRIMARY_IMPROVEMENT_PERCENT, **key
                ))
        if "one_edit" in row["case"]:
            for metric in ("timing.workflow.p50", "timing.workflow.mean"):
                latency.append(gate_check(
                    index, lane="native", metric=metric,
                    max_change=LATENCY_GUARD_PERCENT, **key
                ))
    memory: list[dict[str, Any]] = []
    allocation: list[dict[str, Any]] = []
    for row in alloc_rows:
        key = {name: row[name] for name in ("case", "shape", "repeat")}
        for metric in (
            "allocation.workflow.attributable_region_peak_live_bytes.max",
        ):
            memory.append(gate_check(
                index, lane="alloc", metric=metric,
                max_change=LATENCY_GUARD_PERCENT, **key
            ))
        for metric in (
            "allocation.workflow.allocated_bytes.max",
        ):
            allocation.append(gate_check(
                index, lane="alloc", metric=metric,
                max_change=LATENCY_GUARD_PERCENT, **key
            ))
    rss: list[dict[str, Any]] = []
    for row in native_rows:
        key = {name: row[name] for name in ("case", "shape", "repeat")}
        rss.append(gate_check(
            index, lane="native", metric="time_v.max_rss_bytes",
            max_change=LATENCY_GUARD_PERCENT, **key
        ))
    memory.extend(rss)
    exact_records = [
        record
        for lane_result in comparisons.values()
        for pair in lane_result["rows"]
        for record in pair["source_exact"]
    ]
    correctness = {
        "name": "matched_identity_and_exact_refusal",
        "pass": identity_equal and all(record["equal"] for record in exact_records),
        "identity_equal": identity_equal,
        "exact_check_count": len(exact_records),
        "exact_checks": exact_records,
    }
    groups = {
        "primary_one_percent": gate_group(
            "primary_one_percent",
            primary,
            plan["admission"]["primary"],
        ),
        "one_cell_latency": gate_group(
            "one_cell_latency",
            latency,
            plan["admission"]["latency_guards"],
        ),
        "workflow_memory": gate_group(
            "workflow_memory",
            memory,
            plan["admission"]["workflow_memory"],
        ),
        "allocation": gate_group(
            "allocation",
            allocation,
            plan["admission"]["allocation"],
        ),
        "correctness_identity": correctness,
    }
    groups["all_frozen_main_gates_pass"] = all(
        group["pass"] for group in groups.values()
    )
    groups["external_controls_required"] = {
        "status": "pending",
        "validated_here": False,
        "required": ["guard", "cap", "quality", "profile"],
        "reason": "main metrics analyzer does not own guard, cap, quality, or profile evidence",
    }
    return groups


def analyze() -> dict[str, Any]:
    plan = plan_data()
    frozen_baseline, baseline_manifest, host = validate_stage_metadata("baseline", plan)
    frozen_candidate, candidate_manifest, _ = validate_stage_metadata("candidate", plan)
    require(frozen_baseline == frozen_candidate,
            "baseline and candidate frozen input envelopes differ")
    binaries: dict[str, dict[str, dict[str, Any]]] = {}
    for stage in ("baseline", "candidate"):
        binaries[stage] = {
            NORMAL_BINARY: binary_metadata(stage, NORMAL_BINARY, plan),
            ALLOC_BINARY: binary_metadata(stage, ALLOC_BINARY, plan),
        }

    stage_rows: dict[str, list[dict[str, Any]]] = {"baseline": [], "candidate": []}
    for stage in stage_rows:
        for lane in LANES:
            stage_rows[stage].extend(validate_stage(
                stage, plan, lane, binaries[stage]
            ))
        counts = {
            lane: len([row for row in stage_rows[stage] if row["lane"] == lane])
            for lane in LANES
        }
        require(counts == {"preflight": 16, "native": 32, "alloc": 32},
                f"{stage} row counts differ: {counts}")

    rows = stage_rows["baseline"] + stage_rows["candidate"]
    identity_groups: dict[tuple[str, str], dict[str, Any]] = {}
    for row in rows:
        key = (row["case"], row["shape"])
        current = row["identity"]
        if key in identity_groups:
            require(identity_groups[key] == current,
                    f"matched source/sink/output/semantic identity differs for {key}")
        else:
            identity_groups[key] = current

    receipts = [read_json(HERE / row["receipt"]) for row in rows]
    ordered = sorted(receipts, key=lambda item: item["start_utc"])
    require(all(
        timestamp(left["end_utc"], "receipt end")
        <= timestamp(right["start_utc"], "receipt start")
        for left, right in zip(ordered, ordered[1:])
    ), "capture receipts overlap")

    lane_comparisons = {
        lane: compare_lane(stage_rows["baseline"], stage_rows["candidate"], lane)
        for lane in LANES
    }
    identity_equal = all(
        pair["identity_equal"]
        for lane_result in lane_comparisons.values()
        for pair in lane_result["rows"]
    )
    all_adverse = [
        record for lane_result in lane_comparisons.values()
        for record in lane_result["adverse"]
    ]
    drift = repeat_drift(rows)
    drift_over_five = [record for record in drift if record["over_five_percent"]]
    gates = main_gates(plan, stage_rows["baseline"], stage_rows["candidate"],
                       lane_comparisons, identity_equal)
    main_pass = gates["all_frozen_main_gates_pass"]
    disposition = (
        "pending external guard/cap/quality/profile evidence"
        if main_pass else "reject: one or more frozen main metrics gates failed"
    )
    by_stage_lane = {
        stage: {
            lane: [row for row in stage_rows[stage] if row["lane"] == lane]
            for lane in LANES
        }
        for stage in stage_rows
    }
    all_numeric_comparisons = [
        record for lane_result in lane_comparisons.values()
        for record in lane_result["comparisons"]
    ]
    all_exact_comparisons = [
        record
        for lane_result in lane_comparisons.values()
        for pair in lane_result["rows"]
        for record in pair["source_exact"]
    ]
    return {
        "schema": SCHEMA,
        "status": "pass",
        "stage": "matched baseline/candidate",
        "scope": plan["scope"],
        "priority": plan["priority"],
        "performance_claim": (
            "descriptive matched comparison; adoption remains pending external guard, cap, quality, and profile gates"
        ),
        "disposition": disposition,
        "plan_sha256": sha256(HERE / "plan.json"),
        "capture_sha256": sha256(HERE / "capture.py"),
        "run_sha256": sha256(HERE / "run.py"),
        "numerical_helper": {
            "path": str(_helper_path.relative_to(REPO)),
            "sha256": sha256(_helper_path),
            "report_statistics_dependency": {
                "path": str(_helper_base_path.relative_to(REPO)),
                "sha256": sha256(_helper_base_path),
            },
        },
        "corpus_binding_validator": {
            "path": str(_binding_path.relative_to(REPO)),
            "sha256": sha256(_binding_path),
        },
        "frozen_inputs": frozen_baseline,
        "source_manifests": {
            "baseline": baseline_manifest,
            "candidate": candidate_manifest,
        },
        "host": host,
        "binaries": binaries,
        "job_counts": {
            stage: {lane: len(by_stage_lane[stage][lane]) for lane in LANES}
            for stage in by_stage_lane
        },
        "sample_counts": {
            stage: {
                lane: sum(row["samples"] for row in by_stage_lane[stage][lane])
                for lane in LANES
            }
            for stage in by_stage_lane
        },
        "preflight": {
            "baseline_rows": by_stage_lane["baseline"]["preflight"],
            "candidate_rows": by_stage_lane["candidate"]["preflight"],
            "comparison": lane_comparisons["preflight"],
        },
        "native": {
            "baseline_rows": by_stage_lane["baseline"]["native"],
            "candidate_rows": by_stage_lane["candidate"]["native"],
            "timing_scope": (
                "workflow is elapsed_ns and equals open + planning + commit + publication; "
                "commit_ns includes edit staging and MultiSourceEdit::commit; reopen is diagnostic and excluded"
            ),
            "reported_statistics": list(REPORT_STATS),
            "comparison": lane_comparisons["native"],
        },
        "allocation": {
            "baseline_rows": by_stage_lane["baseline"]["alloc"],
            "candidate_rows": by_stage_lane["candidate"]["alloc"],
            "scope": "allocator binary operation_global_system_allocator",
            "reported_metrics": list(ALLOCATION_REPORT_FIELDS),
            "reported_workflow_metrics": [
                "allocation_calls", "deallocation_calls", "reallocation_calls",
                "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
                "retained_live_delta_bytes", "absolute_region_peak_live_bytes",
                "attributable_region_peak_live_bytes",
            ],
            "instrumented_elapsed_excluded": True,
            "comparison": lane_comparisons["alloc"],
        },
        "matched_identity": {
            "keys": ["corpus", "sink", "source", "output_sha256",
                     "semantic_sha256", "untouched_member_count",
                     "untouched_member_sha256", "catalog"],
            "case_shape_count": len(identity_groups),
            "all_lanes_and_repeats_equal": identity_equal,
        },
        "comparisons": {
            "numeric": all_numeric_comparisons,
            "exact_source": all_exact_comparisons,
            "adverse": all_adverse,
        },
        "repeat_drift": drift,
        "repeat_drift_over_five_percent": drift_over_five,
        "main_gates": gates,
        "field_paths": {
            "allocation_metrics": {
                phase: f"results[0].source.xlsx_cell_values.{phase}_allocation_metrics"
                for phase in ALLOCATION_PHASES
            },
            "absolute_live_after": "results[0].source.xlsx_cell_values.<phase>_allocation_metrics[*].live_bytes_after",
            "absolute_region_peak": "results[0].source.xlsx_cell_values.<phase>_allocation_metrics[*].region_peak_live_bytes",
            "publication_retention": "results[0].source.xlsx_cell_values.post_publication_budget[*]",
            "cache_retention": "results[0].source.xlsx_cell_values.cache_retained_bytes[*]",
            "output_retention": "results[0].sink.retained_output_bytes (optional; omitted by current CountingSink)",
        },
        "limits": [
            "Allocator-instrumented elapsed samples are excluded from latency interpretation.",
            "GNU time peak RSS is process-lifetime high water, not operation-local peak.",
            "Source/profile inference is not a counter for copied or reduced bytes.",
            "Guard, cap, quality, and profile evidence are external to this analyzer.",
            "No physical-cold, provider, native Office, concurrency, scaling, or fuzz claim.",
            "ODF remains deferred until the OLE2/OOXML optimization goal completes; iWork is excluded.",
        ],
    }


def write_identical(path: Path, value: dict[str, Any]) -> None:
    encoded = (json.dumps(value, indent=2, sort_keys=True, allow_nan=False) + "\n").encode()
    require(path.parent.exists(), f"output parent does not exist: {path.parent}")
    try:
        with path.open("xb") as stream:
            stream.write(encoded)
    except FileExistsError:
        try:
            existing = path.read_bytes()
        except OSError as error:
            raise EvidenceError(f"cannot replay output {path}: {error}") from error
        require(existing == encoded,
                f"existing output differs from deterministic replay: {path}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--stage", choices=("matched",), default="matched",
                        help="retained for interface symmetry; this analyzer always matches both stages")
    parser.add_argument("--output", type=Path)
    parser.add_argument("output_positional", nargs="?", type=Path)
    args = parser.parse_args(argv)
    output = args.output or args.output_positional or (HERE / "metrics-analysis.json")
    try:
        result = analyze()
        write_identical(output, result)
    except EvidenceError as error:
        print(f"evidence check failed: {error}", file=sys.stderr)
        return 1
    print(f"0553 matched metrics {result['status']}: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
