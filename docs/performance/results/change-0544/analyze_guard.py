#!/usr/bin/env python3
"""Validate and compare the standalone XLSX planning refusal guard.

The refusal guard is an evidence lane for the private shared worksheet
completion path.  It deliberately has a smaller report than the main benchmark:
the measured operation is only ``SourceBackedEditor::edit_sheets``.  This
analyzer authenticates both normal and allocator captures, checks the
correctness and source identities, and compares the complete timing and
allocation vectors.  It never starts a build or a benchmark.

If a capture is incomplete, ``analyze`` returns ``pending`` and does not make
up rows, samples, or gate results.  Once both ABBA lanes are complete, all
matched rows and all material allocation fields remain in the JSON output;
only the explicitly frozen guard thresholds decide admission.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import statistics
import sys
from datetime import datetime
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
PLAN_PATH = HERE / "plan.json"
RUN_PATH = HERE / "run.py"
GUARD_RUN_PATH = HERE / "guard_run.py"
ANALYZER_PATH = Path(__file__).resolve()

STAGES = ("baseline", "candidate")
LANES = ("normal", "alloc")
SHAPES = ("medium", "dense-sparse")
CASES = ("valid", "late-validator", "late-raw")
REPEATS = (1, 2)
TIMING_STATS = ("p50", "p95", "p99", "mean")
ALL_STATS = ("p50", "p95", "p99", "mean", "min", "max",
             "standard_deviation")
ALLOCATION_FIELDS = (
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
EXPECTED_BINARY = "xlsx_planning_guard"
EXPECTED_SCHEMA = "litchi.xlsx.planning-refusal-guard.v1"
EXPECTED_SCOPE = "operation_global_system_allocator"
ADVERSE_THRESHOLD_PERCENT = 5.0

EXPECTED_ERRORS = {
    "late-validator": "value-only edits refuse attribute 'future' on 'c'",
    "late-raw": "invalid worksheet boolean 'maybe'",
}
EXPECTED_DIMENSIONS = {
    "medium": (96, 96),
    "dense-sparse": (128, 128),
}


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory guard artifact."""


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def _read_json(path: Path) -> Any:
    try:
        with path.open(encoding="utf-8") as stream:
            return json.load(stream)
    except FileNotFoundError:
        raise
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {path}: {error}") from error


def _sha(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def _digest(value: Any, label: str) -> str:
    _require(isinstance(value, str) and re.fullmatch(r"[0-9a-f]{64}", value) is not None,
             f"{label} is not a lowercase SHA-256 digest")
    return value


def _nonempty_string(value: Any, label: str) -> str:
    _require(isinstance(value, str) and value, f"{label} is not a nonempty string")
    return value


def _nonnegative_integer(value: Any, label: str) -> int:
    _require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
             f"{label} is not a nonnegative integer")
    return value


def _positive_integer(value: Any, label: str) -> int:
    result = _nonnegative_integer(value, label)
    _require(result > 0, f"{label} is not positive")
    return result


def _finite_number(value: Any, label: str) -> float:
    _require(isinstance(value, (int, float)) and not isinstance(value, bool)
             and math.isfinite(float(value)), f"{label} is not finite")
    return float(value)


def _path(path: Path, label: str) -> None:
    _require(path.is_file() and not path.is_symlink(),
             f"{label} is not a regular file")


def _relative(path: Path) -> str:
    try:
        return str(path.relative_to(HERE))
    except ValueError as error:
        raise EvidenceError(f"path is outside guard evidence directory: {path}") from error


def _plan() -> dict[str, Any]:
    plan = _read_json(PLAN_PATH)
    _require(isinstance(plan, dict), "plan is not an object")
    revision = plan.get("revision")
    _require(isinstance(revision, str) and re.fullmatch(r"[0-9a-f]{40}", revision) is not None,
             "plan.revision is not a lowercase commit digest")
    cpu = plan.get("cpu")
    _require(cpu == 2, "plan.cpu must be 2 for the refusal guard")
    owned = plan.get("owned_paths")
    _require(isinstance(owned, list) and len(owned) == 1
             and isinstance(owned[0], str) and owned[0],
             "plan.owned_paths must contain one target directory")
    config = plan.get("refusal_guard")
    _require(isinstance(config, dict), "plan.refusal_guard is not an object")
    expected_keys = {
        "bin", "shapes", "cases", "repeats", "native_warmup", "native_samples",
        "allocation_warmup", "allocation_samples",
        "native_invalid_max_baseline_valid_ratio",
        "allocation_invalid_peak_max_baseline_valid_ratio",
        "ordinary_valid_max_regression_percent", "scope",
    }
    _require(set(config) == expected_keys,
             "plan.refusal_guard field inventory differs")
    _require(config["bin"] == EXPECTED_BINARY, "plan refusal guard binary differs")
    _require(config["shapes"] == list(SHAPES), "plan refusal guard shapes differ")
    _require(config["cases"] == list(CASES), "plan refusal guard cases differ")
    for key in ("repeats", "native_warmup", "native_samples",
                "allocation_warmup", "allocation_samples"):
        _positive_integer(config[key], f"plan.refusal_guard.{key}")
    _require(config["repeats"] == len(REPEATS),
             "plan refusal guard repeat count differs")
    for key, expected in (
        ("native_invalid_max_baseline_valid_ratio", 2.0),
        ("allocation_invalid_peak_max_baseline_valid_ratio", 1.1),
        ("ordinary_valid_max_regression_percent", 5.0),
    ):
        value = _finite_number(config[key], f"plan.refusal_guard.{key}")
        _require(value == expected, f"plan.refusal_guard.{key} differs")
    _nonempty_string(config["scope"], "plan.refusal_guard.scope")
    _path(GUARD_RUN_PATH, "guard_run.py")
    _path(RUN_PATH, "run.py")
    return plan


def _expected_jobs(plan: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    _require(lane in LANES, f"unsupported refusal guard lane {lane}")
    config = plan["refusal_guard"]
    warmup_key = "native_warmup" if lane == "normal" else "allocation_warmup"
    samples_key = "native_samples" if lane == "normal" else "allocation_samples"
    jobs: list[dict[str, Any]] = []
    # This is the order emitted by guard_run.capture: repeat two reverses the
    # shape loop to provide the second half of the ABBA sequence.
    for repeat in REPEATS:
        shapes = SHAPES if repeat == 1 else tuple(reversed(SHAPES))
        for shape in shapes:
            for case in CASES:
                jobs.append({
                    "name": (
                        f"guard-{'native' if lane == 'normal' else 'alloc'}-"
                        f"r{repeat}-{shape}-{case}"
                    ),
                    "stage": None,
                    "lane": lane,
                    "repeat": repeat,
                    "shape": shape,
                    "case": case,
                    "warmup": config[warmup_key],
                    "samples": config[samples_key],
                })
    return jobs


def _expected_job_map(plan: dict[str, Any], lane: str) -> dict[str, dict[str, Any]]:
    jobs = _expected_jobs(plan, lane)
    return {job["name"]: job for job in jobs}


def _parse_timestamp(value: Any, label: str) -> datetime:
    _require(isinstance(value, str) and value, f"{label} is missing")
    try:
        parsed = datetime.fromisoformat(value)
    except ValueError as error:
        raise EvidenceError(f"{label} is not an ISO-8601 timestamp") from error
    _require(parsed.tzinfo is not None, f"{label} has no timezone")
    return parsed


def _validate_build_command(command: Any, lane: str, target: str, label: str) -> None:
    _require(isinstance(command, list) and all(isinstance(item, str) for item in command),
             f"{label}.command is not a string list")
    required = [
        "cargo", "build", "--release", "--locked", "--manifest-path",
        "tools/perf-baseline/Cargo.toml", "--bin", EXPECTED_BINARY,
        "--target-dir",
    ]
    for item in required:
        _require(item in command, f"{label}.command is missing {item}")
    _require(command[command.index("--target-dir") + 1] == target,
             f"{label}.command target directory differs")
    if lane == "alloc":
        _require("--features" in command
                 and command[command.index("--features") + 1] == "allocator-metrics",
                 f"{label}.command is missing allocator-metrics")
    else:
        _require("--features" not in command,
                 f"{label}.command unexpectedly enables allocator metrics")
    _require(command.count("--bin") == 1 and command.count("--manifest-path") == 1,
             f"{label}.command has duplicate cargo selectors")


def _check_binary(stage: str, lane: str, plan: dict[str, Any],
                  manifest_sha: str) -> dict[str, Any]:
    folder = HERE / stage
    suffix = "normal" if lane == "normal" else "alloc"
    identity_path = folder / f"binary-guard-{suffix}.json"
    identity = _read_json(identity_path)
    label = f"{stage}/binary-guard-{suffix}"
    _require(isinstance(identity, dict), f"{label} identity is not an object")
    expected_keys = {"path", "sha256", "bytes", "build_receipt_sha256",
                     "source_manifest_sha256"}
    _require(expected_keys <= set(identity), f"{label} identity is incomplete")
    binary_digest = _digest(identity["sha256"], f"{label}.sha256")
    byte_count = _positive_integer(identity["bytes"], f"{label}.bytes")
    binary_path = Path(_nonempty_string(identity["path"], f"{label}.path"))
    expected_path = Path(plan["owned_paths"][0]) / "retained-binaries" / \
        f"{stage}-guard-{suffix}"
    _require(binary_path == expected_path,
             f"{label}.path differs from retained guard path")
    if binary_path.exists():
        _path(binary_path, label)
        _require(_sha(binary_path) == binary_digest, f"{label} digest differs")
        _require(binary_path.stat().st_size == byte_count, f"{label} byte count differs")
    else:
        # A completed cleanup may remove the retained binaries after capture.
        cleanup_path = HERE / "cleanup.json"
        _path(cleanup_path, "cleanup.json")
        cleanup = _read_json(cleanup_path)
        _require(isinstance(cleanup, dict)
                 and cleanup.get("owned_paths_absent") is True
                 and cleanup.get("accessible_process_references") == []
                 and cleanup.get("removed") == plan["owned_paths"]
                 and all(not Path(item).exists() for item in plan["owned_paths"]),
                 f"{label} is missing without completed cleanup custody")
    _require(identity["source_manifest_sha256"] == manifest_sha,
             f"{label} source manifest digest differs")

    build_receipt_path = folder / f"build-guard-{suffix}.receipt.json"
    build_receipt = _read_json(build_receipt_path)
    build_label = f"{stage}/build-guard-{suffix}"
    _require(identity["build_receipt_sha256"] == _sha(build_receipt_path),
             f"{label} build receipt digest differs")
    _require(isinstance(build_receipt, dict), f"{build_label} receipt is not an object")
    _require(build_receipt.get("exit_code") == 0
             and build_receipt.get("binary_sha256") is None,
             f"{build_label} receipt did not complete as a build")
    _require(build_receipt.get("plan_sha256") == _sha(PLAN_PATH),
             f"{build_label} plan digest differs")
    _require(build_receipt.get("script_sha256") == _sha(RUN_PATH),
             f"{build_label} run script digest differs")
    _require(build_receipt.get("source_manifest_sha256") == manifest_sha,
             f"{build_label} source manifest digest differs")
    _validate_build_command(build_receipt.get("command"), lane,
                            plan["owned_paths"][0], build_label)
    artifacts = build_receipt.get("artifacts")
    _require(isinstance(artifacts, dict)
             and set(artifacts) == {f"build-guard-{suffix}.stdout",
                                    f"build-guard-{suffix}.stderr"},
             f"{build_label} artifact inventory differs")
    for name, digest in artifacts.items():
        artifact = folder / name
        _path(artifact, f"{build_label}/{name}")
        _require(_sha(artifact) == digest, f"{build_label}/{name} digest differs")
    start = _parse_timestamp(build_receipt.get("start_utc"), f"{build_label}.start_utc")
    end = _parse_timestamp(build_receipt.get("end_utc"), f"{build_label}.end_utc")
    _require(start <= end, f"{build_label} timestamps are reversed")
    seconds = _finite_number(build_receipt.get("seconds"), f"{build_label}.seconds")
    _require(seconds >= 0.0, f"{build_label}.seconds is negative")
    return {
        "path": str(binary_path),
        "sha256": binary_digest,
        "bytes": byte_count,
        "build_receipt_sha256": identity["build_receipt_sha256"],
        "source_manifest_sha256": manifest_sha,
        "build_receipt": {
            "path": _relative(build_receipt_path),
            "sha256": _sha(build_receipt_path),
            "start_utc": build_receipt["start_utc"],
            "end_utc": build_receipt["end_utc"],
            "seconds": seconds,
        },
    }


def _stats(values: list[int], label: str, unit: str) -> dict[str, Any]:
    _require(isinstance(values, list) and values, f"{label} is empty")
    for index, value in enumerate(values):
        _nonnegative_integer(value, f"{label}[{index}]")
    ordered = sorted(values)
    count = len(ordered)
    mean = sum(ordered) / count
    if count == 1:
        standard_deviation = 0.0
    else:
        standard_deviation = statistics.stdev(ordered)
    p95_index = min(math.ceil(count * 0.95) - 1, count - 1)
    p99_index = min(math.ceil(count * 0.99) - 1, count - 1)
    return {
        "unit": unit,
        "samples": list(values),
        "sorted_samples": ordered,
        "p50": (ordered[(count - 1) // 2] + ordered[count // 2]) // 2,
        "p95": ordered[p95_index],
        "p99": ordered[p99_index],
        "min": ordered[0],
        "max": ordered[-1],
        "mean": mean,
        "standard_deviation": standard_deviation,
    }


def _validate_allocation_sample(sample: Any, lane: str, label: str) -> dict[str, Any]:
    _require(isinstance(sample, dict), f"{label} is not an allocation object")
    status = sample.get("status")
    scope = sample.get("scope")
    _require(scope == EXPECTED_SCOPE, f"{label}.scope differs")
    if lane == "normal":
        _require(status == "unavailable", f"{label}.status must be unavailable")
        for field in ALLOCATION_FIELDS:
            _require(field not in sample or sample[field] is None,
                     f"{label}.{field} is present in normal lane")
        return {"status": status, "scope": scope,
                **{field: None for field in ALLOCATION_FIELDS}}
    _require(status == "measured", f"{label}.status must be measured")
    values: dict[str, int] = {}
    for field in ALLOCATION_FIELDS:
        values[field] = _nonnegative_integer(sample.get(field), f"{label}.{field}")
    _require(values["failed_allocation_calls"] == 0,
             f"{label} recorded a failed allocation")
    _require(values["live_bytes_before"] + values["allocated_bytes"]
             == values["live_bytes_after"] + values["deallocated_bytes"],
             f"{label} live-byte balance does not reconcile")
    _require(values["peak_live_bytes_before"] >= values["live_bytes_before"],
             f"{label} pre-operation peak is below live bytes")
    _require(values["peak_live_bytes_after"] >= values["peak_live_bytes_before"]
             and values["peak_live_bytes_after"] >= values["live_bytes_after"],
             f"{label} post-operation peak is invalid")
    _require(values["region_peak_live_bytes"] >= max(values["live_bytes_before"],
                                                     values["live_bytes_after"])
             and values["region_peak_live_bytes"] <= values["peak_live_bytes_after"],
             f"{label} region peak is outside the process peak envelope")
    return {"status": status, "scope": scope, **values}


def _validate_source(source: Any, shape: str, label: str) -> dict[str, Any]:
    _require(isinstance(source, dict), f"{label}.source is not an object")
    required = {
        "generator", "shape", "rows", "columns", "archive_bytes", "worksheet_bytes",
        "source_sha256", "worksheet_sha256", "worksheet_member", "compression",
        "fixture_kind",
    }
    _require(required <= set(source), f"{label}.source is incomplete")
    _require(source["shape"] == shape, f"{label}.source.shape differs")
    rows, columns = EXPECTED_DIMENSIONS[shape]
    _require(source["rows"] == rows and source["columns"] == columns,
             f"{label}.source dimensions differ")
    for field in ("rows", "columns", "archive_bytes", "worksheet_bytes"):
        _positive_integer(source[field], f"{label}.source.{field}")
    _nonempty_string(source["generator"], f"{label}.source.generator")
    _digest(source["source_sha256"], f"{label}.source.source_sha256")
    _digest(source["worksheet_sha256"], f"{label}.source.worksheet_sha256")
    _require(source["worksheet_member"] == "xl/worksheets/sheet1.xml",
             f"{label}.source.worksheet_member differs")
    _require(source["compression"] == "stored", f"{label}.source.compression differs")
    _require(source["fixture_kind"] == "raw_zip_input_no_opc_authored_xml_validation",
             f"{label}.source.fixture_kind differs")
    return source


def _validate_logical_error(value: Any, case: str, label: str) -> dict[str, Any]:
    _require(isinstance(value, dict), f"{label}.logical_error is not an object")
    expected = EXPECTED_ERRORS.get(case)
    if expected is None:
        _require(value == {"status": "accepted", "variant": None, "message": None},
                 f"{label}.logical_error differs for valid input")
    else:
        _require(value == {"status": "expected_failure", "variant": "Invalid",
                           "message": expected},
                 f"{label}.logical_error differs for invalid input")
    return value


def _validate_correctness(value: Any, case: str, label: str) -> dict[str, Any]:
    _require(isinstance(value, dict), f"{label}.correctness is not an object")
    invalid = case in EXPECTED_ERRORS
    expected = {
        "source_unchanged": True,
        "retry_preserved_error": invalid,
        "valid_snapshot_values": not invalid,
        "empty_commit_is_noop": not invalid,
        "expected_error_exact": invalid,
        # Valid rows perform the empty commit after the measured operation;
        # refusal rows have no transaction to commit, so the oracle is false.
        "commit_outside_timing": not invalid,
    }
    _require(value == expected, f"{label}.correctness oracle differs")
    return value


def _validate_phase(raw: dict[str, Any], job: dict[str, Any], lane: str,
                    label: str) -> tuple[dict[str, Any], dict[str, Any]]:
    phase = raw.get("phase")
    _require(isinstance(phase, dict), f"{label}.phase is not an object")
    _require(phase.get("name") == "edit_sheets", f"{label}.phase.name differs")
    count = job["samples"]
    duration = phase.get("duration_ns")
    _require(isinstance(duration, list) and len(duration) == count,
             f"{label}.phase.duration_ns cardinality differs")
    for index, value in enumerate(duration):
        _nonnegative_integer(value, f"{label}.phase.duration_ns[{index}]")
    order = phase.get("sample_order")
    _require(isinstance(order, list) and order == list(range(count)),
             f"{label}.phase.sample_order differs")
    samples = phase.get("samples")
    _require(isinstance(samples, list) and len(samples) == count,
             f"{label}.phase.samples cardinality differs")
    allocations: list[dict[str, Any]] = []
    for index, sample in enumerate(samples):
        _require(isinstance(sample, dict), f"{label}.phase.samples[{index}] is not an object")
        _require(sample.get("order") == index,
                 f"{label}.phase.samples[{index}].order differs")
        _require(sample.get("duration_ns") == duration[index],
                 f"{label}.phase.samples[{index}] duration disagrees")
        allocations.append(_validate_allocation_sample(
            sample.get("allocation_metrics"), lane,
            f"{label}.phase.samples[{index}].allocation_metrics"))
    vector = phase.get("allocation_metrics")
    _require(isinstance(vector, list) and len(vector) == count,
             f"{label}.phase.allocation_metrics cardinality differs")
    for index, sample in enumerate(vector):
        duplicate = _validate_allocation_sample(
            sample, lane, f"{label}.phase.allocation_metrics[{index}]")
        _require(duplicate == allocations[index],
                 f"{label}.phase.allocation_metrics[{index}] disagrees with sample")
    timing = _stats(duration, f"{label}.phase.duration_ns", "ns")
    allocation_vectors: dict[str, list[int] | list[None]] = {}
    for field in ALLOCATION_FIELDS:
        allocation_vectors[field] = [sample.get(field) for sample in allocations]
    allocation_stats: dict[str, Any] = {}
    if lane == "alloc":
        for field in ALLOCATION_FIELDS:
            values = allocation_vectors[field]
            _require(all(isinstance(value, int) for value in values),
                     f"{label}.allocation.{field} contains unavailable values")
            allocation_stats[field] = _stats(values, f"{label}.allocation.{field}",
                                              "count" if field.endswith("calls")
                                              else "bytes")
    return timing, {
        "status": "unavailable" if lane == "normal" else "measured",
        "scope": EXPECTED_SCOPE,
        "samples": allocations,
        "vectors": allocation_vectors,
        "statistics": allocation_stats,
    }


def _validate_report(report: Any, stage: str, lane: str,
                     job: dict[str, Any], binary: dict[str, Any],
                     manifest_sha: str) -> dict[str, Any]:
    label = f"{stage}/{job['name']}"
    _require(isinstance(report, dict), f"{label} report is not an object")
    _require(report.get("schema") == EXPECTED_SCHEMA, f"{label}.schema differs")
    _require(report.get("tool") == EXPECTED_BINARY, f"{label}.tool differs")
    _require(report.get("case") == job["case"] and report.get("shape") == job["shape"],
             f"{label} case/shape differs")
    _require(report.get("warmup_iterations") == job["warmup"]
             and report.get("samples") == job["samples"],
             f"{label} warmup/sample counts differ")
    source = _validate_source(report.get("source"), job["shape"], label)
    _digest(report.get("source_sha256"), f"{label}.source_sha256")
    _require(report["source_sha256"] == source["source_sha256"],
             f"{label}.source_sha256 disagrees with source identity")
    _validate_logical_error(report.get("logical_error"), job["case"], label)
    correctness = _validate_correctness(report.get("correctness"), job["case"], label)
    timing, allocation = _validate_phase(report, job, lane, label)
    _require(report.get("allocation_scope") == EXPECTED_SCOPE,
             f"{label}.allocation_scope differs")
    expected_allocator = (
        "Rust system allocator" if lane == "normal"
        else "CountingSystemAllocator(std::alloc::System)")
    expected_instrumentation = (
        "none" if lane == "normal" else "system_allocator_operation_scoped")
    expected_counter_revision = None if lane == "normal" else "serialized_region_peak_v3"
    _require(report.get("allocator") == expected_allocator,
             f"{label}.allocator identity differs")
    _require(report.get("instrumentation") == expected_instrumentation,
             f"{label}.instrumentation identity differs")
    _require(report.get("counter_revision") == expected_counter_revision,
             f"{label}.counter_revision differs")
    _nonempty_string(report.get("timing_scope"), f"{label}.timing_scope")
    _require(report.get("performance_claim")
             == "none: correctness and planning diagnostic only; no speedup claim",
             f"{label}.performance_claim differs")

    binary_report = report.get("binary")
    _require(isinstance(binary_report, dict), f"{label}.binary is not an object")
    _require(binary_report.get("path") == binary["path"],
             f"{label}.binary.path differs")
    _require(binary_report.get("sha256") == binary["sha256"],
             f"{label}.binary.sha256 differs")
    _require(binary_report.get("bytes") == binary["bytes"],
             f"{label}.binary.bytes differs")
    runner = report.get("runner")
    _require(isinstance(runner, dict), f"{label}.runner is not an object")
    _require(runner.get("git_revision") == _plan_revision,
             f"{label}.runner.git_revision differs")
    _require(runner.get("profile") == "release", f"{label}.runner.profile differs")

    return {
        "stage": stage,
        "name": job["name"],
        "lane": lane,
        "repeat": job["repeat"],
        "shape": job["shape"],
        "case": job["case"],
        "warmup": job["warmup"],
        "samples": job["samples"],
        "source": source,
        "source_sha256": source["source_sha256"],
        "logical_error": report["logical_error"],
        "correctness": correctness,
        "timing": {"edit_sheets": timing},
        "allocation": allocation,
        "binary": {"path": binary["path"], "sha256": binary["sha256"],
                   "bytes": binary["bytes"]},
        "report_schema": report["schema"],
    }


def _validate_capture_command(receipt: dict[str, Any], job: dict[str, Any],
                             binary: dict[str, Any], report_path: Path,
                             stage: str, label: str) -> None:
    command = receipt.get("command")
    _require(isinstance(command, list) and all(isinstance(item, str) for item in command),
             f"{label}.command is not a string list")
    expected = [
        "taskset", "-c", str(_plan_cpu), binary["path"],
        "--shape", job["shape"], "--case", job["case"],
        "--warmup", str(job["warmup"]), "--samples", str(job["samples"]),
        "--json", str(report_path),
    ]
    _require(command == expected, f"{label}.command differs from guard_run.py")
    _require(stage in STAGES, f"{label} stage is unsupported")


def _check_capture(stage: str, lane: str, job: dict[str, Any],
                   binary: dict[str, Any], manifest_sha: str,
                   candidate_manifest_sha: str, plan: dict[str, Any]) -> dict[str, Any]:
    folder = HERE / stage
    label = f"{stage}/{job['name']}"
    report_path = folder / f"{job['name']}.json"
    receipt_path = folder / f"{job['name']}.receipt.json"
    stdout_path = folder / f"{job['name']}.stdout"
    stderr_path = folder / f"{job['name']}.stderr"
    for path, name in ((report_path, "report"), (receipt_path, "receipt"),
                       (stdout_path, "stdout"), (stderr_path, "stderr")):
        _path(path, f"{label}/{name}")
    receipt = _read_json(receipt_path)
    _require(isinstance(receipt, dict), f"{label}.receipt is not an object")
    _require(receipt.get("exit_code") == 0, f"{label}.receipt did not complete")
    _require(receipt.get("binary_sha256") == binary["sha256"],
             f"{label}.receipt binary digest differs")
    _require(receipt.get("source_manifest_sha256") == manifest_sha,
             f"{label}.receipt source manifest differs")
    working_manifest_sha = manifest_sha
    if stage == "candidate" or job["repeat"] == 2:
        working_manifest_sha = candidate_manifest_sha
    _require(receipt.get("working_source_manifest_sha256") == working_manifest_sha,
             f"{label}.receipt working source manifest differs")
    _require(receipt.get("script_sha256") == _sha(RUN_PATH),
             f"{label}.receipt run script digest differs")
    _require(receipt.get("plan_sha256") == _sha(PLAN_PATH),
             f"{label}.receipt plan digest differs")
    _validate_capture_command(receipt, job, binary, report_path, stage, label)
    start = _parse_timestamp(receipt.get("start_utc"), f"{label}.receipt.start_utc")
    end = _parse_timestamp(receipt.get("end_utc"), f"{label}.receipt.end_utc")
    _require(start <= end, f"{label}.receipt timestamps are reversed")
    seconds = _finite_number(receipt.get("seconds"), f"{label}.receipt.seconds")
    _require(seconds >= 0.0, f"{label}.receipt.seconds is negative")
    artifacts = receipt.get("artifacts")
    expected_artifacts = {report_path.name, stdout_path.name, stderr_path.name}
    _require(isinstance(artifacts, dict) and set(artifacts) == expected_artifacts,
             f"{label}.receipt artifact inventory differs")
    for name, digest in artifacts.items():
        artifact = folder / name
        _require(_digest(digest, f"{label}.receipt.artifacts.{name}") == _sha(artifact),
                 f"{label}.receipt artifact digest differs: {name}")
    report_bytes = report_path.read_bytes()
    stdout_bytes = stdout_path.read_bytes()
    _require(stdout_bytes.rstrip(b"\n") == report_bytes,
             f"{label}.stdout does not reproduce the report")
    report = _read_json(report_path)
    row = _validate_report(report, stage, lane, job, binary, manifest_sha)
    row["report_sha256"] = _sha(report_path)
    row["receipt"] = {
        "path": _relative(receipt_path),
        "sha256": _sha(receipt_path),
        "start_utc": receipt["start_utc"],
        "end_utc": receipt["end_utc"],
        "seconds": seconds,
        "binary_sha256": binary["sha256"],
        "source_manifest_sha256": manifest_sha,
        "working_source_manifest_sha256": working_manifest_sha,
        "artifact_sha256": {name: artifacts[name] for name in sorted(artifacts)},
    }
    return row


def _manifest(stage: str) -> tuple[dict[str, str], str]:
    path = HERE / stage / "source-manifest.json"
    _path(path, f"{stage}/source-manifest.json")
    value = _read_json(path)
    _require(isinstance(value, dict) and value, f"{stage} source manifest is empty")
    for name, digest in value.items():
        _require(isinstance(name, str) and name and not Path(name).is_absolute(),
                 f"{stage} source manifest path is invalid")
        _digest(digest, f"{stage} source manifest {name}")
    guard_sources = (
        "tools/perf-baseline/src/xlsx_planning_guard.rs",
        "tools/perf-baseline/src/bin/xlsx_planning_guard.rs",
    )
    _require(any(name in value for name in guard_sources),
             f"{stage} source manifest omits the standalone guard source")
    return value, _sha(path)


def _check_stage(stage: str, plan: dict[str, Any], lane: str,
                 candidate_manifest_sha: str) -> dict[str, Any]:
    folder = HERE / stage
    _require(folder.is_dir(), f"{stage} stage directory is missing")
    manifest, manifest_sha = _manifest(stage)
    binary = _check_binary(stage, lane, plan, manifest_sha)
    expected = _expected_job_map(plan, lane)
    prefix = "guard-native-" if lane == "normal" else "guard-alloc-"
    actual = {
        path.name[:-len(".receipt.json")]
        for path in folder.glob("*.receipt.json")
        if path.name.startswith(prefix)
    }
    _require(actual == set(expected),
             f"{stage}/{lane} receipt set differs: {sorted(actual ^ set(expected))}")
    rows: dict[str, dict[str, Any]] = {}
    receipts: list[tuple[datetime, datetime, str]] = []
    for name, job in expected.items():
        row = _check_capture(stage, lane, job, binary, manifest_sha,
                             candidate_manifest_sha, plan)
        rows[name] = row
        start = _parse_timestamp(row["receipt"]["start_utc"], f"{stage}/{name}.start_utc")
        end = _parse_timestamp(row["receipt"]["end_utc"], f"{stage}/{name}.end_utc")
        receipts.append((start, end, name))
    ordered = sorted(receipts, key=lambda item: (item[0], item[1], item[2]))
    _require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
             f"{stage}/{lane} child receipts overlap")
    return {
        "stage": stage,
        "lane": lane,
        "manifest_sha256": manifest_sha,
        "manifest_entries": len(manifest),
        "binary": binary,
        "rows": [rows[name] for name in expected],
        "row_count": len(rows),
        "total_samples": sum(job["samples"] for job in expected.values()),
        "receipts_non_overlapping": True,
    }


def _missing_stage(stage: str, lane: str, plan: dict[str, Any]) -> list[str]:
    folder = HERE / stage
    if not folder.is_dir():
        return [stage]
    suffix = "normal" if lane == "normal" else "alloc"
    missing: list[str] = []
    for name in ("source-manifest.json", f"binary-guard-{suffix}.json",
                 f"build-guard-{suffix}.receipt.json"):
        if not (folder / name).is_file():
            missing.append(f"{stage}/{name}")
    for job in _expected_job_map(plan, lane).values():
        for suffix_name in (".json", ".receipt.json", ".stdout", ".stderr"):
            if not (folder / (job["name"] + suffix_name)).is_file():
                missing.append(f"{stage}/{job['name']}{suffix_name}")
    return missing


def _stage_or_pending(stage: str, lane: str, plan: dict[str, Any],
                      candidate_manifest_sha: str | None = None) -> dict[str, Any]:
    missing = _missing_stage(stage, lane, plan)
    if missing:
        return {
            "status": "pending",
            "stage": stage,
            "lane": lane,
            "missing_artifacts": missing,
            "expected_row_count": len(_expected_jobs(plan, lane)),
        }
    if candidate_manifest_sha is None:
        candidate_folder = HERE / "candidate" / "source-manifest.json"
        if not candidate_folder.is_file():
            return {
                "status": "pending",
                "stage": stage,
                "lane": lane,
                "missing_artifacts": ["candidate/source-manifest.json"],
                "expected_row_count": len(_expected_jobs(plan, lane)),
            }
        candidate_manifest_sha = _sha(candidate_folder)
    return {"status": "pass", "evidence": _check_stage(
        stage, plan, lane, candidate_manifest_sha)}


def _identity(row: dict[str, Any]) -> dict[str, Any]:
    return {
        "source": row["source"],
        "source_sha256": row["source_sha256"],
        "logical_error": row["logical_error"],
        "correctness": row["correctness"],
    }


def _comparison_record(baseline: float | int, candidate: float | int,
                      label: str) -> dict[str, Any]:
    _finite_number(baseline, f"{label}.baseline")
    _finite_number(candidate, f"{label}.candidate")
    _require(float(baseline) >= 0.0 and float(candidate) >= 0.0,
             f"{label} contains a negative value")
    if float(baseline) == 0.0:
        change = 0.0 if float(candidate) == 0.0 else None
    else:
        change = (float(candidate) / float(baseline) - 1.0) * 100.0
    return {
        "baseline": baseline,
        "candidate": candidate,
        "delta": candidate - baseline,
        "change_percent": change,
    }


def _adverse(adverse: list[dict[str, Any]], *, lane: str, case: str,
             shape: str, repeat: int, metric: str, value: dict[str, Any]) -> None:
    change = value["change_percent"]
    zero_adverse = value["baseline"] == 0 and value["candidate"] > 0
    if (change is not None and change > ADVERSE_THRESHOLD_PERCENT) or zero_adverse:
        adverse.append({
            "lane": lane,
            "case": case,
            "shape": shape,
            "repeat": repeat,
            "metric": metric,
            "threshold_percent": ADVERSE_THRESHOLD_PERCENT,
            "baseline_zero_adverse": zero_adverse,
            **value,
        })


def _drift(drift: list[dict[str, Any]], *, stage: str, lane: str,
           case: str, shape: str, metric: str, first: dict[str, Any],
           second: dict[str, Any]) -> None:
    value = _comparison_record(first["value"], second["value"],
                               f"{stage}/{lane}/{case}/{shape}/{metric}")
    change = value["change_percent"]
    zero_adverse = value["baseline"] == 0 and value["candidate"] > 0
    if (change is not None and abs(change) > ADVERSE_THRESHOLD_PERCENT) or zero_adverse:
        drift.append({
            "stage": stage,
            "lane": lane,
            "case": case,
            "shape": shape,
            "repeat_first": 1,
            "repeat_second": 2,
            "metric": metric,
            "threshold_percent": ADVERSE_THRESHOLD_PERCENT,
            "baseline_zero_drift": zero_adverse,
            **value,
        })


def _check_abba(baseline: dict[str, Any], candidate: dict[str, Any],
                plan: dict[str, Any], lane: str) -> dict[str, Any]:
    jobs = _expected_jobs(plan, lane)
    by_stage = {
        "baseline": {row["name"]: row for row in baseline["rows"]},
        "candidate": {row["name"]: row for row in candidate["rows"]},
    }
    groups: list[tuple[str, int, dict[str, dict[str, Any]]]] = [
        ("baseline", 1, by_stage["baseline"]),
        ("candidate", 1, by_stage["candidate"]),
        ("candidate", 2, by_stage["candidate"]),
        ("baseline", 2, by_stage["baseline"]),
    ]
    group_records: list[dict[str, Any]] = []
    for stage, repeat, rows in groups:
        names = [job["name"] for job in jobs if job["repeat"] == repeat]
        ordered = sorted(
            ((row["receipt"]["start_utc"], row["receipt"]["end_utc"], name)
             for name, row in rows.items() if row["repeat"] == repeat),
            key=lambda item: (item[0], item[1], item[2]),
        )
        _require([item[2] for item in ordered] == names,
                 f"{lane} {stage}-r{repeat} child order differs from guard_run.py")
        first = _parse_timestamp(ordered[0][0], f"{lane}/{stage}-r{repeat}.first")
        last = _parse_timestamp(ordered[-1][1], f"{lane}/{stage}-r{repeat}.last")
        group_records.append({
            "group": f"{stage}-r{repeat}",
            "stage": stage,
            "repeat": repeat,
            "first_start_utc": ordered[0][0],
            "last_end_utc": ordered[-1][1],
            "job_names": [item[2] for item in ordered],
        })
        if group_records and len(group_records) > 1:
            previous = group_records[-2]
            previous_end = _parse_timestamp(previous["last_end_utc"],
                                            f"{lane}/{previous['group']}.last")
            _require(previous_end <= first,
                     f"{lane} ABBA groups overlap: {previous['group']} and {stage}-r{repeat}")
    return {
        "order": [item["group"] for item in group_records],
        "groups": group_records,
        "passed": True,
    }


def _compare_lane(baseline: dict[str, Any], candidate: dict[str, Any],
                  plan: dict[str, Any], lane: str) -> dict[str, Any]:
    left = {row["name"]: row for row in baseline["rows"]}
    right = {row["name"]: row for row in candidate["rows"]}
    _require(set(left) == set(right), f"{lane} baseline/candidate rows differ")
    rows: list[dict[str, Any]] = []
    adverse: list[dict[str, Any]] = []
    for name in _expected_job_map(plan, lane):
        before, after = left[name], right[name]
        _require(_identity(before) == _identity(after),
                 f"{lane} logical/corpus identity differs for {name}")
        timing_metrics: dict[str, Any] = {}
        for stat in ALL_STATS:
            value = _comparison_record(before["timing"]["edit_sheets"][stat],
                                       after["timing"]["edit_sheets"][stat],
                                       f"{lane}/{name}/timing/{stat}")
            timing_metrics[stat] = value
            _adverse(adverse, lane=lane, case=before["case"], shape=before["shape"],
                     repeat=before["repeat"], metric=f"timing.edit_sheets.{stat}",
                     value=value)
        allocation_metrics: dict[str, Any] = {}
        if lane == "alloc":
            for field in ALLOCATION_FIELDS:
                allocation_metrics[field] = {}
                before_stats = before["allocation"]["statistics"][field]
                after_stats = after["allocation"]["statistics"][field]
                for stat in ALL_STATS:
                    value = _comparison_record(before_stats[stat], after_stats[stat],
                                               f"{lane}/{name}/allocation/{field}/{stat}")
                    allocation_metrics[field][stat] = value
                    _adverse(adverse, lane=lane, case=before["case"], shape=before["shape"],
                             repeat=before["repeat"], metric=f"allocation.{field}.{stat}",
                             value=value)
        rows.append({
            "name": name,
            "case": before["case"],
            "shape": before["shape"],
            "repeat": before["repeat"],
            "identity_equal": True,
            "baseline": {
                "timing": before["timing"],
                "allocation": before["allocation"],
            },
            "candidate": {
                "timing": after["timing"],
                "allocation": after["allocation"],
            },
            "comparison": {
                "timing": timing_metrics,
                "allocation": allocation_metrics,
            },
        })
    return {"rows": rows, "adverse_flags_over_five_percent": adverse}


def _cross_lane_corpus(stage: str, normal: dict[str, Any],
                       allocation: dict[str, Any]) -> dict[str, Any]:
    """Ensure normal and allocator binaries measured the same fixture/oracle."""

    left = {(row["repeat"], row["shape"], row["case"]): row
            for row in normal["rows"]}
    right = {(row["repeat"], row["shape"], row["case"]): row
             for row in allocation["rows"]}
    _require(set(left) == set(right), f"{stage} normal/allocator row keys differ")
    rows: list[dict[str, Any]] = []
    for key in sorted(left):
        _require(_identity(left[key]) == _identity(right[key]),
                 f"{stage} normal/allocator corpus identity differs for {key}")
        rows.append({
            "stage": stage,
            "repeat": key[0],
            "shape": key[1],
            "case": key[2],
            "identity_equal": True,
            "source_sha256": left[key]["source_sha256"],
        })
    return {"rows": rows, "passed": True}


def _cross_lane_receipt_custody(stage: str, normal: dict[str, Any],
                                allocation: dict[str, Any]) -> dict[str, Any]:
    """Require the two capture lanes to remain serial within each stage."""

    intervals = []
    for evidence in (normal, allocation):
        for row in evidence["rows"]:
            intervals.append((
                _parse_timestamp(row["receipt"]["start_utc"],
                                 f"{stage}/{row['name']}.start_utc"),
                _parse_timestamp(row["receipt"]["end_utc"],
                                 f"{stage}/{row['name']}.end_utc"),
                row["name"],
            ))
    ordered = sorted(intervals, key=lambda item: (item[0], item[1], item[2]))
    _require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
             f"{stage} normal/allocator child receipts overlap")
    return {"passed": True, "receipt_count": len(ordered),
            "receipts_non_overlapping": True}


def _valid_gate(normal_baseline: dict[str, Any], normal_candidate: dict[str, Any],
                plan: dict[str, Any]) -> list[dict[str, Any]]:
    before = {(row["repeat"], row["shape"]): row for row in normal_baseline["rows"]
              if row["case"] == "valid"}
    after = {(row["repeat"], row["shape"]): row for row in normal_candidate["rows"]
             if row["case"] == "valid"}
    _require(set(before) == {(repeat, shape) for repeat in REPEATS for shape in SHAPES},
             "normal valid baseline matrix is incomplete")
    _require(set(after) == set(before), "normal valid candidate matrix is incomplete")
    threshold = float(plan["refusal_guard"]["ordinary_valid_max_regression_percent"])
    checks: list[dict[str, Any]] = []
    for key in sorted(before):
        left = before[key]["timing"]["edit_sheets"]
        right = after[key]["timing"]["edit_sheets"]
        value = _comparison_record(left["p50"], right["p50"],
                                   f"valid p50 {key}")
        change = value["change_percent"]
        passed = change is not None and change <= threshold
        checks.append({
            "repeat": key[0], "shape": key[1], "metric": "timing.edit_sheets.p50",
            **value, "max_allowed_regression_percent": threshold, "passed": passed,
        })
    return checks


def _invalid_native_gate(normal_baseline: dict[str, Any], normal_candidate: dict[str, Any],
                         plan: dict[str, Any]) -> list[dict[str, Any]]:
    left = {(row["repeat"], row["shape"], row["case"]): row
            for row in normal_baseline["rows"] if row["case"] == "valid"}
    candidate_valid = {(row["repeat"], row["shape"], row["case"]): row
                       for row in normal_candidate["rows"] if row["case"] == "valid"}
    threshold = float(plan["refusal_guard"]["native_invalid_max_baseline_valid_ratio"])
    checks: list[dict[str, Any]] = []
    for row in normal_candidate["rows"]:
        if row["case"] == "valid":
            continue
        key = (row["repeat"], row["shape"], "valid")
        _require(key in left and key in candidate_valid,
                 f"missing valid baseline/candidate for {row['name']}")
        baseline_value = left[key]["timing"]["edit_sheets"]["p50"]
        candidate_value = row["timing"]["edit_sheets"]["p50"]
        _require(baseline_value >= 0 and candidate_value >= 0,
                 f"invalid native gate has a negative value for {row['name']}")
        ratio = None if baseline_value == 0 else candidate_value / baseline_value
        passed = candidate_value == 0 if baseline_value == 0 else ratio <= threshold
        checks.append({
            "case": row["case"], "shape": row["shape"], "repeat": row["repeat"],
            "metric": "timing.edit_sheets.p50",
            "baseline_valid_p50": baseline_value,
            "candidate_invalid_p50": candidate_value,
            "candidate_to_baseline_valid_ratio": ratio,
            "max_allowed_ratio": threshold,
            "passed": passed,
        })
    return checks


def _incremental_peak(row: dict[str, Any], field: str) -> list[int]:
    allocation = row["allocation"]
    _require(allocation["status"] == "measured", f"{row['name']} allocation is not measured")
    before = allocation["vectors"]["live_bytes_before"]
    peak = allocation["vectors"][field]
    _require(all(isinstance(value, int) for value in before + peak),
             f"{row['name']} allocation peak vector is incomplete")
    result = [current - initial for current, initial in zip(peak, before)]
    _require(all(value >= 0 for value in result), f"{row['name']} incremental peak is negative")
    return result


def _invalid_allocation_gate(allocation_baseline: dict[str, Any],
                             allocation_candidate: dict[str, Any],
                             plan: dict[str, Any]) -> list[dict[str, Any]]:
    rows_left = {(row["repeat"], row["shape"], row["case"]): row
                 for row in allocation_baseline["rows"]}
    rows_right = {(row["repeat"], row["shape"], row["case"]): row
                  for row in allocation_candidate["rows"]}
    threshold = float(plan["refusal_guard"]["allocation_invalid_peak_max_baseline_valid_ratio"])
    checks: list[dict[str, Any]] = []
    for (repeat, shape, case), row in sorted(rows_right.items()):
        if case == "valid":
            continue
        valid_key = (repeat, shape, "valid")
        _require(valid_key in rows_left, f"missing valid allocation baseline for {row['name']}")
        valid_values = _incremental_peak(rows_left[valid_key], "region_peak_live_bytes")
        invalid_values = _incremental_peak(row, "region_peak_live_bytes")
        baseline_min = min(valid_values)
        candidate_max = max(invalid_values)
        ratio = None if baseline_min == 0 else candidate_max / baseline_min
        passed = candidate_max == 0 if baseline_min == 0 else ratio <= threshold
        checks.append({
            "case": case, "shape": shape, "repeat": repeat,
            "metric": "allocation.region_peak_live_bytes_incremental",
            "baseline_valid_min_incremental_peak": baseline_min,
            "candidate_invalid_max_incremental_peak": candidate_max,
            "candidate_to_baseline_valid_ratio": ratio,
            "max_allowed_ratio": threshold,
            "passed": passed,
        })
    return checks


def _allocation_drift(stage: str, evidence: dict[str, Any],
                      plan: dict[str, Any]) -> list[dict[str, Any]]:
    rows = {(row["repeat"], row["shape"], row["case"]): row
            for row in evidence["rows"]}
    drift: list[dict[str, Any]] = []
    for shape in SHAPES:
        for case in CASES:
            first = rows[(1, shape, case)]
            second = rows[(2, shape, case)]
            for field in ALLOCATION_FIELDS:
                first_stats = first["allocation"]["statistics"].get(field)
                second_stats = second["allocation"]["statistics"].get(field)
                if first_stats is None or second_stats is None:
                    continue
                for stat in ALL_STATS:
                    _drift(drift, stage=stage, lane="alloc", case=case, shape=shape,
                           metric=f"allocation.{field}.{stat}",
                           first={"value": first_stats[stat]},
                           second={"value": second_stats[stat]})
    return drift


def _timing_drift(stage: str, evidence: dict[str, Any]) -> list[dict[str, Any]]:
    rows = {(row["repeat"], row["shape"], row["case"]): row
            for row in evidence["rows"]}
    drift: list[dict[str, Any]] = []
    for shape in SHAPES:
        for case in CASES:
            first = rows[(1, shape, case)]["timing"]["edit_sheets"]
            second = rows[(2, shape, case)]["timing"]["edit_sheets"]
            for stat in ALL_STATS:
                _drift(drift, stage=stage, lane="normal", case=case, shape=shape,
                       metric=f"timing.edit_sheets.{stat}",
                       first={"value": first[stat]}, second={"value": second[stat]})
    return drift


def _compare(baseline: dict[str, Any], candidate: dict[str, Any],
             plan: dict[str, Any]) -> dict[str, Any]:
    corpus_equality = {
        "baseline": _cross_lane_corpus("baseline", baseline["normal"],
                                        baseline["alloc"]),
        "candidate": _cross_lane_corpus("candidate", candidate["normal"],
                                         candidate["alloc"]),
    }
    receipt_custody = {
        "baseline": _cross_lane_receipt_custody("baseline", baseline["normal"],
                                                 baseline["alloc"]),
        "candidate": _cross_lane_receipt_custody("candidate", candidate["normal"],
                                                  candidate["alloc"]),
    }
    normal = _compare_lane(baseline["normal"], candidate["normal"], plan, "normal")
    allocation = _compare_lane(baseline["alloc"], candidate["alloc"], plan, "alloc")
    valid = _valid_gate(baseline["normal"], candidate["normal"], plan)
    invalid_native = _invalid_native_gate(baseline["normal"], candidate["normal"], plan)
    invalid_alloc = _invalid_allocation_gate(baseline["alloc"], candidate["alloc"], plan)
    adverse = (normal["adverse_flags_over_five_percent"]
               + allocation["adverse_flags_over_five_percent"])
    drift = (_timing_drift("baseline", baseline["normal"])
             + _timing_drift("candidate", candidate["normal"])
             + _allocation_drift("baseline", baseline["alloc"], plan)
             + _allocation_drift("candidate", candidate["alloc"], plan))
    abba = {
        "normal": _check_abba(baseline["normal"], candidate["normal"], plan, "normal"),
        "alloc": _check_abba(baseline["alloc"], candidate["alloc"], plan, "alloc"),
    }
    gates = {
        "valid_p50_regression": valid,
        "invalid_native_p50": invalid_native,
        "invalid_allocation_peak": invalid_alloc,
    }
    passed = (all(row["passed"] for row in valid)
              and all(row["passed"] for row in invalid_native)
              and all(row["passed"] for row in invalid_alloc)
              and all(item["passed"] for item in abba.values()))
    return {
        "normal": normal,
        "allocation": allocation,
        "corpus_equality": corpus_equality,
        "receipt_custody": receipt_custody,
        "gates": gates,
        "admission_passed": passed,
        "abba": abba,
        "adverse_flags_over_five_percent": adverse,
        "same_build_drift_over_five_percent": drift,
        "thresholds": {
            "adverse_percent": ADVERSE_THRESHOLD_PERCENT,
            "valid_p50_max_regression_percent": float(
                plan["refusal_guard"]["ordinary_valid_max_regression_percent"]),
            "invalid_native_p50_max_baseline_valid_ratio": float(
                plan["refusal_guard"]["native_invalid_max_baseline_valid_ratio"]),
            "invalid_allocation_peak_max_baseline_valid_ratio": float(
                plan["refusal_guard"]["allocation_invalid_peak_max_baseline_valid_ratio"]),
        },
    }


_plan_revision = ""
_plan_cpu = 2


def analyze() -> dict[str, Any]:
    """Return deterministic guard evidence, or pending until all jobs exist."""

    global _plan_revision, _plan_cpu
    plan = _plan()
    _plan_revision = plan["revision"]
    _plan_cpu = plan["cpu"]
    plan_sha = _sha(PLAN_PATH)
    driver = {"path": _relative(GUARD_RUN_PATH), "sha256": _sha(GUARD_RUN_PATH)}
    run_driver = {"path": _relative(RUN_PATH), "sha256": _sha(RUN_PATH)}

    candidate_manifest_path = HERE / "candidate" / "source-manifest.json"
    candidate_manifest_sha = (
        _sha(candidate_manifest_path) if candidate_manifest_path.is_file() else None
    )
    stages: dict[str, dict[str, Any]] = {}
    for stage in STAGES:
        stages[stage] = {}
        for lane in LANES:
            stages[stage][lane] = _stage_or_pending(
                stage, lane, plan, candidate_manifest_sha)

    missing = [
        f"{stage}/{lane}: {item}"
        for stage in STAGES
        for lane in LANES
        for item in stages[stage][lane].get("missing_artifacts", [])
    ]
    base_result: dict[str, Any] = {
        "schema": "litchi.xlsx.planning-refusal-guard-analysis.v1",
        "stage": "compare",
        "plan_sha256": plan_sha,
        "capture_driver": driver,
        "run_driver": run_driver,
        "analyzer": {"path": _relative(ANALYZER_PATH),
                     "sha256": _sha(ANALYZER_PATH)},
        "refusal_guard": dict(plan["refusal_guard"]),
        "expected": {
            "stages": list(STAGES), "lanes": list(LANES),
            "shapes": list(SHAPES), "cases": list(CASES),
            "repeats": list(REPEATS),
            "normal_job_count_per_stage": len(_expected_jobs(plan, "normal")),
            "allocation_job_count_per_stage": len(_expected_jobs(plan, "alloc")),
        },
        "stages": stages,
        "scope": plan["refusal_guard"]["scope"],
    }
    if missing:
        base_result.update({
            "status": "pending",
            "admission_status": "pending",
            "missing_artifacts": missing,
            "note": (
                "Both normal and allocator ABBA lanes must validate before "
                "comparison; no synthetic evidence was created."
            ),
        })
        return base_result

    baseline = {lane: stages["baseline"][lane]["evidence"] for lane in LANES}
    candidate = {lane: stages["candidate"][lane]["evidence"] for lane in LANES}
    comparison = _compare(baseline, candidate, plan)
    base_result.update({
        "status": "pass",
        "admission_status": "pass" if comparison["admission_passed"] else "reject",
        "baseline": baseline,
        "candidate": candidate,
        "comparison": comparison,
    })
    return base_result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path,
                        default=HERE / "guard-analysis.json")
    args = parser.parse_args()
    try:
        result = analyze()
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n",
                               encoding="utf-8")
    except FileNotFoundError as error:
        print(f"evidence check failed: missing {error.filename}", file=sys.stderr)
        return 1
    except EvidenceError as error:
        print(f"evidence check failed: {error}", file=sys.stderr)
        return 1
    print(f"0544 refusal guard {result['status']}: {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
