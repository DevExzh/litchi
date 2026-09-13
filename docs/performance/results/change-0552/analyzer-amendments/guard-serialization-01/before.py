#!/usr/bin/env python3
"""Validate and compare the supplemental 0552 XLSX guard captures.

This report owns the two small, independent evidence lanes used by the 0552
experiment: ``xlsx_planning_guard`` and ``perf_cap_boundary``.  It authenticates
the frozen plan and inputs, every build and capture receipt, source/execution
manifest binding, fixture/corpus identity, correctness oracles, complete
sample vectors, and the serial ABBA order.  It then emits the native timing
comparisons and repeat drift plus the allocator-derived peaks and retained
bytes.  Allocator-instrumented elapsed time is validated for sample custody but
is deliberately excluded from performance comparisons.

The analyzer is read-only by default.  Without ``--output`` it writes the
derived document to stdout.  An explicit output path uses exclusive-create or
identical-replay semantics, so a replay cannot replace a different report.
Missing later-stage evidence produces a pending document; no synthetic rows or
zero samples are made.  The guard/cap result is supplemental and never claims
that the complete 0552 optimization or quality gate passed.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import math
import re
import statistics
import sys
from pathlib import Path
from typing import Any


sys.dont_write_bytecode = True

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PLAN_PATH = HERE / "plan.json"
RUN_PATH = HERE / "run.py"
CAPTURE_PATH = HERE / "capture.py"
GUARDED_CAPTURE_PATH = HERE / "guarded_capture.py"
FROZEN_INPUTS_PATH = HERE / "frozen-inputs.json"
SUPPLEMENTAL_INPUTS_PATH = HERE / "supplemental-inputs.json"
WORKSPACE_LOCK_PATH = HERE / "workspace-lock.json"
RETAINED_WORKSPACE_LOCK_PATH = HERE / "workspace-Cargo.lock"
ADR_MANIFEST_PATH = HERE / "adr-manifest.json"

STAGES = ("baseline", "candidate")
LANES = ("normal", "alloc")
REPEATS = (1, 2)
GUARD_SHAPES = ("medium", "dense-sparse")
GUARD_CASES = ("valid", "late-validator", "late-raw")
CAP_SIZES = (1, 2, 160, 164, 256)

GUARD_SCHEMA = "litchi.xlsx.planning-refusal-guard.v1"
GUARD_TOOL = "xlsx_planning_guard"
CAP_SCHEMA = "litchi.xlsx.cap-boundary-guard.v1"
CAP_TOOL = "perf_cap_boundary"
GUARD_SCOPE = "operation_global_system_allocator"
HOST_SCOPE = "Accessible compiler processes; no host quiescence guarantee"
CPU = 2
TARGET = Path("/home/zhuhe/litchi-goal-0552-target")
WORKSPACE_LOCK_SHA256 = "9111221ee9d100daf90328a544613cb3f70287611dcc55a37d3b1b7a5d99c91a"
SHARED_EVENT_CAP = 131_072
ORDINARY_EVENT_CAP = 1_000_000
SOURCE_STREAM_BYTE_LIMIT = 8 * 1024 * 1024
SPARSE_COMMENT_BYTES = 1024 * 1024
DRIFT_THRESHOLD_PERCENT = 5.0
VALID_MAX_RATIO = 1.05
INVALID_NATIVE_MAX_RATIO = 2.0
INVALID_PEAK_MAX_RATIO = 1.10
STATS = ("p50", "p95", "p99", "mean", "min", "max", "standard_deviation")
GATE_STATS = ("p50", "mean")
ALLOC_FIELDS = (
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
ALLOC_CALL_FIELDS = {
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
}
GUARD_REPORT_KEYS = {
    "schema", "tool", "case", "shape", "warmup_iterations", "samples",
    "source", "source_sha256", "logical_error", "phase", "correctness",
    "binary", "runner", "allocation_scope", "allocator", "instrumentation",
    "counter_revision", "timing_scope", "performance_claim",
}
GUARD_SOURCE_KEYS = {
    "generator", "shape", "rows", "columns", "archive_bytes", "worksheet_bytes",
    "source_sha256", "worksheet_sha256", "worksheet_member", "compression",
    "fixture_kind",
}
GUARD_PHASE_KEYS = {"name", "sample_order", "duration_ns", "allocation_metrics", "samples"}
GUARD_SAMPLE_KEYS = {"order", "duration_ns", "allocation_metrics"}
UNAVAILABLE_KEYS = {"status", "scope"}
MEASURED_KEYS = {"status", "scope", *ALLOC_FIELDS}
GUARD_BINARY_KEYS = {"path", "sha256", "bytes"}
GUARD_RUNNER_KEYS = {"git_revision", "git_dirty", "rustc_vv", "os", "arch", "profile"}
CAP_REPORT_KEYS = {
    "archive_bytes", "binary", "case", "cells", "columns", "comment_bytes",
    "correctness", "event_cap_relation", "event_count", "event_count_formula",
    "expected_event_count", "fixture_out", "last_cell", "ordinary_parser_event_cap",
    "performance_claim", "phase", "rows", "samples", "schema",
    "shared_provisional_event_cap", "size", "source", "source_stream_byte_limit",
    "source_stream_eligible", "source_xml_bytes", "tool", "warmup_iterations",
}
CAP_SOURCE_KEYS = {
    "archive_bytes", "columns", "compression", "encoding", "fixture_dump",
    "fixture_kind", "format", "generator", "identity_method", "marker_free",
    "rows", "shape", "source_xml_bytes", "worksheet_bytes", "worksheet_member",
    "worksheet_xml_bytes",
}
CAP_BINARY_KEYS = {"bytes", "identity_method", "path", "profile"}
CAP_PHASE_KEYS = {"duration_ns", "name", "sample_order", "samples", "timing_scope",
                   "warmup_iterations"}
CAP_CORRECTNESS_KEYS = {
    "commit_outside_timing", "empty_commit_is_noop", "no_op_publication_exact",
    "snapshot_a1", "snapshot_last_cell", "source_bytes_unchanged",
    "source_unchanged", "source_xml_unchanged", "valid_snapshot_values",
}
RECEIPT_KEYS = {
    "command", "start_utc", "end_utc", "seconds", "exit_code",
    "execution_stage", "execution_manifest_sha256", "binary_sha256",
    "source_manifest_sha256", "script_sha256", "plan_sha256", "environment",
    "artifacts",
}
ENVIRONMENT_KEYS = {
    "RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "LD_PRELOAD", "MALLOC_CONF",
    "GLIBC_TUNABLES",
}
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
REVISION_RE = re.compile(r"^[0-9a-f]{40}$")

EXPECTED_ERRORS = {
    "late-validator": "value-only edits refuse attribute 'future' on 'c'",
    "late-raw": "invalid worksheet boolean 'maybe'",
}
GUARD_DIMENSIONS = {"medium": (96, 96), "dense-sparse": (128, 128)}


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory evidence artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def read_json(path: Path) -> Any:
    try:
        with path.open(encoding="utf-8") as stream:
            return json.load(stream)
    except FileNotFoundError:
        raise
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


def relative(path: Path) -> str:
    try:
        return path.relative_to(HERE).as_posix()
    except ValueError as error:
        raise EvidenceError(f"path is outside 0552 evidence: {path}") from error


def regular(path: Path, label: str) -> None:
    require(path.is_file() and not path.is_symlink(), f"{label} is not a regular file")


def hash_value(value: Any, label: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
            f"{label} is not a lowercase SHA-256 digest")
    return value


def revision(value: Any, label: str) -> str:
    require(isinstance(value, str) and REVISION_RE.fullmatch(value) is not None,
            f"{label} is not a lowercase commit digest")
    return value


def nonempty_string(value: Any, label: str) -> str:
    require(isinstance(value, str) and value, f"{label} is not a nonempty string")
    return value


def nonnegative_integer(value: Any, label: str) -> int:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a nonnegative integer")
    return value


def positive_integer(value: Any, label: str) -> int:
    result = nonnegative_integer(value, label)
    require(result > 0, f"{label} is not positive")
    return result


def finite_number(value: Any, label: str) -> float:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")
    return float(value)


def timestamp(value: Any, label: str) -> _datetime.datetime:
    require(isinstance(value, str) and value, f"{label} is missing")
    try:
        parsed = _datetime.datetime.fromisoformat(value)
    except ValueError as error:
        raise EvidenceError(f"{label} is not an ISO-8601 timestamp") from error
    require(parsed.tzinfo is not None, f"{label} has no timezone")
    return parsed


def path_name(path: Path, label: str) -> None:
    require(path.name == path.as_posix() and not path.is_absolute(),
            f"{label} is not a local artifact name")


def stats(values: list[int], unit: str) -> dict[str, Any]:
    require(isinstance(values, list) and values, "statistics vector is empty")
    for index, value in enumerate(values):
        nonnegative_integer(value, f"statistics[{index}]")
    ordered = sorted(values)
    count = len(ordered)
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
        "mean": sum(ordered) / count,
        "standard_deviation": 0.0 if count == 1 else statistics.stdev(ordered),
    }


def comparison(before: int | float, after: int | float, label: str) -> dict[str, Any]:
    before_f = finite_number(before, f"{label}.baseline")
    after_f = finite_number(after, f"{label}.candidate")
    require(before_f >= 0.0 and after_f >= 0.0, f"{label} contains a negative value")
    change = 0.0 if before_f == 0.0 and after_f == 0.0 else (
        None if before_f == 0.0 else (after_f / before_f - 1.0) * 100.0
    )
    return {
        "baseline": before,
        "candidate": after,
        "delta": after - before,
        "candidate_to_baseline_ratio": None if before_f == 0.0 else after_f / before_f,
        "change_percent": change,
    }


def over_five(record: dict[str, Any]) -> bool:
    change = record.get("change_percent")
    return (change is None and record["baseline"] == 0 and record["candidate"] > 0) or (
        change is not None and abs(float(change)) > DRIFT_THRESHOLD_PERCENT
    )


def _plan() -> dict[str, Any]:
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict), "plan is not an object")
    expected = {
        "revision", "created_utc", "previous_turn", "priority", "scope",
        "hypothesis", "cpu", "owned_paths", "cases", "shapes", "native",
        "alloc", "profile", "guard", "cap", "native_order", "allocator_order",
        "admission", "limitations",
    }
    require(set(plan) == expected, "plan field inventory differs")
    revision(plan["revision"], "plan.revision")
    require(plan["priority"] == "OLE2/OOXML first; ODF deferred until that goal completes; iWork excluded",
            "plan priority differs")
    require(plan["scope"] == "Matched compact source-cell proof experiment for source-backed XLSX MultiSourceEdit",
            "plan scope differs")
    require(plan["cpu"] == CPU, "plan CPU differs")
    require(plan["owned_paths"] == [str(TARGET)], "plan owned target differs")
    require(plan["cases"] == [
        "xlsx_source_backed_cell_values_one_edit_save",
        "xlsx_source_backed_cell_values_one_percent_edit_save",
        "xlsx_source_backed_managed_cell_values_one_edit_save",
        "xlsx_source_backed_managed_cell_values_one_percent_edit_save",
    ], "plan main case matrix differs")
    require(plan["shapes"] == ["medium", "dense-sparse", "noncompact", "vendor-extension"],
            "plan main shape matrix differs")
    for lane in ("native", "alloc"):
        config = plan[lane]
        require(isinstance(config, dict) and set(config) == {"repeats", "samples", "warmup"},
                f"plan.{lane} inventory differs")
        require(config == ({"repeats": 2, "samples": 200, "warmup": 20}
                          if lane == "native" else
                          {"repeats": 2, "samples": 20, "warmup": 3}),
                f"plan.{lane} counts differ")
    profile = plan["profile"]
    require(isinstance(profile, dict)
            and profile == {
                "repeats": 2,
                "samples": 1,
                "warmup": 0,
                "case": "xlsx_source_backed_cell_values_one_percent_edit_save",
                "owner": "litchi_xlsx::cell_values::source::MultiSourceEdit::commit",
                "parent": "litchi_perf_baseline::run_xlsx_cell_values_edit_save",
            }, "plan profile differs")
    require(isinstance(plan["guard"], dict) and plan["guard"] == {
        "shapes": list(GUARD_SHAPES), "cases": list(GUARD_CASES), "repeats": 2,
        "native_samples": 200, "native_warmup": 20,
        "alloc_samples": 20, "alloc_warmup": 3,
    }, "plan guard matrix differs")
    require(isinstance(plan["cap"], dict) and plan["cap"] == {
        "sizes": list(CAP_SIZES), "repeats": 2, "samples": 200, "warmup": 20,
    }, "plan cap matrix differs")
    require(plan["native_order"] == [
        "baseline r1", "candidate r1", "candidate r2",
        "retained baseline r2 under candidate source manifest",
    ], "plan native ABBA order differs")
    require(plan["allocator_order"] == plan["native_order"],
            "plan allocator ABBA order differs")
    admission = plan["admission"]
    require(isinstance(admission, dict), "plan admission is not an object")
    require(admission["latency_guards"] ==
            "Every one-cell case/shape/repeat and valid-noop planning/cap control p50 and mean <= 1.05x matched baseline.",
            "plan latency guard text differs")
    require(admission["refusals"] ==
            "Exact errors/retry identities unchanged; each invalid guard p50 and mean <= max(1.05x corresponding baseline invalid,2x baseline valid); each invalid attributable peak <= 1.10x baseline valid attributable peak.",
            "plan refusal gate text differs")
    require(admission["noop_memory"] ==
            "Valid planning guard region peak minus live_bytes_before <= 1.05x baseline. Decline if durable metadata cannot fit representative no-op memory envelope.",
            "plan no-op memory text differs")
    return plan


def _frozen_inputs() -> dict[str, Any]:
    frozen = read_json(FROZEN_INPUTS_PATH)
    require(isinstance(frozen, dict) and set(frozen) == {"frozen_utc", "files"},
            "frozen-inputs inventory differs")
    files = frozen["files"]
    require(isinstance(files, dict) and set(files) == {
        "docs/performance/results/change-0552/plan.json",
        "docs/performance/results/change-0552/run.py",
        "docs/performance/results/change-0552/capture.py",
        "docs/performance/results/change-0552/adr-manifest.json",
    }, "frozen-inputs file inventory differs")
    timestamp(frozen["frozen_utc"], "frozen-inputs.frozen_utc")
    for name, expected in files.items():
        hash_value(expected, f"frozen-inputs.files.{name}")
        path = REPO / name
        regular(path, name)
        require(sha256(path) == expected, f"frozen input changed: {name}")
    adr = read_json(ADR_MANIFEST_PATH)
    require(isinstance(adr, dict) and set(adr) == {"files", "checked_utc", "status"},
            "ADR manifest inventory differs")
    timestamp(adr["checked_utc"], "adr-manifest.checked_utc")
    require(adr["status"] == "all previously read ADRs unchanged",
            "ADR manifest status differs")
    require(isinstance(adr["files"], dict) and adr["files"],
            "ADR manifest files are missing")
    for name, expected in adr["files"].items():
        require(isinstance(name, str) and not Path(name).is_absolute(),
                "ADR manifest path is invalid")
        hash_value(expected, f"adr-manifest.files.{name}")
        regular(REPO / name, f"ADR {name}")
        require(sha256(REPO / name) == expected, f"ADR changed: {name}")
    return {"files": dict(sorted(files.items())), "adr_manifest_sha256": sha256(ADR_MANIFEST_PATH),
            "adr_entries": len(adr["files"]), "frozen_utc": frozen["frozen_utc"]}


def _supplemental_inputs() -> dict[str, Any]:
    supplemental = read_json(SUPPLEMENTAL_INPUTS_PATH)
    require(isinstance(supplemental, dict)
            and set(supplemental) == {"frozen_utc", "scope", "files"},
            "supplemental-inputs inventory differs")
    timestamp(supplemental["frozen_utc"], "supplemental-inputs.frozen_utc")
    require(supplemental["scope"] ==
            "Workspace-lock custody before first workspace build; per-child lock checks on later pipelines; quality commands frozen before use. Original capture inputs unchanged.",
            "supplemental-inputs scope differs")
    files = supplemental["files"]
    expected_names = {
        "docs/performance/results/change-0552/guarded_capture.py",
        "docs/performance/results/change-0552/workspace-lock.json",
        "docs/performance/results/change-0552/workspace-Cargo.lock",
        "docs/performance/results/change-0552/quality-plan.json",
    }
    require(isinstance(files, dict) and set(files) == expected_names,
            "supplemental-inputs file inventory differs")
    for name, expected in files.items():
        hash_value(expected, f"supplemental-inputs.files.{name}")
        path = REPO / name
        regular(path, name)
        require(sha256(path) == expected, f"supplemental input changed: {name}")
    lock = read_json(WORKSPACE_LOCK_PATH)
    require(isinstance(lock, dict) and set(lock) == {
        "frozen_utc", "path", "sha256", "retained", "reason", "baseline_cap_not_started",
    }, "workspace-lock inventory differs")
    timestamp(lock["frozen_utc"], "workspace-lock.frozen_utc")
    require(lock["path"] == "Cargo.lock" and lock["retained"] == "workspace-Cargo.lock",
            "workspace-lock paths differ")
    hash_value(lock["sha256"], "workspace-lock.sha256")
    require(lock["sha256"] == WORKSPACE_LOCK_SHA256,
            "workspace-lock digest differs from frozen supplemental binding")
    require(lock["baseline_cap_not_started"] is True,
            "workspace-lock baseline custody marker differs")
    regular(REPO / lock["path"], "workspace Cargo.lock")
    regular(HERE / lock["retained"], "retained workspace Cargo.lock")
    require(sha256(REPO / lock["path"]) == lock["sha256"],
            "workspace Cargo.lock changed")
    require(sha256(HERE / lock["retained"]) == lock["sha256"],
            "retained workspace Cargo.lock changed")
    return {
        "frozen_utc": supplemental["frozen_utc"],
        "scope": supplemental["scope"],
        "files": dict(sorted(files.items())),
        "workspace_lock_sha256": lock["sha256"],
    }


def _manifest(stage: str) -> tuple[dict[str, str], str]:
    path = HERE / stage / "source-manifest.json"
    regular(path, f"{stage}/source-manifest.json")
    value = read_json(path)
    require(isinstance(value, dict) and value, f"{stage} source manifest is empty")
    for name, digest in value.items():
        require(isinstance(name, str) and name and not Path(name).is_absolute(),
                f"{stage} source manifest path is invalid")
        parts = Path(name).parts
        require(".." not in parts and Path(name).as_posix() == name,
                f"{stage} source manifest path escapes its root")
        hash_value(digest, f"{stage} source manifest {name}")
    require("tools/perf-baseline/src/xlsx_planning_guard.rs" in value,
            f"{stage} source manifest omits planning guard source")
    require("crates/litchi-xlsx/examples/perf_cap_boundary.rs" in value,
            f"{stage} source manifest omits cap example")
    return dict(sorted(value.items())), sha256(path)


def _cleanup_binary_hash(path: Path, digest: str) -> None:
    if path.exists():
        regular(path, f"retained binary {path}")
        require(sha256(path) == digest, f"retained binary digest differs: {path}")
        return
    cleanup_path = HERE / "cleanup.json"
    regular(cleanup_path, "cleanup.json")
    cleanup = read_json(cleanup_path)
    require(isinstance(cleanup, dict)
            and cleanup.get("owned_paths_absent") is True
            and cleanup.get("accessible_process_references") == []
            and cleanup.get("removed") == [str(TARGET)]
            and not TARGET.exists(),
            f"missing retained binary has no completed cleanup custody: {path}")
    by_kind = cleanup.get("binary_sha256_by_kind")
    require(isinstance(by_kind, dict), "cleanup binary hash map is missing")
    basename = path.name
    accepted = [str(path), basename, f"{path.parent.name}/{basename}"]
    require(any(by_kind.get(key) == digest for key in accepted),
            f"cleanup does not retain digest for {path}")


def _binary_descriptor(stage: str, kind: str, manifest_sha: str) -> dict[str, Any]:
    names = {
        "guard-normal": "binary-guard-normal.json",
        "guard-alloc": "binary-guard-alloc.json",
        "cap": "binary-cap.json",
    }
    files = {
        "guard-normal": "guard-normal",
        "guard-alloc": "guard-alloc",
        "cap": "cap",
    }
    require(kind in names, f"unknown binary kind {kind}")
    path = HERE / stage / names[kind]
    regular(path, f"{stage}/{path.name}")
    value = read_json(path)
    expected_keys = {"path", "sha256", "bytes", "build_receipt_sha256", "source_manifest_sha256"}
    require(isinstance(value, dict) and set(value) == expected_keys,
            f"{stage}/{path.name} inventory differs")
    actual_path = Path(nonempty_string(value["path"], f"{stage}/{path.name}.path"))
    expected_path = TARGET / "retained" / stage / files[kind]
    require(actual_path == expected_path, f"{stage}/{path.name} path differs")
    digest = hash_value(value["sha256"], f"{stage}/{path.name}.sha256")
    byte_count = positive_integer(value["bytes"], f"{stage}/{path.name}.bytes")
    require(value["source_manifest_sha256"] == manifest_sha,
            f"{stage}/{path.name} source manifest differs")
    build_receipt_name = {
        "guard-normal": "build-guard-normal.receipt.json",
        "guard-alloc": "build-guard-alloc.receipt.json",
        "cap": "build-cap.receipt.json",
    }[kind]
    build_path = HERE / stage / build_receipt_name
    regular(build_path, f"{stage}/{build_receipt_name}")
    require(value["build_receipt_sha256"] == sha256(build_path),
            f"{stage}/{path.name} build receipt digest differs")
    _cleanup_binary_hash(actual_path, digest)
    if actual_path.exists():
        require(actual_path.stat().st_size == byte_count,
                f"{stage}/{path.name} byte count differs")
    return {
        "path": str(actual_path), "sha256": digest, "bytes": byte_count,
        "build_receipt_sha256": value["build_receipt_sha256"],
        "source_manifest_sha256": manifest_sha,
    }


def _expected_build_command(kind: str) -> list[str]:
    prefix = [
        "env", f"TMPDIR={TARGET / 'tmp'}", "CARGO_BUILD_JOBS=2", "CARGO_INCREMENTAL=0",
        "cargo", "build", "--release", "--locked",
    ]
    if kind.startswith("guard-"):
        command = prefix + ["--manifest-path", "tools/perf-baseline/Cargo.toml",
                            "--bin", "xlsx_planning_guard", "--target-dir", str(TARGET)]
        if kind == "guard-alloc":
            command += ["--features", "allocator-metrics"]
        return command
    require(kind == "cap", f"unknown build kind {kind}")
    return prefix + ["-p", "litchi-xlsx", "--example", "perf_cap_boundary",
                     "--target-dir", str(TARGET)]


def _validate_host(path: Path, label: str) -> None:
    regular(path, label)
    value = read_json(path)
    require(isinstance(value, dict) and set(value) == {
        "observed_utc", "compiler_processes", "scope",
    }, f"{label} inventory differs")
    timestamp(value["observed_utc"], f"{label}.observed_utc")
    require(value["scope"] == HOST_SCOPE, f"{label}.scope differs")
    require(isinstance(value["compiler_processes"], list),
            f"{label}.compiler_processes is not a list")
    for index, process in enumerate(value["compiler_processes"]):
        require(isinstance(process, dict) and set(process) == {"pid", "comm", "cwd"},
                f"{label}.compiler_processes[{index}] differs")
        positive_integer(process["pid"], f"{label}.compiler_processes[{index}].pid")
        require(process["comm"] in ("cargo", "rustc"),
                f"{label}.compiler_processes[{index}].comm differs")
        nonempty_string(process["cwd"], f"{label}.compiler_processes[{index}].cwd")


def _receipt_common(value: Any, path: Path, label: str) -> tuple[_datetime.datetime, _datetime.datetime]:
    require(isinstance(value, dict) and set(value) == RECEIPT_KEYS,
            f"{label} receipt inventory differs")
    start = timestamp(value["start_utc"], f"{label}.start_utc")
    end = timestamp(value["end_utc"], f"{label}.end_utc")
    require(start < end, f"{label} receipt timestamps are not increasing")
    seconds = finite_number(value["seconds"], f"{label}.seconds")
    require(seconds > 0.0, f"{label}.seconds is not positive")
    wall = (end - start).total_seconds()
    require(abs(seconds - wall) <= max(0.25, wall * 0.02 + 0.05),
            f"{label}.seconds does not match UTC interval")
    environment = value["environment"]
    require(isinstance(environment, dict) and set(environment) == ENVIRONMENT_KEYS,
            f"{label}.environment inventory differs")
    require(all(item is None for item in environment.values()),
            f"{label}.environment contains uncontrolled values")
    require(isinstance(value["command"], list)
            and all(isinstance(item, str) and item for item in value["command"]),
            f"{label}.command is malformed")
    require(isinstance(value["exit_code"], int) and not isinstance(value["exit_code"], bool),
            f"{label}.exit_code is malformed")
    for field in ("execution_manifest_sha256", "source_manifest_sha256",
                  "script_sha256", "plan_sha256"):
        hash_value(value[field], f"{label}.{field}")
    artifacts = value["artifacts"]
    require(isinstance(artifacts, dict), f"{label}.artifacts is not an object")
    for name, digest in artifacts.items():
        require(isinstance(name, str) and Path(name).name == name,
                f"{label}.artifacts contains a non-local name")
        hash_value(digest, f"{label}.artifacts.{name}")
    return start, end


def _build_receipt(stage: str, kind: str, manifest_sha: str) -> dict[str, Any]:
    receipt_name = {
        "guard-normal": "build-guard-normal.receipt.json",
        "guard-alloc": "build-guard-alloc.receipt.json",
        "cap": "build-cap.receipt.json",
    }[kind]
    path = HERE / stage / receipt_name
    value = read_json(path)
    start, end = _receipt_common(value, path, f"{stage}/{receipt_name}")
    label = f"{stage}/{receipt_name}"
    require(value["exit_code"] == 0 and value["binary_sha256"] is None,
            f"{label} is not a successful build receipt")
    require(value["execution_stage"] == stage,
            f"{label}.execution_stage differs")
    require(value["execution_manifest_sha256"] == manifest_sha,
            f"{label}.execution_manifest_sha256 differs")
    require(value["source_manifest_sha256"] == manifest_sha,
            f"{label}.source_manifest_sha256 differs")
    require(value["script_sha256"] == sha256(RUN_PATH), f"{label}.script_sha256 differs")
    require(value["plan_sha256"] == sha256(PLAN_PATH), f"{label}.plan_sha256 differs")
    require(value["command"] == _expected_build_command(kind),
            f"{label}.command differs from frozen capture.py")
    expected_artifacts = {receipt_name[:-len(".receipt.json")] + suffix
                          for suffix in (".host.json", ".stdout", ".stderr")}
    require(set(value["artifacts"]) == expected_artifacts,
            f"{label}.artifacts inventory differs")
    for name, digest in value["artifacts"].items():
        artifact = HERE / stage / name
        regular(artifact, f"{label}/{name}")
        require(sha256(artifact) == digest, f"{label}/{name} digest differs")
        if name.endswith(".host.json"):
            _validate_host(artifact, f"{label}/{name}")
    return {"path": relative(path), "sha256": sha256(path),
            "start_utc": value["start_utc"], "end_utc": value["end_utc"],
            "seconds": value["seconds"], "command": value["command"],
            "interval": (start, end)}


def _guard_jobs(lane: str) -> list[dict[str, Any]]:
    require(lane in LANES, f"unknown guard lane {lane}")
    samples, warmup = ((200, 20) if lane == "normal" else (20, 3))
    name_lane = "native" if lane == "normal" else "alloc"
    jobs: list[dict[str, Any]] = []
    for repeat in REPEATS:
        shapes = GUARD_SHAPES if repeat == 1 else tuple(reversed(GUARD_SHAPES))
        for shape in shapes:
            for case in GUARD_CASES:
                jobs.append({
                    "name": f"guard-{name_lane}-r{repeat}-{shape}-{case}",
                    "stage": None, "lane": lane, "repeat": repeat, "shape": shape,
                    "case": case, "samples": samples, "warmup": warmup,
                })
    return jobs


def _cap_jobs() -> list[dict[str, Any]]:
    jobs: list[dict[str, Any]] = []
    for repeat in REPEATS:
        sizes = CAP_SIZES if repeat == 1 else tuple(reversed(CAP_SIZES))
        for size in sizes:
            jobs.append({"name": f"cap-r{repeat}-{size}", "stage": None,
                         "repeat": repeat, "size": size, "samples": 200, "warmup": 20})
    return jobs


def _execution(stage: str, repeat: int, candidate_manifest_sha: str) -> tuple[str, str]:
    if stage == "candidate" or repeat == 2:
        return "candidate", candidate_manifest_sha
    return "baseline", sha256(HERE / "baseline" / "source-manifest.json")


def _allocation_sample(value: Any, measured: bool, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is not an allocation object")
    require(value.get("scope") == GUARD_SCOPE, f"{label}.scope differs")
    if not measured:
        require(set(value) == UNAVAILABLE_KEYS and value["status"] == "unavailable",
                f"{label} unavailable sample differs")
        return {"status": "unavailable", "scope": GUARD_SCOPE}
    require(set(value) == MEASURED_KEYS and value["status"] == "measured",
            f"{label} measured sample inventory differs")
    result: dict[str, Any] = {"status": "measured", "scope": GUARD_SCOPE}
    for field in ALLOC_FIELDS:
        result[field] = nonnegative_integer(value[field], f"{label}.{field}")
    require(result["failed_allocation_calls"] == 0,
            f"{label} recorded a failed allocation")
    require(result["live_bytes_before"] + result["allocated_bytes"] ==
            result["live_bytes_after"] + result["deallocated_bytes"],
            f"{label} live-byte balance does not reconcile")
    require(result["peak_live_bytes_before"] >= result["live_bytes_before"],
            f"{label} pre-operation peak is below live bytes")
    require(result["peak_live_bytes_after"] >= result["peak_live_bytes_before"]
            and result["peak_live_bytes_after"] >= result["live_bytes_after"],
            f"{label} post-operation peak is invalid")
    require(result["region_peak_live_bytes"] >= max(result["live_bytes_before"],
                                                     result["live_bytes_after"])
            and result["region_peak_live_bytes"] <= result["peak_live_bytes_after"],
            f"{label} region peak is outside process peak envelope")
    return result


def _guard_source(value: Any, shape: str, label: str) -> dict[str, Any]:
    require(isinstance(value, dict) and set(value) == GUARD_SOURCE_KEYS,
            f"{label}.source inventory differs")
    rows, columns = GUARD_DIMENSIONS[shape]
    require(value["generator"] == "litchi-xlsx-planning-refusal-guard-fixed-grid-v1",
            f"{label}.source.generator differs")
    require(value["shape"] == shape and value["rows"] == rows and value["columns"] == columns,
            f"{label}.source dimensions differ")
    for field in ("rows", "columns", "archive_bytes", "worksheet_bytes"):
        positive_integer(value[field], f"{label}.source.{field}")
    hash_value(value["source_sha256"], f"{label}.source.source_sha256")
    hash_value(value["worksheet_sha256"], f"{label}.source.worksheet_sha256")
    require(value["worksheet_member"] == "xl/worksheets/sheet1.xml",
            f"{label}.source.worksheet_member differs")
    require(value["compression"] == "stored", f"{label}.source.compression differs")
    require(value["fixture_kind"] == "raw_zip_input_no_opc_authored_xml_validation",
            f"{label}.source.fixture_kind differs")
    return value


def _guard_correctness(value: Any, case: str, label: str) -> dict[str, Any]:
    expected = {
        "source_unchanged": True,
        "retry_preserved_error": case in EXPECTED_ERRORS,
        "valid_snapshot_values": case not in EXPECTED_ERRORS,
        "empty_commit_is_noop": case not in EXPECTED_ERRORS,
        "expected_error_exact": case in EXPECTED_ERRORS,
        "commit_outside_timing": case not in EXPECTED_ERRORS,
    }
    require(value == expected, f"{label}.correctness oracle differs")
    return dict(expected)


def _guard_phase(value: Any, lane: str, count: int, label: str) -> tuple[dict[str, Any], dict[str, Any]]:
    require(isinstance(value, dict) and set(value) == GUARD_PHASE_KEYS,
            f"{label}.phase inventory differs")
    require(value["name"] == "edit_sheets", f"{label}.phase.name differs")
    require(value["sample_order"] == list(range(count)), f"{label}.phase.sample_order differs")
    durations = value["duration_ns"]
    require(isinstance(durations, list) and len(durations) == count,
            f"{label}.phase.duration_ns cardinality differs")
    for index, duration in enumerate(durations):
        nonnegative_integer(duration, f"{label}.phase.duration_ns[{index}]")
    measured = lane == "alloc"
    allocation_vectors = value["allocation_metrics"]
    require(isinstance(allocation_vectors, list) and len(allocation_vectors) == count,
            f"{label}.phase.allocation_metrics cardinality differs")
    samples = value["samples"]
    require(isinstance(samples, list) and len(samples) == count,
            f"{label}.phase.samples cardinality differs")
    allocations: list[dict[str, Any]] = []
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict) and set(sample) == GUARD_SAMPLE_KEYS,
                f"{label}.phase.samples[{index}] inventory differs")
        require(sample["order"] == index and sample["duration_ns"] == durations[index],
                f"{label}.phase.samples[{index}] disagrees with phase vectors")
        item = _allocation_sample(sample["allocation_metrics"], measured,
                                  f"{label}.phase.samples[{index}].allocation_metrics")
        vector_item = _allocation_sample(allocation_vectors[index], measured,
                                         f"{label}.phase.allocation_metrics[{index}]")
        require(item == vector_item,
                f"{label}.phase.allocation_metrics[{index}] disagrees with sample")
        allocations.append(item)
    timing = stats(durations, "ns")
    if not measured:
        return timing, {"status": "unavailable", "scope": GUARD_SCOPE}
    vectors = {field: [item[field] for item in allocations] for field in ALLOC_FIELDS}
    allocation_stats = {
        field: stats(vectors[field], "count" if field in ALLOC_CALL_FIELDS else "bytes")
        for field in ALLOC_FIELDS
    }
    incremental = [
        peak - before
        for peak, before in zip(vectors["region_peak_live_bytes"],
                                vectors["live_bytes_before"])
    ]
    require(all(value >= 0 for value in incremental), f"{label} incremental peak is negative")
    allocation_stats["incremental_region_peak_live_bytes"] = stats(incremental, "bytes")
    allocation_stats["retained_after_live_bytes"] = stats(vectors["live_bytes_after"], "bytes")
    return timing, {
        "status": "measured", "scope": GUARD_SCOPE,
        "samples": allocations, "vectors": vectors,
        "statistics": allocation_stats,
        "attributable": {
            "allocation_calls": allocation_stats["allocation_calls"],
            "allocated_bytes": allocation_stats["allocated_bytes"],
            "reallocation_calls": allocation_stats["reallocation_calls"],
            "region_peak_live_bytes": allocation_stats["region_peak_live_bytes"],
            "incremental_region_peak_live_bytes": allocation_stats[
                "incremental_region_peak_live_bytes"],
            "retained_after_live_bytes": allocation_stats["retained_after_live_bytes"],
        },
    }


def _guard_report(report: Any, job: dict[str, Any], binary: dict[str, Any],
                  label: str) -> dict[str, Any]:
    require(isinstance(report, dict) and set(report) == GUARD_REPORT_KEYS,
            f"{label} report inventory differs")
    require(report["schema"] == GUARD_SCHEMA and report["tool"] == GUARD_TOOL,
            f"{label} schema/tool differs")
    require(report["case"] == job["case"] and report["shape"] == job["shape"],
            f"{label} case/shape differs")
    require(report["warmup_iterations"] == job["warmup"]
            and report["samples"] == job["samples"], f"{label} counts differ")
    source = _guard_source(report["source"], job["shape"], label)
    hash_value(report["source_sha256"], f"{label}.source_sha256")
    require(report["source_sha256"] == source["source_sha256"],
            f"{label}.source_sha256 disagrees with source")
    expected_error = EXPECTED_ERRORS.get(job["case"])
    expected_logical = ({"status": "accepted", "variant": None, "message": None}
                        if expected_error is None else
                        {"status": "expected_failure", "variant": "Invalid",
                         "message": expected_error})
    require(report["logical_error"] == expected_logical,
            f"{label}.logical_error differs")
    correctness = _guard_correctness(report["correctness"], job["case"], label)
    timing, allocation = _guard_phase(report["phase"], job["lane"], job["samples"], label)
    require(report["allocation_scope"] == GUARD_SCOPE,
            f"{label}.allocation_scope differs")
    if job["lane"] == "normal":
        expected_allocator, expected_instrumentation, expected_counter = (
            "Rust system allocator", "none", None)
    else:
        expected_allocator, expected_instrumentation, expected_counter = (
            "CountingSystemAllocator(std::alloc::System)",
            "system_allocator_operation_scoped", "serialized_region_peak_v3")
    require(report["allocator"] == expected_allocator,
            f"{label}.allocator differs")
    require(report["instrumentation"] == expected_instrumentation,
            f"{label}.instrumentation differs")
    require(report["counter_revision"] == expected_counter,
            f"{label}.counter_revision differs")
    require(report["timing_scope"] ==
            "source_open_and_fixture_and_selector_setup_excluded; edit_sheets_result_retained_until_clock_and_allocation_end; inspection_commit_and_drop_excluded",
            f"{label}.timing_scope differs")
    require(report["performance_claim"] ==
            "none: correctness and planning diagnostic only; no speedup claim",
            f"{label}.performance_claim differs")
    binary_report = report["binary"]
    require(isinstance(binary_report, dict) and set(binary_report) == GUARD_BINARY_KEYS,
            f"{label}.binary inventory differs")
    require(binary_report == {
        "path": binary["path"], "sha256": binary["sha256"], "bytes": binary["bytes"]
    }, f"{label}.binary identity differs")
    runner = report["runner"]
    require(isinstance(runner, dict) and set(runner) == GUARD_RUNNER_KEYS,
            f"{label}.runner inventory differs")
    require(runner["git_revision"] == _plan_revision,
            f"{label}.runner.git_revision differs")
    require(isinstance(runner["git_dirty"], bool), f"{label}.runner.git_dirty differs")
    nonempty_string(runner["rustc_vv"], f"{label}.runner.rustc_vv")
    require(runner["os"] == "linux" and runner["arch"] == "x86_64"
            and runner["profile"] == "release", f"{label}.runner platform differs")
    return {
        "name": job["name"], "stage": job["stage"], "lane": job["lane"],
        "repeat": job["repeat"], "shape": job["shape"], "case": job["case"],
        "warmup": job["warmup"], "samples": job["samples"],
        "source": source, "source_sha256": report["source_sha256"],
        "logical_error": expected_logical, "correctness": correctness,
        "timing": {"edit_sheets": timing}, "allocation": allocation,
        "binary": dict(binary), "report_sha256": None, "receipt": None,
    }


def _validate_capture_receipt(stage: str, job: dict[str, Any], binary: dict[str, Any],
                              manifest_sha: str, candidate_manifest_sha: str,
                              report_path: Path, extra_artifacts: set[str]) -> tuple[dict[str, Any], _datetime.datetime, _datetime.datetime]:
    folder = HERE / stage
    receipt_path = folder / f"{job['name']}.receipt.json"
    receipt = read_json(receipt_path)
    start, end = _receipt_common(receipt, receipt_path, f"{stage}/{job['name']}")
    label = f"{stage}/{job['name']}"
    require(receipt["exit_code"] == 0, f"{label} did not complete successfully")
    require(receipt["binary_sha256"] == binary["sha256"], f"{label}.binary_sha256 differs")
    expected_execution_stage, expected_execution_manifest = _execution(
        stage, job["repeat"], candidate_manifest_sha)
    require(receipt["execution_stage"] == expected_execution_stage,
            f"{label}.execution_stage differs")
    require(receipt["execution_manifest_sha256"] == expected_execution_manifest,
            f"{label}.execution_manifest_sha256 differs")
    require(receipt["source_manifest_sha256"] == manifest_sha,
            f"{label}.source_manifest_sha256 differs")
    require(receipt["script_sha256"] == sha256(RUN_PATH), f"{label}.script_sha256 differs")
    require(receipt["plan_sha256"] == sha256(PLAN_PATH), f"{label}.plan_sha256 differs")
    if "size" in job:
        expected_command = [
            "taskset", "-c", str(CPU), binary["path"], "--size", str(job["size"]),
            "--samples", str(job["samples"]), "--warmup", str(job["warmup"]),
            "--json", str(report_path), "--fixture-out",
            str(folder / f"{job['name']}.zip"),
        ]
    else:
        expected_command = [
            "taskset", "-c", str(CPU), binary["path"], "--shape", job["shape"],
            "--case", job["case"], "--samples", str(job["samples"]),
            "--warmup", str(job["warmup"]), "--json", str(report_path),
        ]
    require(receipt["command"] == expected_command,
            f"{label}.command differs from frozen capture.py")
    expected_artifacts = {report_path.name, f"{job['name']}.stdout", f"{job['name']}.stderr",
                          f"{job['name']}.host.json"} | extra_artifacts
    require(set(receipt["artifacts"]) == expected_artifacts,
            f"{label}.artifacts inventory differs")
    for name, digest in receipt["artifacts"].items():
        artifact = folder / name
        regular(artifact, f"{label}/{name}")
        require(sha256(artifact) == digest, f"{label}/{name} digest differs")
        if name.endswith(".host.json"):
            _validate_host(artifact, f"{label}/{name}")
    stdout = folder / f"{job['name']}.stdout"
    require(stdout.read_bytes().rstrip(b"\n") == report_path.read_bytes(),
            f"{label}.stdout does not reproduce report")
    return receipt, start, end


def _guard_capture(stage: str, lane: str, job: dict[str, Any], binary: dict[str, Any],
                   manifest_sha: str, candidate_manifest_sha: str) -> tuple[dict[str, Any], tuple[_datetime.datetime, _datetime.datetime]]:
    job = dict(job)
    job["stage"] = stage
    folder = HERE / stage
    report_path = folder / f"{job['name']}.json"
    regular(report_path, f"{stage}/{job['name']}.json")
    receipt, start, end = _validate_capture_receipt(
        stage, job, binary, manifest_sha, candidate_manifest_sha, report_path, set())
    report = _guard_report(read_json(report_path), job, binary, f"{stage}/{job['name']}")
    report["report_sha256"] = sha256(report_path)
    report["receipt"] = {
        "path": relative(HERE / stage / f"{job['name']}.receipt.json"),
        "sha256": sha256(HERE / stage / f"{job['name']}.receipt.json"),
        "start_utc": receipt["start_utc"], "end_utc": receipt["end_utc"],
        "seconds": receipt["seconds"], "execution_stage": receipt["execution_stage"],
        "execution_manifest_sha256": receipt["execution_manifest_sha256"],
        "source_manifest_sha256": receipt["source_manifest_sha256"],
    }
    return report, (start, end)


def _column_name(column: int) -> str:
    value = column + 1
    output = ""
    while value:
        value, remainder = divmod(value - 1, 26)
        output = chr(ord("A") + remainder) + output
    return output


def _cap_fixture(report: dict[str, Any], fixture_path: Path, receipt: dict[str, Any],
                 label: str) -> dict[str, Any]:
    regular(fixture_path, f"{label}.fixture")
    actual_digest = sha256(fixture_path)
    actual_bytes = fixture_path.stat().st_size
    hash_value(receipt["artifacts"][fixture_path.name], f"{label}.receipt.fixture")
    require(receipt["artifacts"][fixture_path.name] == actual_digest,
            f"{label}.fixture receipt digest differs")
    return {"path": relative(fixture_path), "sha256": actual_digest, "bytes": actual_bytes}


def _cap_source(value: Any, size: int, fixture: dict[str, Any], label: str) -> dict[str, Any]:
    require(isinstance(value, dict) and set(value) == CAP_SOURCE_KEYS,
            f"{label}.source inventory differs")
    expected_kind = ("single_worksheet_sparse_numeric_with_comment"
                     if size <= 2 else "single_worksheet_dense_numeric_grid")
    expected = {
        "shape": f"{size}x{size}", "rows": size, "columns": size,
        "format": "OOXML/XLSX", "fixture_kind": expected_kind,
        "generator": "litchi-xlsx-cap-boundary-stored-grid-v1",
        "worksheet_member": "xl/worksheets/sheet1.xml", "compression": "stored",
        "encoding": "UTF-8", "marker_free": True,
        "identity_method": "parent binds SHA-256 of fixture_dump bytes",
    }
    for field, expected_value in expected.items():
        require(value[field] == expected_value, f"{label}.source.{field} differs")
    for field in ("rows", "columns", "source_xml_bytes", "worksheet_bytes",
                  "worksheet_xml_bytes", "archive_bytes"):
        positive_integer(value[field], f"{label}.source.{field}")
    require(value["archive_bytes"] == fixture["bytes"],
            f"{label}.source.archive_bytes differs from fixture")
    require(Path(value["fixture_dump"]) == (HERE / fixture["path"]).resolve(),
            f"{label}.source.fixture_dump differs")
    return value


def _cap_phase(value: Any, count: int, warmup: int, label: str) -> dict[str, Any]:
    require(isinstance(value, dict) and set(value) == CAP_PHASE_KEYS,
            f"{label}.phase inventory differs")
    require(value["name"] == "edit_sheets" and value["warmup_iterations"] == warmup
            and value["samples"] == count, f"{label}.phase metadata differs")
    require(value["timing_scope"] ==
            "source/editor/selector setup excluded; transaction retained through clock stop; inspection/commit/drop excluded",
            f"{label}.phase.timing_scope differs")
    require(value["sample_order"] == list(range(count)), f"{label}.phase.sample_order differs")
    durations = value["duration_ns"]
    require(isinstance(durations, list) and len(durations) == count,
            f"{label}.phase.duration_ns cardinality differs")
    for index, item in enumerate(durations):
        nonnegative_integer(item, f"{label}.phase.duration_ns[{index}]")
    return stats(durations, "ns")


def _cap_report(report: Any, job: dict[str, Any], binary: dict[str, Any], fixture: dict[str, Any],
                label: str) -> dict[str, Any]:
    require(isinstance(report, dict) and set(report) == CAP_REPORT_KEYS,
            f"{label} report inventory differs")
    size = job["size"]
    require(report["schema"] == CAP_SCHEMA and report["tool"] == CAP_TOOL
            and report["case"] == "valid" and report["size"] == size,
            f"{label} schema/case/size differs")
    require(report["warmup_iterations"] == job["warmup"]
            and report["samples"] == job["samples"], f"{label} counts differ")
    cells = size * size
    require(report["rows"] == size and report["columns"] == size
            and report["cells"] == cells
            and report["last_cell"] == f"{_column_name(size - 1)}{size}",
            f"{label} dimensions differ")
    expected_events = 5 * cells + 2 * size + 6 + int(size <= 2)
    require(report["event_count"] == expected_events
            and report["expected_event_count"] == expected_events,
            f"{label} event count differs")
    require(report["event_count_formula"] == "5*N*N+2*N+6+I(N<=2) (including EOF)",
            f"{label} event formula differs")
    require(report["comment_bytes"] == (SPARSE_COMMENT_BYTES if size <= 2 else 0),
            f"{label} comment bytes differs")
    require(report["shared_provisional_event_cap"] == SHARED_EVENT_CAP
            and report["ordinary_parser_event_cap"] == ORDINARY_EVENT_CAP,
            f"{label} event caps differ")
    require(report["event_cap_relation"] == ("below" if expected_events <= SHARED_EVENT_CAP else "above"),
            f"{label} event-cap relation differs")
    require(report["source_stream_byte_limit"] == SOURCE_STREAM_BYTE_LIMIT
            and report["source_stream_eligible"] is True,
            f"{label} source-stream eligibility differs")
    require(report["archive_bytes"] == fixture["bytes"],
            f"{label}.archive_bytes differs from fixture")
    expected_fixture_path = (HERE / fixture["path"]).resolve()
    require(Path(report["fixture_out"]) == expected_fixture_path,
            f"{label}.fixture_out differs")
    source = _cap_source(report["source"], size, fixture, label)
    require(source["archive_bytes"] == report["archive_bytes"]
            and source["source_xml_bytes"] == report["source_xml_bytes"],
            f"{label} source byte fields disagree")
    binary_report = report["binary"]
    require(isinstance(binary_report, dict) and set(binary_report) == CAP_BINARY_KEYS,
            f"{label}.binary inventory differs")
    require(binary_report == {
        "path": binary["path"], "bytes": binary["bytes"], "profile": "release",
        "identity_method": "parent binds SHA-256 of the captured binary",
    }, f"{label}.binary identity differs")
    correctness = report["correctness"]
    expected_correctness = {
        "source_unchanged": True, "source_bytes_unchanged": True,
        "source_xml_unchanged": True, "valid_snapshot_values": True,
        "snapshot_a1": 1, "snapshot_last_cell": cells,
        "empty_commit_is_noop": True, "commit_outside_timing": True,
        "no_op_publication_exact": True,
    }
    require(isinstance(correctness, dict) and set(correctness) == CAP_CORRECTNESS_KEYS
            and correctness == expected_correctness, f"{label}.correctness differs")
    timing = _cap_phase(report["phase"], job["samples"], job["warmup"], label)
    return {
        "name": job["name"], "stage": job["stage"], "repeat": job["repeat"],
        "size": size, "warmup": job["warmup"], "samples": job["samples"],
        "fixture": fixture, "source": source, "correctness": expected_correctness,
        "timing": {"edit_sheets": timing}, "binary": dict(binary),
        "report_sha256": None, "receipt": None,
    }


def _cap_capture(stage: str, job: dict[str, Any], binary: dict[str, Any], manifest_sha: str,
                 candidate_manifest_sha: str) -> tuple[dict[str, Any], tuple[_datetime.datetime, _datetime.datetime]]:
    job = dict(job)
    job["stage"] = stage
    folder = HERE / stage
    report_path = folder / f"{job['name']}.json"
    fixture_path = folder / f"{job['name']}.zip"
    regular(report_path, f"{stage}/{job['name']}.json")
    regular(fixture_path, f"{stage}/{job['name']}.zip")
    receipt, start, end = _validate_capture_receipt(
        stage, job, binary, manifest_sha, candidate_manifest_sha, report_path,
        {fixture_path.name})
    fixture = _cap_fixture(read_json(report_path), fixture_path, receipt, f"{stage}/{job['name']}")
    report = _cap_report(read_json(report_path), job, binary, fixture, f"{stage}/{job['name']}")
    report["report_sha256"] = sha256(report_path)
    receipt_path = HERE / stage / f"{job['name']}.receipt.json"
    report["receipt"] = {
        "path": relative(receipt_path), "sha256": sha256(receipt_path),
        "start_utc": receipt["start_utc"], "end_utc": receipt["end_utc"],
        "seconds": receipt["seconds"], "execution_stage": receipt["execution_stage"],
        "execution_manifest_sha256": receipt["execution_manifest_sha256"],
        "source_manifest_sha256": receipt["source_manifest_sha256"],
    }
    return report, (start, end)


def _missing_stage(stage: str) -> list[str]:
    folder = HERE / stage
    names = ["source-manifest.json", "source.patch", "binary-guard-normal.json",
             "binary-guard-alloc.json", "binary-cap.json",
             "build-guard-normal.receipt.json", "build-guard-alloc.receipt.json",
             "build-cap.receipt.json"]
    for build in ("build-guard-normal", "build-guard-alloc", "build-cap"):
        names += [f"{build}.host.json", f"{build}.stdout", f"{build}.stderr"]
    for lane in LANES:
        for job in _guard_jobs(lane):
            for suffix in (".json", ".receipt.json", ".stdout", ".stderr", ".host.json"):
                names.append(job["name"] + suffix)
    for job in _cap_jobs():
        for suffix in (".json", ".receipt.json", ".stdout", ".stderr", ".host.json", ".zip"):
            names.append(job["name"] + suffix)
    return [f"{stage}/{name}" for name in names if not (folder / name).is_file()]


def _check_order(intervals: list[tuple[_datetime.datetime, _datetime.datetime, str]], label: str) -> None:
    ordered = sorted(intervals, key=lambda item: (item[0], item[1], item[2]))
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            f"{label} receipts overlap")


def _stage(stage: str, candidate_manifest_sha: str) -> dict[str, Any]:
    manifest, manifest_sha = _manifest(stage)
    binaries = {
        "guard-normal": _binary_descriptor(stage, "guard-normal", manifest_sha),
        "guard-alloc": _binary_descriptor(stage, "guard-alloc", manifest_sha),
        "cap": _binary_descriptor(stage, "cap", manifest_sha),
    }
    builds = {
        kind: _build_receipt(stage, kind, manifest_sha)
        for kind in ("guard-normal", "guard-alloc", "cap")
    }
    guard: dict[str, dict[str, Any]] = {}
    cap_rows: list[dict[str, Any]] = []
    all_intervals: list[tuple[_datetime.datetime, _datetime.datetime, str]] = []
    for lane, binary_kind in (("normal", "guard-normal"), ("alloc", "guard-alloc")):
        expected = _guard_jobs(lane)
        prefix = "guard-native-" if lane == "normal" else "guard-alloc-"
        actual = {path.name[:-len(".receipt.json")] for path in (HERE / stage).glob(
            f"{prefix}*.receipt.json")}
        require(actual == {job["name"] for job in expected},
                f"{stage}/{lane} receipt matrix differs")
        rows: list[dict[str, Any]] = []
        lane_intervals = []
        for job in expected:
            row, interval = _guard_capture(stage, lane, job, binaries[binary_kind],
                                            manifest_sha, candidate_manifest_sha)
            rows.append(row)
            lane_intervals.append((interval[0], interval[1], job["name"]))
            all_intervals.append((interval[0], interval[1], job["name"]))
        _check_order(lane_intervals, f"{stage}/{lane}")
        guard[lane] = {"rows": rows, "row_count": len(rows),
                       "samples": sum(row["samples"] for row in rows),
                       "manifest_sha256": manifest_sha,
                       "binary": binaries[binary_kind],
                       "build": builds[binary_kind]}
    expected_cap = _cap_jobs()
    actual_cap = {path.name[:-len(".receipt.json")] for path in (HERE / stage).glob(
        "cap-r*.receipt.json")}
    require(actual_cap == {job["name"] for job in expected_cap},
            f"{stage}/cap receipt matrix differs")
    cap_intervals = []
    for job in expected_cap:
        row, interval = _cap_capture(stage, job, binaries["cap"], manifest_sha,
                                     candidate_manifest_sha)
        cap_rows.append(row)
        cap_intervals.append((interval[0], interval[1], job["name"]))
        all_intervals.append((interval[0], interval[1], job["name"]))
    _check_order(cap_intervals, f"{stage}/cap")
    _check_order(all_intervals, f"{stage} guard/cap captures")
    # Each build is serially completed before its own capture lane begins.
    for kind, build in builds.items():
        rows = guard["normal"]["rows"] if kind == "guard-normal" else (
            guard["alloc"]["rows"] if kind == "guard-alloc" else cap_rows)
        first = min(timestamp(row["receipt"]["start_utc"], "capture start") for row in rows)
        end = timestamp(build["end_utc"], f"{stage}/{kind} build end")
        require(end <= first, f"{stage}/{kind} build overlaps its capture lane")
    return {
        "stage": stage, "status": "pass", "manifest_sha256": manifest_sha,
        "manifest_entries": len(manifest), "binaries": binaries, "builds": builds,
        "guard": guard, "cap": {"rows": cap_rows, "row_count": len(cap_rows),
                                  "samples": sum(row["samples"] for row in cap_rows),
                                  "manifest_sha256": manifest_sha,
                                  "binary": binaries["cap"], "build": builds["cap"]},
    }


def _identity_guard(row: dict[str, Any]) -> dict[str, Any]:
    return {key: row[key] for key in ("source", "source_sha256", "logical_error", "correctness")}


def _guard_cross_identity(stage: dict[str, Any]) -> dict[str, Any]:
    normal = {(row["repeat"], row["shape"], row["case"]): row
              for row in stage["guard"]["normal"]["rows"]}
    alloc = {(row["repeat"], row["shape"], row["case"]): row
             for row in stage["guard"]["alloc"]["rows"]}
    require(set(normal) == set(alloc), "guard normal/allocator row keys differ")
    rows = []
    for key in sorted(normal):
        require(_identity_guard(normal[key]) == _identity_guard(alloc[key]),
                f"guard normal/allocator identity differs for {key}")
        rows.append({"repeat": key[0], "shape": key[1], "case": key[2],
                     "identity_equal": True, "source_sha256": normal[key]["source_sha256"]})
    return {"rows": rows, "passed": True}


def _all_guard_map(stage: dict[str, Any], lane: str) -> dict[tuple[int, str, str], dict[str, Any]]:
    return {(row["repeat"], row["shape"], row["case"]): row
            for row in stage["guard"][lane]["rows"]}


def _drift_record(stage_name: str, lane: str, repeat_first: int, repeat_second: int,
                  shape: str, case: str | None, metric: str, first: Any, second: Any) -> dict[str, Any]:
    value = comparison(first, second, f"{stage_name}/{lane}/{shape}/{case}/{metric}")
    return {"stage": stage_name, "lane": lane, "shape": shape, "case": case,
            "repeat_first": repeat_first, "repeat_second": repeat_second,
            "metric": metric, **value,
            "threshold_percent": DRIFT_THRESHOLD_PERCENT,
            "over_five_percent": over_five(value)}


def _guard_comparison(base: dict[str, Any], cand: dict[str, Any]) -> dict[str, Any]:
    base_normal = _all_guard_map(base, "normal")
    cand_normal = _all_guard_map(cand, "normal")
    base_alloc = _all_guard_map(base, "alloc")
    cand_alloc = _all_guard_map(cand, "alloc")
    require(set(base_normal) == set(cand_normal) == set(base_alloc) == set(cand_alloc),
            "guard baseline/candidate matrices differ")
    native_rows: list[dict[str, Any]] = []
    alloc_rows: list[dict[str, Any]] = []
    adverse: list[dict[str, Any]] = []
    for key in sorted(base_normal):
        before, after = base_normal[key], cand_normal[key]
        require(_identity_guard(before) == _identity_guard(after),
                f"guard baseline/candidate identity differs for {key}")
        metrics = {}
        for metric in STATS:
            item = comparison(before["timing"]["edit_sheets"][metric],
                              after["timing"]["edit_sheets"][metric],
                              f"guard/normal/{key}/{metric}")
            metrics[metric] = item
            if item["change_percent"] is None or (
                item["change_percent"] > DRIFT_THRESHOLD_PERCENT):
                adverse.append({"lane": "normal", "repeat": key[0], "shape": key[1],
                                "case": key[2], "metric": f"timing.edit_sheets.{metric}",
                                **item})
        native_rows.append({"repeat": key[0], "shape": key[1], "case": key[2],
                            "baseline": before["timing"], "candidate": after["timing"],
                            "comparisons": metrics, "identity_equal": True})
        before_a, after_a = base_alloc[key], cand_alloc[key]
        require(_identity_guard(before_a) == _identity_guard(after_a),
                f"guard allocation baseline/candidate identity differs for {key}")
        alloc_metrics = {}
        for field in ALLOC_FIELDS + ("incremental_region_peak_live_bytes",
                                     "retained_after_live_bytes"):
            field_metrics = {}
            left = before_a["allocation"]["statistics"][field]
            right = after_a["allocation"]["statistics"][field]
            for metric in STATS:
                item = comparison(left[metric], right[metric],
                                  f"guard/alloc/{key}/{field}/{metric}")
                field_metrics[metric] = item
                if item["change_percent"] is None or item["change_percent"] > DRIFT_THRESHOLD_PERCENT:
                    adverse.append({"lane": "alloc", "repeat": key[0], "shape": key[1],
                                    "case": key[2], "metric": f"allocation.{field}.{metric}",
                                    **item})
            alloc_metrics[field] = field_metrics
        alloc_rows.append({"repeat": key[0], "shape": key[1], "case": key[2],
                           "baseline": {"allocation": before_a["allocation"]},
                           "candidate": {"allocation": after_a["allocation"]},
                           "comparisons": alloc_metrics, "identity_equal": True,
                           "instrumented_elapsed_excluded": True})
    drift: list[dict[str, Any]] = []
    for stage_name, stage in (("baseline", base), ("candidate", cand)):
        normal = _all_guard_map(stage, "normal")
        alloc = _all_guard_map(stage, "alloc")
        for shape in GUARD_SHAPES:
            for case in GUARD_CASES:
                first, second = normal[(1, shape, case)], normal[(2, shape, case)]
                for metric in STATS:
                    drift.append(_drift_record(stage_name, "normal", 1, 2, shape, case,
                                               f"timing.edit_sheets.{metric}",
                                               first["timing"]["edit_sheets"][metric],
                                               second["timing"]["edit_sheets"][metric]))
                first, second = alloc[(1, shape, case)], alloc[(2, shape, case)]
                for field in ALLOC_FIELDS + ("incremental_region_peak_live_bytes",
                                             "retained_after_live_bytes"):
                    for metric in STATS:
                        drift.append(_drift_record(
                            stage_name, "alloc", 1, 2, shape, case,
                            f"allocation.{field}.{metric}",
                            first["allocation"]["statistics"][field][metric],
                            second["allocation"]["statistics"][field][metric]))
    gates: dict[str, list[dict[str, Any]]] = {
        "valid_native_p50_mean": [], "invalid_native_p50_mean": [],
        "valid_noop_incremental_peak": [], "invalid_incremental_peak": [],
    }
    for repeat in REPEATS:
        for shape in GUARD_SHAPES:
            base_valid = base_normal[(repeat, shape, "valid")]
            cand_valid = cand_normal[(repeat, shape, "valid")]
            for metric in GATE_STATS:
                item = comparison(base_valid["timing"]["edit_sheets"][metric],
                                  cand_valid["timing"]["edit_sheets"][metric],
                                  f"guard valid {repeat}/{shape}/{metric}")
                allowed = base_valid["timing"]["edit_sheets"][metric] * VALID_MAX_RATIO
                gates["valid_native_p50_mean"].append({"repeat": repeat, "shape": shape,
                    "metric": f"timing.edit_sheets.{metric}", **item,
                    "max_allowed_ratio": VALID_MAX_RATIO,
                    "passed": item["candidate"] <= allowed})
            base_valid_a = base_alloc[(repeat, shape, "valid")]
            cand_valid_a = cand_alloc[(repeat, shape, "valid")]
            base_peak = base_valid_a["allocation"]["statistics"][
                "incremental_region_peak_live_bytes"]["max"]
            cand_peak = cand_valid_a["allocation"]["statistics"][
                "incremental_region_peak_live_bytes"]["max"]
            gates["valid_noop_incremental_peak"].append({
                "repeat": repeat, "shape": shape,
                "metric": "allocation.incremental_region_peak_live_bytes.max",
                **comparison(base_peak, cand_peak, f"guard no-op peak {repeat}/{shape}"),
                "max_allowed_ratio": VALID_MAX_RATIO,
                "passed": cand_peak <= base_peak * VALID_MAX_RATIO,
            })
            for case in ("late-validator", "late-raw"):
                base_invalid = base_normal[(repeat, shape, case)]
                cand_invalid = cand_normal[(repeat, shape, case)]
                for metric in GATE_STATS:
                    base_invalid_value = base_invalid["timing"]["edit_sheets"][metric]
                    cand_invalid_value = cand_invalid["timing"]["edit_sheets"][metric]
                    base_valid_value = base_valid["timing"]["edit_sheets"][metric]
                    bound = max(base_invalid_value * VALID_MAX_RATIO,
                                base_valid_value * INVALID_NATIVE_MAX_RATIO)
                    gates["invalid_native_p50_mean"].append({
                        "repeat": repeat, "shape": shape, "case": case,
                        "metric": f"timing.edit_sheets.{metric}",
                        **comparison(base_invalid_value, cand_invalid_value,
                                     f"guard invalid {repeat}/{shape}/{case}/{metric}"),
                        "baseline_valid": base_valid_value,
                        "baseline_invalid": base_invalid_value,
                        "max_allowed_value": bound,
                        "max_allowed_ratio_to_invalid": VALID_MAX_RATIO,
                        "max_allowed_ratio_to_valid": INVALID_NATIVE_MAX_RATIO,
                        "passed": cand_invalid_value <= bound,
                    })
                base_peak_values = base_alloc[(repeat, shape, "valid")]["allocation"][
                    "statistics"]["incremental_region_peak_live_bytes"]
                cand_peak_values = cand_alloc[(repeat, shape, case)]["allocation"][
                    "statistics"]["incremental_region_peak_live_bytes"]
                base_min = base_peak_values["min"]
                cand_max = cand_peak_values["max"]
                gates["invalid_incremental_peak"].append({
                    "repeat": repeat, "shape": shape, "case": case,
                    "metric": "allocation.incremental_region_peak_live_bytes",
                    "baseline_valid_min": base_min, "candidate_invalid_max": cand_max,
                    "candidate_to_baseline_valid_ratio": None if base_min == 0 else cand_max / base_min,
                    "max_allowed_ratio": INVALID_PEAK_MAX_RATIO,
                    "passed": cand_max == 0 if base_min == 0 else cand_max <= base_min * INVALID_PEAK_MAX_RATIO,
                })
    gate_passed = all(item["passed"] for values in gates.values() for item in values)
    return {
        "native": {"rows": native_rows, "allocator_elapsed_excluded": True},
        "allocation": {"rows": alloc_rows, "allocator_elapsed_excluded": True},
        "gates": gates, "admission_passed": gate_passed,
        "adverse_flags_over_five_percent": adverse,
        "same_build_drift": drift,
        "same_build_drift_over_five_percent": [item for item in drift if item["over_five_percent"]],
        "thresholds": {"adverse_percent": DRIFT_THRESHOLD_PERCENT,
                       "valid_max_ratio": VALID_MAX_RATIO,
                       "invalid_native_max_ratio": INVALID_NATIVE_MAX_RATIO,
                       "invalid_peak_max_ratio": INVALID_PEAK_MAX_RATIO},
    }


def _cap_comparison(base: dict[str, Any], cand: dict[str, Any]) -> dict[str, Any]:
    left = {(row["repeat"], row["size"]): row for row in base["cap"]["rows"]}
    right = {(row["repeat"], row["size"]): row for row in cand["cap"]["rows"]}
    require(set(left) == set(right), "cap baseline/candidate matrices differ")
    rows: list[dict[str, Any]] = []
    adverse: list[dict[str, Any]] = []
    gates: list[dict[str, Any]] = []
    fixture_identities: dict[str, Any] = {}
    for size in CAP_SIZES:
        identities = [left[(repeat, size)]["fixture"] for repeat in REPEATS] + \
            [right[(repeat, size)]["fixture"] for repeat in REPEATS]
        keys = {(item["sha256"], item["bytes"]) for item in identities}
        require(len(keys) == 1, f"cap fixture identity differs for size {size}")
        digest, byte_count = next(iter(keys))
        fixture_identities[str(size)] = {
            "sha256": digest, "bytes": byte_count,
            "files": [item["path"] for item in identities],
        }
    for key in sorted(left):
        before, after = left[key], right[key]
        require((before["fixture"]["sha256"], before["fixture"]["bytes"]) ==
                (after["fixture"]["sha256"], after["fixture"]["bytes"]),
                f"cap fixture identity differs for {key}")
        metric_values: dict[str, Any] = {}
        for metric in STATS:
            item = comparison(before["timing"]["edit_sheets"][metric],
                              after["timing"]["edit_sheets"][metric],
                              f"cap/{key}/{metric}")
            metric_values[metric] = item
            if item["change_percent"] is None or item["change_percent"] > DRIFT_THRESHOLD_PERCENT:
                adverse.append({"repeat": key[0], "size": key[1],
                                "metric": f"timing.edit_sheets.{metric}", **item})
        gate_metrics = {}
        for metric in GATE_STATS:
            item = metric_values[metric]
            gate_metrics[metric] = {
                **item, "max_allowed_ratio": VALID_MAX_RATIO,
                "passed": item["candidate"] <= item["baseline"] * VALID_MAX_RATIO,
            }
            gates.append({"repeat": key[0], "size": key[1], "metric": metric,
                          **gate_metrics[metric]})
        rows.append({"repeat": key[0], "size": key[1],
                     "baseline": before["timing"], "candidate": after["timing"],
                     "comparisons": metric_values, "gate": gate_metrics,
                     "fixture": fixture_identities[str(key[1])]})
    drift: list[dict[str, Any]] = []
    for stage_name, stage in (("baseline", base), ("candidate", cand)):
        stage_rows = {(row["repeat"], row["size"]): row for row in stage["cap"]["rows"]}
        for size in CAP_SIZES:
            first, second = stage_rows[(1, size)], stage_rows[(2, size)]
            for metric in STATS:
                drift.append(_drift_record(stage_name, "cap", 1, 2, str(size), None,
                                           f"timing.edit_sheets.{metric}",
                                           first["timing"]["edit_sheets"][metric],
                                           second["timing"]["edit_sheets"][metric]))
    passed = all(item["passed"] for item in gates)
    return {
        "rows": rows, "gates": gates, "admission_passed": passed,
        "fixture_identities_by_size": fixture_identities,
        "adverse_flags_over_five_percent": adverse,
        "same_build_drift": drift,
        "same_build_drift_over_five_percent": [item for item in drift if item["over_five_percent"]],
        "thresholds": {"valid_max_ratio": VALID_MAX_RATIO,
                       "adverse_percent": DRIFT_THRESHOLD_PERCENT},
    }


def _abba(base: dict[str, Any], cand: dict[str, Any], kind: str, lane: str | None = None) -> dict[str, Any]:
    if kind == "guard":
        base_rows = base["guard"][lane]["rows"]
        cand_rows = cand["guard"][lane]["rows"]
        jobs = _guard_jobs(lane)
    else:
        base_rows = base["cap"]["rows"]
        cand_rows = cand["cap"]["rows"]
        jobs = _cap_jobs()
    groups = []
    for stage_name, repeat, rows in (("baseline", 1, base_rows), ("candidate", 1, cand_rows),
                                      ("candidate", 2, cand_rows), ("baseline", 2, base_rows)):
        selected = [row for row in rows if row["repeat"] == repeat]
        ordered = sorted(selected, key=lambda row: (row["receipt"]["start_utc"], row["name"]))
        expected_names = [job["name"] for job in jobs if job["repeat"] == repeat]
        require([row["name"] for row in ordered] == expected_names,
                f"{kind}/{lane or 'cap'} {stage_name}-r{repeat} order differs")
        group = {"group": f"{stage_name}-r{repeat}", "stage": stage_name,
                 "repeat": repeat, "job_names": expected_names,
                 "first_start_utc": ordered[0]["receipt"]["start_utc"],
                 "last_end_utc": ordered[-1]["receipt"]["end_utc"]}
        if groups:
            require(timestamp(groups[-1]["last_end_utc"], "ABBA previous end") <=
                    timestamp(group["first_start_utc"], "ABBA current start"),
                    f"{kind}/{lane or 'cap'} ABBA groups overlap")
        groups.append(group)
    return {"order": [group["group"] for group in groups], "groups": groups, "passed": True}


def _missing_document(plan: dict[str, Any], missing: dict[str, list[str]],
                      frozen: dict[str, Any], supplemental: dict[str, Any]) -> dict[str, Any]:
    return {
        "schema": "litchi.xlsx.guard-cap-analysis.v1",
        "status": "pending", "admission_status": "pending",
        "plan_sha256": sha256(PLAN_PATH), "capture_sha256": sha256(CAPTURE_PATH),
        "run_sha256": sha256(RUN_PATH), "guarded_capture_sha256": sha256(GUARDED_CAPTURE_PATH),
        "analyzer_sha256": sha256(Path(__file__)), "frozen_inputs": frozen,
        "supplemental_inputs": supplemental,
        "expected": {"stages": list(STAGES), "guard_lanes": list(LANES),
                      "guard_shapes": list(GUARD_SHAPES), "guard_cases": list(GUARD_CASES),
                      "repeats": list(REPEATS), "cap_sizes": list(CAP_SIZES),
                      "guard_native_samples": 200, "guard_alloc_samples": 20,
                      "cap_samples": 200},
        "stages": {stage: {"status": "pending", "missing_artifacts": missing[stage]}
                   for stage in STAGES},
        "performance_claim": "none: supplemental guard/cap evidence only; no final adoption or speedup claim",
        "note": "All guard and cap artifacts must validate before comparison; no synthetic evidence was created.",
    }


_plan_revision = ""


def analyze() -> dict[str, Any]:
    global _plan_revision
    plan = _plan()
    _plan_revision = plan["revision"]
    frozen = _frozen_inputs()
    supplemental = _supplemental_inputs()
    missing = {stage: _missing_stage(stage) for stage in STAGES}
    if any(missing.values()):
        return _missing_document(plan, missing, frozen, supplemental)
    candidate_manifest_sha = sha256(HERE / "candidate" / "source-manifest.json")
    stages = {stage: _stage(stage, candidate_manifest_sha) for stage in STAGES}
    for stage in STAGES:
        _guard_cross_identity(stages[stage])
    base, cand = stages["baseline"], stages["candidate"]
    guard_comparison = _guard_comparison(base, cand)
    cap_comparison = _cap_comparison(base, cand)
    abba = {
        "guard_normal": _abba(base, cand, "guard", "normal"),
        "guard_alloc": _abba(base, cand, "guard", "alloc"),
        "cap": _abba(base, cand, "cap"),
    }
    return {
        "schema": "litchi.xlsx.guard-cap-analysis.v1", "status": "pass",
        "admission_status": "pass" if guard_comparison["admission_passed"]
        and cap_comparison["admission_passed"] else "reject",
        "plan_sha256": sha256(PLAN_PATH), "capture_sha256": sha256(CAPTURE_PATH),
        "run_sha256": sha256(RUN_PATH), "guarded_capture_sha256": sha256(GUARDED_CAPTURE_PATH),
        "analyzer_sha256": sha256(Path(__file__)), "frozen_inputs": frozen,
        "supplemental_inputs": supplemental,
        "expected": {"stages": list(STAGES), "guard_lanes": list(LANES),
                      "guard_shapes": list(GUARD_SHAPES), "guard_cases": list(GUARD_CASES),
                      "repeats": list(REPEATS), "cap_sizes": list(CAP_SIZES),
                      "guard_native_samples": 200, "guard_alloc_samples": 20,
                      "cap_samples": 200},
        "stages": stages,
        "comparison": {"guard": guard_comparison, "cap": cap_comparison, "abba": abba},
        "performance_claim": "none: supplemental guard/cap gates only; no final adoption or speedup claim",
        "limitations": [
            "Allocator-instrumented elapsed samples are excluded from latency comparisons.",
            "The supplemental lanes do not establish the main workflow, profile, quality, hardware, cold, provider, or scaling gates.",
            "The current single-sheet SourceEdit API remains outside the measured owner.",
        ],
    }


def write_exclusive_or_identical(path: Path, value: dict[str, Any]) -> None:
    encoded = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode()
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("xb") as stream:
            stream.write(encoded)
    except FileExistsError:
        require(path.is_file() and not path.is_symlink(),
                f"output exists but is not a regular file: {path}")
        require(path.read_bytes() == encoded,
                f"existing output differs; refusing replacement: {path}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path,
                        help="explicit output path; exclusive-create/identical-replay")
    args = parser.parse_args()
    try:
        result = analyze()
        if args.output is None:
            print(json.dumps(result, indent=2, sort_keys=True))
        else:
            write_exclusive_or_identical(args.output, result)
            print(f"0552 guard/cap analysis {result['status']}: {args.output}")
    except FileNotFoundError as error:
        print(f"0552 guard/cap analyzer: missing {error.filename}", file=sys.stderr)
        return 2
    except (EvidenceError, OSError, ValueError, KeyError) as error:
        print(f"0552 guard/cap analyzer: error: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
