#!/usr/bin/env python3
"""Validate and summarize the 0550 XLSX commit attribution capture.

This analyzer is intentionally a baseline diagnostic.  It consumes the
preflight, native, and allocator children emitted by ``capture.jobs`` and
does not compare a candidate, assign an admission gate, or manufacture a
speedup claim.  Native elapsed samples are reported with the workflow and
its four source-backed phases.  The allocator executable is a separate
instrumentation binary: only its plan/commit/publication allocation vectors
are reported, and its instrumented elapsed samples are never interpreted as
latency evidence.

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
import capture as CAP  # noqa: E402
import run as RUN  # noqa: E402


_binding_path = REPO / "tools" / "validate_perf_corpus_binding.py"
_binding_spec = importlib.util.spec_from_file_location(
    "perf_corpus_binding_for_0550", _binding_path
)
if _binding_spec is None or _binding_spec.loader is None:
    raise ImportError(f"cannot load corpus binding validator: {_binding_path}")
_BINDING = importlib.util.module_from_spec(_binding_spec)
_binding_spec.loader.exec_module(_BINDING)


# Reuse the established 0546 numerical helper for the report-statistics
# calculation.  Its source-bound report validator is tied to the older plan,
# so only the pure arithmetic helper is used here; this analyzer owns the
# 0550 matrix and all evidence bindings below.
_helper_path = REPO / "docs" / "performance" / "results" / "change-0546" / "integration" / "analyze.py"
_helper_spec = importlib.util.spec_from_file_location(
    "xlsx_0546_numerical_helper_for_0550", _helper_path
)
if _helper_spec is None or _helper_spec.loader is None:
    raise ImportError(f"cannot load numerical helper: {_helper_path}")
_HELPER = importlib.util.module_from_spec(_helper_spec)
_helper_spec.loader.exec_module(_HELPER)
_REPORT_STATS = _HELPER.BASE.report_stats
_helper_base_path = Path(_HELPER.HELPER).resolve()


SCHEMA = "xlsx_multisource_edit_metrics_v1"
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
REVISION_RE = re.compile(r"^[0-9a-f]{40}$")
STAGE = "baseline"
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
ALLOCATION_FIELD_NAMES = (
    "allocation_calls",
    "allocated_bytes",
    "incremental_region_peak_live_bytes",
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
        plan.get("scope") == "Current source-backed XLSX MultiSourceEdit commit attribution; no optimization or speedup claim",
        "plan scope differs from the frozen diagnostic",
    )
    require(plan.get("priority") == "OLE2/OOXML first; ODF deferred until that optimization goal completes; iWork excluded",
            "plan priority differs")
    require(plan.get("cpu") == 2, "plan CPU differs from the pinned CPU")
    require(plan.get("cases") == [
        "xlsx_source_backed_cell_values_one_edit_save",
        "xlsx_source_backed_cell_values_one_percent_edit_save",
    ], "plan case matrix differs")
    require(plan.get("shapes") == ["medium", "dense-sparse", "noncompact", "vendor-extension"],
            "plan shape matrix differs")
    for lane in ("native", "alloc"):
        config = plan.get(lane)
        require(isinstance(config, dict), f"plan.{lane} is missing")
        require(config.get("repeats") == 2 and config.get("warmup") == 3
                and config.get("samples") == 30,
                f"plan.{lane} counts differ")
    require(plan.get("drift_review_percent") == 5,
            "plan repeat-drift threshold differs")
    profile = plan.get("profile")
    require(isinstance(profile, dict), "plan.profile is missing")
    require(profile.get("case") == "xlsx_source_backed_cell_values_one_percent_edit_save",
            "plan profile case differs")
    require(profile.get("owner") == "litchi_xlsx::cell_values::source::MultiSourceEdit::commit",
            "plan profile owner differs")
    return plan


def frozen_inputs() -> dict[str, str]:
    frozen = read_json(HERE / "frozen-inputs.json")
    require(isinstance(frozen, dict), "frozen-inputs.json is not an object")
    for name in ("run.py", "capture.py", "plan.json", "adr-manifest.json"):
        expected = frozen.get(name)
        require(is_hash(expected), f"frozen input hash for {name} is malformed")
        require(expected == sha256(HERE / name), f"frozen input {name} has changed")
    return {name: frozen[name] for name in sorted(frozen)}


def _cleanup_binary_hash(kind: str, expected_path: Path) -> str | None:
    cleanup_path = HERE / "cleanup.json"
    if not cleanup_path.is_file():
        return None
    cleanup = read_json(cleanup_path)
    if not isinstance(cleanup, dict):
        return None
    if cleanup.get("owned_paths_absent") is not True:
        return None
    if cleanup.get("accessible_process_references") != []:
        return None
    removed = cleanup.get("removed")
    owned = plan_data().get("owned_paths")
    if removed != owned or not isinstance(owned, list):
        return None
    by_kind = cleanup.get("binary_sha256_by_kind")
    if not isinstance(by_kind, dict):
        return None
    for key in (f"baseline/{kind}", str(expected_path), expected_path.name, kind):
        value = by_kind.get(key)
        if is_hash(value):
            return value
    return None


def binary_metadata(kind: str, plan: dict[str, Any]) -> dict[str, Any]:
    require(kind in (NORMAL_BINARY, ALLOC_BINARY), f"unknown binary kind {kind}")
    folder = HERE / STAGE
    descriptor_path = folder / f"binary-{kind}.json"
    descriptor = read_json(descriptor_path)
    require(isinstance(descriptor, dict), f"{descriptor_path.name} is not an object")
    path = Path(descriptor.get("path", ""))
    expected_path = Path(RUN.SCRATCH) / kind
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
        require(_cleanup_binary_hash(kind, path) == digest,
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
    require(build.get("execution_stage") == STAGE,
            f"{build_path.name} execution stage differs")
    require(build.get("execution_manifest_sha256") == sha256(manifest),
            f"{build_path.name} execution manifest differs")
    require(build.get("script_sha256") == sha256(HERE / "run.py"),
            f"{build_path.name} script hash differs")
    require(build.get("plan_sha256") == sha256(HERE / "plan.json"),
            f"{build_path.name} plan hash differs")
    require(build.get("source_manifest_sha256") == sha256(manifest),
            f"{build_path.name} source manifest hash differs")
    target = Path(plan["owned_paths"][0])
    expected_command = [
        "env",
        "TMPDIR=" + str(target / "tmp"),
        "CARGO_BUILD_JOBS=2",
        "CARGO_INCREMENTAL=0",
        "cargo",
        "build",
        "--release",
        "--locked",
        "--manifest-path",
        "tools/perf-baseline/Cargo.toml",
        "--bin",
        "litchi-perf-baseline-alloc" if kind == ALLOC_BINARY else "litchi-perf-baseline",
        "--target-dir",
        str(target),
    ]
    if kind == ALLOC_BINARY:
        expected_command += ["--features", "allocator-metrics"]
    require(build.get("command") == expected_command,
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
        "binary": "litchi-perf-baseline-alloc" if kind == ALLOC_BINARY else "litchi-perf-baseline",
    }


def expected_command(job: dict[str, Any], kind: str, binary: dict[str, Any]) -> list[str]:
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
        "--warmup",
        str(job["warmup"]),
        "--samples",
        str(job["samples"]),
        "--json",
        str(HERE / STAGE / (job["name"] + ".json")),
        "--corpus-manifest",
        str(HERE / STAGE / (job["name"] + ".catalog.json")),
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
    """Return source identity after excluding measured phase/allocation vectors."""

    require(isinstance(source, dict), "result.source is not an object")
    result: dict[str, Any] = {}
    for key in sorted(source):
        value = source[key]
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
    return {
        "accepted_bytes": value["accepted_bytes"],
        "write_calls": value["write_calls"],
        "largest_write": value["largest_write"],
        "write_size_buckets": {key: buckets[key] for key in sorted(buckets)},
    }


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


def validate_source(source: Any, count: int, measured: bool, label: str) -> tuple[dict[str, Any], dict[str, Any], dict[str, list[int]], dict[str, list[dict[str, Any]]]]:
    require(isinstance(source, dict), f"{label}.source is not an object")
    root_vector_names = (
        "read_calls",
        "read_bytes",
        "ordinary_payload_read_calls",
        "ordinary_payload_read_bytes",
        "max_in_flight_reads",
        "ordinary_payload_materializations",
    )
    for name in root_vector_names:
        integer_vector(source.get(name), count, f"{label}.source.{name}")
    xlsx = source.get("xlsx_cell_values")
    require(isinstance(xlsx, dict), f"{label}.source.xlsx_cell_values is missing")
    require(xlsx.get("implementation") == "source-backed",
            f"{label} implementation is not source-backed")
    require(xlsx.get("cache_mode") == "unmanaged-control",
            f"{label} cache mode is not unmanaged-control")
    require(xlsx.get("cache_budget_managed") is False,
            f"{label} unexpectedly reports a managed budget")
    nonnegative_integer(xlsx.get("update_count"), f"{label}.update_count")
    nonnegative_integer(xlsx.get("selected_worksheet_count"),
                       f"{label}.selected_worksheet_count")
    require(isinstance(xlsx.get("timing_scope"), str) and "commit" in xlsx["timing_scope"],
            f"{label}.timing_scope does not include commit")

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

    # All remaining source arrays are diagnostics, but they must still have
    # exact sample cardinality and JSON-safe nonnegative values.
    for name, value in xlsx.items():
        if name in PHASES or name.endswith("_allocation_metrics"):
            continue
        if isinstance(value, list):
            vector(value, count, f"{label}.source.xlsx_cell_values.{name}")
            for index, item in enumerate(value):
                _validate_recursive_value(item,
                                          f"{label}.source.xlsx_cell_values.{name}[{index}]")
    output_values = xlsx.get("output_sha256")
    output_values = vector(output_values, count,
                           f"{label}.source.xlsx_cell_values.output_sha256")
    for index, item in enumerate(output_values):
        check_hash(item, f"{label}.source.xlsx_cell_values.output_sha256[{index}]")
    semantic_values = xlsx.get("semantic_sha256")
    semantic_values = vector(semantic_values, count,
                             f"{label}.source.xlsx_cell_values.semantic_sha256")
    for index, item in enumerate(semantic_values):
        check_hash(item, f"{label}.source.xlsx_cell_values.semantic_sha256[{index}]")
    nonnegative_integer(xlsx.get("untouched_member_count"),
                       f"{label}.untouched_member_count")
    untouched_values = vector(xlsx.get("untouched_member_sha256"), count,
                              f"{label}.source.xlsx_cell_values.untouched_member_sha256")
    for index, item in enumerate(untouched_values):
        check_hash(item,
                   f"{label}.source.xlsx_cell_values.untouched_member_sha256[{index}]")

    # `source_identity` proves each non-measured source field is constant in a
    # child.  Keep it separate from phase/allocation observations in output.
    stable = source_identity(source, count)
    logical = {
        "source": stable,
        "source_output_sha256": output_values[0],
        "semantic_sha256": semantic_values[0],
        "untouched_member_count": xlsx["untouched_member_count"],
        "untouched_member_sha256": untouched_values[0],
    }
    observations = {
        "source_read_calls": list(source["read_calls"]),
        "source_read_bytes": list(source["read_bytes"]),
        "ordinary_payload_read_calls": list(source["ordinary_payload_read_calls"]),
        "ordinary_payload_read_bytes": list(source["ordinary_payload_read_bytes"]),
        "max_in_flight_reads": list(source["max_in_flight_reads"]),
        "ordinary_payload_materializations": list(source["ordinary_payload_materializations"]),
    }
    return logical, observations, phase_values, allocation_values


def validate_report(path: Path, catalog_path: Path, job: dict[str, Any],
                    kind: str, binary: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    label = job["name"]
    report = read_json(path)
    require(isinstance(report, dict), f"{label} report is not an object")
    require(report.get("schema_version") == 1, f"{label} schema version differs")
    tool = report.get("tool")
    require(isinstance(tool, dict), f"{label}.tool is not an object")
    expected_binary = binary["binary"]
    require(tool.get("binary") == expected_binary, f"{label} tool binary differs")
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

    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{label} must contain one result")
    result = results[0]
    require(isinstance(result, dict), f"{label} result is not an object")
    require(result.get("case") == job["case"], f"{label} result case differs")
    corpus = result.get("corpus")
    require(isinstance(corpus, dict), f"{label}.corpus is missing")
    require(corpus.get("shape") == job["shape"], f"{label} corpus shape differs")
    require(corpus.get("package_format") == "XLSX/OPC/ZIP",
            f"{label} corpus package format differs")
    check_hash(corpus.get("archive_sha256"), f"{label}.corpus.archive_sha256")
    nonnegative_integer(corpus.get("archive_bytes"), f"{label}.corpus.archive_bytes")
    nonnegative_integer(corpus.get("archive_member_count"),
                       f"{label}.corpus.archive_member_count")
    nonnegative_integer(corpus.get("entry_count"), f"{label}.corpus.entry_count")
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
        result.get("source"), job["samples"], kind == ALLOC_BINARY, label
    )
    require(logical["source_output_sha256"] == output,
            f"{label} source output hash differs from result output")
    workflow_values = [
        sum(phase_values[phase][index] for phase in TIMED_PHASES)
        for index in range(job["samples"])
    ]
    require(
        [workflow_values[index] for index in sample_order] == elapsed_values,
        f"{label} phase vectors do not reconcile with elapsed samples",
    )

    # The corpus catalog is independently canonicalized and must bind to the
    # same report result.  This also checks its source archive and case
    # binding, rather than merely checking that a sidecar exists.
    catalog = read_json(catalog_path)
    try:
        _BINDING.validate_binding(report, catalog)
    except Exception as error:  # validator uses its own public exception type
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
        "corpus_ids": sorted(
            item.get("id") for item in catalog.get("corpora", [])
            if isinstance(item, dict)
        ),
        "case_bindings": sorted(
            (
                item.get("case"),
                item.get("corpus_id"),
                item.get("legacy_archive_sha256"),
            )
            for item in catalog.get("case_bindings", [])
            if isinstance(item, dict)
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
    row = {
        "lane": job["lane"],
        "name": label,
        "repeat": job["repeat"],
        "case": job["case"],
        "shape": job["shape"],
        "warmup": job["warmup"],
        "samples": job["samples"],
        "report": str(path.relative_to(HERE)),
        "receipt": str((HERE / STAGE / (label + ".receipt.json")).relative_to(HERE)),
        "catalog": str(catalog_path.relative_to(HERE)),
        "report_sha256": sha256(path),
        "receipt_sha256": sha256(HERE / STAGE / (label + ".receipt.json")),
        "catalog_sha256": catalog_identity["catalog_sha256"],
        "binary_sha256": binary["sha256"],
        "identity": row_identity,
        "identity_sha256": hashlib.sha256(
            json.dumps(row_identity, sort_keys=True, separators=(",", ":")).encode()
        ).hexdigest(),
        "source_observations": source_observations,
        "commit_scope": (
            "commit_ns includes edit staging and MultiSourceEdit::commit; "
            "open, selector planning, and stream publication are separate phases"
        ),
        "allocation_instrumented_elapsed_excluded": kind == ALLOC_BINARY,
        "allocation": {
            PHASE_LABELS[f"{phase}_ns"]: {
                field: stats(
                    [item[field] for item in allocation_values[phase]],
                    "count" if field == "allocation_calls" else "bytes",
                )
                for field in ALLOCATION_FIELD_NAMES
            }
            for phase in ALLOCATION_PHASES
        } if kind == ALLOC_BINARY else None,
    }
    if kind == NORMAL_BINARY:
        row["timing"] = timing
        row["phase_samples"] = {
            PHASE_LABELS[phase]: list(phase_values[phase]) for phase in PHASES
        }
        row["elapsed_sample_order"] = list(sample_order)
        row["elapsed_samples"] = list(elapsed_values)
    return row


def check_receipt(job: dict[str, Any], binary: dict[str, Any],
                  plan: dict[str, Any]) -> tuple[dict[str, Any], Path, Path]:
    folder = HERE / STAGE
    path = folder / f"{job['name']}.receipt.json"
    receipt = read_json(path)
    require(isinstance(receipt, dict), f"{path.name} is not an object")
    validate_receipt_common(receipt, path.name)
    require(receipt.get("exit_code") == 0, f"{path.name} did not exit successfully")
    require(receipt.get("binary_sha256") == binary["sha256"],
            f"{path.name} binary hash differs")
    require(receipt.get("execution_stage") == STAGE,
            f"{path.name} execution stage differs")
    require(receipt.get("execution_manifest_sha256") == sha256(folder / "source-manifest.json"),
            f"{path.name} execution manifest differs")
    require(receipt.get("source_manifest_sha256") == sha256(folder / "source-manifest.json"),
            f"{path.name} source manifest differs")
    require(receipt.get("script_sha256") == sha256(HERE / "run.py"),
            f"{path.name} script hash differs")
    require(receipt.get("plan_sha256") == sha256(HERE / "plan.json"),
            f"{path.name} plan hash differs")
    require(receipt.get("command") == expected_command(job, ALLOC_BINARY if job["lane"] == "alloc" else NORMAL_BINARY, binary),
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


def validate_stage(plan: dict[str, Any], lane: str,
                   binaries: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    jobs = list(CAP.jobs(lane))
    expected_count = {"preflight": 8, "native": 16, "alloc": 16}[lane]
    require(len(jobs) == expected_count,
            f"{lane} capture.jobs produced {len(jobs)} jobs, expected {expected_count}")
    expected_names = {job["name"] for job in jobs}
    prefix = f"{lane}-"
    actual_names = {
        path.name[: -len(".receipt.json")]
        for path in (HERE / STAGE).glob(f"{prefix}*.receipt.json")
    }
    require(actual_names == expected_names,
            f"{lane} receipt matrix differs: {sorted(actual_names ^ expected_names)}")
    kind = ALLOC_BINARY if lane == "alloc" else NORMAL_BINARY
    rows: list[dict[str, Any]] = []
    for job in jobs:
        receipt, report_path, catalog_path = check_receipt(job, binaries[kind], plan)
        row = validate_report(report_path, catalog_path, job, kind, binaries[kind], plan)
        rows.append(row)
    return rows


def validate_stage_metadata(plan: dict[str, Any]) -> tuple[dict[str, str], dict[str, Any], dict[str, Any]]:
    frozen = frozen_inputs()
    folder = HERE / STAGE
    manifest_path = folder / "source-manifest.json"
    require(manifest_path.is_file() and not manifest_path.is_symlink(),
            "baseline source manifest is missing")
    manifest = read_json(manifest_path)
    require(isinstance(manifest, dict) and manifest,
            "baseline source manifest is empty")
    for relative, digest in manifest.items():
        require(isinstance(relative, str) and relative and not Path(relative).is_absolute(),
                "baseline source manifest has an invalid path")
        check_hash(digest, f"baseline source manifest {relative}")
    host = read_json(HERE / "host.json")
    require(isinstance(host, dict), "host.json is not an object")
    return frozen, {
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


def drift_record(first: Any, second: Any, *, lane: str, case: str,
                 shape: str, metric: str) -> dict[str, Any]:
    finite_number(first, f"drift {metric}.first")
    finite_number(second, f"drift {metric}.second")
    change = percent_change(float(first), float(second))
    return {
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
    by_key = {(row["lane"], row["case"], row["shape"], row["repeat"]): row
              for row in rows if row["lane"] in ("native", "alloc")}
    output: list[dict[str, Any]] = []
    for lane in ("native", "alloc"):
        groups = sorted({(row["case"], row["shape"]) for row in rows
                         if row["lane"] == lane})
        for case, shape in groups:
            first = by_key[(lane, case, shape, 1)]
            second = by_key[(lane, case, shape, 2)]
            if lane == "native":
                for phase in ("workflow", "open", "planning", "commit", "publication", "reopen"):
                    for metric in REPORT_STATS:
                        record = drift_record(
                            first["timing"][phase][metric],
                            second["timing"][phase][metric],
                            lane=lane,
                            case=case,
                            shape=shape,
                            metric=f"timing.{phase}.{metric}",
                        )
                        if record["over_five_percent"]:
                            output.append(record)
            else:
                for phase in ("planning", "commit", "publication"):
                    for metric_name in ALLOCATION_FIELD_NAMES:
                        for metric in REPORT_STATS:
                            record = drift_record(
                                first["allocation"][phase][metric_name][metric],
                                second["allocation"][phase][metric_name][metric],
                                lane=lane,
                                case=case,
                                shape=shape,
                                metric=f"allocation.{phase}.{metric_name}.{metric}",
                            )
                            if record["over_five_percent"]:
                                output.append(record)
    return output


def analyze() -> dict[str, Any]:
    plan = plan_data()
    frozen, manifest, host = validate_stage_metadata(plan)
    binaries = {
        NORMAL_BINARY: binary_metadata(NORMAL_BINARY, plan),
        ALLOC_BINARY: binary_metadata(ALLOC_BINARY, plan),
    }
    rows: list[dict[str, Any]] = []
    for lane in LANES:
        rows.extend(validate_stage(plan, lane, binaries))
    require(len([row for row in rows if row["lane"] == "preflight"]) == 8,
            "preflight row count differs")
    require(len([row for row in rows if row["lane"] == "native"]) == 16,
            "native row count differs")
    require(len([row for row in rows if row["lane"] == "alloc"]) == 16,
            "allocation row count differs")
    # Matched identity includes corpus, sink, output, semantic, untouched
    # member, and stable source evidence.  It is checked across all three
    # lanes and both native/allocator repeats for each case/shape.
    identity_groups: dict[tuple[str, str], dict[str, Any]] = {}
    for row in rows:
        key = (row["case"], row["shape"])
        current = row["identity"]
        if key in identity_groups:
            require(identity_groups[key] == current,
                    f"matched source/sink/output/semantic identity differs for {key}")
        else:
            identity_groups[key] = current
    receipts = [
        read_json(HERE / STAGE / (row["name"] + ".receipt.json")) for row in rows
    ]
    ordered = sorted(receipts, key=lambda item: item["start_utc"])
    require(all(
        timestamp(left["end_utc"], "receipt end")
        <= timestamp(right["start_utc"], "receipt start")
        for left, right in zip(ordered, ordered[1:])
    ), "capture receipts overlap")
    by_lane = {
        lane: [row for row in rows if row["lane"] == lane] for lane in LANES
    }
    drift = repeat_drift(rows)
    return {
        "schema": SCHEMA,
        "status": "pass",
        "stage": STAGE,
        "scope": plan["scope"],
        "priority": plan["priority"],
        "performance_claim": "none: descriptive baseline attribution only; no admission gate or speedup claim",
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
        "frozen_inputs": frozen,
        "source_manifest": manifest,
        "host": host,
        "binaries": binaries,
        "job_counts": {
            "preflight": len(by_lane["preflight"]),
            "native": len(by_lane["native"]),
            "alloc": len(by_lane["alloc"]),
        },
        "sample_counts": {
            "preflight": sum(row["samples"] for row in by_lane["preflight"]),
            "native": sum(row["samples"] for row in by_lane["native"]),
            "alloc": sum(row["samples"] for row in by_lane["alloc"]),
        },
        "preflight": {"rows": by_lane["preflight"]},
        "native": {
            "rows": by_lane["native"],
            "timing_scope": (
                "workflow is elapsed_ns and equals open + planning + commit + publication; "
                "commit_ns includes edit staging and MultiSourceEdit::commit; reopen is diagnostic and excluded"
            ),
            "reported_statistics": list(REPORT_STATS),
        },
        "allocation": {
            "rows": by_lane["alloc"],
            "scope": "allocator binary operation_global_system_allocator",
            "reported_metrics": list(ALLOCATION_FIELD_NAMES),
            "instrumented_elapsed_excluded": True,
        },
        "matched_identity": {
            "keys": ["corpus", "sink", "source", "output_sha256",
                     "semantic_sha256", "untouched_member_count",
                     "untouched_member_sha256", "catalog"],
            "case_shape_count": len(identity_groups),
            "all_lanes_and_repeats_equal": True,
        },
        "same_build_drift_over_five_percent": drift,
        "repeat_drift_over_five_percent": drift,
        "limits": [
            "30 native samples per child are descriptive and do not support a registered latency claim",
            "allocator-instrumented elapsed samples are excluded from latency interpretation",
            "source/profile inference is not a counter for copied or reduced bytes",
            "no physical-cold, provider, native Office, concurrency, scaling, or fuzz claim",
            "ODF remains deferred until the OLE2/OOXML optimization goal completes; iWork is excluded",
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
    parser.add_argument("--stage", choices=("baseline",), default="baseline")
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
    print(f"0550 baseline metrics {result['status']}: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
