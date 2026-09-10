#!/usr/bin/env python3
"""Bounded phase evidence for the opened DOCX edit benchmark.

This driver is intentionally a custody layer around the existing 0495 report
validator.  It does not build Rust, reconstruct a source checkout, or infer a
source identity from the mutable worktree.  The four executable and gate
records in ``builds.json`` are the source of truth for the two source phases.

The 0496 harness adds one field to every retained sample row,
``phase_diagnostics``.  The field is checked here and then removed from a
temporary in-memory projection before the unchanged 0495 validator is called.
The retained report is never rewritten.  This keeps the old semantic and
oracle checks authoritative while making the new phase contract fail closed.
"""

from __future__ import annotations

import argparse
import copy
import datetime as _datetime
import fcntl
import hashlib
import importlib.util
import json
import math
import os
from pathlib import Path
import re
import shutil
import signal
import statistics
import subprocess
import sys
import tempfile
from typing import Any, Callable, Iterable


ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
TEMP = Path("/home/zhuhe/.cache/litchi-goal-0496")
CPU_LOCK = Path("/home/zhuhe/.cache/litchi-goal-0484/cpu.lock")
CPU = 2

SCHEMA = "docx-phase-diagnostic-v1"
VERSION = 1
BUILDS_SCHEMA = "docx-phase-builds-v1"
CAPTURE_SCHEMA = "docx-phase-diagnostic-capture-v1"
TERMINAL_SCHEMA = "docx-phase-diagnostic-terminal-v1"
ANALYSIS_SCHEMA = "docx-phase-diagnostic-analysis-v1"
PROTOCOL_SCHEMA = "docx-phase-diagnostic-protocol-v1"
PHASE_SCHEMA = "docx_managed_edit_phase_diagnostics_v1"

CASE = "docx_opened_document_managed_vs_unmanaged_one_paragraph_edit_save"
ROLES = ("normal", "allocator")
PHASES = ("before", "after")
APIS = ("unmanaged-api", "managed-api")
ARMS = ("owned", "file-warm", "short")
UNMANAGED = "unmanaged-api"
MANAGED = "managed-api"
REPEATS = (1, 2)
FORMAL_SAMPLES = 30
FORMAL_WARMUPS = 3
DEFAULT_TIMEOUT_SECONDS = 180
TERM_GRACE_SECONDS = 10
BOOTSTRAP_REPETITIONS = 10_000
BOOTSTRAP_SEED = 496
ADVERSE_THRESHOLD_PERCENT = 5.0
SHORT_RANGE_BYTES = 4_096
ATTEMPT_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]*$")
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
REVISION_RE = re.compile(r"^[0-9a-f]{40}$")
EXPECTED_REVISIONS = {
    "before": "de8ee88b0727ae59e4d2b4c8b8a6c24349724ae8",
    "after": "44a4710699ef17041d5969240c30984dffbc3319",
}
HARNESS_CUSTODY_FILES = (
    "tools/perf-baseline/src/docx_managed_edit.rs",
    "tools/perf-baseline/src/lib.rs",
    "tools/perf-baseline/src/main.rs",
    "tools/perf-baseline/src/bin/litchi-perf-baseline-alloc.rs",
)

OLD_ROOT = ROOT.parent / "change-0495"
OLD_MEASURE = OLD_ROOT / "measure.py"
OLD_SUPPORT = OLD_ROOT / "support.py"
OLD_SEAL = OLD_ROOT / "seal.json"
BUILDS_FILE = ROOT / "builds.json"
PROTOCOL_FILE = ROOT / "protocol.json"
MACHINE_FILE = ROOT / "machine.json"
PROVENANCE_FILE = ROOT / "provenance.json"

# The Rust phase instrument deliberately keeps these labels and scopes stable.
PHASE_TIMING_SCOPE = (
    "opt-in wall-clock Instant intervals nested inside the full lifecycle clock: "
    "open, edit staging, commit, diagnostics/XML identity, publication, "
    "published snapshot drop, and commit drop; these are not CPU-time measurements"
)
PHASE_RESIDUAL_SCOPE = (
    "full lifecycle time minus the listed phase intervals; includes "
    "phase-boundary arithmetic, budget evidence, result handling, and any "
    "package work not assigned to a named phase"
)
PHASE_INSTRUMENTATION_SCOPE = (
    "phase-clock overhead is included in the full lifecycle and is not "
    "isolated by a second control clock; phase fields are absent unless "
    "--phase-diagnostics is enabled"
)
PHASE_ALLOCATION_SCOPE = (
    "no nested phase allocation regions are reported: the shared allocator "
    "region is non-reentrant, so allocation remains one full-lifecycle sample"
)
PHASE_DURATION_FIELDS = (
    "open_ns",
    "edit_staging_ns",
    "commit_ns",
    "diagnostics_xml_identity_ns",
    "publication_ns",
    "published_snapshot_drop_ns",
    "commit_drop_ns",
)
PHASE_FIELDS = (
    "schema",
    "timing_scope",
    *PHASE_DURATION_FIELDS,
    "phase_sum_ns",
    "lifecycle_residual_ns",
    "residual_scope",
    "instrumentation_overhead_ns",
    "instrumentation_scope",
    "allocation_scope",
)

# These are the exact full-lifecycle allocator fields retained by the sealed
# 0495 report validator.  They are reported only for allocator binaries.  A
# normal binary has no allocator sample, so its analysis carries an explicit
# unavailable marker rather than synthetic zero vectors.
SAMPLE_ALLOCATION_SCOPE = "operation_global_system_allocator"
ALLOCATION_VECTOR_FIELDS = (
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
    "allocation_peak_increment_bytes",
)
ALLOCATION_RAW_FIELDS = ALLOCATION_VECTOR_FIELDS[:-1]

# 0495 imports this same fixed report field list.  Keeping the fallback here
# makes the projection shape explicit even when the sealed validator cannot be
# imported by a lightweight unit-test harness.
FALLBACK_REPORT_FIELDS = (
    "schema", "version", "case_name", "benchmark", "api", "api_name",
    "provider_scope", "timing_scope", "setup_scope", "allocation_scope",
    "physical_scope", "range_scope", "zero_length_scope", "file_scope",
    "budget_scope", "corpus_version", "corpus_generator",
    "source_archive_sha256", "source_archive_bytes", "source_bytes",
    "source_sha256", "expected_text_bytes", "expected_text_sha256",
    "expected_text_scope", "expected_archive_members", "expected_output_sha256",
    "expected_output_bytes", "source_revision", "requested_source_revision",
    "limits", "binary_sha256", "binary_bytes", "current_exe", "provider",
    "corpus", "preflight", "warmup", "samples", "allocator",
    "instrumentation", "rows",
)

CAPTURE_ARTIFACTS = ("stdout.txt", "stderr.txt", "resource.txt", "report.json", "replay-cleanup.json")
ENV_KEYS = (
    "PATH", "HOME", "USER", "LANG", "LC_ALL", "RUSTUP_TOOLCHAIN", "RUSTFLAGS",
    "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL", "CARGO_PROFILE_RELEASE_DEBUG",
    "CARGO_TARGET_DIR", "DEBUGINFOD_URLS", "RUSTDOCFLAGS",
)


class PhaseDiagnosticError(RuntimeError):
    """A fail-closed protocol, custody, report, or analysis error."""


def fail(message: str) -> None:
    raise PhaseDiagnosticError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _exact(value: Any, keys: Iterable[str], path: str) -> None:
    require(isinstance(value, dict), f"{path}: expected object")
    expected = set(keys)
    actual = set(value)
    require(actual == expected, f"{path}: fields differ (expected {sorted(expected)}, got {sorted(actual)})")


def _uint(value: Any, path: str, *, positive: bool = False) -> int:
    require(type(value) is int and value >= (1 if positive else 0), f"{path}: expected unsigned integer")
    return value


def _text(value: Any, path: str, *, nonempty: bool = True) -> str:
    require(isinstance(value, str) and (not nonempty or bool(value)), f"{path}: expected text")
    return value


def _hash(value: Any, path: str) -> str:
    value = _text(value, path)
    require(SHA256_RE.fullmatch(value) is not None, f"{path}: malformed SHA-256")
    return value


def _finite(value: Any, path: str = "json") -> None:
    if isinstance(value, float):
        require(math.isfinite(value), f"{path}: non-finite number")
    elif isinstance(value, dict):
        for key, item in value.items():
            require(isinstance(key, str), f"{path}: non-string object key")
            _finite(item, f"{path}.{key}")
    elif isinstance(value, list):
        for index, item in enumerate(value):
            _finite(item, f"{path}[{index}]")


def _read_json(path: Path) -> Any:
    try:
        with path.open("r", encoding="utf-8") as stream:
            value = json.load(stream)
    except (OSError, ValueError, json.JSONDecodeError) as error:
        fail(f"{path}: invalid JSON: {error}")
    _finite(value, str(path))
    return value


def _write_new(path: Path, value: Any) -> None:
    _finite(value, str(path))
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("x", encoding="utf-8") as stream:
            json.dump(value, stream, indent=2, sort_keys=True)
            stream.write("\n")
    except FileExistsError:
        fail(f"refusing to replace immutable artifact: {path}")


def _sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _meta(path: Path, *, executable: bool = False) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    stat = path.stat()
    if executable:
        require(os.access(path, os.X_OK), f"binary is not executable: {path}")
    return {"path": str(path), "bytes": stat.st_size, "sha256": _sha(path)}


def _recorded_meta(value: Any, path: str, *, executable: bool = False) -> dict[str, Any]:
    _exact(value, ("path", "bytes", "sha256"), path)
    file_path = Path(_text(value["path"], f"{path}.path"))
    require(file_path.is_absolute(), f"{path}.path: path must be absolute")
    _uint(value["bytes"], f"{path}.bytes")
    _hash(value["sha256"], f"{path}.sha256")
    actual = _meta(file_path, executable=executable)
    require(actual == value, f"{path}: retained artifact changed")
    return dict(value)


def _timestamp(value: Any, path: str) -> _datetime.datetime:
    text = _text(value, path)
    try:
        parsed = _datetime.datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError as error:
        fail(f"{path}: malformed timestamp: {error}")
    require(parsed.tzinfo is not None, f"{path}: timestamp has no timezone")
    return parsed.astimezone(_datetime.timezone.utc)


def _now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat(timespec="microseconds").replace("+00:00", "Z")


def _attempt(value: str) -> str:
    require(isinstance(value, str) and ATTEMPT_RE.fullmatch(value) is not None,
            "attempt must be a path-safe token")
    return value


def _build_key(phase: str, role: str) -> str:
    require(phase in PHASES and role in ROLES, f"unknown build coordinate {phase}/{role}")
    return f"{phase}/{role}"


def _validate_source_manifest(value: dict[str, Any], path: Path) -> None:
    """Validate recorded source hashes without requiring the disposable tree."""

    require(isinstance(value, dict) and value, f"{path}: source manifest is empty")
    for name, item in value.items():
        _text(name, f"{path}.member", nonempty=True)
        _exact(item, ("path", "bytes", "sha256"), f"{path}.{name}")
        require(Path(_text(item["path"], f"{path}.{name}.path")).is_absolute(),
                f"{path}.{name}.path: source path must be absolute")
        _uint(item["bytes"], f"{path}.{name}.bytes")
        _hash(item["sha256"], f"{path}.{name}.sha256")
    required_names = {
        "tools/perf-baseline/src/docx_managed_edit.rs",
        "tools/perf-baseline/Cargo.toml",
    }
    require(required_names.issubset(value), f"{path}: source manifest omits harness custody members")


def _validate_gate(gate_meta: dict[str, Any], source_meta: dict[str, Any], path: Path) -> None:
    _recorded_meta(gate_meta, f"{path}.gate")
    gate = _read_json(Path(gate_meta["path"]))
    require(isinstance(gate, dict), f"{path}: gate receipt is not an object")
    require(gate.get("exit_code") == 0, f"{path}: original build gate did not exit zero")
    require(gate.get("source_unchanged") is True, f"{path}: original build gate source changed")
    require(gate.get("source_manifest") == source_meta,
            f"{path}: original build gate source manifest binding differs")
    manifest = _read_json(Path(source_meta["path"]))
    cargo_member = manifest.get("tools/perf-baseline/Cargo.toml")
    require(isinstance(cargo_member, dict), f"{path}: source manifest omits Cargo.toml")
    source_root = Path(_text(cargo_member.get("path"), f"{path}.source_manifest.Cargo.toml.path")).resolve()
    # The disposable checkout may be removed after the build, so bind the
    # recorded checkout root from its retained Cargo.toml path rather than
    # requiring that source tree to remain present.
    source_root = source_root.parent.parent.parent
    cwd = Path(_text(gate.get("cwd"), f"{path}.cwd"))
    require(cwd.is_absolute() and cwd.resolve() == source_root,
            f"{path}: original build gate cwd differs from recorded source root")
    role = path.name.rsplit("/", 1)[-1]
    role = role if role in ROLES else None
    require(role is not None, f"{path}: build gate role is not identifiable")
    expected_argv = [
        "cargo", "build", "--release", "--locked", "--offline", "--manifest-path",
        "tools/perf-baseline/Cargo.toml",
    ]
    if role == "normal":
        expected_argv += ["--bin", "litchi-perf-baseline"]
    else:
        expected_argv += ["--features", "allocator-metrics", "--bin", "litchi-perf-baseline-alloc"]
    require(gate.get("argv") == expected_argv, f"{path}: original build gate argv differs")
    environment = gate.get("environment")
    require(isinstance(environment, dict), f"{path}: original build gate environment is missing")
    require(isinstance(environment.get("CARGO_TARGET_DIR"), str)
            and bool(environment["CARGO_TARGET_DIR"])
            and isinstance(environment.get("TMPDIR"), str)
            and bool(environment["TMPDIR"]),
            f"{path}: original build gate target/temp roots are missing")
    expected_environment = {
        "RUSTUP_TOOLCHAIN": "1.98.1", "CARGO_BUILD_JOBS": "4", "CARGO_INCREMENTAL": "0",
        "CARGO_PROFILE_RELEASE_DEBUG": "1", "RUSTFLAGS": "-C force-frame-pointers=yes -C force-unwind-tables=yes",
        "CARGO_TARGET_DIR": environment.get("CARGO_TARGET_DIR"), "TMPDIR": environment.get("TMPDIR"),
        "DEBUGINFOD_URLS": "", "LC_ALL": "C", "RUSTDOCFLAGS": "-Dwarnings",
    }
    require(set(environment) == set(expected_environment)
            and all(environment.get(key) == value for key, value in expected_environment.items()),
            f"{path}: original build gate environment differs")
    for name in ("stdout", "stderr"):
        _recorded_meta(gate.get(name), f"{path}.{name}")
    _uint(gate.get("started_ns"), f"{path}.started_ns", positive=True)
    _uint(gate.get("finished_ns"), f"{path}.finished_ns", positive=True)
    require(gate["finished_ns"] > gate["started_ns"], f"{path}: build gate chronology is invalid")
    _uint(gate.get("pid"), f"{path}.pid", positive=True)


def _validate_build_record(value: Any, phase: str, role: str, path: Path) -> dict[str, Any]:
    require(isinstance(value, dict), f"{path}: build record is missing")
    _exact(value, ("binary", "git_revision", "source_manifest", "gate"), str(path))
    revision = _text(value["git_revision"], f"{path}.git_revision")
    require(REVISION_RE.fullmatch(revision) is not None, f"{path}: malformed git revision")
    require(revision == EXPECTED_REVISIONS[phase],
            f"{path}: source revision is not the frozen {phase} revision")
    binary = _recorded_meta(value["binary"], f"{path}.binary", executable=True)
    source_manifest = _recorded_meta(value["source_manifest"], f"{path}.source_manifest")
    manifest = _read_json(Path(source_manifest["path"]))
    _validate_source_manifest(manifest, Path(source_manifest["path"]))
    _validate_gate(value["gate"], source_manifest, path)
    return {
        "phase": phase,
        "role": role,
        "key": _build_key(phase, role),
        "git_revision": revision,
        "binary": binary,
        "source_manifest": source_manifest,
        "gate": dict(value["gate"]),
    }


def load_builds(path: Path | None = None) -> dict[str, dict[str, Any]]:
    """Load and authenticate all four root-owned build records."""

    path = Path(path or BUILDS_FILE)
    value = _read_json(path)
    _exact(value, ("schema", "builds"), str(path))
    require(value["schema"] == BUILDS_SCHEMA, f"{path}: build schema differs")
    builds = value["builds"]
    require(isinstance(builds, dict), f"{path}.builds: expected object")
    expected_keys = {_build_key(phase, role) for phase in PHASES for role in ROLES}
    require(set(builds) == expected_keys,
            f"{path}.builds: inventory differs (expected {sorted(expected_keys)})")
    result: dict[str, dict[str, Any]] = {}
    for phase in PHASES:
        for role in ROLES:
            key = _build_key(phase, role)
            result[key] = _validate_build_record(builds[key], phase, role, path / key)
    for phase in PHASES:
        normal = result[_build_key(phase, "normal")]
        allocator = result[_build_key(phase, "allocator")]
        require(normal["git_revision"] == allocator["git_revision"], f"{phase}: role revisions differ")
        require(normal["source_manifest"] == allocator["source_manifest"],
                f"{phase}: role source manifests differ")
    before_manifest = _read_json(Path(result["before/normal"]["source_manifest"]["path"]))
    after_manifest = _read_json(Path(result["after/normal"]["source_manifest"]["path"]))
    for name in HARNESS_CUSTODY_FILES:
        require(name in before_manifest and name in after_manifest,
                f"source manifests omit shared harness file {name}")
        require(before_manifest[name]["bytes"] == after_manifest[name]["bytes"]
                and before_manifest[name]["sha256"] == after_manifest[name]["sha256"],
                f"shared harness file changed between before/after: {name}")
    return result


def _build_binding(build: dict[str, Any]) -> dict[str, Any]:
    return {
        "key": build["key"],
        "phase": build["phase"],
        "role": build["role"],
        "git_revision": build["git_revision"],
        "binary": build["binary"],
        "source_manifest": build["source_manifest"],
        "gate": build["gate"],
    }


def _builds_binding(builds: dict[str, dict[str, Any]]) -> dict[str, Any]:
    return {key: _build_binding(builds[key]) for key in sorted(builds)}


def _phase_arms(api: str) -> tuple[str, ...]:
    require(api in APIS, f"unknown API {api}")
    return ("file-warm", "short") if api == MANAGED else ARMS


def formal_inventory() -> list[dict[str, Any]]:
    """Return the 32-child formal inventory in its prescribed ABBA order."""

    result: list[dict[str, Any]] = []
    ordinal = 0
    for repeat in REPEATS:
        phases = PHASES if repeat == 1 else tuple(reversed(PHASES))
        roles = ROLES if repeat == 1 else tuple(reversed(ROLES))
        for phase in phases:
            phase_apis = (UNMANAGED,) if phase == "before" else (UNMANAGED, MANAGED)
            apis = phase_apis if repeat == 1 else tuple(reversed(phase_apis))
            for api in apis:
                arms = _phase_arms(api)
                arms = arms if repeat == 1 else tuple(reversed(arms))
                for role in roles:
                    for arm in arms:
                        label = f"r{repeat}-{phase}-{api}-{role}-{arm}"
                        result.append({
                            "ordinal": ordinal,
                            "repeat": repeat,
                            "phase": phase,
                            "api": api,
                            "role": role,
                            "arm": arm,
                            "provider": {"owned": "owned", "file-warm": "file", "short": "short"}[arm],
                            "samples": FORMAL_SAMPLES,
                            "warmups": FORMAL_WARMUPS,
                            "label": label,
                        })
                        ordinal += 1
    require(len(result) == 32, "formal inventory must contain exactly 32 children")
    require(sum(1 for item in result if item["phase"] == "before") == 12,
            "formal inventory before count differs")
    require(sum(1 for item in result if item["phase"] == "after" and item["api"] == UNMANAGED) == 12,
            "formal inventory unmanaged after count differs")
    require(sum(1 for item in result if item["phase"] == "after" and item["api"] == MANAGED) == 8,
            "formal inventory managed after count differs")
    return result


def _env_snapshot() -> dict[str, str]:
    return {key: os.environ[key] for key in ENV_KEYS if key in os.environ}


def _custody_binding(path: Path, name: str) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"{name} custody file is missing: {path}")
    return _meta(path)


def _validate_machine_and_provenance(protocol: dict[str, Any]) -> None:
    machine = protocol["machine"]
    provenance = protocol["provenance"]
    _exact(machine, ("path", "bytes", "sha256"), "protocol.machine")
    _exact(provenance, ("path", "bytes", "sha256"), "protocol.provenance")
    _recorded_meta(machine, "protocol.machine")
    _recorded_meta(provenance, "protocol.provenance")
    machine_value = _read_json(Path(machine["path"]))
    require(isinstance(machine_value, dict)
            and CPU in machine_value.get("allowed_cpus", []),
            "machine record does not bind the pinned CPU")
    provenance_value = _read_json(Path(provenance["path"]))
    require(isinstance(provenance_value, dict), "provenance record is not an object")
    _exact(provenance_value, ("schema", "inputs", "production_change", "historical_evidence"),
           "provenance")
    require(provenance_value["schema"] == "docx-phase-provenance-v1",
            "provenance schema differs")
    inputs = provenance_value["inputs"]
    require(isinstance(inputs, dict) and inputs, "provenance inputs are missing")
    for name, item in inputs.items():
        _recorded_meta(item, f"provenance.inputs.{name}")


def _validator_binding() -> dict[str, Any]:
    for path in (OLD_MEASURE, OLD_SUPPORT, OLD_SEAL):
        require(path.is_file(), f"sealed 0495 validator input is missing: {path}")
    return {
        "measure": {"path": str(OLD_MEASURE), "sha256": _sha(OLD_MEASURE)},
        "support": {"path": str(OLD_SUPPORT), "sha256": _sha(OLD_SUPPORT)},
        "seal": {"path": str(OLD_SEAL), "sha256": _sha(OLD_SEAL)},
    }


def _test_binding() -> dict[str, Any]:
    return {"path": str(Path(__file__).resolve().parent / "test_capture.py"),
            "sha256": _sha(Path(__file__).resolve().parent / "test_capture.py")}


def protocol_value(builds: dict[str, dict[str, Any]], builds_path: Path = BUILDS_FILE) -> dict[str, Any]:
    inventory = formal_inventory()
    return {
        "schema": PROTOCOL_SCHEMA,
        "version": VERSION,
        "case": CASE,
        "cpu": CPU,
        "cpu_lock": str(CPU_LOCK),
        "timeout_seconds": DEFAULT_TIMEOUT_SECONDS,
        "term_grace_seconds": TERM_GRACE_SECONDS,
        "samples": FORMAL_SAMPLES,
        "warmups": FORMAL_WARMUPS,
        "expected_children": len(inventory),
        "expected_samples": len(inventory) * FORMAL_SAMPLES,
        "formal_runs": inventory,
        "builds_input": {"path": str(Path(builds_path).resolve()), "sha256": _sha(builds_path)},
        "builds": _builds_binding(builds),
        "driver": {"path": str(Path(__file__).resolve()), "sha256": _sha(Path(__file__))},
        "validator": _validator_binding(),
        "tests": _test_binding(),
        "machine": _custody_binding(MACHINE_FILE, "machine"),
        "provenance": _custody_binding(PROVENANCE_FILE, "provenance"),
        "environment": _env_snapshot(),
        "temporary_root": str(TEMP),
        "capture_root": str(ROOT / "captures"),
        "report_contract": {
            "canonical_validator": "change-0495/measure.py:validate_report",
            "phase_field": "phase_diagnostics",
            "phase_schema": PHASE_SCHEMA,
            "phase_fields": list(PHASE_FIELDS),
            "projection": "remove phase_diagnostics from every row in memory before 0495 validation",
        },
        "claims": {
            "scope": "descriptive opt-in phase diagnostics with paired before/after unmanaged review",
            "no_historical_flags_resolved": True,
            "managed_after_is_capability_only": True,
        },
    }


def create_freeze(builds_path: Path = BUILDS_FILE) -> Path:
    """Authenticate root build inputs and write the immutable 0496 protocol."""

    builds_path = Path(builds_path).resolve()
    builds = load_builds(builds_path)
    value = protocol_value(builds, builds_path)
    _write_new(PROTOCOL_FILE, value)
    return PROTOCOL_FILE


def _load_protocol() -> tuple[dict[str, Any], str, dict[str, dict[str, Any]]]:
    require(PROTOCOL_FILE.is_file(), f"frozen protocol is missing: {PROTOCOL_FILE}")
    protocol = _read_json(PROTOCOL_FILE)
    _exact(protocol, (
        "schema", "version", "case", "cpu", "cpu_lock", "timeout_seconds",
        "term_grace_seconds", "samples", "warmups", "expected_children",
        "expected_samples", "formal_runs", "builds_input", "builds", "driver",
        "validator", "tests", "machine", "provenance", "environment", "temporary_root",
        "capture_root", "report_contract", "claims",
    ), "protocol")
    require(protocol["schema"] == PROTOCOL_SCHEMA and protocol["version"] == VERSION,
            "protocol schema/version differs")
    require(protocol["case"] == CASE and protocol["cpu"] == CPU
            and protocol["cpu_lock"] == str(CPU_LOCK)
            and protocol["timeout_seconds"] == DEFAULT_TIMEOUT_SECONDS
            and protocol["term_grace_seconds"] == TERM_GRACE_SECONDS
            and protocol["samples"] == FORMAL_SAMPLES and protocol["warmups"] == FORMAL_WARMUPS,
            "protocol constants differ")
    inventory = formal_inventory()
    require(protocol["formal_runs"] == inventory and protocol["expected_children"] == len(inventory)
            and protocol["expected_samples"] == len(inventory) * FORMAL_SAMPLES,
            "protocol inventory differs")
    builds_input = protocol["builds_input"]
    _exact(builds_input, ("path", "sha256"), "protocol.builds_input")
    builds_path = Path(_text(builds_input["path"], "protocol.builds_input.path"))
    require(_sha(builds_path) == _hash(builds_input["sha256"], "protocol.builds_input.sha256"),
            "builds.json changed after freeze")
    builds = load_builds(builds_path)
    require(_builds_binding(builds) == protocol["builds"], "build records changed after freeze")
    _exact(protocol["driver"], ("path", "sha256"), "protocol.driver")
    require(Path(protocol["driver"]["path"]).resolve() == Path(__file__).resolve()
            and _sha(Path(__file__)) == protocol["driver"]["sha256"],
            "capture driver changed after freeze")
    _exact(protocol["validator"], ("measure", "support", "seal"), "protocol.validator")
    for name, expected_path in (("measure", OLD_MEASURE), ("support", OLD_SUPPORT), ("seal", OLD_SEAL)):
        _exact(protocol["validator"][name], ("path", "sha256"), f"protocol.validator.{name}")
        require(Path(protocol["validator"][name]["path"]).resolve() == expected_path
                and _sha(expected_path) == protocol["validator"][name]["sha256"],
                f"sealed 0495 validator {name} changed after freeze")
    _exact(protocol["tests"], ("path", "sha256"), "protocol.tests")
    require(Path(protocol["tests"]["path"]).resolve() == Path(__file__).resolve().parent / "test_capture.py"
            and _sha(Path(__file__).resolve().parent / "test_capture.py") == protocol["tests"]["sha256"],
            "0496 focused tests changed after freeze")
    require(protocol["report_contract"]["phase_field"] == "phase_diagnostics"
            and protocol["report_contract"]["phase_schema"] == PHASE_SCHEMA
            and protocol["report_contract"]["phase_fields"] == list(PHASE_FIELDS),
            "phase report contract changed")
    require(protocol["temporary_root"] == str(TEMP)
            and protocol["capture_root"] == str(ROOT / "captures"),
            "protocol roots differ")
    _validate_machine_and_provenance(protocol)
    return protocol, _sha(PROTOCOL_FILE), builds


def _arm_config(arm: str) -> dict[str, Any]:
    if arm == "owned":
        return {"provider": "owned", "trace_ranges": False}
    if arm == "file-warm":
        return {"provider": "file", "trace_ranges": False}
    if arm == "short":
        return {"provider": "short", "trace_ranges": True, "max_range": SHORT_RANGE_BYTES}
    fail(f"unknown arm {arm}")


def _command(spec: dict[str, Any], build: dict[str, Any], report: Path, resource: Path) -> list[str]:
    arm = _arm_config(spec["arm"])
    command = [
        "/usr/bin/time", "-v", "-o", str(resource),
        "/usr/bin/taskset", "-c", str(CPU), str(build["binary"]["path"]),
        "docx-managed-edit", "--edit-api", spec["api"], "--provider", arm["provider"],
        "--phase-diagnostics",
    ]
    if "max_range" in arm:
        command += ["--max-range", str(arm["max_range"])]
    if arm.get("trace_ranges"):
        command += ["--trace-ranges"]
    command += [
        "--samples", str(spec["samples"]), "--warmup", str(spec["warmups"]),
        "--source-revision", build["git_revision"], "--output", str(report),
    ]
    return command


def _run_root(attempt: str, label: str) -> Path:
    return TEMP / "runs" / attempt / label


def _capture_root(attempt: str, label: str) -> Path:
    return ROOT / "captures" / attempt / label


def _cpu_lock(action: Callable[[], Any]) -> Any:
    CPU_LOCK.parent.mkdir(parents=True, exist_ok=True)
    with CPU_LOCK.open("a+") as stream:
        fcntl.flock(stream.fileno(), fcntl.LOCK_EX)
        try:
            return action()
        finally:
            fcntl.flock(stream.fileno(), fcntl.LOCK_UN)


def _kill_group(process: subprocess.Popen[bytes]) -> str | None:
    if process.poll() is not None:
        return None
    try:
        os.killpg(process.pid, signal.SIGTERM)
    except ProcessLookupError:
        return None
    try:
        process.communicate(timeout=TERM_GRACE_SECONDS)
    except subprocess.TimeoutExpired:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
        process.communicate()
        return "SIGKILL"
    return "SIGTERM"


def _cleanup_private(root: Path) -> dict[str, Any]:
    expected_parent = TEMP / "runs"
    require(root.parent.parent == expected_parent, f"private scratch escaped 0496 root: {root}")
    require(root.is_dir() and not root.is_symlink(), f"private scratch is not a directory: {root}")
    removed: list[str] = []
    remaining: list[str] = []
    # Every child, including file-provider staging directories, is owned by
    # this one run.  Refuse to follow a symlink and remove only this exact root.
    for child in list(root.iterdir()):
        require(not child.is_symlink(), f"private scratch contains symlink: {child}")
    shutil.rmtree(root)
    removed.append(str(root))
    attempt_parent = root.parent
    if attempt_parent.exists() and attempt_parent.is_dir() and not attempt_parent.is_symlink() and not list(attempt_parent.iterdir()):
        attempt_parent.rmdir()
        removed.append(str(attempt_parent))
    runs_parent = expected_parent
    if runs_parent.exists() and runs_parent.is_dir() and not runs_parent.is_symlink() and not list(runs_parent.iterdir()):
        runs_parent.rmdir()
        removed.append(str(runs_parent))
    if root.exists():
        remaining.append(str(root))
    return {"schema": "docx-phase-private-cleanup-v1", "status": "pass" if not remaining else "failed",
            "root": str(root), "removed": removed, "remaining": remaining}


def _resource(path: Path) -> dict[str, Any]:
    require(path.is_file(), f"missing GNU time resource receipt: {path}")
    raw = path.read_bytes()
    require(raw, f"empty GNU time resource receipt: {path}")
    result: dict[str, Any] = {"raw_bytes": len(raw), "raw_sha256": hashlib.sha256(raw).hexdigest()}
    for line in raw.decode("utf-8", errors="replace").splitlines():
        if ":" not in line:
            continue
        key, value = (part.strip() for part in line.split(":", 1))
        key = key.lower().replace(" ", "_")
        first = value.split()[0] if value.split() else ""
        if first.isdigit():
            result[key] = int(first)
        else:
            try:
                result[key] = float(first)
            except ValueError:
                result[key] = value
    _uint(result.get("maximum_resident_set_size_(kbytes)"), f"{path}: maximum RSS", positive=True)
    return result


def _phase_evidence(value: Any, row: dict[str, Any], path: str) -> dict[str, Any]:
    _exact(value, PHASE_FIELDS, path)
    require(value["schema"] == PHASE_SCHEMA, f"{path}.schema: phase schema differs")
    require(value["timing_scope"] == PHASE_TIMING_SCOPE, f"{path}.timing_scope: scope differs")
    for name in PHASE_DURATION_FIELDS + ("phase_sum_ns", "lifecycle_residual_ns"):
        _uint(value[name], f"{path}.{name}")
    expected_sum = sum(value[name] for name in PHASE_DURATION_FIELDS)
    require(value["phase_sum_ns"] == expected_sum, f"{path}: phase sum does not conserve intervals")
    require(value["phase_sum_ns"] + value["lifecycle_residual_ns"] == row["latency_ns"],
            f"{path}: phases do not conserve lifecycle latency")
    require(value["residual_scope"] == PHASE_RESIDUAL_SCOPE, f"{path}.residual_scope: scope differs")
    require(value["instrumentation_overhead_ns"] is None,
            f"{path}.instrumentation_overhead_ns: unsupported measurement present")
    require(value["instrumentation_scope"] == PHASE_INSTRUMENTATION_SCOPE,
            f"{path}.instrumentation_scope: scope differs")
    require(value["allocation_scope"] == PHASE_ALLOCATION_SCOPE,
            f"{path}.allocation_scope: scope differs")
    return dict(value)


def _contains_nested_phase_field(value: Any, path: str, *, top: bool = True) -> None:
    if isinstance(value, dict):
        for key, item in value.items():
            if key == "phase_diagnostics" and not top:
                fail(f"{path}.{key}: phase field is only allowed on a report row")
            _contains_nested_phase_field(item, f"{path}.{key}", top=False)
    elif isinstance(value, list):
        for index, item in enumerate(value):
            _contains_nested_phase_field(item, f"{path}[{index}]", top=False)


def _load_0495() -> Any:
    require(OLD_MEASURE.is_file(), f"0495 validator is missing: {OLD_MEASURE}")
    module_name = "_sealed_change0495_measure"
    loaded = sys.modules.get(module_name)
    if loaded is not None:
        return loaded
    old_path = str(OLD_ROOT)
    inserted = old_path not in sys.path
    if inserted:
        sys.path.insert(0, old_path)
    try:
        spec = importlib.util.spec_from_file_location(module_name, OLD_MEASURE)
        require(spec is not None and spec.loader is not None, "cannot load sealed 0495 validator")
        module = importlib.util.module_from_spec(spec)
        sys.modules[module_name] = module
        spec.loader.exec_module(module)
        return module
    except Exception:
        sys.modules.pop(module_name, None)
        raise
    finally:
        if inserted:
            sys.path.remove(old_path)


def _canonical_fields() -> tuple[str, ...]:
    try:
        fields = tuple(_load_0495().REPORT_FIELDS)
    except (ImportError, AttributeError, OSError):
        fields = FALLBACK_REPORT_FIELDS
    return fields


def _phase_projection(value: dict[str, Any], path: str) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    canonical = set(_canonical_fields())
    require(set(value) == canonical, f"{path}: top-level report fields differ from 0495")
    rows = value.get("rows")
    require(isinstance(rows, list) and len(rows) == FORMAL_SAMPLES, f"{path}.rows: sample count differs")
    projected_rows: list[dict[str, Any]] = []
    phase_rows: list[dict[str, Any]] = []
    for index, row in enumerate(rows):
        require(isinstance(row, dict), f"{path}.rows[{index}]: expected object")
        require("phase_diagnostics" in row, f"{path}.rows[{index}]: phase diagnostics missing")
        phase_rows.append(_phase_evidence(row["phase_diagnostics"], row, f"{path}.rows[{index}].phase_diagnostics"))
        stripped = dict(row)
        del stripped["phase_diagnostics"]
        projected_rows.append(stripped)
    projection = copy.deepcopy(value)
    projection["rows"] = projected_rows
    # The recursive check is intentionally run on the stripped canonical
    # projection: a phase field hidden in budget/oracle/corpus must still be
    # rejected even when a future canonical validator is loosened.
    _contains_nested_phase_field(projection, path)
    return projection, phase_rows


def _validate_report(path: Path, spec: dict[str, Any], build: dict[str, Any]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    raw = _read_json(path)
    require(isinstance(raw, dict), f"{path}: report is not an object")
    projection, phase_rows = _phase_projection(raw, str(path))
    # 0495's validator accepts a path and intentionally has exact row fields.
    # Write only the explicit projection to a private temporary directory; the
    # retained report remains byte-for-byte untouched.
    projection_root = TEMP / "projections"
    projection_root.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="report-", dir=projection_root) as scratch:
        projected_path = Path(scratch) / "report.json"
        projected_path.write_text(json.dumps(projection, sort_keys=True), encoding="utf-8")
        validator = _load_0495()
        canonical = validator.validate_report(
            projected_path,
            role=spec["role"], api=spec["api"], arm_name=spec["arm"],
            samples=spec["samples"], warmups=spec["warmups"],
            source_revision=build["git_revision"],
            binary_sha256=build["binary"]["sha256"], binary_bytes=build["binary"]["bytes"],
        )
    return canonical, phase_rows


def _artifact_meta(path: Path) -> dict[str, Any]:
    return _meta(path)


def _expected_environment(protocol: dict[str, Any], tmpdir: Path) -> dict[str, str]:
    environment = dict(protocol["environment"])
    environment["TMPDIR"] = str(tmpdir)
    return environment


def _validate_terminal(started: dict[str, Any], terminal: dict[str, Any], spec: dict[str, Any],
                       build: dict[str, Any], protocol: dict[str, Any], protocol_hash: str,
                       directory: Path, *, require_pass: bool = True) -> None:
    started_fields = {
        "schema", "version", "status", "attempt", "run", "protocol", "build", "binary",
        "source_manifest", "argv", "cwd", "environment", "machine", "started_utc",
        "timeout_seconds", "tmpdir", "file_root", "driver",
    }
    terminal_fields = started_fields | {
        "exit_code", "timed_out", "termination", "launch_error", "validation_error",
        "finished_utc", "artifacts", "missing_artifacts", "cleanup", "started_artifact",
    }
    _exact(started, started_fields, f"{directory}/started.json")
    _exact(terminal, terminal_fields, f"{directory}/terminal.json")
    require(started["schema"] == CAPTURE_SCHEMA and started["version"] == VERSION
            and started["status"] == "running", f"{directory}: started identity differs")
    require(terminal["schema"] == TERMINAL_SCHEMA and terminal["version"] == VERSION,
            f"{directory}: terminal identity differs")
    # ``schema`` changes from the capture receipt to the terminal receipt;
    # every other started field is immutable, including the environment and
    # machine binding.
    for key in started_fields - {"status", "schema"}:
        require(terminal[key] == started[key], f"{directory}: terminal {key} changed")
    require(started["attempt"] == spec["attempt"]
            and started["run"] == {key: spec[key] for key in
                                    ("ordinal", "repeat", "phase", "api", "role", "arm", "provider", "samples", "warmups", "label")},
            f"{directory}: run specification differs")
    require(started["protocol"] == {"path": "protocol.json", "sha256": protocol_hash},
            f"{directory}: protocol binding differs")
    require(started["build"] == _build_binding(build) and started["binary"] == build["binary"]
            and started["source_manifest"] == build["source_manifest"],
            f"{directory}: build/source binding differs")
    require(started["cwd"] == str(REPO), f"{directory}: cwd differs")
    require(started["machine"] == protocol["machine"], f"{directory}: machine binding differs")
    require(started["driver"] == protocol["driver"], f"{directory}: driver binding differs")
    tmpdir = Path(_text(started["tmpdir"], f"{directory}.tmpdir"))
    expected_private = _run_root(spec["attempt"], spec["label"])
    expected_tmpdir = expected_private / ("file" if spec["arm"] == "file-warm" else "tmp")
    require(tmpdir == expected_tmpdir, f"{directory}: private TMPDIR differs")
    file_root = started["file_root"]
    if spec["arm"] == "file-warm":
        require(file_root == str(expected_private / "file"),
                f"{directory}: file staging root differs")
    else:
        require(file_root is None, f"{directory}: non-file arm has file root")
    require(started["environment"] == _expected_environment(protocol, tmpdir)
            and terminal["environment"] == started["environment"],
            f"{directory}: child environment differs")
    require(started["argv"] == terminal["argv"] == _command(spec, build, directory / "report.json", directory / "resource.txt"),
            f"{directory}: child argv differs")
    _timestamp(started["started_utc"], f"{directory}.started_utc")
    finished = _timestamp(terminal["finished_utc"], f"{directory}.finished_utc")
    started_at = _timestamp(started["started_utc"], f"{directory}.started_utc")
    require(finished > started_at, f"{directory}: chronology invalid")
    require(type(terminal["timed_out"]) is bool, f"{directory}: timed_out is not boolean")
    require(isinstance(terminal["missing_artifacts"], list)
            and all(isinstance(item, str) for item in terminal["missing_artifacts"]),
            f"{directory}: missing artifact list malformed")
    cleanup = terminal["cleanup"]
    _exact(cleanup, ("schema", "status", "root", "removed", "remaining"), f"{directory}.cleanup")
    require(cleanup["schema"] == "docx-phase-private-cleanup-v1"
            and cleanup["root"] == str(_run_root(spec["attempt"], spec["label"]))
            and cleanup["status"] == "pass" and cleanup["remaining"] == [],
            f"{directory}: private cleanup failed")
    require(terminal["started_artifact"] == _artifact_meta(directory / "started.json"),
            f"{directory}: started artifact hash differs")
    artifacts = terminal["artifacts"]
    _exact(artifacts, CAPTURE_ARTIFACTS, f"{directory}.artifacts")
    for name in CAPTURE_ARTIFACTS:
        _exact(artifacts[name], ("path", "bytes", "sha256"), f"{directory}.artifacts.{name}")
        require(Path(artifacts[name]["path"]) == directory / name,
                f"{directory}: artifact escaped capture directory")
        actual = _artifact_meta(directory / name)
        require(actual == artifacts[name], f"{directory}: artifact changed: {name}")
    require(_read_json(directory / "replay-cleanup.json") == cleanup,
            f"{directory}: cleanup receipt differs")
    require(set(item.name for item in directory.iterdir()) == set(("started.json", "terminal.json", *CAPTURE_ARTIFACTS)),
            f"{directory}: retained artifact inventory differs")
    if require_pass:
        _require_success_terminal(terminal, str(directory))


def _require_success_terminal(terminal: dict[str, Any], path: str = "terminal") -> None:
    """Reject typed process failures before any report/statistical analysis."""

    require(terminal.get("status") == "pass" and terminal.get("exit_code") == 0
            and terminal.get("timed_out") is False and terminal.get("termination") is None
            and terminal.get("launch_error") is None and terminal.get("validation_error") is None
            and terminal.get("missing_artifacts") == [], f"{path}: terminal did not pass")


def _launch(spec: dict[str, Any], build: dict[str, Any], protocol: dict[str, Any],
            protocol_hash: str, timeout_seconds: int) -> Path:
    require(timeout_seconds == DEFAULT_TIMEOUT_SECONDS, "formal timeout differs from frozen 180-second contract")
    attempt = _attempt(spec["attempt"])
    directory = _capture_root(attempt, spec["label"])
    require(not directory.exists(), f"refusing to replace immutable capture: {directory}")
    directory.parent.mkdir(parents=True, exist_ok=True)
    directory.mkdir()
    private = _run_root(attempt, spec["label"])
    require(not private.exists(), f"refusing to replace private scratch: {private}")
    private.mkdir(parents=True)
    tmpdir = private / "tmp"
    tmpdir.mkdir()
    file_root = private / "file" if spec["arm"] == "file-warm" else None
    if file_root is not None:
        file_root.mkdir()
    report = directory / "report.json"
    resource = directory / "resource.txt"
    stdout = directory / "stdout.txt"
    stderr = directory / "stderr.txt"
    stdout.touch()
    stderr.touch()
    argv = _command(spec, build, report, resource)
    environment_tmpdir = file_root if file_root is not None else tmpdir
    environment = _expected_environment(protocol, environment_tmpdir)
    started = {
        "schema": CAPTURE_SCHEMA,
        "version": VERSION,
        "status": "running",
        "attempt": attempt,
        "run": {key: spec[key] for key in
                ("ordinal", "repeat", "phase", "api", "role", "arm", "provider", "samples", "warmups", "label")},
        "protocol": {"path": "protocol.json", "sha256": protocol_hash},
        "build": _build_binding(build),
        "binary": build["binary"],
        "source_manifest": build["source_manifest"],
        "argv": argv,
        "cwd": str(REPO),
        "environment": environment,
        "machine": protocol["machine"],
        "started_utc": _now(),
        "timeout_seconds": timeout_seconds,
        "tmpdir": str(environment_tmpdir),
        "file_root": str(file_root) if file_root is not None else None,
        "driver": protocol["driver"],
    }
    _write_new(directory / "started.json", started)
    process: subprocess.Popen[bytes] | None = None
    exit_code: int | None = None
    timed_out = False
    termination: str | None = None
    launch_error: str | None = None
    validation_error: str | None = None
    try:
        with stdout.open("wb") as out, stderr.open("wb") as err:
            process = subprocess.Popen(
                argv, cwd=REPO, env=environment, stdin=subprocess.DEVNULL,
                stdout=out, stderr=err, start_new_session=True,
            )
            try:
                process.communicate(timeout=timeout_seconds)
            except subprocess.TimeoutExpired:
                timed_out = True
                termination = _kill_group(process)
            exit_code = process.returncode
    except (OSError, subprocess.SubprocessError) as error:
        launch_error = f"{type(error).__name__}: {error}"
        if process is not None and process.poll() is None:
            termination = _kill_group(process)
            exit_code = process.returncode
    if exit_code == 0 and not timed_out and launch_error is None:
        try:
            _validate_report(report, spec, build)
        except (PhaseDiagnosticError, OSError, ValueError, ImportError) as error:
            validation_error = f"{type(error).__name__}: {error}"
    try:
        cleanup = _cleanup_private(private)
    except (PhaseDiagnosticError, OSError) as error:
        cleanup = {"schema": "docx-phase-private-cleanup-v1", "status": "failed",
                   "root": str(private), "removed": [], "remaining": [str(error)]}
    _write_new(directory / "replay-cleanup.json", cleanup)
    expected_artifacts = [stdout, stderr, resource, report, directory / "replay-cleanup.json"]
    artifacts = {item.name: _meta(item) for item in expected_artifacts if item.is_file()}
    missing = [item.name for item in expected_artifacts if not item.is_file()]
    passed = (exit_code == 0 and timed_out is False and termination is None
              and launch_error is None and validation_error is None
              and not missing and cleanup["status"] == "pass")
    terminal = dict(started)
    terminal.update({
        "schema": TERMINAL_SCHEMA,
        "status": "pass" if passed else "failed",
        "exit_code": exit_code,
        "timed_out": timed_out,
        "termination": termination,
        "launch_error": launch_error,
        "validation_error": validation_error,
        "finished_utc": _now(),
        "artifacts": artifacts,
        "missing_artifacts": missing,
        "cleanup": cleanup,
        "started_artifact": _artifact_meta(directory / "started.json"),
    })
    _write_new(directory / "terminal.json", terminal)
    if not passed:
        fail(f"{spec['label']} failed; retained terminal receipt: {directory / 'terminal.json'}")
    return directory / "terminal.json"


def capture_one(attempt: str, spec: dict[str, Any], *, timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS) -> Path:
    protocol, protocol_hash, builds = _load_protocol()
    spec = dict(spec, attempt=_attempt(attempt))
    build = builds[_build_key(spec["phase"], spec["role"])]
    return _cpu_lock(lambda: _launch(spec, build, protocol, protocol_hash, timeout_seconds))


def capture_all(attempt: str, *, timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS) -> None:
    protocol, protocol_hash, builds = _load_protocol()
    attempt = _attempt(attempt)
    inventory = [dict(item, attempt=attempt) for item in protocol["formal_runs"]]

    def run() -> None:
        for spec in inventory:
            _launch(spec, builds[_build_key(spec["phase"], spec["role"])], protocol, protocol_hash, timeout_seconds)

    # Hold the shared lock across the whole 32-child capture, including all
    # reversed blocks, so diagnostics cannot interleave with another lane.
    _cpu_lock(run)


def _percentiles(values: list[int | float]) -> dict[str, Any]:
    require(values, "cannot summarize an empty vector")
    ordered = sorted(values)
    n = len(ordered)
    p50 = (ordered[n // 2 - 1] + ordered[n // 2]) / 2 if n % 2 == 0 else ordered[n // 2]
    return {
        "n": n, "min": ordered[0], "max": ordered[-1], "mean": statistics.fmean(values),
        "p50": p50, "p95": ordered[math.ceil(n * 0.95) - 1],
        "p99": ordered[math.ceil(n * 0.99) - 1],
    }


def paired_bootstrap(values: list[int | float], *, repetitions: int = BOOTSTRAP_REPETITIONS,
                     seed: int = BOOTSTRAP_SEED) -> dict[str, Any]:
    """Bootstrap paired block deltas, never individual samples."""

    require(values, "cannot bootstrap an empty paired vector")
    require(type(repetitions) is int and repetitions > 0, "bootstrap repetitions must be positive")
    require(type(seed) is int and seed >= 0, "bootstrap seed must be nonnegative")
    rng = seed & 0xFFFFFFFFFFFFFFFF
    medians: list[float] = []
    for _ in range(repetitions):
        sample: list[int | float] = []
        for _ in values:
            rng = (rng * 6364136223846793005 + 1442695040888963407) & 0xFFFFFFFFFFFFFFFF
            sample.append(values[(rng >> 32) % len(values)])
        sample.sort()
        at = len(sample) // 2
        medians.append((sample[at - 1] + sample[at]) / 2 if len(sample) % 2 == 0 else sample[at])
    medians.sort()
    median = _percentiles(values)["p50"]
    low_index = max(0, int(repetitions * 0.025))
    high_index = min(repetitions - 1, int(repetitions * 0.975) - 1)
    return {
        "method": "paired_block_bootstrap_median",
        "seed": seed,
        "resamples": repetitions,
        "confidence": 0.95,
        "blocks": len(values),
        "median": median,
        "ci_low": medians[low_index],
        "ci_high": medians[high_index],
    }


def _within_child_bootstrap(values: list[int | float], *, repetitions: int = BOOTSTRAP_REPETITIONS,
                            seed: int = BOOTSTRAP_SEED) -> dict[str, Any]:
    """Descriptive median CI from one child's 30 samples.

    This is intentionally separate from ``paired_bootstrap``.  Samples from
    distinct child processes are not paired observations; the cross-phase
    comparison resamples two repeat blocks, while this helper only describes
    the within-child sample vector.
    """

    require(values, "cannot bootstrap an empty child vector")
    require(type(repetitions) is int and repetitions > 0, "bootstrap repetitions must be positive")
    rng = seed & 0xFFFFFFFFFFFFFFFF
    medians: list[float] = []
    for _ in range(repetitions):
        sample: list[int | float] = []
        for _ in values:
            rng = (rng * 6364136223846793005 + 1442695040888963407) & 0xFFFFFFFFFFFFFFFF
            sample.append(values[(rng >> 32) % len(values)])
        sample.sort()
        at = len(sample) // 2
        medians.append((sample[at - 1] + sample[at]) / 2 if len(sample) % 2 == 0 else sample[at])
    medians.sort()
    low_index = max(0, int(repetitions * 0.025))
    high_index = min(repetitions - 1, int(repetitions * 0.975) - 1)
    return {
        "method": "independent_within_child_bootstrap_median",
        "seed": seed,
        "resamples": repetitions,
        "confidence": 0.95,
        "samples": len(values),
        "median": _percentiles(values)["p50"],
        "ci_low": medians[low_index],
        "ci_high": medians[high_index],
    }


def _phase_vectors(report: dict[str, Any]) -> dict[str, list[int]]:
    vectors = {name: [] for name in PHASE_DURATION_FIELDS + ("phase_sum_ns", "lifecycle_residual_ns")}
    for row in report["rows"]:
        phase = row["phase_diagnostics"]
        for name in vectors:
            vectors[name].append(phase[name])
    return vectors


def _allocation_metrics(entry: dict[str, Any], rows: list[dict[str, Any]]) -> dict[str, Any]:
    """Summarize the one full-lifecycle allocator region when available."""

    role = entry.get("spec", {}).get("role", "normal")
    if role == "normal":
        # The canonical 0495 validator requires ``allocation`` to be null for
        # normal binaries.  Keep that distinction visible in analysis rather
        # than turning unavailable counters into zeroes.
        require(all(row.get("allocation") is None for row in rows),
                f"{entry.get('directory', 'entry')}: normal row contains allocator metrics")
        return {
            "availability": "unavailable",
            "scope": SAMPLE_ALLOCATION_SCOPE,
            "vectors": None,
            "percentiles": None,
            "median_ci": None,
            "derived_fields": {
                "allocation_peak_increment_bytes":
                    "region_peak_live_bytes - live_bytes_before (unavailable for normal binaries)",
            },
        }

    require(role == "allocator", f"{entry.get('directory', 'entry')}: unknown allocator role")
    vectors = {name: [] for name in ALLOCATION_VECTOR_FIELDS}
    for index, row in enumerate(rows):
        allocation = row.get("allocation")
        path = f"{entry.get('directory', 'entry')}.rows[{index}].allocation"
        _exact(allocation, ("status", "scope", *ALLOCATION_RAW_FIELDS), path)
        require(allocation["status"] == "measured" and allocation["scope"] == SAMPLE_ALLOCATION_SCOPE,
                f"{path}: full-lifecycle allocator sample is unavailable")
        for name in ALLOCATION_RAW_FIELDS:
            vectors[name].append(_uint(allocation[name], f"{path}.{name}"))
        increment = allocation["region_peak_live_bytes"] - allocation["live_bytes_before"]
        require(increment >= 0, f"{path}: region peak increment is negative")
        vectors["allocation_peak_increment_bytes"].append(increment)
    return {
        "availability": "measured",
        "scope": SAMPLE_ALLOCATION_SCOPE,
        "vectors": vectors,
        "percentiles": {name: _percentiles(values) for name, values in vectors.items()},
        "median_ci": {name: _within_child_bootstrap(values) for name, values in vectors.items()},
        "derived_fields": {
            "allocation_peak_increment_bytes":
                "region_peak_live_bytes - live_bytes_before for each sample",
        },
    }


def _entry_metrics(entry: dict[str, Any]) -> dict[str, Any]:
    rows = entry["report"]["rows"]
    latency = [row["latency_ns"] for row in rows]
    rss = entry["resource"]["maximum_resident_set_size_(kbytes)"] * 1024
    phase_rows = entry.get("phase_rows")
    require(isinstance(phase_rows, list) and len(phase_rows) == len(rows),
            f"{entry.get('directory', 'entry')}: validated phase rows are missing")
    phase_vectors = {name: [phase[name] for phase in phase_rows]
                     for name in PHASE_DURATION_FIELDS + ("phase_sum_ns", "lifecycle_residual_ns")}
    return {
        "latency_ns": _percentiles(latency),
        "latency_median_ci": _within_child_bootstrap(latency),
        "rss_bytes": rss,
        "phases": {key: _percentiles(value) for key, value in phase_vectors.items()},
        "phase_median_ci": {key: _within_child_bootstrap(value) for key, value in phase_vectors.items()},
        "allocation": _allocation_metrics(entry, rows),
    }


def _relative(before: int | float, delta: int | float) -> float | None:
    return None if before == 0 else (delta / abs(before)) * 100.0


def _comparison_metric_specs(role: str) -> list[tuple[str, Callable[[dict[str, Any]], int | float]]]:
    specs: list[tuple[str, Callable[[dict[str, Any]], int | float]]] = []
    for metric in ("p50", "p95", "p99"):
        specs.append((f"latency_ns.{metric}", lambda m, metric=metric: m["latency_ns"][metric]))
    specs.append(("rss_bytes", lambda m: m["rss_bytes"]))
    for phase in PHASE_DURATION_FIELDS + ("phase_sum_ns", "lifecycle_residual_ns"):
        for metric in ("p50", "p95", "p99"):
            specs.append((f"phase.{phase}.{metric}",
                          lambda m, phase=phase, metric=metric: m["phases"][phase][metric]))
    if role == "allocator":
        for field in ALLOCATION_VECTOR_FIELDS:
            for metric in ("p50", "p95", "p99"):
                specs.append((f"allocation.{field}.{metric}",
                              lambda m, field=field, metric=metric:
                              m["allocation"]["percentiles"][field][metric]))
    return specs


def _paired_comparisons(entries: list[dict[str, Any]]) -> list[dict[str, Any]]:
    by_key: dict[tuple[str, str, str, int], dict[str, dict[str, Any]]] = {}
    for entry in entries:
        spec = entry["spec"]
        if spec["api"] != UNMANAGED:
            continue
        key = (spec["role"], spec["arm"], spec["api"], spec["repeat"])
        metrics = entry.get("_metrics")
        if metrics is None:
            metrics = _entry_metrics(entry)
        by_key.setdefault(key, {})[spec["phase"]] = metrics
    groups: dict[tuple[str, str, str], list[tuple[int, dict[str, Any], dict[str, Any]]]] = {}
    for (role, arm, api, repeat), phases in by_key.items():
        require(set(phases) == set(PHASES), f"missing before/after pair for {role}/{arm}/r{repeat}")
        groups.setdefault((role, arm, api), []).append((repeat, phases["before"], phases["after"]))
    comparisons: list[dict[str, Any]] = []
    for (role, arm, api), blocks in sorted(groups.items()):
        for metric, getter in _comparison_metric_specs(role):
            deltas: list[int | float] = []
            raw_blocks: list[dict[str, Any]] = []
            for repeat, before, after in sorted(blocks):
                before_value, after_value = getter(before), getter(after)
                delta = after_value - before_value
                deltas.append(delta)
                block = {"repeat": repeat, "before": before_value, "after": after_value,
                         "delta": delta, "absolute_delta": delta,
                         "relative_percent": _relative(before_value, delta),
                         "adverse_flag": (_relative(before_value, delta) is not None
                                          and _relative(before_value, delta) > ADVERSE_THRESHOLD_PERCENT),
                         "threshold_percent": ADVERSE_THRESHOLD_PERCENT}
                if metric == "latency_ns.p50":
                    block["before_within_child_median_ci"] = before["latency_median_ci"]
                    block["after_within_child_median_ci"] = after["latency_median_ci"]
                elif metric.startswith("phase.") and metric.rsplit(".", 1)[1] == "p50":
                    phase_name = metric.split(".")[1]
                    block["before_within_child_median_ci"] = before["phase_median_ci"][phase_name]
                    block["after_within_child_median_ci"] = after["phase_median_ci"][phase_name]
                    # This is a ratio of the two p50 values.  It is not an
                    # additive share of the lifecycle because medians do not
                    # distribute over sums.
                    block["before_phase_median_ratio_percent"] = _relative(before["latency_ns"]["p50"], before_value)
                    block["after_phase_median_ratio_percent"] = _relative(after["latency_ns"]["p50"], after_value)
                elif metric.startswith("allocation.") and metric.rsplit(".", 1)[1] == "p50":
                    field = metric.split(".")[1]
                    block["before_within_child_median_ci"] = before["allocation"]["median_ci"][field]
                    block["after_within_child_median_ci"] = after["allocation"]["median_ci"][field]
                raw_blocks.append(block)
            bootstrap = paired_bootstrap(deltas)
            median_before = statistics.median(item["before"] for item in raw_blocks)
            median_delta = statistics.median(deltas)
            relative = _relative(median_before, median_delta)
            aggregate_adverse = relative is not None and relative > ADVERSE_THRESHOLD_PERCENT
            individual_adverse = any(item["adverse_flag"] for item in raw_blocks)
            comparisons.append({
                "role": role, "arm": arm, "api": api, "metric": metric,
                "blocks": raw_blocks, "delta_values": deltas,
                "median_before": median_before, "median_after": statistics.median(item["after"] for item in raw_blocks),
                "median_delta": median_delta, "absolute_delta": median_delta, "relative_percent": relative,
                # Preserve an adverse repeat-level cell even when the median
                # of the two repeat deltas falls below five percent.
                "adverse_flag": individual_adverse,
                "aggregate_adverse_flag": aggregate_adverse,
                "individual_adverse_blocks": [item["repeat"] for item in raw_blocks if item["adverse_flag"]],
                "threshold_percent": ADVERSE_THRESHOLD_PERCENT,
                "bootstrap": bootstrap,
                "interpretation": (
                    "descriptive paired block flag; phase percentages require absolute/share context; "
                    "no causal or historical-flag resolution claim"
                ),
            })
    return comparisons


def _repeat_visibility(entries: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Expose descriptive repeat-to-repeat changes for every matrix cell."""

    grouped: dict[tuple[str, str, str, str], dict[int, dict[str, Any]]] = {}
    for entry in entries:
        spec = entry["spec"]
        metrics = entry.get("_metrics")
        if metrics is None:
            metrics = _entry_metrics(entry)
        key = (spec["phase"], spec["api"], spec["role"], spec["arm"])
        require(spec["repeat"] not in grouped.setdefault(key, {}),
                f"duplicate repeat for {key}")
        grouped[key][spec["repeat"]] = metrics
    result: list[dict[str, Any]] = []
    for (phase, api, role, arm), repeats in sorted(grouped.items()):
        require(set(repeats) == set(REPEATS), f"repeat inventory incomplete for {phase}/{api}/{role}/{arm}")
        first, second = repeats[1], repeats[2]
        metrics: dict[str, Any] = {}
        for name, getter in _comparison_metric_specs(role):
            first_value, second_value = getter(first), getter(second)
            delta = second_value - first_value
            relative = _relative(first_value, delta)
            metrics[name] = {
                "repeat1": first_value,
                "repeat2": second_value,
                "absolute_delta": delta,
                "relative_percent": relative,
                "flag_over_5_percent": relative is not None and abs(relative) > ADVERSE_THRESHOLD_PERCENT,
            }
        allocation = first["allocation"]
        require(allocation["availability"] == second["allocation"]["availability"],
                f"repeat allocator availability differs for {phase}/{api}/{role}/{arm}")
        result.append({
            "phase": phase,
            "api": api,
            "role": role,
            "arm": arm,
            "metrics": metrics,
            "allocation": {
                "availability": allocation["availability"],
                "scope": allocation["scope"],
                "metrics": None if allocation["availability"] == "unavailable" else {
                    name: metrics[f"allocation.{name}.p50"]
                    for name in ALLOCATION_VECTOR_FIELDS
                },
            },
            "scope": "descriptive repeat variance within one phase/API/role/arm; no optimization regression claim",
        })
    return result


def _collect(attempt: str, protocol: dict[str, Any], builds: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    attempt = _attempt(attempt)
    root = ROOT / "captures" / attempt
    require(root.is_dir() and not root.is_symlink(), f"capture attempt is missing: {root}")
    expected = protocol["formal_runs"]
    labels = {item["label"] for item in expected}
    require({item.name for item in root.iterdir()} == labels, f"{root}: capture inventory differs")
    entries: list[dict[str, Any]] = []
    for item in expected:
        spec = dict(item, attempt=attempt)
        directory = root / spec["label"]
        started = _read_json(directory / "started.json")
        terminal = _read_json(directory / "terminal.json")
        build = builds[_build_key(spec["phase"], spec["role"])]
        _validate_terminal(started, terminal, spec, build, protocol, _sha(PROTOCOL_FILE), directory)
        report, phase_rows = _validate_report(directory / "report.json", spec, build)
        resource = _resource(directory / "resource.txt")
        entries.append({
            "spec": spec, "directory": directory, "started": started, "terminal": terminal,
            "report": report, "phase_rows": phase_rows, "resource": resource,
            "started_at": _timestamp(started["started_utc"], "started_utc"),
            "finished_at": _timestamp(terminal["finished_utc"], "finished_utc"),
        })
    ordered = _chronological_entries(entries, expected, attempt)
    require(not (TEMP / "runs" / attempt).exists(), f"private attempt scratch remains: {TEMP / 'runs' / attempt}")
    return ordered


def _chronological_entries(entries: list[dict[str, Any]], expected: list[dict[str, Any]], attempt: str) -> list[dict[str, Any]]:
    """Match actual receipt timestamps to formal order, never directory order."""

    ordered = sorted(entries, key=lambda item: item["started_at"])
    previous_finished: _datetime.datetime | None = None
    for expected_spec, entry in zip(expected, ordered):
        require(entry["spec"] == dict(expected_spec, attempt=attempt),
                "actual chronological capture order differs from frozen formal_runs")
        if previous_finished is not None:
            require(entry["started_at"] >= previous_finished,
                    "capture children overlap; CPU lock/order custody is invalid")
        previous_finished = entry["finished_at"]
    require(len(ordered) == len(expected), "actual receipt count differs from formal inventory")
    return ordered


def analyze_data(entries: list[dict[str, Any]], *, protocol: dict[str, Any] | None = None) -> dict[str, Any]:
    require(len(entries) == 32, "analysis requires exactly 32 retained children")
    children = []
    for entry in entries:
        metrics = _entry_metrics(entry)
        entry["_metrics"] = metrics
        children.append({
            "label": entry["spec"]["label"],
            "phase": entry["spec"]["phase"], "api": entry["spec"]["api"],
            "role": entry["spec"]["role"], "arm": entry["spec"]["arm"],
            "repeat": entry["spec"]["repeat"], "metrics": metrics,
        })
    managed = [item for item in children if item["api"] == MANAGED]
    comparisons = _paired_comparisons(entries)
    repeat_variance = _repeat_visibility(entries)
    return {
        "schema": ANALYSIS_SCHEMA,
        "version": VERSION,
        "case": CASE,
        "children": children,
        "child_count": len(children),
        "sample_count": len(children) * FORMAL_SAMPLES,
        "managed_after_capability": managed,
        "paired_unmanaged_comparisons": comparisons,
        "repeat_variance": repeat_variance,
        "adverse_flags": [item for item in comparisons if item["adverse_flag"]],
        "claims": {
            "scope": "descriptive phase latency/RSS/full-lifecycle allocation evidence and paired/repeat block review",
            "threshold_percent": ADVERSE_THRESHOLD_PERCENT,
            "bootstrap": {
                "paired_delta": {"method": "paired_block_bootstrap_median", "resamples": BOOTSTRAP_REPETITIONS, "seed": BOOTSTRAP_SEED},
                "within_child": {"method": "independent_within_child_bootstrap_median", "resamples": BOOTSTRAP_REPETITIONS, "seed": BOOTSTRAP_SEED},
                "uncertainty_scope": "descriptive within-child and two-repeat block resampling; not independent-host uncertainty",
            },
            "sample_pseudoreplication": False,
            "historical_flags_resolved": False,
            "managed_after_is_capability_only": True,
        },
        "protocol": protocol,
    }


def analyze(attempt: str) -> Path:
    protocol, _protocol_hash, builds = _load_protocol()
    entries = _collect(attempt, protocol, builds)
    result = analyze_data(entries, protocol={"path": "protocol.json", "sha256": _sha(PROTOCOL_FILE)})
    output = ROOT / "analysis" / f"{_attempt(attempt)}.json"
    _write_new(output, result)
    return output


def verify(attempt: str) -> Path:
    protocol, _protocol_hash, builds = _load_protocol()
    entries = _collect(attempt, protocol, builds)
    result = {"schema": "docx-phase-diagnostic-verification-v1", "version": VERSION,
              "attempt": _attempt(attempt), "children": len(entries),
              "samples": len(entries) * FORMAL_SAMPLES, "status": "pass",
              "protocol": {"path": "protocol.json", "sha256": _sha(PROTOCOL_FILE)}}
    output = ROOT / "verification" / f"{_attempt(attempt)}.json"
    _write_new(output, result)
    return output


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    freeze = sub.add_parser("freeze")
    freeze.add_argument("--builds", type=Path, default=BUILDS_FILE)
    for name in ("capture", "analyze", "verify"):
        command = sub.add_parser(name)
        command.add_argument("--attempt", required=True)
        if name == "capture":
            command.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT_SECONDS)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        if args.command == "freeze":
            print(create_freeze(args.builds))
        elif args.command == "capture":
            require(args.timeout == DEFAULT_TIMEOUT_SECONDS, "formal capture timeout is fixed at 180 seconds")
            capture_all(args.attempt, timeout_seconds=args.timeout)
        elif args.command == "analyze":
            print(analyze(args.attempt))
        elif args.command == "verify":
            print(verify(args.attempt))
        return 0
    except PhaseDiagnosticError as error:
        print(f"0496: FAIL: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
