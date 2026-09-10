#!/usr/bin/env python3
"""Bounded publication-route evidence for the DOCX tail-append harness.

The driver is a custody and analysis layer.  It never builds a binary or
reconstructs a checkout.  The root-owned build records authenticate the two
retained executables; the sealed 0484/0489 validators remain the source of
truth for the legacy report and replay/oracle contract.  Counting and atomic
reports are validated with an explicit publication extension so no atomic
write-call or sink digest is invented.
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
from typing import Any, Callable, Iterable, Mapping

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
TEMP = Path("/home/zhuhe/.cache/litchi-goal-0497")
CPU_LOCK = Path("/home/zhuhe/.cache/litchi-goal-0484/cpu.lock")
CPU = 2

VERSION = 1
CASE = "docx_replayable_tail_append_publication_routes"
ROLES = ("normal", "allocator")
PHASES = ("before", "after")
PUBLICATIONS = ("hashing_sink", "counting_sink", "atomic_path")
REPEATS = (1, 2)
FORMAL_SAMPLES = 30
FORMAL_WARMUPS = 3
PILOT_SAMPLES = 3
PILOT_WARMUPS = 1
DEFAULT_TIMEOUT_SECONDS = 1_800
TERM_GRACE_SECONDS = 10
BOOTSTRAP_REPETITIONS = 10_000
BOOTSTRAP_SEED = 497
ADVERSE_THRESHOLD_PERCENT = 5.0
SINK_WRITE_BYTES = 4 * 1024
ATTEMPT_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]*$")
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
REVISION_RE = re.compile(r"^[0-9a-f]{40}$")
EXPECTED_REVISION = "66af25e2fc3c208e6819fffa5fa29e102942bc8e"
EXPECTED_REVISIONS = {phase: EXPECTED_REVISION for phase in PHASES}

FORMAL_ORDERING = {
    "legacy": {
        "repeat1": "before_hash_then_after_hash_per_pair",
        "repeat2": "after_hash_then_before_hash_per_pair_reversed_pairs",
    },
    "after_only": {
        "repeat1": "counting_then_atomic_per_pair",
        "repeat2": "atomic_then_counting_per_pair_reversed_pairs",
    },
    "chronology": "receipts_must_follow_formal_runs_without_overlap",
    "purpose": "balance before_after and after_only route drift while retaining reversed pair order",
}

BUILDS_SCHEMA = "docx-replayable-tail-publication-builds-v1"
CAPTURE_SCHEMA = "docx-replayable-tail-publication-capture-v1"
TERMINAL_SCHEMA = "docx-replayable-tail-publication-terminal-v1"
ANALYSIS_SCHEMA = "docx-replayable-tail-publication-analysis-v1"
PROTOCOL_SCHEMA = "docx-replayable-tail-publication-protocol-v1"
CLEANUP_SCHEMA = "docx-replayable-tail-publication-private-cleanup-v1"

BUILDS_FILE = ROOT / "builds.json"
PROTOCOL_FILE = ROOT / "protocol.json"
MACHINE_FILE = ROOT / "machine.json"
PROVENANCE_FILE = ROOT / "provenance.json"
FIXTURE_INPUTS_FILE = ROOT / "fixture-inputs.json"
CAPTURE_ROOT = ROOT / "captures"

# The build script records this exact environment.  The capture environment
# is the frozen host snapshot plus a per-child private TMPDIR.
BUILD_ENV_KEYS = (
    "RUSTUP_TOOLCHAIN", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
    "CARGO_PROFILE_RELEASE_DEBUG", "RUSTFLAGS", "CARGO_TARGET_DIR",
    "TMPDIR", "DEBUGINFOD_URLS", "LC_ALL", "RUSTDOCFLAGS",
)
CAPTURE_ENV_KEYS = (
    "PATH", "HOME", "USER", "LANG", "LC_ALL", "RUSTUP_TOOLCHAIN",
    "RUSTFLAGS", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
    "CARGO_PROFILE_RELEASE_DEBUG", "CARGO_TARGET_DIR", "DEBUGINFOD_URLS",
    "RUSTDOCFLAGS",
)

# These are the common compilation inputs which must remain identical between
# the two source manifests.  The same revision is required, while the source
# manifest itself still authenticates every selected input.
HARNESS_CUSTODY_FILES = (
    "tools/perf-baseline/Cargo.toml",
    "tools/perf-baseline/src/lib.rs",
    "tools/perf-baseline/src/main.rs",
)
CANDIDATE_SOURCE_FILE = ROOT / "candidate-source.json"
CANDIDATE_PATCH_FILE = ROOT / "candidate.patch"

OLD_COMPARE_ROOT = REPO / "docs/performance/results/change-0489"
OLD_ROUTE_ROOT = REPO / "docs/performance/results/change-0484"
OLD_COMPARE = OLD_COMPARE_ROOT / "compare.py"
OLD_ROUTE_MEASURE = OLD_ROUTE_ROOT / "measure_routes.py"
OLD_MEASURE = OLD_ROUTE_ROOT / "measure.py"
OLD_COMMON = OLD_ROUTE_ROOT / "common.py"
OLD_ROUTE_PROTOCOL = OLD_ROUTE_ROOT / "route-protocol.json"

class MeasureError(RuntimeError):
    """A protocol, custody, report, or analysis failure."""


def fail(message: str) -> None:
    raise MeasureError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _exact(value: Any, keys: Iterable[str], path: str) -> None:
    require(isinstance(value, dict), f"{path}: expected object")
    expected = set(keys)
    actual = set(value)
    require(actual == expected,
            f"{path}: fields differ (expected {sorted(expected)}, got {sorted(actual)})")


def _uint(value: Any, path: str, *, positive: bool = False) -> int:
    require(type(value) is int and value >= (1 if positive else 0),
            f"{path}: expected unsigned integer")
    return value


def _text(value: Any, path: str, *, nonempty: bool = True) -> str:
    require(isinstance(value, str) and (not nonempty or bool(value)),
            f"{path}: expected text")
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
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError, json.JSONDecodeError) as error:
        fail(f"{path}: invalid JSON: {error}")
    _finite(value, str(path))
    return value


def _write_new(path: Path, value: Any) -> None:
    _finite(value, str(path))
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("x", encoding="utf-8") as stream:
            json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
            stream.write("\n")
    except FileExistsError as error:
        fail(f"refusing to replace immutable artifact: {path}")
        raise error


def _sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _meta(path: Path, *, executable: bool = False) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    if executable:
        require(os.access(path, os.X_OK), f"file is not executable: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": _sha(path)}


def _recorded_meta(value: Any, path: str, *, executable: bool = False) -> dict[str, Any]:
    _exact(value, ("path", "bytes", "sha256"), path)
    recorded_path = Path(_text(value["path"], f"{path}.path"))
    require(recorded_path.is_absolute(), f"{path}.path: path must be absolute")
    _uint(value["bytes"], f"{path}.bytes")
    _hash(value["sha256"], f"{path}.sha256")
    require(_meta(recorded_path, executable=executable) == value,
            f"{path}: retained artifact changed")
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


_SEALED: Any = None


def _load_sealed() -> Any:
    global _SEALED
    if _SEALED is not None:
        return _SEALED
    require(OLD_COMPARE.is_file() and OLD_ROUTE_MEASURE.is_file() and OLD_MEASURE.is_file(),
            "sealed 0484/0489 route validator inputs are missing")
    old_paths = [str(OLD_COMPARE_ROOT), str(OLD_ROUTE_ROOT)]
    inserted = [path for path in old_paths if path not in sys.path]
    for path in reversed(inserted):
        sys.path.insert(0, path)
    try:
        name = "_sealed_change0497_compare"
        spec = importlib.util.spec_from_file_location(name, OLD_COMPARE)
        require(spec is not None and spec.loader is not None,
                "cannot load sealed 0489 comparison validator")
        module = importlib.util.module_from_spec(spec)
        sys.modules[name] = module
        spec.loader.exec_module(module)
        _SEALED = module
        return module
    except Exception:
        sys.modules.pop("_sealed_change0497_compare", None)
        raise
    finally:
        for path in inserted:
            try:
                sys.path.remove(path)
            except ValueError:
                pass


# Importing the sealed helper is read-only.  The fallback exists only so unit
# tests can still import this module when an intentionally incomplete checkout
# omits the historical evidence; formal plan/freeze/capture commands fail
# closed through _load_sealed().
try:
    _INITIAL_SEALED = _load_sealed()
    ARMS = tuple(copy.deepcopy(item) for item in _INITIAL_SEALED.ARMS)
except Exception:  # pragma: no cover - formal environments always have the seal
    _INITIAL_SEALED = None
    _WORKLOADS = ("s64-a64-short-c64", "s64-a16384-short-c64", "s131072-a64-short-c64")
    ARMS = tuple({"id": f"deterministic-owned-{workload}"} for workload in _WORKLOADS) + tuple(
        {"id": f"deterministic-file-{workload}"} for workload in _WORKLOADS
    ) + tuple({"id": f"deterministic-short-read-{workload}"} for workload in _WORKLOADS) + tuple(
        {"id": f"deterministic-latency-{workload}"} for workload in _WORKLOADS
    ) + tuple({"id": f"memory_store-owned-{workload}"} for workload in _WORKLOADS) + tuple(
        {"id": f"file_store-owned-{workload}"} for workload in _WORKLOADS
    )
ARM_BY_ID = {item["id"]: item for item in ARMS}
require(len(ARMS) == 18 and len(ARM_BY_ID) == 18, "0497 arm inventory must contain 18 unique arms")


def _sealed() -> Any:
    return _INITIAL_SEALED or _load_sealed()


def _build_key(phase: str, role: str) -> str:
    require(phase in PHASES and role in ROLES, f"unknown build coordinate {phase}/{role}")
    return f"{phase}/{role}"


def _publication_for_phase(phase: str) -> tuple[str, ...]:
    require(phase in PHASES, f"unknown phase {phase}")
    return ("hashing_sink",) if phase == "before" else PUBLICATIONS


def formal_inventory() -> list[dict[str, Any]]:
    """Return 288 children using balanced per-pair formal publication order."""

    result: list[dict[str, Any]] = []
    ordinal = 0
    pairs = [(arm["id"], role) for arm in ARMS for role in ROLES]

    def append_run(repeat: int, phase: str, publication: str,
                   arm_id: str, role: str) -> None:
        nonlocal ordinal
        arm = ARM_BY_ID[arm_id]
        result.append({
            "ordinal": ordinal,
            "repeat": repeat,
            "phase": phase,
            "publication": publication,
            "role": role,
            "arm": arm_id,
            "workload": arm.get("workload"),
            "route": arm.get("route"),
            "input_mode": arm.get("input_mode"),
            "samples": FORMAL_SAMPLES,
            "warmups": FORMAL_WARMUPS,
            "label": f"r{repeat}-{phase}-{publication}-{role}-{arm_id}",
        })
        ordinal += 1

    # Pair each legacy before/after observation.  Repeat two reverses both
    # the pair order and phase order, producing ABBA-style drift control.
    for repeat, sequence, phase_order in (
        (1, pairs, ("before", "after")),
        (2, tuple(reversed(pairs)), ("after", "before")),
    ):
        for arm_id, role in sequence:
            for phase in phase_order:
                append_run(repeat, phase, "hashing_sink", arm_id, role)

    # Counting and atomic routes are after-only.  Their per-pair order is
    # likewise reversed in repeat two to avoid making one route a time-bound
    # proxy for the other.
    for repeat, sequence, publication_order in (
        (1, pairs, ("counting_sink", "atomic_path")),
        (2, tuple(reversed(pairs)), ("atomic_path", "counting_sink")),
    ):
        for arm_id, role in sequence:
            for publication in publication_order:
                append_run(repeat, "after", publication, arm_id, role)
    require(len(result) == 288, "formal inventory must contain exactly 288 children")
    require(sum(item["phase"] == "before" for item in result) == 72,
            "formal inventory before count differs")
    require(sum(item["phase"] == "after" and item["publication"] == "hashing_sink" for item in result) == 72,
            "formal inventory after hashing count differs")
    require(sum(item["phase"] == "after" and item["publication"] == "counting_sink" for item in result) == 72,
            "formal inventory after counting count differs")
    require(sum(item["phase"] == "after" and item["publication"] == "atomic_path" for item in result) == 72,
            "formal inventory after atomic count differs")
    require(len({item["label"] for item in result}) == len(result),
            "formal labels are not unique")
    for index, (arm_id, role) in enumerate(pairs):
        first = 2 * index
        second = first + 1
        require(result[first]["repeat"] == 1 and result[first]["phase"] == "before"
                and result[first]["publication"] == "hashing_sink"
                and result[first]["arm"] == arm_id and result[first]["role"] == role,
                "repeat-one legacy ABBA schedule differs")
        require(result[second]["repeat"] == 1 and result[second]["phase"] == "after"
                and result[second]["publication"] == "hashing_sink"
                and result[second]["arm"] == arm_id and result[second]["role"] == role,
                "repeat-one legacy after order differs")
    reversed_pairs = tuple(reversed(pairs))
    for index, (arm_id, role) in enumerate(reversed_pairs):
        first = 72 + 2 * index
        second = first + 1
        require(result[first]["repeat"] == 2 and result[first]["phase"] == "after"
                and result[first]["publication"] == "hashing_sink"
                and result[first]["arm"] == arm_id and result[first]["role"] == role,
                "repeat-two legacy ABBA schedule differs")
        require(result[second]["repeat"] == 2 and result[second]["phase"] == "before"
                and result[second]["publication"] == "hashing_sink"
                and result[second]["arm"] == arm_id and result[second]["role"] == role,
                "repeat-two legacy before order differs")
    for index, (arm_id, role) in enumerate(pairs):
        first = 144 + 2 * index
        second = first + 1
        require(result[first]["repeat"] == 1 and result[first]["publication"] == "counting_sink"
                and result[first]["arm"] == arm_id and result[first]["role"] == role,
                "repeat-one after-only counting order differs")
        require(result[second]["repeat"] == 1 and result[second]["publication"] == "atomic_path"
                and result[second]["arm"] == arm_id and result[second]["role"] == role,
                "repeat-one after-only atomic order differs")
    for index, (arm_id, role) in enumerate(reversed_pairs):
        first = 216 + 2 * index
        second = first + 1
        require(result[first]["repeat"] == 2 and result[first]["publication"] == "atomic_path"
                and result[first]["arm"] == arm_id and result[first]["role"] == role,
                "repeat-two after-only atomic order differs")
        require(result[second]["repeat"] == 2 and result[second]["publication"] == "counting_sink"
                and result[second]["arm"] == arm_id and result[second]["role"] == role,
                "repeat-two after-only counting order differs")
    return result


def pilot_inventory() -> list[dict[str, Any]]:
    """Return the separate 72-child 3-sample smoke lane.

    Every selected arm and publication route is exercised once per phase by
    the normal binary.  This is intentionally separate from the 288-child
    formal inventory and cannot be used as formal evidence.
    """

    result: list[dict[str, Any]] = []
    ordinal = 0
    for phase in PHASES:
        for publication in _publication_for_phase(phase):
            for arm in ARMS:
                result.append({
                    "ordinal": ordinal, "repeat": 1, "phase": phase,
                    "publication": publication, "role": "normal", "arm": arm["id"],
                    "workload": arm.get("workload"), "route": arm.get("route"),
                    "input_mode": arm.get("input_mode"), "samples": PILOT_SAMPLES,
                    "warmups": PILOT_WARMUPS,
                    "label": f"pilot-{phase}-{publication}-{arm['id']}", "pilot": True,
                })
                ordinal += 1
    require(len(result) == 72, "pilot inventory must contain exactly 72 children")
    return result


def _validate_source_manifest(value: Any, path: Path) -> dict[str, Any]:
    require(isinstance(value, dict) and value, f"{path}: source manifest is empty")
    for name, item in value.items():
        _text(name, f"{path}.member")
        _exact(item, ("path", "bytes", "sha256"), f"{path}.{name}")
        require(Path(_text(item["path"], f"{path}.{name}.path")).is_absolute(),
                f"{path}.{name}.path: path must be absolute")
        _uint(item["bytes"], f"{path}.{name}.bytes")
        _hash(item["sha256"], f"{path}.{name}.sha256")
    require(set(HARNESS_CUSTODY_FILES).issubset(value),
            f"{path}: source manifest omits harness custody members")
    return value


def _gate_role(path: str) -> str:
    role = path.rsplit("/", 1)[-1]
    require(role in ROLES, f"{path}: build gate role is not identifiable")
    return role


def _validate_gate(gate_meta: Any, source_meta: dict[str, Any], phase: str, role: str, path: Path) -> dict[str, Any]:
    recorded = _recorded_meta(gate_meta, f"{path}.gate")
    gate = _read_json(Path(recorded["path"]))
    require(isinstance(gate, dict), f"{path}: build gate receipt is not an object")
    required = {"argv", "cwd", "environment", "source_manifest", "started_ns", "driver",
                "pid", "exit_code", "finished_ns", "source_unchanged", "stdout", "stderr"}
    require(set(gate) == required, f"{path}: build gate receipt fields differ")
    require(gate["exit_code"] == 0 and gate["source_unchanged"] is True,
            f"{path}: original build gate failed or source changed")
    require(gate["source_manifest"] == source_meta,
            f"{path}: original gate source manifest binding differs")
    manifest = _read_json(Path(source_meta["path"]))
    cargo = manifest.get("tools/perf-baseline/Cargo.toml")
    require(isinstance(cargo, dict), f"{path}: source manifest omits Cargo.toml")
    cargo_path = Path(_text(cargo.get("path"), f"{path}.Cargo.toml.path"))
    source_root = cargo_path.resolve().parents[2]
    require(Path(_text(gate["cwd"], f"{path}.cwd")).resolve() == source_root,
            f"{path}: original build gate cwd differs from source root")
    expected = ["cargo", "build", "--release", "--locked", "--offline", "--manifest-path",
                "tools/perf-baseline/Cargo.toml"]
    if role == "allocator":
        expected += ["--features", "allocator-metrics"]
    expected += ["--bin", "docx_replayable_tail_append"]
    require(gate["argv"] == expected, f"{path}: original build gate argv differs")
    environment = gate["environment"]
    require(isinstance(environment, dict) and set(environment) == set(BUILD_ENV_KEYS),
            f"{path}: original build gate environment fields differ")
    require(environment["RUSTUP_TOOLCHAIN"] == "1.98.1"
            and environment["CARGO_BUILD_JOBS"] == "4"
            and environment["CARGO_INCREMENTAL"] == "0"
            and environment["CARGO_PROFILE_RELEASE_DEBUG"] == "1"
            and environment["RUSTFLAGS"] == "-C force-frame-pointers=yes -C force-unwind-tables=yes"
            and environment["DEBUGINFOD_URLS"] == ""
            and environment["LC_ALL"] == "C"
            and environment["RUSTDOCFLAGS"] == "-Dwarnings",
            f"{path}: original build gate toolchain environment differs")
    for name in ("CARGO_TARGET_DIR", "TMPDIR"):
        _text(environment[name], f"{path}.environment.{name}")
    for name in ("stdout", "stderr"):
        _recorded_meta(gate[name], f"{path}.{name}")
    _uint(gate["started_ns"], f"{path}.started_ns", positive=True)
    _uint(gate["finished_ns"], f"{path}.finished_ns", positive=True)
    require(gate["finished_ns"] > gate["started_ns"], f"{path}: build gate chronology is invalid")
    _uint(gate["pid"], f"{path}.pid", positive=True)
    driver = gate["driver"]
    _recorded_meta(driver, f"{path}.driver")
    require(Path(driver["path"]).name == "build.py", f"{path}: build driver identity differs")
    return gate


def _validate_build_record(value: Any, phase: str, role: str, path: Path) -> dict[str, Any]:
    require(isinstance(value, dict), f"{path}: build record is missing")
    _exact(value, ("binary", "git_revision", "source_manifest", "gate"), str(path))
    revision = _text(value["git_revision"], f"{path}.git_revision")
    require(REVISION_RE.fullmatch(revision) and revision == EXPECTED_REVISIONS[phase],
            f"{path}: source revision differs from frozen 66af baseline")
    binary = _recorded_meta(value["binary"], f"{path}.binary", executable=True)
    source_meta = _recorded_meta(value["source_manifest"], f"{path}.source_manifest")
    manifest = _validate_source_manifest(_read_json(Path(source_meta["path"])), Path(source_meta["path"]))
    gate = _validate_gate(value["gate"], source_meta, phase, role, path)
    return {"phase": phase, "role": role, "key": _build_key(phase, role),
            "git_revision": revision, "binary": binary, "source_manifest": source_meta,
            "manifest": manifest, "gate": gate,
            "gate_meta": dict(value["gate"])}


def _build_binding(build: dict[str, Any]) -> dict[str, Any]:
    return {"key": build["key"], "phase": build["phase"], "role": build["role"],
            "git_revision": build["git_revision"], "binary": build["binary"],
            "source_manifest": build["source_manifest"], "gate": build["gate_meta"]}


def _builds_binding(builds: dict[str, dict[str, Any]]) -> dict[str, Any]:
    return {key: _build_binding(builds[key]) for key in sorted(builds)}


def load_builds(path: Path | None = None) -> dict[str, dict[str, Any]]:
    path = Path(path or BUILDS_FILE).resolve()
    value = _read_json(path)
    _exact(value, ("schema", "builds"), str(path))
    require(value["schema"] == BUILDS_SCHEMA, f"{path}: build schema differs")
    builds = value["builds"]
    expected_keys = {_build_key(phase, role) for phase in PHASES for role in ROLES}
    require(isinstance(builds, dict) and set(builds) == expected_keys,
            f"{path}.builds: inventory differs")
    result = {}
    for phase in PHASES:
        for role in ROLES:
            result[_build_key(phase, role)] = _validate_build_record(
                builds[_build_key(phase, role)], phase, role, path / _build_key(phase, role))
    for phase in PHASES:
        require(result[f"{phase}/normal"]["source_manifest"] == result[f"{phase}/allocator"]["source_manifest"],
                f"{phase}: role source manifests differ")
        require(result[f"{phase}/normal"]["git_revision"] == result[f"{phase}/allocator"]["git_revision"],
                f"{phase}: role revisions differ")
    before = result["before/normal"]["manifest"]
    after = result["after/normal"]["manifest"]
    for name in HARNESS_CUSTODY_FILES:
        require(before[name]["bytes"] == after[name]["bytes"]
                and before[name]["sha256"] == after[name]["sha256"],
                f"shared harness source changed between phases: {name}")
    _, candidate = _candidate_source_binding()
    changed = set(before) ^ set(after)
    changed.update(name for name in set(before) & set(after)
                   if before[name]["bytes"] != after[name]["bytes"]
                   or before[name]["sha256"] != after[name]["sha256"])
    require(changed.issubset(set(candidate)),
            f"source manifests differ outside reviewed overlay: {sorted(changed - set(candidate))}")
    for name, expected in candidate.items():
        require(name in after and after[name]["bytes"] == expected["bytes"]
                and after[name]["sha256"] == expected["sha256"],
                f"after source manifest does not match candidate overlay: {name}")
    return result


def _custody(path: Path, label: str) -> dict[str, Any]:
    return _meta(path) if path.is_file() else fail(f"{label} is missing: {path}")


def _validate_machine_provenance() -> tuple[dict[str, Any], dict[str, Any]]:
    machine_meta = _custody(MACHINE_FILE, "machine.json")
    provenance_meta = _custody(PROVENANCE_FILE, "provenance.json")
    machine = _read_json(MACHINE_FILE)
    require(isinstance(machine, dict), "machine record is not an object")
    allowed = machine.get("allowed_cpus", machine.get("coordinator_affinity"))
    if allowed is not None:
        require(isinstance(allowed, list) and CPU in allowed,
                "machine record does not bind CPU 2")
    selected = machine.get("selected_cpu")
    if selected is not None:
        require(selected == CPU, "machine selected CPU differs")
    provenance = _read_json(PROVENANCE_FILE)
    require(isinstance(provenance, dict), "provenance record is not an object")
    inputs = provenance.get("inputs", {})
    require(isinstance(inputs, dict) and inputs, "provenance input custody is missing")
    for name, item in inputs.items():
        _recorded_meta(item, f"provenance.inputs.{name}")
    return machine_meta, provenance_meta


def _sealed_bindings() -> dict[str, Any]:
    for path in (OLD_COMPARE, OLD_ROUTE_MEASURE, OLD_MEASURE, OLD_COMMON, OLD_ROUTE_PROTOCOL):
        require(path.is_file(), f"sealed validator input is missing: {path}")
    return {"compare": _meta(OLD_COMPARE), "measure_routes": _meta(OLD_ROUTE_MEASURE),
            "measure": _meta(OLD_MEASURE), "common": _meta(OLD_COMMON),
            "route_protocol": _meta(OLD_ROUTE_PROTOCOL)}


def _fixture_inputs_binding() -> dict[str, Any]:
    """Bind root's supplemental non-Rust fixture inventory."""

    metadata = _meta(FIXTURE_INPUTS_FILE)
    value = _read_json(FIXTURE_INPUTS_FILE)
    require(isinstance(value, dict) and set(value) == {"scope", "files"},
            "fixture-inputs.json schema differs")
    _text(value["scope"], "fixture-inputs.scope")
    files = value["files"]
    require(isinstance(files, dict) and files, "fixture-inputs.files is empty")
    for name, item in files.items():
        _text(name, "fixture-inputs.file-name")
        _exact(item, ("bytes", "sha256"), f"fixture-inputs.files.{name}")
        _uint(item["bytes"], f"fixture-inputs.files.{name}.bytes")
        _hash(item["sha256"], f"fixture-inputs.files.{name}.sha256")
    return metadata


def _candidate_source_binding() -> tuple[dict[str, Any], dict[str, dict[str, Any]]]:
    """Authenticate the reviewed source overlay and its patch receipt.

    The manifest is the allowlist: the current review has five harness files,
    while the source-budget fixture correction may add a sixth.  Keeping the
    allowlist in the frozen manifest lets the verifier bind that reviewed
    change without baking a stale file count into the driver.
    """

    metadata = _meta(CANDIDATE_SOURCE_FILE)
    value = _read_json(CANDIDATE_SOURCE_FILE)
    require(isinstance(value, dict) and value, "candidate-source.json is empty")
    for name, item in value.items():
        _text(name, "candidate-source.file-name")
        _exact(item, ("bytes", "sha256"), f"candidate-source.{name}")
        _uint(item["bytes"], f"candidate-source.{name}.bytes")
        _hash(item["sha256"], f"candidate-source.{name}.sha256")
    # The reviewed overlay currently contains five files; the source-budget
    # fixture fix may add the sixth explicitly reviewed test file.  The
    # frozen candidate-source manifest, rather than a stale driver constant,
    # is the exact allowlist for the phase comparison below.
    require(len(value) >= 5, "candidate source overlay is unexpectedly small")
    patch = _meta(CANDIDATE_PATCH_FILE)
    return {"manifest": metadata, "patch": patch}, value


def _environment_snapshot() -> dict[str, str]:
    result: dict[str, str] = {}
    for key in CAPTURE_ENV_KEYS:
        if key in os.environ:
            result[key] = os.environ[key]
    # These are protocol constants even when the coordinator shell omitted a
    # variable.  Build.py records the same values.
    result.update({
        "RUSTUP_TOOLCHAIN": "1.98.1", "RUSTFLAGS": "-C force-frame-pointers=yes -C force-unwind-tables=yes",
        "CARGO_BUILD_JOBS": "4", "CARGO_INCREMENTAL": "0", "CARGO_PROFILE_RELEASE_DEBUG": "1",
        "CARGO_TARGET_DIR": str(TEMP / "target"), "DEBUGINFOD_URLS": "", "LC_ALL": "C",
        "RUSTDOCFLAGS": "-Dwarnings",
    })
    return result


def protocol_value(builds: dict[str, dict[str, Any]], builds_path: Path = BUILDS_FILE) -> dict[str, Any]:
    machine_meta, provenance_meta = _validate_machine_provenance()
    candidate_binding, _ = _candidate_source_binding()
    inventory = formal_inventory()
    pilots = pilot_inventory()
    return {
        "schema": PROTOCOL_SCHEMA, "version": VERSION, "case": CASE, "cpu": CPU,
        "cpu_lock": str(CPU_LOCK), "timeout_seconds": DEFAULT_TIMEOUT_SECONDS,
        "term_grace_seconds": TERM_GRACE_SECONDS, "samples": FORMAL_SAMPLES,
        "warmups": FORMAL_WARMUPS, "expected_children": len(inventory),
        "expected_samples": len(inventory) * FORMAL_SAMPLES, "formal_runs": inventory,
        "pilot_samples": PILOT_SAMPLES, "pilot_warmups": PILOT_WARMUPS,
        "expected_pilot_children": len(pilots),
        "expected_pilot_samples": len(pilots) * PILOT_SAMPLES,
        "pilot_runs": pilots,
        "ordering": copy.deepcopy(FORMAL_ORDERING),
        "builds_input": {"path": str(Path(builds_path).resolve()), "sha256": _sha(builds_path)},
        "builds": _builds_binding(builds), "driver": _meta(Path(__file__)),
        "tests": _meta(ROOT / "test_measure.py"), "sealed": _sealed_bindings(),
        "fixture_inputs": _fixture_inputs_binding(),
        "candidate_source": candidate_binding,
        "machine": machine_meta, "provenance": provenance_meta,
        "environment": _environment_snapshot(), "temporary_root": str(TEMP),
        "capture_root": str(CAPTURE_ROOT),
        "report_contract": {
            "legacy_validator": "change-0489/compare.py using change-0484 measure_routes.py",
            "legacy_publication": None,
            "counting_publication": {
                "route": "counting_sink", "timed_proof": True,
                "digest_source": "production ParagraphStreamPublication artifact proof",
                "write_counts": "reported only when emitted by CountingSink",
            },
            "atomic_publication": {
                "route": "atomic_path", "timed_output": "production atomic destination lifecycle",
                "readback_scope": "outside timed lifecycle", "write_counts": "unavailable; never synthesized",
            },
        },
        "claims": {
            "legacy_comparison": "before/after default hashing_sink only",
            "after_only_capabilities": ["counting_sink", "atomic_path"],
            "formal_ordering": "balanced per-pair ABBA schedule; chronology is custody, not a performance claim",
            "atomic_speedup_claim": False,
            "sample_pseudoreplication": False,
        },
    }


def create_freeze(builds_path: Path = BUILDS_FILE) -> Path:
    builds_path = Path(builds_path).resolve()
    builds = load_builds(builds_path)
    value = protocol_value(builds, builds_path)
    _write_new(PROTOCOL_FILE, value)
    return PROTOCOL_FILE


def _load_protocol() -> tuple[dict[str, Any], str, dict[str, dict[str, Any]]]:
    require(PROTOCOL_FILE.is_file(), f"frozen protocol is missing: {PROTOCOL_FILE}")
    protocol = _read_json(PROTOCOL_FILE)
    required = {"schema", "version", "case", "cpu", "cpu_lock", "timeout_seconds",
                "term_grace_seconds", "samples", "warmups", "expected_children",
                "expected_samples", "formal_runs", "pilot_samples", "pilot_warmups",
                "expected_pilot_children", "expected_pilot_samples", "pilot_runs",
                "builds_input", "builds", "driver",
                "tests", "sealed", "fixture_inputs", "candidate_source", "machine", "provenance", "environment",
                "temporary_root", "capture_root", "ordering", "report_contract", "claims"}
    require(set(protocol) == required, "protocol top-level fields differ")
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
    pilots = pilot_inventory()
    require(protocol["pilot_samples"] == PILOT_SAMPLES
            and protocol["pilot_warmups"] == PILOT_WARMUPS
            and protocol["pilot_runs"] == pilots
            and protocol["expected_pilot_children"] == len(pilots)
            and protocol["expected_pilot_samples"] == len(pilots) * PILOT_SAMPLES,
            "protocol pilot inventory differs")
    require(protocol["ordering"] == FORMAL_ORDERING,
            "protocol formal ordering differs")
    builds_input = protocol["builds_input"]
    _exact(builds_input, ("path", "sha256"), "protocol.builds_input")
    builds_path = Path(_text(builds_input["path"], "protocol.builds_input.path"))
    require(_sha(builds_path) == _hash(builds_input["sha256"], "protocol.builds_input.sha256"),
            "builds input changed after freeze")
    builds = load_builds(builds_path)
    require(_builds_binding(builds) == protocol["builds"], "build records changed after freeze")
    _exact(protocol["driver"], ("path", "bytes", "sha256"), "protocol.driver")
    require(Path(protocol["driver"]["path"]).resolve() == Path(__file__).resolve()
            and _meta(Path(__file__)) == protocol["driver"], "measurement driver changed after freeze")
    _exact(protocol["tests"], ("path", "bytes", "sha256"), "protocol.tests")
    require(Path(protocol["tests"]["path"]).resolve() == (ROOT / "test_measure.py").resolve()
            and _meta(ROOT / "test_measure.py") == protocol["tests"], "measurement tests changed after freeze")
    require(protocol["sealed"] == _sealed_bindings(), "sealed validator input changed after freeze")
    require(protocol["fixture_inputs"] == _fixture_inputs_binding(),
            "supplemental fixture inventory changed after freeze")
    candidate_binding, _ = _candidate_source_binding()
    require(protocol["candidate_source"] == candidate_binding,
            "reviewed candidate source overlay changed after freeze")
    machine, provenance = _validate_machine_provenance()
    require(protocol["machine"] == machine and protocol["provenance"] == provenance,
            "machine or provenance custody changed after freeze")
    require(protocol["temporary_root"] == str(TEMP) and protocol["capture_root"] == str(CAPTURE_ROOT),
            "protocol roots differ")
    return protocol, _sha(PROTOCOL_FILE), builds


def _run_root(attempt: str, label: str, *, pilot: bool = False) -> Path:
    lane = TEMP / "runs" / "pilots" if pilot else TEMP / "runs"
    return lane / attempt / label


def _capture_dir(attempt: str, label: str, *, pilot: bool = False) -> Path:
    lane = CAPTURE_ROOT / "pilots" if pilot else CAPTURE_ROOT
    return lane / attempt / label


def _private_input(arm: Mapping[str, Any], private: Path) -> dict[str, Any] | None:
    if arm.get("input_mode") != "file":
        return None
    sealed = _sealed()
    old_arm = sealed._old_arm(arm)
    relative = old_arm.get("input_file")
    require(isinstance(relative, str), "file input arm has no sealed fixture")
    source = (OLD_ROUTE_ROOT / relative).resolve()
    require(source.is_file() and not source.is_symlink(), f"sealed input fixture is missing: {source}")
    destination = private / "input" / Path(relative).name
    destination.parent.mkdir(parents=True, exist_ok=False)
    shutil.copyfile(source, destination)
    source_meta = _meta(source)
    copied_meta = _meta(destination)
    require(copied_meta["bytes"] == source_meta["bytes"]
            and copied_meta["sha256"] == source_meta["sha256"],
            "private input fixture copy changed")
    return {"path": str(destination), "bytes": copied_meta["bytes"],
            "sha256": copied_meta["sha256"],
            "source_fixture": {"path": str(source), "bytes": source_meta["bytes"],
                               "sha256": source_meta["sha256"]},
            "identity": "private_prepared_file_capability_fingerprint_matches_source_archive"}


def _command(spec: Mapping[str, Any], build: Mapping[str, Any], report: Path,
             resource: Path, replay_dir: Path | None, private_input: dict[str, Any] | None) -> list[str]:
    sealed = _sealed()
    arm = ARM_BY_ID[spec["arm"]]
    argv = list(sealed._argv_for(
        arm, build["binary"], report=report, resource=resource,
        samples=spec["samples"], warmups=spec["warmups"], replay_dir=replay_dir,
    ))
    if private_input is not None:
        require("--input-file" in argv, "file input command omitted --input-file")
        index = argv.index("--input-file")
        require(index + 1 < len(argv), "file input command has no path")
        argv[index + 1] = private_input["path"]
    publication = spec["publication"]
    if publication != "hashing_sink":
        # The Rust harness accepts publication aliases; freeze the canonical
        # snake_case spelling in every terminal argv receipt.
        insert = argv.index("--json") if "--json" in argv else len(argv)
        argv[insert:insert] = ["--publication", publication]
    return argv


def _expected_environment(protocol: Mapping[str, Any], tmpdir: Path) -> dict[str, str]:
    environment = {str(key): str(value) for key, value in protocol["environment"].items()}
    environment["TMPDIR"] = str(tmpdir)
    return environment


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


def _failure_inventory(private: Path) -> list[dict[str, Any]]:
    inventory: list[dict[str, Any]] = []
    for path in sorted(private.rglob("*")):
        require(not path.is_symlink(), f"private scratch contains symlink: {path}")
        if path.is_file():
            inventory.append({"path": str(path), "bytes": path.stat().st_size,
                              "sha256": _sha(path)})
    return inventory


def _cleanup_private(private: Path, *, preserve_failure: bool = False) -> dict[str, Any]:
    runs_root = TEMP / "runs"
    expected_parent = runs_root / "pilots" if private.parent.parent.name == "pilots" else runs_root
    require(private.parent.parent == expected_parent,
            f"private scratch escaped 0497 root: {private}")
    require(private.is_dir() and not private.is_symlink(),
            f"private scratch is not a directory: {private}")
    for child in private.rglob("*"):
        require(not child.is_symlink(), f"private scratch contains symlink: {child}")
    if preserve_failure:
        return {"schema": CLEANUP_SCHEMA, "status": "preserved_failure",
                "root": str(private), "removed": [], "remaining": [str(private)],
                "preserved": True, "inventory": _failure_inventory(private)}
    direct = list(private.iterdir())
    require(all(child.name in {"tmp", "replay"} and child.is_dir()
                and not list(child.iterdir()) for child in direct),
            f"private scratch has unexpected successful-run output: {[child.name for child in direct]}")
    shutil.rmtree(private)
    removed = [str(private)]
    attempt_parent = private.parent
    if attempt_parent.is_dir() and not attempt_parent.is_symlink() and not list(attempt_parent.iterdir()):
        attempt_parent.rmdir()
        removed.append(str(attempt_parent))
    runs_parent = expected_parent
    if runs_parent.is_dir() and not runs_parent.is_symlink() and not list(runs_parent.iterdir()):
        runs_parent.rmdir()
        removed.append(str(runs_parent))
    remaining = [str(path) for path in (private, attempt_parent, runs_parent) if path.exists()]
    return {"schema": CLEANUP_SCHEMA, "status": "pass" if not remaining else "failed",
            "root": str(private), "removed": removed, "remaining": remaining,
            "preserved": False, "inventory": []}


def _resource(path: Path) -> dict[str, Any]:
    require(path.is_file() and path.stat().st_size > 0, f"missing GNU time resource receipt: {path}")
    raw = path.read_bytes()
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


def _arm_case(arm: Mapping[str, Any]) -> dict[str, Any]:
    return dict(_sealed()._case_for(arm))


def _route_spec(arm: Mapping[str, Any]) -> Any:
    return _sealed().old_routes.ROUTE_BY_NAME[arm["route"]]


def _report_path_for_projection(report: dict[str, Any], *, counting: bool) -> dict[str, Any]:
    """Return a temporary legacy projection; never write the raw report back."""

    projection = copy.deepcopy(report)
    config = projection.get("config")
    require(isinstance(config, dict), "report config is missing")
    publication = config.pop("publication", None)
    if counting:
        require(publication == "counting_sink", "counting report publication config differs")
    else:
        require(publication == "atomic_path", "atomic report publication config differs")
    config["sink"] = _sealed().old_routes.base.SINK_CONTRACT
    for index, sample in enumerate(projection["cases"][0]["samples"]):
        require(isinstance(sample, dict), f"projection sample {index} is not an object")
        sample.pop("publication", None)
        if counting:
            sink = sample.get("sink")
            require(isinstance(sink, dict), f"counting sample {index}: sink is missing")
            # The counting sink's actual write calls/histogram/accepted bytes
            # remain raw.  Only the production artifact proof supplies the
            # digest which the historical validator requires.
            require(sink.get("sha256") is None, f"counting sample {index}: sink fabricated a digest")
            proof = sample.get("_publication_digest_for_projection")
            require(isinstance(proof, str), "internal counting proof is missing")
            sink["sha256"] = proof
            sample.pop("_publication_digest_for_projection", None)
    return projection


def _write_projection(value: dict[str, Any]) -> Path:
    root = TEMP / "projections"
    root.mkdir(parents=True, exist_ok=True)
    handle = tempfile.NamedTemporaryFile("w", encoding="utf-8", suffix=".json",
                                         prefix="projection-", dir=root, delete=False)
    path = Path(handle.name)
    with handle:
        json.dump(value, handle, sort_keys=True, allow_nan=False)
    return path


def _publication_record(sample: Any, path: str) -> dict[str, Any]:
    require(isinstance(sample, dict), f"{path}: sample is not an object")
    value = sample.get("publication")
    require(isinstance(value, dict), f"{path}.publication: missing publication proof")
    required = {"schema", "route", "timing_scope", "timed_candidate_artifact_bytes",
                "timed_candidate_artifact_sha256", "timed_candidate_matches_oracle",
                "verification_scope"}
    require(set(value) in (required, required | {"atomic"}),
            f"{path}.publication: fields differ")
    require(value["route"] in ("counting_sink", "atomic_path"),
            f"{path}.publication.route: not an after-only route")
    require(value["schema"] == "docx-replayable-tail-append-publication-v1",
            f"{path}.publication.schema: schema differs")
    _uint(value["timed_candidate_artifact_bytes"], f"{path}.publication.timed_candidate_artifact_bytes", positive=True)
    _hash(value["timed_candidate_artifact_sha256"], f"{path}.publication.timed_candidate_artifact_sha256")
    require(value["timed_candidate_matches_oracle"] is True,
            f"{path}.publication: timed artifact does not match oracle")
    expected_scope = {
        "counting_sink": "timed_production_artifact_proof_plus_untimed_candidate_oracle",
        "atomic_path": "timed_production_artifact_proof_plus_post_timer_path_oracle",
    }[value["route"]]
    expected_timing_scope = {
        "counting_sink": "source_admission_prepare_sequential_sink_publication_drop",
        "atomic_path": "source_admission_prepare_atomic_write_data_sync_rename_parent_directory_sync_publication_drop",
    }[value["route"]]
    require(value["timing_scope"] == expected_timing_scope,
            f"{path}.publication.timing_scope: scope differs")
    require(value["verification_scope"] == expected_scope,
            f"{path}.publication.verification_scope: scope differs")
    if value["route"] == "counting_sink":
        require("atomic" not in value, f"{path}.publication.atomic: counting route has atomic record")
    else:
        require(set(value) == required | {"atomic"}
                and isinstance(value["atomic"], dict),
                f"{path}.publication.atomic: atomic record is missing")
    return value


def _oracle_digest(report: dict[str, Any]) -> tuple[int, str]:
    observed = report["cases"][0]
    oracle = observed["oracle"]
    return int(oracle["candidate_archive_bytes"]), str(oracle["candidate_archive_sha256"])


def _validate_publication_sample(sample: dict[str, Any], publication: str,
                                  candidate_bytes: int, candidate_sha: str, path: str,
                                  private_tmpdir: Path | None = None) -> dict[str, Any]:
    value = _publication_record(sample, path)
    require(value["route"] == publication, f"{path}.publication.route: differs from run route")
    require(value["timed_candidate_artifact_bytes"] == candidate_bytes
            and value["timed_candidate_artifact_sha256"] == candidate_sha,
            f"{path}.publication: timed artifact proof differs from candidate oracle")
    if publication == "counting_sink":
        sink = sample.get("sink")
        require(isinstance(sink, dict), f"{path}.sink: counting sink is missing")
        require(set(sink) == {"accepted_bytes", "write_calls", "largest_write", "histogram", "sha256"},
                f"{path}.sink: counting observation fields differ")
        require("sha256" not in sink or sink["sha256"] is None,
                f"{path}.sink.sha256: counting route must not claim a local digest")
        for field in ("accepted_bytes", "write_calls", "largest_write", "histogram"):
            require(field in sink and sink[field] is not None,
                    f"{path}.sink.{field}: counting observation is missing")
        require(sink["accepted_bytes"] == candidate_bytes,
                f"{path}.sink.accepted_bytes: differs from candidate oracle")
        return value
    sink = sample.get("sink")
    require(isinstance(sink, dict), f"{path}.sink: atomic sink observation is missing")
    # SinkRecord serializes its unavailable Option fields as JSON nulls in the
    # current harness.  Older harness receipts omitted them, so accept either
    # representation while rejecting every fabricated count/digest.
    atomic_sink_fields = {"accepted_bytes", "write_calls", "largest_write", "histogram", "sha256"}
    require(set(sink).issubset(atomic_sink_fields),
            f"{path}.sink: atomic route has unknown sink fields")
    require(all(field_value is None for field_value in sink.values()),
            f"{path}.sink: atomic route must not invent sink write counts")
    atomic = value["atomic"]
    _exact(atomic, ("destination_path", "private_parent_path", "before", "after",
                    "post_timer_archive_bytes", "post_timer_archive_sha256",
                    "output_bytes_exact", "output_sha256_exact", "inverse_oracle_scope",
                    "post_timer_oracle", "cleanup"), f"{path}.publication.atomic")
    for name in ("destination_path", "private_parent_path"):
        recorded = Path(_text(atomic[name], f"{path}.publication.atomic.{name}"))
        require(recorded.is_absolute(), f"{path}.publication.atomic.{name}: path is not absolute")
        require(not recorded.exists(), f"{path}.publication.atomic.{name}: cleaned path still exists")
    require(private_tmpdir is not None,
            f"{path}.publication.atomic: private TMPDIR binding is missing")
    destination = Path(atomic["destination_path"])
    private_parent = Path(atomic["private_parent_path"])
    require(private_parent.parent == private_tmpdir
            and destination.parent == private_parent
            and destination.name == "published.docx",
            f"{path}.publication.atomic: destination escaped private TMPDIR")
    for name in ("before", "after"):
        state = atomic[name]
        _exact(state, ("exists", "regular_file", "bytes"), f"{path}.publication.atomic.{name}")
        require(type(state["exists"]) is bool and type(state["regular_file"]) is bool,
                f"{path}.publication.atomic.{name}: state booleans malformed")
        if name == "before":
            require(state == {"exists": False, "regular_file": False, "bytes": None},
                    f"{path}.publication.atomic.before: destination was not absent")
        else:
            require(state["exists"] is True and state["regular_file"] is True
                    and state["bytes"] == candidate_bytes,
                    f"{path}.publication.atomic.after: output state differs")
    require(atomic["post_timer_archive_bytes"] == candidate_bytes,
            f"{path}.publication.atomic.post_timer_archive_bytes: differs")
    require(atomic["post_timer_archive_sha256"] == candidate_sha,
            f"{path}.publication.atomic.post_timer_archive_sha256: differs")
    require(atomic["output_bytes_exact"] is True and atomic["output_sha256_exact"] is True,
            f"{path}.publication.atomic: output identity failed")
    require(atomic["inverse_oracle_scope"] ==
            "untimed_fixture_publication_inverse_exact; timed_atomic_publication_inverse_not_reexecuted",
            f"{path}.publication.atomic.inverse_oracle_scope: scope differs")
    oracle = atomic["post_timer_oracle"]
    require(isinstance(oracle, dict), f"{path}.publication.atomic.post_timer_oracle: missing")
    _exact(oracle, (
        "candidate_archive_bytes", "candidate_archive_sha256", "candidate_main_xml_bytes",
        "candidate_main_xml_sha256", "candidate_semantic", "candidate_member_count",
        "candidate_xml_exact", "candidate_semantic_exact", "untouched_member_metadata_exact",
        "untouched_raw_members_preserved", "physical_order_exact", "opaque_member_exact",
        "source_unchanged", "inverse_exact",
    ), f"{path}.publication.atomic.post_timer_oracle")
    _uint(oracle["candidate_archive_bytes"],
          f"{path}.publication.atomic.post_timer_oracle.candidate_archive_bytes", positive=True)
    _hash(oracle["candidate_archive_sha256"],
          f"{path}.publication.atomic.post_timer_oracle.candidate_archive_sha256")
    require(oracle["candidate_archive_bytes"] == candidate_bytes
            and oracle["candidate_archive_sha256"] == candidate_sha,
            f"{path}.publication.atomic.post_timer_oracle: candidate archive identity differs")
    _uint(oracle["candidate_main_xml_bytes"],
          f"{path}.publication.atomic.post_timer_oracle.candidate_main_xml_bytes", positive=True)
    _hash(oracle["candidate_main_xml_sha256"],
          f"{path}.publication.atomic.post_timer_oracle.candidate_main_xml_sha256")
    _uint(oracle["candidate_member_count"],
          f"{path}.publication.atomic.post_timer_oracle.candidate_member_count", positive=True)
    semantic = oracle["candidate_semantic"]
    _exact(semantic, ("paragraph_count", "order_sha256", "text_sha256", "text_bytes"),
           f"{path}.publication.atomic.post_timer_oracle.candidate_semantic")
    _uint(semantic["paragraph_count"],
          f"{path}.publication.atomic.post_timer_oracle.candidate_semantic.paragraph_count",
          positive=True)
    _hash(semantic["order_sha256"],
          f"{path}.publication.atomic.post_timer_oracle.candidate_semantic.order_sha256")
    _hash(semantic["text_sha256"],
          f"{path}.publication.atomic.post_timer_oracle.candidate_semantic.text_sha256")
    _uint(semantic["text_bytes"],
          f"{path}.publication.atomic.post_timer_oracle.candidate_semantic.text_bytes")
    oracle_flags = (
        "candidate_xml_exact", "candidate_semantic_exact", "untouched_member_metadata_exact",
        "untouched_raw_members_preserved", "physical_order_exact", "opaque_member_exact",
        "source_unchanged", "inverse_exact",
    )
    for flag in oracle_flags:
        require(oracle.get(flag) is True,
                f"{path}.publication.atomic.post_timer_oracle.{flag}: failed")
    cleanup = atomic["cleanup"]
    _exact(cleanup, ("destination_removed", "parent_removed"),
           f"{path}.publication.atomic.cleanup")
    require(cleanup["destination_removed"] is True and cleanup["parent_removed"] is True,
            f"{path}.publication.atomic.cleanup: destination cleanup failed")
    return value


def _validate_non_sink_sample(sample: Any, index: int, role: str, case: dict[str, Any],
                              oracle: dict[str, Any], authored: dict[str, Any],
                              expected_authored_opens: int, path: str) -> None:
    """Apply every sealed sample check except the publication sink fields."""

    base = _sealed().old_routes.base
    item = base._object(sample, path)
    require(base._int_field(item.get("sample"), f"{path}.sample") == index,
            f"{path}.sample index differs")
    require(base._positive_int(item.get("source_count"), f"{path}.source_count") == case["source_count"],
            f"{path}.source_count differs")
    require(base._positive_int(item.get("authored_count"), f"{path}.authored_count") == case["authored_count"],
            f"{path}.authored_count differs")
    require(item.get("chunk_mode") == base.REPORT_CHUNK_MODES[case["chunk_mode"]],
            f"{path}.chunk_mode differs")
    require(item.get("text_mode") == base.REPORT_TEXT_MODES[case["text_mode"]],
            f"{path}.text_mode differs")
    base._positive_int(item.get("elapsed_ns"), f"{path}.elapsed_ns")
    reads = base._object(item.get("source_reads"), f"{path}.source_reads")
    calls = base._positive_int(reads.get("calls"), f"{path}.source_reads.calls")
    requested = base._int_field(reads.get("requested_bytes"), f"{path}.source_reads.requested_bytes")
    returned = base._int_field(reads.get("returned_bytes"), f"{path}.source_reads.returned_bytes")
    require(requested > 0 and returned > 0 and returned <= requested,
            f"{path}.source_reads totals are invalid")
    base._check_histogram(reads.get("request_histogram"), f"{path}.source_reads.request_histogram",
                          calls, observed_bytes=requested)
    base._check_histogram(reads.get("returned_histogram"), f"{path}.source_reads.returned_histogram",
                          calls, observed_bytes=returned)
    observed_authored = base._object(item.get("authored"), f"{path}.authored")
    opens = base._int_field(observed_authored.get("opens"), f"{path}.authored.opens")
    require(opens == expected_authored_opens, f"{path}.authored.opens differs")
    event_passes = 1 if expected_authored_opens == 0 else expected_authored_opens
    require(observed_authored.get("events") == authored["event_count"] * event_passes,
            f"{path}.authored.events differs")
    require(observed_authored.get("text_chunks") == (authored["event_count"] - 2 * authored["authored_count"]) * event_passes,
            f"{path}.authored.text_chunks differs")
    require(observed_authored.get("text_bytes") == authored["text_bytes"] * event_passes,
            f"{path}.authored.text_bytes differs")
    for field in ("events", "text_chunks", "text_bytes"):
        base._int_field(observed_authored.get(field), f"{path}.authored.{field}")
    base._check_allocator_sample(item, role, path)
    base._check_process_sample(item.get("process"), f"{path}.process")


def _expected_new_report_config(arm: Mapping[str, Any], route_spec: Any,
                                publication: str, samples: int, warmups: int,
                                replay_dir: Path | None) -> dict[str, Any]:
    """Build the complete selected-arm configuration contract.

    Atomic and counting reports deliberately skip only the sink observation:
    every other ConfigRecord field remains observable and is bound to the
    sealed arm/route metadata here.
    """

    input_profile = arm.get("input_profile")
    require(input_profile is None or isinstance(input_profile, dict),
            "selected arm input profile is malformed")
    input_profile = input_profile or {}
    expected_replay_dir = (
        str(replay_dir.resolve()) if route_spec.name == "file_store" and replay_dir is not None else None
    )
    require(route_spec.name != "file_store" or replay_dir is not None,
            "file-store selected arm has no replay directory")
    sink = {
        "counting_sink": "non_seek_counting_short_write_no_archive_retention_production_artifact_proof",
        "atomic_path": "production_atomic_path_output_proof_post_timer_destination_oracle",
    }[publication]
    return {
        "source_counts": [arm["source_count"]],
        "authored_counts": [arm["authored_count"]],
        "chunk_modes": [_sealed().old_routes.base.REPORT_CHUNK_MODES[arm["chunk_mode"]]],
        "text_modes": [_sealed().old_routes.base.REPORT_TEXT_MODES[arm["text_mode"]]],
        "samples": samples,
        "warmups": warmups,
        "sink_write_bytes": arm["sink_write_bytes"],
        "lifecycle": list(_sealed().old_routes.base.LIFECYCLE),
        "expected_authored_opens": route_spec.expected_authored_opens,
        "expected_replay_opens": route_spec.expected_replay_opens,
        "source": arm["source_contract"],
        "authored_provider": route_spec.report_provider,
        "provider": route_spec.name,
        "replay_max_bytes": route_spec.replay_max_bytes,
        "replay_dir": expected_replay_dir,
        "replay_sync": route_spec.replay_sync,
        "compression": arm["compression"],
        "input_mode": arm["input_mode"],
        "input_storage_kind": arm["input_backing"],
        "input_identity_validation": arm["input_identity_validation"],
        "input_max_range_bytes": input_profile.get("max_range_bytes"),
        "input_delay_us": input_profile.get("delay_us", 0),
        "input_overhead_us": input_profile.get("overhead_us", 0),
        "input_bytes_per_second": input_profile.get("bytes_per_second"),
        "sink": sink,
        "fixture_dir": None,
        "publication": publication,
    }


def _validate_new_report_config(config: Any, arm: Mapping[str, Any], route_spec: Any,
                                publication: str, samples: int, warmups: int,
                                replay_dir: Path | None, path: str) -> None:
    expected = _expected_new_report_config(arm, route_spec, publication, samples,
                                           warmups, replay_dir)
    _exact(config, expected.keys(), path)
    require(config == expected, f"{path}: selected-arm configuration differs")


def _validate_common_new_report(raw: dict[str, Any], arm: Mapping[str, Any], role: str,
                                publication: str, samples: int, warmups: int,
                                binary: dict[str, Any], argv: list[str], replay_dir: Path | None,
                                input_metadata: dict[str, Any] | None, tmpdir: Path,
                                report_path: Path) -> dict[str, Any]:
    """Validate new publication reports with sealed source/replay/oracle checks."""

    sealed = _sealed()
    base = sealed.old_routes.base
    require(raw.get("schema") == base.REPORT_SCHEMA and raw.get("version") == 1,
            f"{report_path}: report schema differs")
    expected_binary = {
        "binary": "litchi-perf-baseline" if role == "normal" else "litchi-perf-baseline-alloc",
        "allocator": "Rust system allocator" if role == "normal" else "CountingSystemAllocator(std::alloc::System)",
        "instrumentation": "none" if role == "normal" else "system_allocator_operation_scoped",
        "counter_revision": None if role == "normal" else "serialized_region_peak_v3",
    }
    require(raw.get("binary") == expected_binary, f"{report_path}: binary identity differs")
    case = _arm_case(arm)
    route_spec = _route_spec(arm)
    config = raw.get("config")
    _validate_new_report_config(config, arm, route_spec, publication, samples, warmups,
                                replay_dir, f"{report_path}.config")
    expected_provider = route_spec.name
    expected_authored_provider = route_spec.report_provider
    cases = raw.get("cases")
    require(isinstance(cases, list) and len(cases) == 1, f"{report_path}: expected one case")
    observed = cases[0]
    require(isinstance(observed, dict), f"{report_path}.cases[0]: expected object")
    require(observed.get("provider") == expected_provider
            and observed.get("source_count") == case["source_count"]
            and observed.get("authored_count") == case["authored_count"]
            and observed.get("chunk_mode") == base.REPORT_CHUNK_MODES[case["chunk_mode"]]
            and observed.get("text_mode") == base.REPORT_TEXT_MODES[case["text_mode"]]
            and observed.get("replay_max_bytes") == route_spec.replay_max_bytes
            and observed.get("compression") == arm["compression"]
            and observed.get("input_mode") == arm["input_mode"]
            and observed.get("input_storage_kind") == arm["input_backing"]
            and observed.get("input_identity_validation") == arm["input_identity_validation"]
            and observed.get("sink_write_bytes") == arm["sink_write_bytes"],
            f"{report_path}: case identity differs")
    source = base._object(observed.get("source"), f"{report_path}.cases[0].source")
    authored = base._object(observed.get("authored"), f"{report_path}.cases[0].authored")
    limits = base._object(observed.get("limits"), f"{report_path}.cases[0].limits")
    oracle = base._object(observed.get("oracle"), f"{report_path}.cases[0].oracle")
    proof = base._object(observed.get("proof"), f"{report_path}.cases[0].proof")
    base._check_source_identity(source, case, f"{report_path}.cases[0].source")
    base._check_authored_identity(authored, case, f"{report_path}.cases[0].authored")
    base._check_limits(limits, authored, case, f"{report_path}.cases[0].limits",
                       expected_max_replay_bytes=route_spec.replay_max_bytes if route_spec.store else None)
    base._check_oracle_and_proof(source, authored, limits, oracle, proof, case,
                                 f"{report_path}.cases[0]")
    # These sealed profile checks are common to route and axis arms.  The
    # publication sink itself is the sole field intentionally handled by the
    # route-specific validator below.
    sealed.old_routes._check_corpus_case(observed, f"{report_path}.cases[0]")
    sealed.old_routes._check_compression_profile(observed, arm["compression"],
                                                 f"{report_path}.cases[0]")
    samples_value = observed.get("samples")
    require(isinstance(samples_value, list) and len(samples_value) == samples,
            f"{report_path}: sample cardinality differs")
    candidate_bytes = int(oracle["candidate_archive_bytes"])
    candidate_sha = str(oracle["candidate_archive_sha256"])
    for index, sample in enumerate(samples_value):
        path = f"{report_path}.cases[0].samples[{index}]"
        _validate_non_sink_sample(sample, index, role, case, oracle, proof["authored"],
                                  route_spec.expected_authored_opens, path)
        _validate_publication_sample(sample, publication, candidate_bytes, candidate_sha,
                                     path, tmpdir)
        if publication == "atomic_path":
            require(sample["publication"]["atomic"]["post_timer_oracle"] == oracle,
                    f"{path}.publication.atomic.post_timer_oracle: differs from case oracle")
        if route_spec.store:
            sealed.old_routes._check_replay_observation(sample, route_spec, proof, path)
        else:
            require(sample.get("replay") is None, f"{path}.replay: deterministic route reported replay")
    require(len(argv) > 7 and argv[7] == binary["path"],
            f"{report_path}: argv binary binding differs")
    if arm.get("input_mode") == "file":
        require(input_metadata is not None, f"{report_path}: file input metadata missing")
        require(source["archive_bytes"] == input_metadata["bytes"]
                and source["archive_sha256"] == input_metadata["sha256"],
                f"{report_path}: private input differs from source archive")
    else:
        require(input_metadata is None, f"{report_path}: non-file input has file metadata")
    if route_spec.name == "file_store":
        require(replay_dir is not None and config.get("replay_dir") == str(replay_dir.resolve()),
                f"{report_path}: file-store replay path differs")
    else:
        require(config.get("replay_dir") is None, f"{report_path}: non-file route has replay path")
    return raw


def _validate_report(path: Path, spec: Mapping[str, Any], build: dict[str, Any],
                     argv: list[str], input_metadata: dict[str, Any] | None,
                     replay_dir: Path | None, tmpdir: Path) -> dict[str, Any]:
    raw = _read_json(path)
    require(isinstance(raw, dict), f"{path}: report is not an object")
    arm = ARM_BY_ID[spec["arm"]]
    publication = spec["publication"]
    if publication == "hashing_sink":
        # The sealed validator checks the complete legacy report shell, sink,
        # replay route, source identity, and every oracle/proof field.
        sealed = _sealed()
        return sealed._validate_report(
            path, arm, spec["role"], samples=spec["samples"], warmups=spec["warmups"],
            binary=build["binary"], argv=argv, input_metadata=input_metadata,
            replay_dir=replay_dir,
        )
    candidate_bytes, candidate_sha = _oracle_digest(raw)
    samples = raw.get("cases", [{}])[0].get("samples", [])
    require(isinstance(samples, list), f"{path}: samples are not a list")
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict), f"{path}: sample {index} is not an object")
        _validate_publication_sample(sample, spec["publication"], candidate_bytes,
                                     candidate_sha, f"{path}.cases[0].samples[{index}]", tmpdir)
    if publication == "counting_sink":
        augmented = copy.deepcopy(raw)
        for sample in augmented["cases"][0]["samples"]:
            sample["_publication_digest_for_projection"] = sample["publication"][
                "timed_candidate_artifact_sha256"]
        projection = _report_path_for_projection(augmented, counting=True)
        projection_path = _write_projection(projection)
        try:
            _sealed()._validate_report(
                projection_path, arm, spec["role"], samples=spec["samples"], warmups=spec["warmups"],
                binary=build["binary"], argv=argv, input_metadata=input_metadata,
                replay_dir=replay_dir,
            )
            return raw
        finally:
            projection_path.unlink(missing_ok=True)
    # Atomic has no meaningful sink write-call/histogram observation.  Apply
    # source/replay/oracle and process/allocation checks directly, then the
    # publication proof check above; this is deliberately not a synthetic
    # hashing projection.
    return _validate_common_new_report(
        raw, arm, spec["role"], publication, spec["samples"], spec["warmups"],
        build["binary"], argv, replay_dir, input_metadata, tmpdir, path,
    )


def _artifact_meta(path: Path) -> dict[str, Any]:
    return _meta(path)


def _validate_terminal(started: dict[str, Any], terminal: dict[str, Any], spec: Mapping[str, Any],
                       build: dict[str, Any], protocol: Mapping[str, Any], protocol_hash: str,
                       directory: Path, *, require_pass: bool = True) -> None:
    started_fields = {
        "schema", "version", "status", "attempt", "run", "protocol", "build", "binary",
        "source_manifest", "argv", "cwd", "environment", "machine", "started_utc",
        "timeout_seconds", "tmpdir", "private_root", "input_file", "replay_dir", "driver",
    }
    terminal_fields = started_fields | {
        "exit_code", "timed_out", "termination", "launch_error", "validation_error",
        "finished_utc", "artifacts", "missing_artifacts", "cleanup", "started_artifact",
    }
    _exact(started, started_fields, f"{directory}/started.json")
    _exact(terminal, terminal_fields, f"{directory}/terminal.json")
    require(started["schema"] == CAPTURE_SCHEMA and started["version"] == VERSION
            and started["status"] == "running", f"{directory}: started receipt differs")
    require(terminal["schema"] == TERMINAL_SCHEMA and terminal["version"] == VERSION,
            f"{directory}: terminal receipt differs")
    for key in started_fields - {"schema", "status"}:
        require(terminal[key] == started[key], f"{directory}: terminal {key} changed")
    run_fields = ("ordinal", "repeat", "phase", "publication", "role", "arm", "workload",
                  "route", "input_mode", "samples", "warmups", "label")
    require(started["attempt"] == spec["attempt"]
            and started["run"] == {key: spec[key] for key in run_fields},
            f"{directory}: run specification differs")
    require(started["protocol"] == {"path": str(PROTOCOL_FILE), "sha256": protocol_hash},
            f"{directory}: protocol binding differs")
    require(started["build"] == _build_binding(build)
            and started["binary"] == build["binary"]
            and started["source_manifest"] == build["source_manifest"],
            f"{directory}: build/source binding differs")
    require(started["cwd"] == str(REPO), f"{directory}: cwd differs")
    require(started["machine"] == protocol["machine"], f"{directory}: machine binding differs")
    require(started["driver"] == protocol["driver"], f"{directory}: driver binding differs")
    pilot = bool(spec.get("pilot", False))
    private = Path(_text(started["private_root"], f"{directory}.private_root"))
    require(private == _run_root(spec["attempt"], spec["label"], pilot=pilot),
            f"{directory}: private root differs")
    tmpdir = Path(_text(started["tmpdir"], f"{directory}.tmpdir"))
    require(tmpdir == private / "tmp", f"{directory}: private TMPDIR differs")
    require(started["environment"] == _expected_environment(protocol, tmpdir)
            and terminal["environment"] == started["environment"],
            f"{directory}: child environment differs")
    replay_dir = started["replay_dir"]
    if ARM_BY_ID[spec["arm"]].get("route") == "file_store":
        require(replay_dir == str(private / "replay"), f"{directory}: replay path differs")
    else:
        require(replay_dir is None, f"{directory}: non-file-store route has replay path")
    input_file = started["input_file"]
    if ARM_BY_ID[spec["arm"]].get("input_mode") == "file":
        require(isinstance(input_file, dict), f"{directory}: file input metadata missing")
        _hash(input_file.get("sha256"), f"{directory}.input_file.sha256")
        _uint(input_file.get("bytes"), f"{directory}.input_file.bytes", positive=True)
        input_path = Path(_text(input_file.get("path"), f"{directory}.input_file.path"))
        require(input_path == private / "input" / input_path.name,
                f"{directory}: private input escaped root")
        if terminal["cleanup"]["status"] == "pass":
            require(not input_path.exists(), f"{directory}: private input was not cleaned")
        else:
            require(input_path.is_file() and _meta(input_path)["sha256"] == input_file["sha256"],
                    f"{directory}: preserved failure input custody differs")
    else:
        require(input_file is None, f"{directory}: non-file arm has input metadata")
    require(started["argv"] == terminal["argv"]
            and started["argv"] == _command(spec, build, directory / "report.json",
                                             directory / "resource.txt",
                                             Path(replay_dir) if replay_dir else None,
                                             input_file),
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
    _exact(cleanup, ("schema", "status", "root", "removed", "remaining", "preserved", "inventory"),
           f"{directory}.cleanup")
    require(cleanup["schema"] == CLEANUP_SCHEMA and cleanup["root"] == str(private),
            f"{directory}: private cleanup identity failed")
    if cleanup["status"] == "pass":
        require(cleanup["preserved"] is False and cleanup["remaining"] == []
                and cleanup["inventory"] == [], f"{directory}: private cleanup failed")
    else:
        require(cleanup["status"] in ("preserved_failure", "failed")
                and cleanup["preserved"] is True and isinstance(cleanup["inventory"], list),
                f"{directory}: private failure custody is malformed")
        if cleanup["status"] == "preserved_failure":
            require(cleanup["remaining"] == [str(private)],
                    f"{directory}: preserved failure tree identity differs")
            require(private.is_dir() and _failure_inventory(private) == cleanup["inventory"],
                    f"{directory}: preserved failure inventory changed")
    require(terminal["started_artifact"] == _artifact_meta(directory / "started.json"),
            f"{directory}: started artifact hash differs")
    artifact_names = ("stdout.txt", "stderr.txt", "resource.txt", "report.json", "replay-cleanup.json")
    artifacts = terminal["artifacts"]
    expected_artifact_names = set(artifact_names) if not terminal["missing_artifacts"] else {
        name for name in artifact_names if name not in terminal["missing_artifacts"]
    }
    require(set(artifacts) == expected_artifact_names,
            f"{directory}.artifacts: retained inventory differs")
    for name in expected_artifact_names:
        _exact(artifacts[name], ("path", "bytes", "sha256"), f"{directory}.artifacts.{name}")
        require(Path(artifacts[name]["path"]) == directory / name,
                f"{directory}: artifact escaped capture directory")
        require(_artifact_meta(directory / name) == artifacts[name],
                f"{directory}: artifact changed: {name}")
    require(_read_json(directory / "replay-cleanup.json") == cleanup,
            f"{directory}: cleanup receipt differs")
    require({item.name for item in directory.iterdir()} == {
        "started.json", "terminal.json", *expected_artifact_names},
            f"{directory}: retained artifact inventory differs")
    if require_pass:
        _require_success_terminal(terminal, str(directory))


def _require_success_terminal(terminal: Mapping[str, Any], path: str = "terminal") -> None:
    require(terminal.get("status") == "pass" and terminal.get("exit_code") == 0
            and terminal.get("timed_out") is False and terminal.get("termination") is None
            and terminal.get("launch_error") is None and terminal.get("validation_error") is None
            and terminal.get("missing_artifacts") == [], f"{path}: terminal did not pass")


def _launch(spec: Mapping[str, Any], build: dict[str, Any], protocol: dict[str, Any],
            protocol_hash: str, timeout_seconds: int) -> Path:
    require(timeout_seconds == DEFAULT_TIMEOUT_SECONDS,
            "formal timeout is fixed at 1800 seconds")
    attempt = _attempt(spec["attempt"])
    pilot = bool(spec.get("pilot", False))
    directory = _capture_dir(attempt, spec["label"], pilot=pilot)
    require(not directory.exists(), f"refusing to replace existing capture: {directory}")
    directory.mkdir(parents=True, exist_ok=False)
    private = _run_root(attempt, spec["label"], pilot=pilot)
    require(not private.exists(), f"refusing to replace private scratch: {private}")
    (private / "tmp").mkdir(parents=True)
    arm = ARM_BY_ID[spec["arm"]]
    replay_dir = private / "replay" if arm.get("route") == "file_store" else None
    if replay_dir is not None:
        replay_dir.mkdir()
    input_metadata = _private_input(arm, private)
    report = directory / "report.json"
    resource = directory / "resource.txt"
    stdout = directory / "stdout.txt"
    stderr = directory / "stderr.txt"
    stdout.touch()
    stderr.touch()
    argv = _command(spec, build, report, resource, replay_dir, input_metadata)
    environment = _expected_environment(protocol, private / "tmp")
    started = {
        "schema": CAPTURE_SCHEMA, "version": VERSION, "status": "running",
        "attempt": attempt,
        "run": {key: spec[key] for key in ("ordinal", "repeat", "phase", "publication", "role",
                                             "arm", "workload", "route", "input_mode", "samples",
                                             "warmups", "label")},
        "protocol": {"path": str(PROTOCOL_FILE), "sha256": protocol_hash},
        "build": _build_binding(build), "binary": build["binary"],
        "source_manifest": build["source_manifest"], "argv": argv, "cwd": str(REPO),
        "environment": environment, "machine": protocol["machine"], "started_utc": _now(),
        "timeout_seconds": timeout_seconds, "tmpdir": str(private / "tmp"),
        "private_root": str(private), "input_file": input_metadata,
        "replay_dir": str(replay_dir) if replay_dir is not None else None,
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
            process = subprocess.Popen(argv, cwd=REPO, env=environment,
                                       stdin=subprocess.DEVNULL, stdout=out, stderr=err,
                                       start_new_session=True)
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
            _validate_report(report, spec, build, argv, input_metadata, replay_dir, private / "tmp")
            if input_metadata is not None:
                input_path = Path(input_metadata["path"])
                actual_input = _meta(input_path)
                require(actual_input == {"path": input_metadata["path"],
                                        "bytes": input_metadata["bytes"],
                                        "sha256": input_metadata["sha256"]},
                        "private input changed after report validation")
                input_path.unlink()
                input_path.parent.rmdir()
        except Exception as error:
            validation_error = f"{type(error).__name__}: {error}"
    try:
        preserve_failure = (exit_code != 0 or timed_out or launch_error is not None
                            or validation_error is not None)
        cleanup = _cleanup_private(private, preserve_failure=preserve_failure)
    except (MeasureError, OSError) as error:
        cleanup = {"schema": CLEANUP_SCHEMA, "status": "failed", "root": str(private),
                   "removed": [], "remaining": [str(error)], "preserved": True,
                   "inventory": _failure_inventory(private) if private.exists() else []}
    _write_new(directory / "replay-cleanup.json", cleanup)
    required = (stdout, stderr, resource, report, directory / "replay-cleanup.json")
    missing = [path.name for path in required if not path.is_file()]
    artifacts = {path.name: _meta(path) for path in required if path.is_file()}
    passed = (exit_code == 0 and not timed_out and termination is None
              and launch_error is None and validation_error is None and not missing
              and cleanup["status"] == "pass")
    terminal = dict(started)
    terminal.update({
        "schema": TERMINAL_SCHEMA, "status": "pass" if passed else "failed",
        "exit_code": exit_code, "timed_out": timed_out, "termination": termination,
        "launch_error": launch_error, "validation_error": validation_error,
        "finished_utc": _now(), "artifacts": artifacts, "missing_artifacts": missing,
        "cleanup": cleanup, "started_artifact": _meta(directory / "started.json"),
    })
    _write_new(directory / "terminal.json", terminal)
    if not passed:
        fail(f"{spec['label']} failed; retained terminal receipt: {directory / 'terminal.json'}")
    return directory / "terminal.json"


def capture_one(attempt: str, spec: Mapping[str, Any], *, timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS) -> Path:
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
            _launch(spec, builds[_build_key(spec["phase"], spec["role"])],
                    protocol, protocol_hash, timeout_seconds)
    # The lock spans every block and every child, including atomic routes.
    _cpu_lock(run)


def capture_pilot(attempt: str, *, timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS) -> None:
    """Run the separate 3-sample smoke inventory under the same CPU lock."""

    protocol, protocol_hash, builds = _load_protocol()
    attempt = _attempt(attempt)
    inventory = [dict(item, attempt=attempt) for item in protocol["pilot_runs"]]

    def run() -> None:
        for spec in inventory:
            _launch(spec, builds[_build_key(spec["phase"], spec["role"])],
                    protocol, protocol_hash, timeout_seconds)

    _cpu_lock(run)


def _percentiles(values: Iterable[int | float]) -> dict[str, Any]:
    vector = list(values)
    require(vector, "cannot summarize empty metric vector")
    ordered = sorted(vector)
    n = len(ordered)
    p50 = ((ordered[n // 2 - 1] + ordered[n // 2]) / 2
           if n % 2 == 0 else ordered[n // 2])
    return {"n": n, "min": ordered[0], "max": ordered[-1],
            "mean": statistics.fmean(vector), "p50": p50,
            "p95": ordered[max(0, math.ceil(n * 0.95) - 1)],
            "p99": ordered[max(0, math.ceil(n * 0.99) - 1)]}


def _bootstrap_median(values: list[int | float], *, repetitions: int = BOOTSTRAP_REPETITIONS,
                      seed: int = BOOTSTRAP_SEED, method: str = "independent_within_child_bootstrap_median") -> dict[str, Any]:
    require(values, "cannot bootstrap empty metric vector")
    require(type(repetitions) is int and repetitions > 0, "bootstrap repetitions must be positive")
    require(type(seed) is int and seed >= 0, "bootstrap seed must be nonnegative")
    rng = seed & 0xFFFFFFFFFFFFFFFF
    medians: list[float] = []
    count = len(values)
    for _ in range(repetitions):
        sample = [values[((rng := (rng * 6364136223846793005 + 1442695040888963407)
                           & 0xFFFFFFFFFFFFFFFF) >> 32) % count] for _ in values]
        sample.sort()
        middle = len(sample) // 2
        medians.append((sample[middle - 1] + sample[middle]) / 2
                       if len(sample) % 2 == 0 else sample[middle])
    medians.sort()
    low = max(0, int(repetitions * 0.025))
    high = min(repetitions - 1, int(repetitions * 0.975) - 1)
    return {"method": method, "seed": seed, "resamples": repetitions,
            "confidence": 0.95, "samples": len(values),
            "median": _percentiles(values)["p50"], "ci_low": medians[low],
            "ci_high": medians[high]}


def paired_bootstrap(values: list[int | float], *, repetitions: int = BOOTSTRAP_REPETITIONS,
                     seed: int = BOOTSTRAP_SEED) -> dict[str, Any]:
    result = _bootstrap_median(values, repetitions=repetitions, seed=seed,
                               method="paired_repeat_block_bootstrap_median")
    result["blocks"] = len(values)
    result["uncertainty_scope"] = (
        "two repeat block deltas; before and after child sample vectors are not paired observations")
    return result


def _within_child_bootstrap(values: list[int | float], *, repetitions: int = BOOTSTRAP_REPETITIONS,
                            seed: int = BOOTSTRAP_SEED) -> dict[str, Any]:
    result = _bootstrap_median(values, repetitions=repetitions, seed=seed)
    result["uncertainty_scope"] = "descriptive within-child sample resampling"
    return result


ALLOCATION_RAW_FIELDS = (
    "allocation_calls", "deallocation_calls", "reallocation_calls", "failed_allocation_calls",
    "allocated_bytes", "deallocated_bytes", "live_bytes_before", "live_bytes_after",
    "peak_live_bytes_before", "peak_live_bytes_after", "region_peak_live_bytes",
)
ALLOCATION_FIELDS = ALLOCATION_RAW_FIELDS + ("allocation_peak_increment_bytes",)


def _allocation_metrics(entry: dict[str, Any], rows: list[dict[str, Any]]) -> dict[str, Any]:
    role = entry["spec"]["role"]
    if role == "normal":
        require(all(row.get("allocation") is None for row in rows),
                f"{entry['directory']}: normal report contains allocator metrics")
        return {"availability": "unavailable", "scope": "operation_global_system_allocator",
                "vectors": None, "percentiles": None, "median_ci": None,
                "derived_fields": {
                    "allocation_peak_increment_bytes":
                    "region_peak_live_bytes - live_bytes_before (unavailable for normal role)"}}
    vectors = {name: [] for name in ALLOCATION_FIELDS}
    base = _sealed().old_routes.base
    for index, row in enumerate(rows):
        allocation = row.get("allocation")
        path = f"{entry['directory']}.rows[{index}].allocation"
        base._check_allocator_sample(row, "allocator", path)
        require(isinstance(allocation, dict), f"{path}: allocator sample missing")
        for name in ALLOCATION_RAW_FIELDS:
            vectors[name].append(base._int_field(allocation.get(name), f"{path}.{name}"))
        increment = allocation["region_peak_live_bytes"] - allocation["live_bytes_before"]
        require(increment >= 0, f"{path}: region peak increment is negative")
        vectors["allocation_peak_increment_bytes"].append(increment)
    return {"availability": "measured", "scope": "operation_global_system_allocator",
            "vectors": vectors,
            "percentiles": {name: _percentiles(values) for name, values in vectors.items()},
            "median_ci": {name: _within_child_bootstrap(values) for name, values in vectors.items()},
            "derived_fields": {
                "allocation_peak_increment_bytes":
                "region_peak_live_bytes - live_bytes_before for each sample"}}


def _entry_metrics(entry: dict[str, Any]) -> dict[str, Any]:
    rows = entry["report"]["cases"][0]["samples"]
    latency = [int(row["elapsed_ns"]) for row in rows]
    rss = int(entry["resource"]["maximum_resident_set_size_(kbytes)"]) * 1024
    metrics: dict[str, Any] = {
        "latency_ns": _percentiles(latency),
        "sample_elapsed_ns": latency,
        "latency_median_ci": _within_child_bootstrap(latency),
        "rss_bytes": {"n": 1, "value": rss, "p50": rss, "p95": rss, "p99": rss,
                      "observation_scope": "whole_child_gnu_time_single_observation",
                      "median_ci": {"availability": "unavailable",
                                    "scope": "one whole-child observation"}},
        "allocation": _allocation_metrics(entry, rows),
        "publication": {"route": entry["spec"]["publication"]},
    }
    publication = entry["spec"]["publication"]
    if publication != "hashing_sink":
        proof_bytes = [int(row["publication"]["timed_candidate_artifact_bytes"]) for row in rows]
        proof_hashes = [row["publication"]["timed_candidate_artifact_sha256"] for row in rows]
        metrics["publication"].update({
            "timed_candidate_artifact_bytes": _percentiles(proof_bytes),
            "timed_candidate_artifact_sha256_unique": sorted(set(proof_hashes)),
            "timed_candidate_matches_oracle": all(
                row["publication"]["timed_candidate_matches_oracle"] is True for row in rows),
            "verification_scope": sorted(set(row["publication"]["verification_scope"] for row in rows)),
            "atomic_readback": publication == "atomic_path",
        })
    return metrics


def _relative(before: int | float, delta: int | float) -> float | None:
    return None if before == 0 else float(delta / abs(before) * 100.0)


def _metric_specs(role: str) -> list[tuple[str, Callable[[dict[str, Any]], int | float]]]:
    result: list[tuple[str, Callable[[dict[str, Any]], int | float]]] = []
    for percentile in ("p50", "p95", "p99"):
        result.append((f"latency_ns.{percentile}",
                       lambda metrics, percentile=percentile: metrics["latency_ns"][percentile]))
    for percentile in ("p50", "p95", "p99"):
        result.append((f"rss_bytes.{percentile}",
                       lambda metrics, percentile=percentile: metrics["rss_bytes"][percentile]))
    if role == "allocator":
        for name in ALLOCATION_FIELDS:
            for percentile in ("p50", "p95", "p99"):
                result.append((f"allocation.{name}.{percentile}",
                               lambda metrics, name=name, percentile=percentile:
                               metrics["allocation"]["percentiles"][name][percentile]))
    return result


def _paired_comparisons(entries: list[dict[str, Any]]) -> list[dict[str, Any]]:
    by_key: dict[tuple[str, str, str, int], dict[str, dict[str, Any]]] = {}
    for entry in entries:
        spec = entry["spec"]
        if spec["publication"] != "hashing_sink":
            continue
        key = (spec["arm"], spec["role"], spec["publication"], spec["repeat"])
        by_key.setdefault(key, {})[spec["phase"]] = entry["_metrics"]
    groups: dict[tuple[str, str, str], list[tuple[int, dict[str, Any], dict[str, Any]]]] = {}
    for (arm, role, publication, repeat), values in by_key.items():
        require(set(values) == set(PHASES), f"missing legacy phase pair for {arm}/{role}/r{repeat}")
        groups.setdefault((arm, role, publication), []).append(
            (repeat, values["before"], values["after"]))
    result: list[dict[str, Any]] = []
    for (arm, role, publication), blocks in sorted(groups.items()):
        require({repeat for repeat, _, _ in blocks} == set(REPEATS),
                f"legacy repeat inventory incomplete for {arm}/{role}")
        for metric, getter in _metric_specs(role):
            raw_blocks = []
            deltas: list[int | float] = []
            for repeat, before, after in sorted(blocks):
                before_value, after_value = getter(before), getter(after)
                delta = after_value - before_value
                relative = _relative(before_value, delta)
                block = {
                    "repeat": repeat, "before": before_value, "after": after_value,
                    "absolute_delta": delta, "relative_percent": relative,
                    "adverse_flag": relative is not None and relative > ADVERSE_THRESHOLD_PERCENT,
                    "threshold_percent": ADVERSE_THRESHOLD_PERCENT,
                }
                if metric.endswith("p50") and metric.startswith("latency_ns"):
                    block["before_within_child_median_ci"] = before["latency_median_ci"]
                    block["after_within_child_median_ci"] = after["latency_median_ci"]
                elif metric.startswith("allocation.") and metric.endswith("p50"):
                    field = metric.split(".")[1]
                    block["before_within_child_median_ci"] = before["allocation"]["median_ci"][field]
                    block["after_within_child_median_ci"] = after["allocation"]["median_ci"][field]
                raw_blocks.append(block)
                deltas.append(delta)
            median_before = statistics.median(block["before"] for block in raw_blocks)
            median_after = statistics.median(block["after"] for block in raw_blocks)
            median_delta = statistics.median(deltas)
            aggregate_relative = _relative(median_before, median_delta)
            result.append({
                "arm": arm, "role": role, "publication": publication, "metric": metric,
                "blocks": raw_blocks, "delta_values": deltas,
                "median_before": median_before, "median_after": median_after,
                "median_delta": median_delta, "absolute_delta": median_delta,
                "relative_percent": aggregate_relative,
                # A single adverse repeat is retained even if the median hides it.
                "adverse_flag": any(block["adverse_flag"] for block in raw_blocks),
                "aggregate_adverse_flag": aggregate_relative is not None
                    and aggregate_relative > ADVERSE_THRESHOLD_PERCENT,
                "individual_adverse_blocks": [block["repeat"] for block in raw_blocks
                                               if block["adverse_flag"]],
                "threshold_percent": ADVERSE_THRESHOLD_PERCENT,
                "bootstrap": paired_bootstrap(deltas),
                "interpretation": "descriptive before/after repeat-block comparison; no causal claim",
            })
    return result


def _repeat_visibility(entries: list[dict[str, Any]]) -> list[dict[str, Any]]:
    grouped: dict[tuple[str, str, str, str], dict[int, dict[str, Any]]] = {}
    for entry in entries:
        spec = entry["spec"]
        key = (spec["phase"], spec["publication"], spec["role"], spec["arm"])
        require(spec["repeat"] not in grouped.setdefault(key, {}), f"duplicate repeat for {key}")
        grouped[key][spec["repeat"]] = entry["_metrics"]
    result = []
    for (phase, publication, role, arm), repeats in sorted(grouped.items()):
        require(set(repeats) == set(REPEATS), f"repeat inventory incomplete for {phase}/{publication}/{role}/{arm}")
        first, second = repeats[1], repeats[2]
        metrics = {}
        for name, getter in _metric_specs(role):
            first_value, second_value = getter(first), getter(second)
            delta = second_value - first_value
            relative = _relative(first_value, delta)
            metrics[name] = {"repeat1": first_value, "repeat2": second_value,
                             "absolute_delta": delta, "relative_percent": relative,
                             "flag_over_5_percent": relative is not None
                             and abs(relative) > ADVERSE_THRESHOLD_PERCENT}
        result.append({"phase": phase, "publication": publication, "role": role, "arm": arm,
                       "metrics": metrics,
                       "allocation": {
                           "availability": first["allocation"]["availability"],
                           "scope": first["allocation"]["scope"],
                           "normal_unavailable": first["allocation"]["availability"] == "unavailable",
                       },
                       "scope": "descriptive repeat variance within one phase/publication/role/arm"})
    return result


def _collect_lane(attempt: str, protocol: dict[str, Any], builds: dict[str, dict[str, Any]],
                  *, pilot: bool = False) -> list[dict[str, Any]]:
    attempt = _attempt(attempt)
    root = CAPTURE_ROOT / "pilots" / attempt if pilot else CAPTURE_ROOT / attempt
    require(root.is_dir() and not root.is_symlink(), f"capture attempt is missing: {root}")
    expected = protocol["pilot_runs"] if pilot else protocol["formal_runs"]
    require({item.name for item in root.iterdir()} == {item["label"] for item in expected},
            f"{root}: capture inventory differs")
    entries: list[dict[str, Any]] = []
    protocol_hash = _sha(PROTOCOL_FILE)
    for expected_spec in expected:
        spec = dict(expected_spec, attempt=attempt)
        directory = root / spec["label"]
        started = _read_json(directory / "started.json")
        terminal = _read_json(directory / "terminal.json")
        build = builds[_build_key(spec["phase"], spec["role"])]
        _validate_terminal(started, terminal, spec, build, protocol, protocol_hash, directory)
        raw = _read_json(directory / "report.json")
        _validate_report(directory / "report.json", spec, build, started["argv"],
                         started["input_file"], Path(started["replay_dir"]) if started["replay_dir"] else None,
                         Path(started["tmpdir"]))
        resource = _resource(directory / "resource.txt")
        entries.append({"spec": spec, "directory": directory, "started": started,
                        "terminal": terminal, "report": raw, "resource": resource,
                        "started_at": _timestamp(started["started_utc"], "started_utc"),
                        "finished_at": _timestamp(terminal["finished_utc"], "finished_utc")})
    ordered = _chronological_entries(entries, expected, attempt)
    private_root = TEMP / "runs" / "pilots" / attempt if pilot else TEMP / "runs" / attempt
    require(not private_root.exists(), f"private capture scratch remains: {private_root}")
    return ordered


def _collect(attempt: str, protocol: dict[str, Any], builds: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    return _collect_lane(attempt, protocol, builds, pilot=False)


def _chronological_entries(entries: list[dict[str, Any]], expected: list[dict[str, Any]],
                           attempt: str) -> list[dict[str, Any]]:
    """Match actual started timestamps to the frozen order and reject overlap."""

    ordered = sorted(entries, key=lambda item: item["started_at"])
    previous_finished: _datetime.datetime | None = None
    for expected_spec, entry in zip(expected, ordered):
        require(entry["spec"] == dict(expected_spec, attempt=attempt),
                "actual receipt chronology differs from frozen formal_runs")
        if previous_finished is not None:
            require(entry["started_at"] >= previous_finished,
                    "capture children overlap; shared CPU lock/order custody is invalid")
        previous_finished = entry["finished_at"]
    require(len(ordered) == len(expected), "actual receipt count differs")
    return ordered


def analyze_data(entries: list[dict[str, Any]], *, protocol: dict[str, Any] | None = None) -> dict[str, Any]:
    require(len(entries) == 288, "analysis requires exactly 288 retained children")
    children = []
    for entry in entries:
        metrics = _entry_metrics(entry)
        entry["_metrics"] = metrics
        spec = entry["spec"]
        children.append({"label": spec["label"], "phase": spec["phase"],
                         "publication": spec["publication"], "role": spec["role"],
                         "arm": spec["arm"], "repeat": spec["repeat"], "metrics": metrics})
    paired = _paired_comparisons(entries)
    repeat_variance = _repeat_visibility(entries)
    capabilities = [item for item in children
                    if item["phase"] == "after" and item["publication"] in ("counting_sink", "atomic_path")]
    return {
        "schema": ANALYSIS_SCHEMA, "version": VERSION, "case": CASE,
        "children": children, "child_count": len(children),
        "sample_count": len(children) * FORMAL_SAMPLES,
        "after_only_capabilities": capabilities,
        "paired_legacy_comparisons": paired,
        "repeat_variance": repeat_variance,
        "adverse_flags": [item for item in paired if item["adverse_flag"]],
        "claims": {
            "scope": "descriptive full-lifecycle latency/RSS/allocation and publication evidence",
            "legacy_comparison": "before/after default hashing_sink only",
            "after_only_routes": ["counting_sink", "atomic_path"],
            "atomic_speedup_claim": False,
            "threshold_percent": ADVERSE_THRESHOLD_PERCENT,
            "bootstrap": {
                "repetitions": BOOTSTRAP_REPETITIONS, "seed": BOOTSTRAP_SEED,
                "confidence": 0.95,
                "paired_scope": "two repeat blocks; no sample-level cross-process pairing",
                "within_child_scope": "descriptive independent resampling of each child sample vector",
            },
            "sample_pseudoreplication": False,
            "historical_flags_resolved": False,
        },
        "protocol": protocol,
    }


def analyze(attempt: str) -> Path:
    protocol, _, builds = _load_protocol()
    entries = _collect(attempt, protocol, builds)
    result = analyze_data(entries, protocol={"path": str(PROTOCOL_FILE), "sha256": _sha(PROTOCOL_FILE)})
    output = ROOT / "analysis" / f"{_attempt(attempt)}.json"
    _write_new(output, result)
    return output


def verify(attempt: str) -> Path:
    protocol, _, builds = _load_protocol()
    entries = _collect(attempt, protocol, builds)
    result = {"schema": "docx-replayable-tail-publication-verification-v1", "version": VERSION,
              "attempt": _attempt(attempt), "children": len(entries),
              "samples": len(entries) * FORMAL_SAMPLES, "status": "pass",
              "protocol": {"path": str(PROTOCOL_FILE), "sha256": _sha(PROTOCOL_FILE)}}
    output = ROOT / "verification" / f"{_attempt(attempt)}.json"
    _write_new(output, result)
    return output


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    freeze = sub.add_parser("freeze")
    freeze.add_argument("--builds", type=Path, default=BUILDS_FILE)
    plan = sub.add_parser("plan")
    plan.add_argument("--json", type=Path)
    for name in ("capture", "pilot", "analyze", "verify"):
        command = sub.add_parser(name)
        command.add_argument("--attempt", required=True)
        if name in ("capture", "pilot"):
            command.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT_SECONDS)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        if args.command == "plan":
            value = {"schema": PROTOCOL_SCHEMA, "version": VERSION, "case": CASE,
                     "formal_runs": formal_inventory(), "expected_children": 288,
                     "expected_samples": 8640}
            if args.json:
                _write_new(args.json, value)
                print(args.json)
            else:
                json.dump(value, sys.stdout, indent=2, sort_keys=True)
                print()
        elif args.command == "freeze":
            print(create_freeze(args.builds))
        elif args.command == "capture":
            require(args.timeout == DEFAULT_TIMEOUT_SECONDS,
                    "formal capture timeout is fixed at 1800 seconds")
            capture_all(args.attempt, timeout_seconds=args.timeout)
        elif args.command == "pilot":
            require(args.timeout == DEFAULT_TIMEOUT_SECONDS,
                    "pilot capture timeout is fixed at 1800 seconds")
            capture_pilot(args.attempt, timeout_seconds=args.timeout)
        elif args.command == "analyze":
            print(analyze(args.attempt))
        elif args.command == "verify":
            print(verify(args.attempt))
        return 0
    except MeasureError as error:
        print(f"0497: FAIL: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
