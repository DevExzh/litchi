#!/usr/bin/env python3
"""Retain and analyze the bounded 0487 OPC splice-audit comparison.

This driver owns only the comparison evidence lane.  It imports the frozen
0484 route and report validators from their sealed directory, and keeps the
new before/after custody records under ``change-0487``.  Importing this file
does not build or run a benchmark.  Captures are deliberately one child at a
time; the coordinator normally invokes the capture commands through the
0487 gate so that the shared CPU lock and source-custody receipt cover the
whole lane.

The selected matrix is intentionally small and explicit:

* three short DOCX workloads;
* deterministic owned, file, short-read, and latency input arms;
* owned memory-store and file-store arms;
* normal and allocator roles, in two forward/reverse process orders; and
* matched ``before`` and ``after`` phases.

The formal matrix is therefore 144 children (18 arms * 2 roles * 2 repeats
* 2 phases), with 30 samples and 3 warmups per child.  A pilot uses the same
arm and role inventory with 3 samples and 1 warmup.  The file-store replay
directory is created per child, checked after the report validator, and then
removed with a retained receipt.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
from pathlib import Path
import re
import signal
import subprocess
import sys
from typing import Any, Iterable, Mapping

import support
from support import ENV, ENV_KEYS, REPO, ROOT, meta, now, read, sha, snapshot, write


# The 0484 directory is a sealed input.  Put it first so that importing the
# old route driver resolves its own ``measure.py`` and ``common.py`` rather
# than the new support module or another checkout module.
ROUTE_ROOT = REPO / "docs" / "performance" / "results" / "change-0484"
BASELINE_ROOT = REPO / "docs" / "performance" / "results" / "change-0485"
if str(ROUTE_ROOT) not in sys.path:
    sys.path.insert(0, str(ROUTE_ROOT))
import analyze_routes as old_analysis  # noqa: E402  (sealed helper import)
import measure_routes as old_routes  # noqa: E402


SCHEMA = "docx-opc-splice-audit-consumed-prefix-comparison-v1"
VERSION = 1
PROTOCOL_FILE = "comparison-protocol.json"
SUMMARY_JSON = "comparison-summary.json"
SUMMARY_MD = "comparison-summary.md"
CAPTURE_SCHEMA = "docx-opc-splice-audit-consumed-prefix-capture-v1"
REPLAY_CLEANUP_SCHEMA = "docx-opc-splice-audit-replay-cleanup-v1"
GATE_SCHEMA = "docx-stream-append-gate-v1"
BUILD_SCHEMA = old_routes.base.BUILD_SCHEMA
ROLE_NAMES = tuple(old_routes.ROLES)
PHASES = ("before", "after")
REPEATS = (1, 2)
FORMAL_SAMPLES = 30
FORMAL_WARMUPS = 3
PILOT_SAMPLES = 3
PILOT_WARMUPS = 1
REVIEW_THRESHOLD_PERCENT = 5.0
DEFAULT_TIMEOUT_SECONDS = 1_800
ATTEMPT_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]*$")
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")

WORKLOADS = (
    "s64-a64-short-c64",
    "s64-a16384-short-c64",
    "s131072-a64-short-c64",
)
INPUT_MODES = ("owned", "file", "short-read", "latency")
STORE_ROUTES = ("memory_store", "file_store")

# The source diff is reviewed separately by the coordinator.  This allowlist
# prevents an OPC batching build from quietly carrying an unrelated crate or
# benchmark change while allowing focused adapter tests beside the change.
ALLOWED_SOURCE_PREFIXES = (
    "crates/litchi-opc/src/source_backed/splice.rs",
    "crates/litchi-opc/tests/source_part_splice.rs",
    "crates/litchi-opc/tests/source_part_splice_replay.rs",
    "tools/perf-baseline/src/docx_replayable_tail_append/",
)


class CompareError(RuntimeError):
    """A retained comparison protocol, capture, or analysis is invalid."""


def fail(message: str) -> None:
    raise CompareError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _attempt(value: str) -> str:
    require(ATTEMPT_RE.fullmatch(value) is not None, "attempt must be a unique path-safe token")
    return value


def _finite(value: Any, path: str = "value") -> None:
    if isinstance(value, float):
        require(math.isfinite(value), f"{path}: non-finite number")
    elif isinstance(value, dict):
        for key, child in value.items():
            _finite(child, f"{path}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            _finite(child, f"{path}[{index}]")


def _read(path: Path) -> Any:
    value = read(path)
    _finite(value, str(path))
    return value


def _object(value: Any, path: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{path}: expected an object")
    return value


def _list(value: Any, path: str) -> list[Any]:
    require(isinstance(value, list), f"{path}: expected a list")
    return value


def _string(value: Any, path: str) -> str:
    require(isinstance(value, str) and value, f"{path}: expected a non-empty string")
    return value


def _sha(value: Any, path: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None, f"{path}: invalid SHA-256")
    return value


def _integer(value: Any, path: str, *, minimum: int = 0) -> int:
    require(type(value) is int and value >= minimum, f"{path}: expected an integer >= {minimum}")
    return value


def _path_metadata(path: Path, label: str) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"{label}: missing regular file: {path}")
    details = meta(path)
    require(details["bytes"] >= 0 and SHA256_RE.fullmatch(str(details["sha256"])) is not None, f"{label}: malformed metadata")
    return {"path": str(path), **details}


def _write_text_exclusive(path: Path, value: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("x", encoding="utf-8", newline="\n") as stream:
            stream.write(value)
    except FileExistsError as error:
        raise CompareError(f"refusing to replace existing artifact: {path}") from error


def _script_hashes() -> dict[str, str]:
    local_names = ["compare.py", "support.py", "gate.py"]
    for optional in ("build_after.py", "record_environment.py"):
        if (ROOT / optional).is_file():
            local_names.append(optional)
    baseline_names = ["compare.py", "support.py", "gate.py"]
    for optional in ("build_after.py", "record_environment.py"):
        if (BASELINE_ROOT / optional).is_file():
            baseline_names.append(optional)
    old_names = (
        "measure.py",
        "measure_routes.py",
        "analyze_routes.py",
        "common.py",
        "corpus_oracle.py",
    )
    values: dict[str, str] = {}
    for name in local_names:
        path = ROOT / name
        require(path.is_file(), f"0487 comparison helper is missing: {path}")
        values[f"0487/{name}"] = sha(path)
    for name in baseline_names:
        path = BASELINE_ROOT / name
        require(path.is_file(), f"sealed 0485 helper is missing: {path}")
        values[f"0485/{name}"] = sha(path)
    for name in old_names:
        path = ROUTE_ROOT / name
        require(path.is_file(), f"sealed 0484 helper is missing: {path}")
        values[f"0484/{name}"] = sha(path)
    return values


def _old_protocol_binding() -> dict[str, Any]:
    path = ROUTE_ROOT / old_routes.ROUTE_PROTOCOL_FILE
    require(path.is_file() and not path.is_symlink(), f"sealed 0484 route protocol is missing: {path}")
    value = _read(path)
    require(value.get("schema") == old_routes.ROUTE_SCHEMA, f"{path}: sealed route protocol schema differs")
    return {"path": str(path), "sha256": sha(path), "schema": value["schema"], "version": value.get("version")}


def _baseline_protocol_binding() -> dict[str, Any]:
    path = BASELINE_ROOT / "comparison-protocol.json"
    require(path.is_file() and not path.is_symlink(), f"sealed 0485 comparison protocol is missing: {path}")
    value = _read(path)
    require(
        value.get("schema") == "docx-opc-splice-audit-comparison-v1" and value.get("version") == 1,
        f"{path}: sealed 0485 comparison protocol schema differs",
    )
    return {"path": str(path), "sha256": sha(path), "schema": value["schema"], "version": value.get("version")}


def _machine_binding() -> dict[str, Any]:
    path = ROOT / "machine.json"
    if not path.is_file():
        return {"path": "machine.json", "sha256": None, "status": "deferred"}
    value = _read(path)
    require(value.get("schema") == "docx-stream-route-machine-v1", f"{path}: machine schema differs")
    require(value.get("driver_sha256") == sha(ROOT / "record_environment.py"), f"{path}: recorder binding differs")
    require(value.get("selected_cpu") == old_routes.base.CPU, f"{path}: selected CPU differs")
    affinity = value.get("coordinator_affinity")
    require(isinstance(affinity, list) and old_routes.base.CPU in affinity, f"{path}: selected CPU is outside coordinator affinity")
    require(value.get("environment") == {key: ENV[key] for key in ENV_KEYS}, f"{path}: environment differs")
    scratch = _object(value.get("scratch"), f"{path}.scratch")
    require(scratch.get("path") == str(support.TEMP.resolve()), f"{path}: scratch capability differs")
    for command_name in ("cpu", "rustc", "cargo", "time", "scratch_mount"):
        command = _object(_object(value.get("commands"), f"{path}.commands").get(command_name), f"{path}.commands.{command_name}")
        require(command.get("status") == "pass" and command.get("exit_code") == 0, f"{path}: required command {command_name} did not pass")
        for stream in ("stdout", "stderr"):
            output = command.get(stream)
            require(isinstance(output, str), f"{path}.commands.{command_name}.{stream}: output is missing")
            require(hashlib.sha256(output.encode("utf-8")).hexdigest() == command.get(f"{stream}_sha256"), f"{path}: {command_name} {stream} hash differs")
    cache_policy = _object(value.get("cache_policy"), f"{path}.cache_policy")
    require(cache_policy.get("cold_cache_claim") is False, f"{path}: cold-cache claim is not explicitly false")
    return {"path": "machine.json", "sha256": sha(path), "status": "ready"}


def _fixture_bindings() -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for workload in WORKLOADS:
        axis_label = f"axis-input-file-{workload}"
        arm = old_routes.AXIS_ARM_BY_LABEL[axis_label]
        relative = str(arm["input_file"])
        path = (ROUTE_ROOT / relative).resolve()
        details = _path_metadata(path, f"file fixture {workload}")
        result[workload] = {
            "path": relative,
            "absolute_path": str(path),
            "bytes": details["bytes"],
            "sha256": details["sha256"],
            "old_arm": axis_label,
        }
    return result


def _arm(arm_id: str, family: str, workload: str, input_mode: str, *, route: str, old_arm: str | None) -> dict[str, Any]:
    case = old_routes.CASE_BY_LABEL[workload]
    result: dict[str, Any] = {
        "id": arm_id,
        "family": family,
        "workload": workload,
        "source_count": case["source_count"],
        "authored_count": case["authored_count"],
        "chunk_mode": case["chunk_mode"],
        "text_mode": case["text_mode"],
        "route": route,
        "input_mode": input_mode,
        "old_route_case": workload,
        "old_axis_arm": old_arm,
        "sink_write_bytes": old_routes.SINK_WRITE_BYTES,
        "compression": old_routes.COMPRESSION,
    }
    if family == "route":
        spec = old_routes.ROUTE_BY_NAME[route]
        result.update({
            "provider": route,
            "cli_provider": spec.cli_provider,
            "report_provider": spec.report_provider,
            "input_backing": old_routes.INPUT_BACKINGS[input_mode],
            "source_contract": old_routes.SOURCE_CONTRACTS[input_mode],
            "input_identity_validation": old_routes.INPUT_IDENTITY_VALIDATION[input_mode],
            "expected_authored_opens": spec.expected_authored_opens,
            "expected_replay_opens": spec.expected_replay_opens,
            "replay_max_bytes": spec.replay_max_bytes,
            "replay_sync": spec.replay_sync,
            "storage_profile": {
                "provider": spec.name,
                "replay_max_bytes": spec.replay_max_bytes,
                "replay_sync": spec.replay_sync,
                "replay_store": spec.name if spec.store else "none",
            },
        })
    else:
        require(input_mode != "owned", "axis family is only used for non-owned input arms")
        selected = old_routes.AXIS_ARM_BY_LABEL[old_arm or ""]
        result.update({
            "provider": selected["provider"],
            "cli_provider": selected["cli_provider"],
            "report_provider": selected["report_provider"],
            "input_backing": selected["input_backing"],
            "source_contract": selected["source_contract"],
            "input_identity_validation": selected["input_identity_validation"],
            "expected_authored_opens": selected["expected_authored_opens"],
            "expected_replay_opens": selected["expected_replay_opens"],
            "replay_max_bytes": selected["replay_max_bytes"],
            "replay_sync": selected["replay_sync"],
            "input_profile": dict(selected["input_profile"]),
            "storage_profile": dict(selected["storage_profile"]),
            "compression_profile": dict(selected["compression_profile"]),
            "value": selected["value"],
        })
    return result


def _arms() -> tuple[dict[str, Any], ...]:
    values: list[dict[str, Any]] = []
    for workload in WORKLOADS:
        values.append(_arm(f"deterministic-owned-{workload}", "route", workload, "owned", route="deterministic", old_arm=None))
        for input_mode in ("file", "short-read", "latency"):
            old_arm = f"axis-input-{input_mode}-{workload}"
            values.append(_arm(f"deterministic-{input_mode}-{workload}", "axis", workload, input_mode, route="deterministic", old_arm=old_arm))
    for route in STORE_ROUTES:
        for workload in WORKLOADS:
            values.append(_arm(f"{route}-owned-{workload}", "route", workload, "owned", route=route, old_arm=None))
    return tuple(values)


ARMS = _arms()
ARM_BY_ID = {arm["id"]: arm for arm in ARMS}
if len(ARMS) != 18 or len(ARM_BY_ID) != len(ARMS):  # pragma: no cover - protocol invariant
    raise RuntimeError("0487 arm matrix did not contain exactly 18 unique arms")


def _runs(*, pilot: bool) -> list[dict[str, Any]]:
    values: list[dict[str, Any]] = []
    for phase in PHASES:
        pairs = [(arm["id"], role) for arm in ARMS for role in ROLE_NAMES]
        sequences = [(None, pairs)] if pilot else [(1, pairs), (2, list(reversed(pairs)))]
        for repeat, sequence in sequences:
            for arm_id, role in sequence:
                if pilot:
                    label = f"{phase}-pilot-{role}-{arm_id}"
                else:
                    label = f"{phase}-r{repeat}-{role}-{arm_id}"
                values.append({
                    "kind": "pilot" if pilot else "formal",
                    "phase": phase,
                    "label": label,
                    "arm": arm_id,
                    "role": role,
                    "repeat": repeat,
                    "workload": ARM_BY_ID[arm_id]["workload"],
                    "route": ARM_BY_ID[arm_id]["route"],
                    "input_mode": ARM_BY_ID[arm_id]["input_mode"],
                })
    return values


def _build_command(role: str) -> list[str]:
    return list(old_routes.base._build_command(role))


def _build_binding(phase: str, role: str) -> dict[str, Any]:
    if phase == "before":
        root_name = "0485"
        root = BASELINE_ROOT
        relative = Path(f"build-{role}.json")
    else:
        root_name = "0487"
        root = ROOT
        relative = Path(f"build-{role}.json")
    path = root / relative
    record_hash = sha(path) if path.is_file() else None
    return {
        "phase": phase,
        "role": role,
        "root": root_name,
        "record_path": relative.as_posix(),
        "record_sha256": record_hash,
        "record_schema": BUILD_SCHEMA,
        "command": _build_command(role),
        "source_binding": "sealed_0485_after" if phase == "before" else "after_build_gate_and_source_manifest",
    }


def _protocol_value(*, include_freeze_source: bool) -> dict[str, Any]:
    fixtures = _fixture_bindings()
    build_bindings = {
        phase: {role: _build_binding(phase, role) for role in ROLE_NAMES}
        for phase in PHASES
    }
    value: dict[str, Any] = {
        "schema": SCHEMA,
        "version": VERSION,
        "change": 487,
        "claim_authorized": False,
        "performance_claim": "none_until_matched_formal_analysis",
        "comparison": "before_after_opc_splice_audit_consumed_prefix_retention",
        "scope": "three exact DOCX workloads, deterministic input profiles, and owned memory/file replay stores",
        "cpu": old_routes.base.CPU,
        "sink_write_bytes": old_routes.SINK_WRITE_BYTES,
        "formal": {"samples": FORMAL_SAMPLES, "warmups": FORMAL_WARMUPS},
        "pilot": {"samples": PILOT_SAMPLES, "warmups": PILOT_WARMUPS},
        "roles": list(ROLE_NAMES),
        "repeats": list(REPEATS),
        "phases": list(PHASES),
        "workloads": list(WORKLOADS),
        "arms": [dict(arm) for arm in ARMS],
        "expected_formal_processes": len(_runs(pilot=False)),
        "expected_pilot_processes": len(_runs(pilot=True)),
        "formal_runs": _runs(pilot=False),
        "pilot_runs": _runs(pilot=True),
        "fixtures": fixtures,
        "build_bindings": build_bindings,
        "sealed_0484": {
            "root": str(ROUTE_ROOT),
            "route_protocol": _old_protocol_binding(),
            "source_manifest": {
                "path": "validation-sources/5a86d19865374a7e914ebb2f168273800b4e04de5ae287bd7ecc4fe37c29b0ae.json",
                "sha256": "5a86d19865374a7e914ebb2f168273800b4e04de5ae287bd7ecc4fe37c29b0ae",
                "files": 7157,
            },
        },
        "baseline_0485": {
            "root": str(BASELINE_ROOT),
            "comparison_protocol": _baseline_protocol_binding(),
            "source_manifest": {
                "path": "validation-sources/c4594536bc2b80c6693be6bf5617ed7e699dd117945bc13183338fb0fd437abe.json",
                "sha256": "c4594536bc2b80c6693be6bf5617ed7e699dd117945bc13183338fb0fd437abe",
                "files": 7157,
            },
        },
        "source_diff_policy": {
            "allowed_prefixes": list(ALLOWED_SOURCE_PREFIXES),
            "require_nonempty": True,
            "description": "after source manifests may differ from the sealed before manifest only under the OPC splice adapter or focused splice/DOCX tests",
        },
        "machine": _machine_binding(),
        "scripts": _script_hashes(),
        "environment": {key: ENV[key] for key in ENV_KEYS},
    }
    if include_freeze_source:
        value["coordinator_source_at_freeze"] = snapshot()
    else:
        value["coordinator_source_at_freeze"] = None
    return value


def _protocol_path() -> Path:
    return ROOT / PROTOCOL_FILE


def _check_source_manifest_binding(value: Mapping[str, Any], root: Path, label: str) -> dict[str, str]:
    record = _object(value, label)
    path_value = _string(record.get("path"), f"{label}.path")
    digest = _sha(record.get("sha256"), f"{label}.sha256")
    manifest_path = Path(path_value)
    if not manifest_path.is_absolute():
        manifest_path = root / manifest_path
    require(manifest_path.is_file() and not manifest_path.is_symlink(), f"{label}: manifest is missing: {manifest_path}")
    require(sha(manifest_path) == digest, f"{label}: manifest hash changed")
    manifest = _read(manifest_path)
    require(isinstance(manifest, dict), f"{label}: manifest payload must be an object")
    require(record.get("files") == len(manifest), f"{label}.files: manifest cardinality differs")
    for name, file_hash in manifest.items():
        require(isinstance(name, str) and isinstance(file_hash, str) and SHA256_RE.fullmatch(file_hash) is not None, f"{label}: malformed source entry")
    return {str(name): str(file_hash) for name, file_hash in manifest.items()}


def _load_protocol() -> tuple[dict[str, Any], str]:
    protocol_path = _protocol_path()
    require(protocol_path.is_file() and not protocol_path.is_symlink(), f"comparison protocol is missing: {protocol_path}")
    value = _read(protocol_path)
    require(value.get("schema") == SCHEMA and value.get("version") == VERSION, f"{protocol_path}: protocol schema/version differs")
    expected = _protocol_value(include_freeze_source=False)
    for field in (
        "change", "cpu", "sink_write_bytes", "roles", "repeats", "phases", "workloads",
        "formal", "pilot", "arms", "formal_runs", "pilot_runs", "expected_formal_processes",
        "expected_pilot_processes", "fixtures", "build_bindings", "sealed_0484", "baseline_0485",
        "source_diff_policy", "environment",
    ):
        # The after build hash is allowed to have been deferred at freeze.  It
        # is checked against the immutable record once capture/analyze starts.
        if field == "build_bindings":
            actual = value.get(field)
            wanted = expected[field]
            require(isinstance(actual, dict), "protocol.build_bindings is missing")
            for phase in PHASES:
                for role in ROLE_NAMES:
                    a = _object(actual.get(phase, {}).get(role), f"protocol.build_bindings.{phase}.{role}")
                    w = wanted[phase][role]
                    for key in ("phase", "role", "root", "record_path", "record_schema", "command", "source_binding"):
                        require(a.get(key) == w[key], f"protocol.build_bindings.{phase}.{role}.{key}: differs from frozen plan")
                    # The sealed before records must already be bound when
                    # this protocol is frozen.  The after build is allowed to
                    # be deferred; once a protocol has recorded an after
                    # hash, that hash is immutable as well.
                    frozen_hash = _object(value["build_bindings"][phase], f"protocol.build_bindings.{phase}")[role].get("record_sha256")
                    if phase == "before":
                        require(frozen_hash is not None and a.get("record_sha256") == w["record_sha256"], f"protocol.build_bindings.{phase}.{role}.record_sha256: sealed before record differs")
                    elif frozen_hash is not None:
                        require(a.get("record_sha256") == frozen_hash and _sha(frozen_hash, f"protocol.build_bindings.{phase}.{role}.record_sha256"), f"protocol.build_bindings.{phase}.{role}.record_sha256: deferred binding differs")
                    else:
                        require(a.get("record_sha256") is None, f"protocol.build_bindings.{phase}.{role}.record_sha256: unexpected deferred hash")
            continue
        require(value.get(field) == expected[field], f"protocol.{field}: differs from frozen matrix or sealed binding")
    scripts = _object(value.get("scripts"), "protocol.scripts")
    require(scripts == _script_hashes(), "protocol canonical helper hashes differ; freeze a new unique attempt")
    machine = _object(value.get("machine"), "protocol.machine")
    if machine.get("sha256") is not None:
        require(sha(ROOT / _string(machine.get("path"), "protocol.machine.path")) == machine["sha256"], "protocol.machine hash changed")
    freeze_source = value.get("coordinator_source_at_freeze")
    if freeze_source is not None:
        _check_source_manifest_binding(freeze_source, ROOT, "protocol.coordinator_source_at_freeze")
    # Fixture files and the sealed old protocol are immutable inputs, so the
    # current bytes must still match their frozen hashes before any capture.
    for workload, fixture in _object(value.get("fixtures"), "protocol.fixtures").items():
        fixture_path = Path(_string(fixture.get("absolute_path"), f"protocol.fixtures.{workload}.absolute_path"))
        actual = _path_metadata(fixture_path, f"protocol.fixtures.{workload}")
        require(actual["bytes"] == fixture.get("bytes") and actual["sha256"] == fixture.get("sha256"), f"protocol.fixtures.{workload}: staged input changed")
    return value, sha(protocol_path)


def freeze() -> None:
    path = _protocol_path()
    require(not path.exists(), f"refusing to replace existing comparison protocol: {path}")
    value = _protocol_value(include_freeze_source=True)
    write(path, value)
    print(f"wrote {path} with {len(ARMS)} arms, {len(value['formal_runs'])} formal processes, and {len(value['pilot_runs'])} pilots")


def plan() -> None:
    print(json.dumps(_protocol_value(include_freeze_source=False), indent=2, sort_keys=True))


def _build_root(phase: str) -> Path:
    return BASELINE_ROOT if phase == "before" else ROOT


def _build_record_path(protocol: Mapping[str, Any], phase: str, role: str) -> Path:
    binding = _object(_object(protocol.get("build_bindings"), "protocol.build_bindings").get(phase), f"protocol.build_bindings.{phase}").get(role)
    binding = _object(binding, f"protocol.build_bindings.{phase}.{role}")
    relative = Path(_string(binding.get("record_path"), f"protocol.build_bindings.{phase}.{role}.record_path"))
    path = relative if relative.is_absolute() else _build_root(phase) / relative
    return path


def _resolve_gate_path(path_value: Any, root: Path, label: str) -> Path:
    path = Path(_string(path_value, f"{label}.path"))
    return path if path.is_absolute() else root / path


def _load_build(protocol: Mapping[str, Any], phase: str, role: str) -> dict[str, Any]:
    path = _build_record_path(protocol, phase, role)
    require(path.is_file() and not path.is_symlink(), f"{phase} {role} build record is missing: {path}")
    binding = _object(_object(protocol["build_bindings"], "protocol.build_bindings")[phase], f"protocol.build_bindings.{phase}")[role]
    expected_hash = binding.get("record_sha256")
    record_hash = sha(path)
    if expected_hash is not None:
        require(record_hash == expected_hash, f"{path}: build record hash differs from frozen binding")
    value = _read(path)
    require(value.get("schema") == BUILD_SCHEMA and value.get("version") == 1, f"{path}: build schema differs")
    require(value.get("role") == role, f"{path}: build role differs")
    require(value.get("command") == _build_command(role), f"{path}: build command differs from the sealed command")
    require(value.get("source_unchanged") is True, f"{path}: source custody is false")
    source_before = _object(value.get("source_before"), f"{path}.source_before")
    source_after = _object(value.get("source_after"), f"{path}.source_after")
    require(source_before == source_after, f"{path}: source changed during build")
    source_map = _check_source_manifest_binding(source_after, _build_root(phase), f"{path}.source_after")
    gate = _object(value.get("gate"), f"{path}.gate")
    gate_path = _resolve_gate_path(gate.get("path"), _build_root(phase), f"{path}.gate")
    gate_hash = _sha(gate.get("sha256"), f"{path}.gate.sha256")
    require(gate_path.is_file() and sha(gate_path) == gate_hash, f"{path}: gate receipt hash changed")
    gate_value = _read(gate_path)
    require(gate_value.get("schema") == "docx-stream-append-gate-v1", f"{gate_path}: gate schema differs")
    require(gate_value.get("exit_code") == 0 and gate_value.get("source_unchanged") is True, f"{gate_path}: build gate failed")
    require(gate_value.get("argv") == _build_command(role), f"{gate_path}: gate command differs from build record")
    require(gate_value.get("cwd") == str(REPO), f"{gate_path}: gate cwd differs")
    require(gate_value.get("environment") == {key: ENV[key] for key in ENV_KEYS}, f"{gate_path}: gate environment differs")
    expected_gate_root = _build_root(phase)
    expected_driver = sha(expected_gate_root / "gate.py")
    expected_common = sha(expected_gate_root / "support.py")
    require(gate_value.get("driver_sha256") == expected_driver, f"{gate_path}: gate driver binding differs")
    require(gate_value.get("common_sha256") == expected_common, f"{gate_path}: gate support binding differs")
    require(gate_value.get("source_before") == source_before and gate_value.get("source_after") == source_after, f"{gate_path}: build source binding differs")
    binary = _object(value.get("binary"), f"{path}.binary")
    binary_path = Path(_string(binary.get("path"), f"{path}.binary.path"))
    actual_binary = _path_metadata(binary_path, f"{path}.binary")
    actual_binary["executable"] = os.access(binary_path, os.X_OK)
    require(actual_binary == binary, f"{path}: copied binary metadata changed")
    environment = _object(value.get("environment"), f"{path}.environment")
    require(environment == {key: ENV[key] for key in ENV_KEYS}, f"{path}: build environment differs")
    return {
        "phase": phase,
        "role": role,
        "path": path,
        "sha256": record_hash,
        "record": value,
        "binary": binary,
        "source": source_after,
        "source_map": source_map,
        "gate": {"path": gate_path, "sha256": gate_hash},
    }


def _load_builds(protocol: Mapping[str, Any]) -> dict[str, dict[str, dict[str, Any]]]:
    result: dict[str, dict[str, dict[str, Any]]] = {}
    for phase in PHASES:
        result[phase] = {role: _load_build(protocol, phase, role) for role in ROLE_NAMES}
        require(result[phase]["normal"]["source"] == result[phase]["allocator"]["source"], f"{phase}: normal/allocator source manifests differ")
    return result


def _capture_gate_path(phase: str, attempt: str) -> Path:
    return ROOT / "validation" / f"capture-{phase}-{attempt}.json"


def _validate_capture_gate(phase: str, attempt: str, expected_source: Mapping[str, Any]) -> dict[str, Any]:
    """Validate the coordinator gate that covered one formal capture lane."""

    path = _capture_gate_path(phase, attempt)
    require(path.is_file() and not path.is_symlink(), f"{phase}: formal capture gate receipt is missing: {path}")
    value = _read(path)
    require(value.get("schema") == GATE_SCHEMA, f"{path}: gate schema differs")
    require(value.get("label") == f"capture-{phase}-{attempt}", f"{path}: gate label differs")
    require(value.get("attempt") == attempt, f"{path}: gate attempt differs")
    require(value.get("exit_code") == 0 and value.get("source_unchanged") is True, f"{path}: formal capture gate failed")
    require(value.get("cwd") == str(REPO), f"{path}: formal capture gate cwd differs")
    require(value.get("environment") == {key: ENV[key] for key in ENV_KEYS}, f"{path}: formal capture gate environment differs")
    require(value.get("driver_sha256") == sha(ROOT / "gate.py"), f"{path}: formal capture gate driver differs")
    require(value.get("common_sha256") == sha(ROOT / "support.py"), f"{path}: formal capture gate support differs")
    require(value.get("source_before") == expected_source and value.get("source_after") == expected_source, f"{path}: formal capture gate source differs from after build")
    argv = _list(value.get("argv"), f"{path}.argv")
    require(len(argv) == 8, f"{path}.argv: formal capture command shape differs")
    python_path = Path(_string(argv[0], f"{path}.argv[0]"))
    require(python_path.name.startswith("python"), f"{path}.argv[0]: capture command is not Python")
    require(argv[1] == "-B", f"{path}.argv[1]: capture command must disable bytecode")
    driver_path = Path(_string(argv[2], f"{path}.argv[2]"))
    if not driver_path.is_absolute():
        driver_path = REPO / driver_path
    require(driver_path.resolve() == Path(__file__).resolve(), f"{path}.argv[2]: capture driver differs")
    require(argv[3:] == ["capture-all", "--phase", phase, "--attempt", attempt], f"{path}.argv: capture command arguments differ")
    return {
        "path": str(path.relative_to(ROOT)),
        "sha256": sha(path),
        "source": value["source_after"],
        "argv": argv,
    }


def _input_metadata(protocol: Mapping[str, Any], arm: Mapping[str, Any]) -> dict[str, Any] | None:
    if arm["input_mode"] != "file":
        return None
    fixture = _object(protocol["fixtures"], "protocol.fixtures")[arm["workload"]]
    path = Path(_string(fixture.get("absolute_path"), f"fixture.{arm['workload']}.absolute_path"))
    actual = _path_metadata(path, f"fixture {arm['workload']}")
    require(actual["bytes"] == fixture["bytes"] and actual["sha256"] == fixture["sha256"], f"fixture {arm['workload']}: hash changed")
    return {
        "path": fixture["path"],
        "absolute_path": str(path),
        "bytes": actual["bytes"],
        "sha256": actual["sha256"],
        "identity": "prepared_file_capability_fingerprint_must_match_report_source_archive",
    }


def _old_arm(arm: Mapping[str, Any]) -> dict[str, Any]:
    if arm["family"] == "axis":
        return dict(old_routes.AXIS_ARM_BY_LABEL[arm["old_axis_arm"]])
    return dict(old_routes.ROUTE_CASE_BY_LABEL[arm["old_route_case"]])


def _case_for(arm: Mapping[str, Any]) -> dict[str, Any]:
    if arm["family"] == "axis":
        return old_routes._axis_case(_old_arm(arm))
    return dict(old_routes.ROUTE_CASE_BY_LABEL[arm["old_route_case"]])


def _argv_for(
    arm: Mapping[str, Any],
    binary: Mapping[str, Any],
    *,
    report: Path,
    resource: Path,
    samples: int,
    warmups: int,
    replay_dir: Path | None,
) -> list[str]:
    case = _case_for(arm)
    if arm["family"] == "axis":
        return old_routes._axis_argv(
            dict(binary), case, _old_arm(arm), samples=samples, warmups=warmups, report=report, resource=resource,
        )
    spec = old_routes.ROUTE_BY_NAME[arm["route"]]
    return old_routes._route_argv(
        dict(binary), case, spec, samples=samples, warmups=warmups, report=report, resource=resource, replay_dir=replay_dir,
    )


def _validate_report(
    report: Path,
    arm: Mapping[str, Any],
    role: str,
    *,
    samples: int,
    warmups: int,
    binary: Mapping[str, Any],
    argv: list[str],
    input_metadata: dict[str, Any] | None,
    replay_dir: Path | None,
) -> dict[str, Any]:
    try:
        if arm["family"] == "axis":
            return old_routes._check_axis_report(
                report,
                role,
                _old_arm(arm),
                samples=samples,
                warmups=warmups,
                binary=dict(binary),
                argv=argv,
                input_metadata=input_metadata,
            )
        return old_routes.check_route_report(
            report,
            role,
            _case_for(arm),
            arm["route"],
            samples=samples,
            warmups=warmups,
            binary=dict(binary),
            argv=argv,
            replay_dir=replay_dir,
        )
    except Exception as error:
        raise CompareError(f"{report}: frozen 0484 report validator rejected the report: {error}") from error


def _run_directory(phase: str, attempt: str, label: str) -> Path:
    directory = ROOT / "captures" / phase / attempt / label
    require(not directory.exists(), f"refusing to replace existing capture directory: {directory}")
    directory.mkdir(parents=True)
    return directory


def _kill_group(process: subprocess.Popen[Any], sig: int) -> None:
    try:
        os.killpg(process.pid, sig)
    except ProcessLookupError:
        pass


def _run_child(argv: list[str], stdout: Path, stderr: Path, timeout_seconds: int) -> dict[str, Any]:
    launch_error: str | None = None
    timed_out = False
    returncode: int | None = None
    try:
        with stdout.open("xb") as out, stderr.open("xb") as err:
            process = subprocess.Popen(
                argv,
                cwd=REPO,
                env=ENV,
                stdout=out,
                stderr=err,
                start_new_session=True,
            )
            try:
                returncode = process.wait(timeout=timeout_seconds)
            except subprocess.TimeoutExpired:
                timed_out = True
                _kill_group(process, signal.SIGTERM)
                try:
                    process.wait(timeout=2)
                except subprocess.TimeoutExpired:
                    pass
                # The wrapper can exit after SIGTERM while a descendant
                # ignores it.  Always kill the process group before treating
                # the timeout as terminal, then reap the wrapper itself.
                _kill_group(process, signal.SIGKILL)
                try:
                    returncode = process.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    _kill_group(process, signal.SIGKILL)
                    returncode = process.wait(timeout=10)
    except (OSError, subprocess.SubprocessError) as error:
        launch_error = f"{type(error).__name__}: {error}"
    return {
        "exit_code": returncode,
        "timed_out": timed_out,
        "launch_error": launch_error,
    }


def _artifact_map(paths: Iterable[Path]) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for path in paths:
        if path.is_file() and not path.is_symlink():
            result[path.name] = meta(path)
    return result


def _cleanup_replay(directory: Path, replay_dir: Path) -> Path:
    require(replay_dir.is_dir() and not replay_dir.is_symlink(), f"{directory}: replay scratch directory is missing or not a directory")
    entries = sorted(item.name for item in replay_dir.iterdir())
    require(not entries, f"{directory}: file-store replay scratch was not empty after validation: {entries}")
    stat = replay_dir.stat()
    receipt_path = directory / "replay-cleanup.json"
    receipt = {
        "schema": REPLAY_CLEANUP_SCHEMA,
        "version": 1,
        "path": str(replay_dir),
        "device": stat.st_dev,
        "inode": stat.st_ino,
        "entries_before_removal": entries,
        "empty_before_removal": True,
        "removed": False,
        "removed_utc": None,
    }
    replay_dir.rmdir()
    receipt["removed"] = True
    receipt["removed_utc"] = now()
    write(receipt_path, receipt)
    return receipt_path


def _capture_one(protocol: Mapping[str, Any], builds: Mapping[str, Mapping[str, Any]], run: Mapping[str, Any], attempt: str, *, pilot: bool, timeout_seconds: int) -> Path:
    phase = _string(run.get("phase"), "run.phase")
    arm = ARM_BY_ID[_string(run.get("arm"), "run.arm")]
    role = _string(run.get("role"), "run.role")
    require(role in ROLE_NAMES, f"run role is invalid: {role}")
    expected_kind = "pilot" if pilot else "formal"
    require(run.get("kind") == expected_kind, f"run kind differs from capture command: {run.get('kind')}")
    if pilot:
        require(run.get("repeat") is None, "pilot run carries a formal repeat")
    else:
        require(run.get("repeat") in REPEATS, "formal run repeat is invalid")
    build = builds[role]
    directory = _run_directory(phase, attempt, _string(run.get("label"), "run.label"))
    samples = PILOT_SAMPLES if pilot else FORMAL_SAMPLES
    warmups = PILOT_WARMUPS if pilot else FORMAL_WARMUPS
    report = directory / "report.json"
    resource = directory / "resource.txt"
    stdout = directory / "stdout.txt"
    stderr = directory / "stderr.txt"
    replay_dir = directory / "replay" if arm["route"] == "file_store" else None
    if replay_dir is not None:
        replay_dir.mkdir()
    argv = _argv_for(
        arm,
        build["binary"],
        report=report,
        resource=resource,
        samples=samples,
        warmups=warmups,
        replay_dir=replay_dir,
    )
    protocol_hash = sha(_protocol_path())
    build_path = build["path"]
    started = {
        "schema": CAPTURE_SCHEMA,
        "version": 1,
        "status": "running",
        "attempt": attempt,
        "phase": phase,
        "run": dict(run),
        "protocol": {"path": PROTOCOL_FILE, "sha256": protocol_hash},
        "build": {"path": str(build_path), "sha256": build["sha256"], "role": role},
        "binary": dict(build["binary"]),
        "machine": protocol.get("machine"),
        "arm": dict(arm),
        "input_file": _input_metadata(protocol, arm),
        "replay_dir": None if replay_dir is None else str(replay_dir),
        "argv": argv,
        "cwd": str(REPO),
        "environment": {key: ENV[key] for key in ENV_KEYS},
        "timeout_seconds": timeout_seconds,
        "started_utc": now(),
    }
    write(directory / "started.json", started)
    process_result = _run_child(argv, stdout, stderr, timeout_seconds)
    validation_error: str | None = None
    if process_result["exit_code"] == 0:
        try:
            for core in (report, resource):
                require(core.is_file() and not core.is_symlink(), f"{core}: required capture artifact is missing")
                require(core.stat().st_size > 0, f"{core}: required capture artifact is empty")
            _validate_report(
                report,
                arm,
                role,
                samples=samples,
                warmups=warmups,
                binary=build["binary"],
                argv=argv,
                input_metadata=started["input_file"],
                replay_dir=replay_dir,
            )
            if replay_dir is not None:
                _cleanup_replay(directory, replay_dir)
        except Exception as error:
            validation_error = str(error)
    required = (stdout, stderr, resource, report)
    missing = [path.name for path in required if not path.is_file()]
    artifacts = _artifact_map((*required, directory / "started.json", directory / "replay-cleanup.json"))
    passed = (
        process_result["exit_code"] == 0
        and not process_result["timed_out"]
        and process_result["launch_error"] is None
        and validation_error is None
        and not missing
    )
    finished = dict(
        started,
        status="pass" if passed else "failed",
        exit_code=process_result["exit_code"],
        timed_out=process_result["timed_out"],
        finished_utc=now(),
        artifacts=artifacts,
        missing_artifacts=missing,
    )
    if process_result["launch_error"] is not None:
        finished["launch_error"] = process_result["launch_error"]
    if validation_error is not None:
        finished["validation_error"] = validation_error
    receipt_path = directory / "receipt.json"
    write(receipt_path, finished)
    if not passed:
        detail = process_result["launch_error"] or validation_error or f"exit {process_result['exit_code']}"
        fail(f"{run['label']} failed ({detail}); receipt retained")
    print(f"captured {run['label']} ({samples} samples, {warmups} warmups)")
    return receipt_path


def _runs_for(protocol: Mapping[str, Any], *, pilot: bool, phase: str) -> list[dict[str, Any]]:
    field = "pilot_runs" if pilot else "formal_runs"
    expected = [run for run in _runs(pilot=pilot) if run["phase"] == phase]
    actual = protocol.get(field)
    require(actual == _runs(pilot=pilot), f"protocol.{field}: run inventory differs")
    return expected


def capture_all(phase: str, attempt: str, *, pilot: bool, timeout_seconds: int) -> None:
    require(phase in PHASES, f"unknown phase: {phase}")
    protocol, _ = _load_protocol()
    builds_all = _load_builds(protocol)
    builds = builds_all[phase]
    for run in _runs_for(protocol, pilot=pilot, phase=phase):
        _capture_one(protocol, builds, run, attempt, pilot=pilot, timeout_seconds=timeout_seconds)


def _parse_resource(path: Path) -> int:
    prefix = "Maximum resident set size (kbytes):"
    values: list[int] = []
    require(path.is_file() and not path.is_symlink(), f"{path}: GNU time resource receipt is missing")
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line.startswith(prefix):
            raw = line[len(prefix):].strip()
            require(raw.isdigit(), f"{path}: malformed maximum RSS")
            values.append(int(raw) * 1024)
    require(len(values) == 1, f"{path}: expected exactly one GNU time maximum RSS observation")
    return values[0]


def _identity(report_case: Mapping[str, Any], label: str) -> dict[str, Any]:
    source = _object(report_case.get("source"), f"{label}.source")
    authored = _object(report_case.get("authored"), f"{label}.authored")
    oracle = _object(report_case.get("oracle"), f"{label}.oracle")
    oracle_flags = {
        field: oracle.get(field)
        for field in (
            "candidate_xml_exact",
            "candidate_semantic_exact",
            "untouched_member_metadata_exact",
            "untouched_raw_members_preserved",
            "physical_order_exact",
            "opaque_member_exact",
            "source_unchanged",
            "inverse_exact",
        )
    }
    candidate_semantic = oracle.get("candidate_semantic")
    require(isinstance(candidate_semantic, dict), f"{label}.oracle.candidate_semantic: expected an object")
    result = {
        "source": {
            "archive_bytes": _integer(source.get("archive_bytes"), f"{label}.source.archive_bytes", minimum=1),
            "archive_sha256": _sha(source.get("archive_sha256"), f"{label}.source.archive_sha256"),
            "main_xml_bytes": _integer(source.get("main_xml_bytes"), f"{label}.source.main_xml_bytes", minimum=1),
            "main_xml_sha256": _sha(source.get("main_xml_sha256"), f"{label}.source.main_xml_sha256"),
        },
        "authored": {
            "encoded_xml_bytes": _integer(authored.get("encoded_xml_bytes"), f"{label}.authored.encoded_xml_bytes", minimum=1),
            "expected_event_sha256": _sha(authored.get("expected_event_sha256"), f"{label}.authored.expected_event_sha256"),
            "expected_encoded_sha256": _sha(authored.get("expected_encoded_sha256"), f"{label}.authored.expected_encoded_sha256"),
        },
        "candidate": {
            "archive_bytes": _integer(oracle.get("candidate_archive_bytes"), f"{label}.oracle.candidate_archive_bytes", minimum=1),
            "archive_sha256": _sha(oracle.get("candidate_archive_sha256"), f"{label}.oracle.candidate_archive_sha256"),
            "main_xml_bytes": _integer(oracle.get("candidate_main_xml_bytes"), f"{label}.oracle.candidate_main_xml_bytes", minimum=1),
            "main_xml_sha256": _sha(oracle.get("candidate_main_xml_sha256"), f"{label}.oracle.candidate_main_xml_sha256"),
            "semantic": candidate_semantic,
            "oracle_flags": oracle_flags,
        },
    }
    return result


def _stats(values: Iterable[int | float]) -> dict[str, Any]:
    try:
        result = old_analysis.stats(values)
    except (old_analysis.AnalysisError, ValueError, TypeError) as error:
        raise CompareError(f"metric cannot be summarized: {error}") from error
    return {str(key): value for key, value in result.items()}


def _row_metrics(report: Mapping[str, Any], role: str, resource_rss: int, label: str) -> dict[str, Any]:
    observed = _object(_list(report.get("cases"), f"{label}.cases")[0], f"{label}.cases[0]")
    samples = _list(observed.get("samples"), f"{label}.cases[0].samples")
    elapsed = [_integer(_object(sample, f"{label}.sample").get("elapsed_ns"), f"{label}.elapsed_ns", minimum=1) for sample in samples]
    result: dict[str, Any] = {
        "elapsed_ns": _stats(elapsed),
        "sample_elapsed_ns": elapsed,
        # GNU time contributes one whole-child observation.  The repeated
        # percentile labels make the comparison table uniform, while ``n``
        # and ``observation_scope`` prevent it from being mistaken for 30
        # independent in-process RSS samples.
        "time_max_rss_bytes": {
            "n": 1,
            "value": resource_rss,
            "p50": resource_rss,
            "p95": resource_rss,
            "p99": resource_rss,
            "observation_scope": "whole_child_gnu_time_single_observation",
        },
    }
    if role == "allocator":
        heap: list[int] = []
        counter_values: dict[str, list[int]] = {
            field: []
            for field in (
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
        }
        for index, sample_value in enumerate(samples):
            sample = _object(sample_value, f"{label}.samples[{index}]")
            allocation = _object(sample.get("allocation"), f"{label}.samples[{index}].allocation")
            region = _integer(allocation.get("region_peak_live_bytes"), f"{label}.samples[{index}].allocation.region_peak_live_bytes")
            live_before = _integer(allocation.get("live_bytes_before"), f"{label}.samples[{index}].allocation.live_bytes_before")
            require(region >= live_before, f"{label}.samples[{index}]: allocator operation peak regressed below live bytes")
            heap.append(region - live_before)
            for field, values in counter_values.items():
                values.append(_integer(allocation.get(field), f"{label}.samples[{index}].allocation.{field}"))
        result["allocator_operation_peak_increment_bytes"] = _stats(heap)
        result["sample_allocator_operation_peak_increment_bytes"] = heap
        result["allocator_counters"] = {field: _stats(values) for field, values in counter_values.items()}
    else:
        result["allocator_operation_peak_increment_bytes"] = None
        result["sample_allocator_operation_peak_increment_bytes"] = None
        result["allocator_counters"] = None
    return result


def _expected_replay_dir(arm: Mapping[str, Any], directory: Path) -> Path | None:
    return directory / "replay" if arm["route"] == "file_store" else None


def _validate_artifacts(receipt: Mapping[str, Any], directory: Path, *, file_store: bool) -> None:
    required_names = {"started.json", "stdout.txt", "stderr.txt", "resource.txt", "report.json"}
    if file_store:
        required_names.add("replay-cleanup.json")
    artifacts = _object(receipt.get("artifacts"), f"{directory}/receipt.json.artifacts")
    require(set(artifacts) == required_names, f"{directory}: artifact inventory differs")
    for name in required_names:
        path = directory / name
        require(path.is_file() and not path.is_symlink(), f"{directory}: artifact is missing: {path}")
        require(artifacts[name] == meta(path), f"{directory}: artifact hash changed: {name}")
    missing = receipt.get("missing_artifacts")
    require(missing == [], f"{directory}: retained receipt reports missing artifacts: {missing}")


def _validate_replay_cleanup(directory: Path) -> None:
    path = directory / "replay-cleanup.json"
    value = _read(path)
    require(value.get("schema") == REPLAY_CLEANUP_SCHEMA and value.get("removed") is True, f"{path}: cleanup receipt failed")
    replay_dir = Path(_string(value.get("path"), f"{path}.path"))
    require(not replay_dir.exists(), f"{path}: replay scratch directory still exists")
    require(value.get("entries_before_removal") == [] and value.get("empty_before_removal") is True, f"{path}: replay directory was not proven empty")


def _validate_started(receipt: Mapping[str, Any], directory: Path, protocol_hash: str, run: Mapping[str, Any], build: Mapping[str, Any], arm: Mapping[str, Any], *, samples: int, warmups: int) -> dict[str, Any]:
    started_path = directory / "started.json"
    started = _read(started_path)
    require(started.get("schema") == CAPTURE_SCHEMA and started.get("status") == "running", f"{started_path}: start receipt differs")
    require(started.get("run") == run and started.get("phase") == run["phase"] and started.get("attempt") == receipt.get("attempt"), f"{started_path}: run identity differs")
    require(_object(started.get("protocol"), f"{started_path}.protocol").get("sha256") == protocol_hash, f"{started_path}: protocol hash differs")
    started_build = _object(started.get("build"), f"{started_path}.build")
    require(started_build.get("sha256") == build["sha256"], f"{started_path}: build hash differs")
    require(started.get("binary") == build["binary"], f"{started_path}: binary binding differs")
    require(started.get("arm") == arm, f"{started_path}: arm binding differs")
    input_metadata = started.get("input_file")
    expected_input = _input_metadata(_read(_protocol_path()), arm)
    require(input_metadata == expected_input, f"{started_path}: input fixture binding differs")
    argv = _argv_for(
        arm,
        build["binary"],
        report=directory / "report.json",
        resource=directory / "resource.txt",
        samples=samples,
        warmups=warmups,
        replay_dir=_expected_replay_dir(arm, directory),
    )
    require(started.get("argv") == argv and receipt.get("argv") == argv, f"{directory}: captured argv differs from sealed arm")
    require(started.get("environment") == {key: ENV[key] for key in ENV_KEYS}, f"{started_path}: environment differs")
    return started


def _validate_formal_phase(protocol: Mapping[str, Any], builds: Mapping[str, Mapping[str, Any]], phase: str, attempt: str) -> list[dict[str, Any]]:
    protocol_hash = sha(_protocol_path())
    expected_runs = [run for run in _runs(pilot=False) if run["phase"] == phase]
    root = ROOT / "captures" / phase / attempt
    require(root.is_dir() and not root.is_symlink(), f"{phase}: formal capture directory is missing: {root}")
    expected_labels = {run["label"] for run in expected_runs}
    actual_labels = {item.name for item in root.iterdir() if item.is_dir() and not item.is_symlink()}
    require(actual_labels == expected_labels, f"{phase}: formal capture inventory differs")
    rows: list[dict[str, Any]] = []
    for run in expected_runs:
        directory = root / run["label"]
        receipt_path = directory / "receipt.json"
        receipt = _read(receipt_path)
        require(receipt.get("schema") == CAPTURE_SCHEMA and receipt.get("status") == "pass" and receipt.get("exit_code") == 0, f"{receipt_path}: formal capture did not pass")
        require(receipt.get("protocol", {}).get("sha256") == protocol_hash, f"{receipt_path}: protocol hash differs")
        require(receipt.get("run") == run and receipt.get("phase") == phase and receipt.get("attempt") == attempt, f"{receipt_path}: run identity differs")
        arm = ARM_BY_ID[run["arm"]]
        build = builds[run["role"]]
        require(receipt.get("build", {}).get("sha256") == build["sha256"], f"{receipt_path}: build binding differs")
        require(receipt.get("binary") == build["binary"], f"{receipt_path}: binary binding differs")
        _validate_started(receipt, directory, protocol_hash, run, build, arm, samples=FORMAL_SAMPLES, warmups=FORMAL_WARMUPS)
        _validate_artifacts(receipt, directory, file_store=arm["route"] == "file_store")
        if arm["route"] == "file_store":
            _validate_replay_cleanup(directory)
        report_path = directory / "report.json"
        replay_dir = _expected_replay_dir(arm, directory)
        report = _validate_report(
            report_path,
            arm,
            run["role"],
            samples=FORMAL_SAMPLES,
            warmups=FORMAL_WARMUPS,
            binary=build["binary"],
            argv=receipt["argv"],
            input_metadata=_input_metadata(protocol, arm),
            replay_dir=replay_dir,
        )
        resource_rss = _parse_resource(directory / "resource.txt")
        observed_case = _object(report["cases"][0], f"{report_path}.cases[0]")
        identity = _identity(observed_case, f"{report_path}.cases[0]")
        rows.append({
            "phase": phase,
            "attempt": attempt,
            "arm": arm["id"],
            "family": arm["family"],
            "route": arm["route"],
            "input_mode": arm["input_mode"],
            "workload": arm["workload"],
            "role": run["role"],
            "repeat": run["repeat"],
            "label": run["label"],
            "directory": str(directory),
            "receipt": str(receipt_path),
            "report": str(report_path),
            "report_sha256": sha(report_path),
            "build_sha256": build["sha256"],
            "binary": build["binary"],
            "source_identity": identity["source"],
            "authored_identity": identity["authored"],
            "candidate_identity": identity["candidate"],
            "samples": FORMAL_SAMPLES,
            "warmups": FORMAL_WARMUPS,
            "metrics": _row_metrics(report, run["role"], resource_rss, str(report_path)),
        })
    require(len(rows) == 72, f"{phase}: formal process cardinality differs")
    return rows


def _manifest_diff(before: Mapping[str, str], after: Mapping[str, str]) -> dict[str, Any]:
    names = sorted(set(before) | set(after))
    changes: list[dict[str, Any]] = []
    for name in names:
        old = before.get(name)
        new = after.get(name)
        if old == new:
            continue
        kind = "added" if old is None else "removed" if new is None else "changed"
        changes.append({"path": name, "kind": kind, "before": old, "after": new})
    allowed = [item["path"] for item in changes if any(item["path"].startswith(prefix) for prefix in ALLOWED_SOURCE_PREFIXES)]
    unexpected = [item["path"] for item in changes if item["path"] not in allowed]
    require(changes, "source manifests are identical; the after build does not identify the requested OPC change")
    require(not unexpected, f"source manifest contains changes outside the OPC batching allowlist: {unexpected}")
    return {
        "changed_files": len(changes),
        "changes": changes,
        "allowed_files": allowed,
        "unexpected_files": unexpected,
        "all_changes_allowed": not unexpected,
    }


def _identity_tuple(row: Mapping[str, Any]) -> tuple[Any, ...]:
    candidate = row["candidate_identity"]
    # ZIP framing, compression metadata, and archive padding can legitimately
    # change between the old and new implementation.  The paired content
    # identity therefore binds the decoded candidate XML, semantic oracle, and
    # preservation flags while retaining the candidate archive hash/length as
    # an independently visible artifact identity in every row.
    candidate_content = {
        "main_xml_bytes": candidate["main_xml_bytes"],
        "main_xml_sha256": candidate["main_xml_sha256"],
        "semantic": candidate["semantic"],
        "oracle_flags": candidate["oracle_flags"],
    }
    return (row["source_identity"], row["authored_identity"], candidate_content)


def _check_content_identity(rows: list[Mapping[str, Any]]) -> None:
    by_arm: dict[str, tuple[Any, ...]] = {}
    for row in rows:
        arm = str(row["arm"])
        identity = _identity_tuple(row)
        previous = by_arm.setdefault(arm, identity)
        require(previous == identity, f"{arm}: source/authored/candidate content identity changed across phases, roles, or repeats")


def _candidate_archive_changes(rows: list[Mapping[str, Any]]) -> list[dict[str, Any]]:
    lookup = {(row["phase"], row["arm"], row["role"], row["repeat"]): row for row in rows}
    result: list[dict[str, Any]] = []
    for arm in ARMS:
        for role in ROLE_NAMES:
            for repeat in REPEATS:
                before = lookup[("before", arm["id"], role, repeat)]["candidate_identity"]
                after = lookup[("after", arm["id"], role, repeat)]["candidate_identity"]
                if before["archive_bytes"] != after["archive_bytes"] or before["archive_sha256"] != after["archive_sha256"]:
                    result.append({
                        "arm": arm["id"],
                        "role": role,
                        "repeat": repeat,
                        "before": {
                            "archive_bytes": before["archive_bytes"],
                            "archive_sha256": before["archive_sha256"],
                        },
                        "after": {
                            "archive_bytes": after["archive_bytes"],
                            "archive_sha256": after["archive_sha256"],
                        },
                        "decoded_content_identity_equal": _identity_tuple(lookup[("before", arm["id"], role, repeat)]) == _identity_tuple(lookup[("after", arm["id"], role, repeat)]),
                    })
    return result


def _percent_change(before: float, after: float) -> float | None:
    if before == 0:
        return None if after == 0 else math.inf
    return (after - before) / before * 100.0


def _comparisons(rows: list[Mapping[str, Any]]) -> list[dict[str, Any]]:
    lookup = {(row["phase"], row["arm"], row["role"], row["repeat"]): row for row in rows}
    result: list[dict[str, Any]] = []
    for arm in ARMS:
        for role in ROLE_NAMES:
            for repeat in REPEATS:
                before = lookup[("before", arm["id"], role, repeat)]
                after = lookup[("after", arm["id"], role, repeat)]
                metrics = ["elapsed_ns", "time_max_rss_bytes"]
                if role == "allocator":
                    metrics.append("allocator_operation_peak_increment_bytes")
                for metric in metrics:
                    before_stats = before["metrics"][metric]
                    after_stats = after["metrics"][metric]
                    require(before_stats is not None and after_stats is not None, f"{arm['id']} {role} {repeat}: comparison metric is unavailable")
                    values: dict[str, Any] = {}
                    for percentile in ("p50", "p95", "p99"):
                        control = float(before_stats[percentile])
                        candidate = float(after_stats[percentile])
                        change = _percent_change(control, candidate)
                        flag = change is None or not math.isfinite(change) or abs(change) > REVIEW_THRESHOLD_PERCENT
                        values[percentile] = {
                            "before": control,
                            "after": candidate,
                            "percent_change": change,
                            "review_flag_5pct": flag,
                            "favorable_lower_value": candidate < control,
                            "adverse_higher_value": candidate > control,
                        }
                    result.append({
                        "arm": arm["id"],
                        "workload": arm["workload"],
                        "route": arm["route"],
                        "input_mode": arm["input_mode"],
                        "role": role,
                        "repeat": repeat,
                        "metric": metric,
                        "before_label": before["label"],
                        "after_label": after["label"],
                        "percent_changes": values,
                    })
    return result


def _source_summary(builds: Mapping[str, Mapping[str, Mapping[str, Any]]]) -> dict[str, Any]:
    before = builds["before"]["normal"]
    after = builds["after"]["normal"]
    before_map = before["source_map"]
    after_map = after["source_map"]
    diff = _manifest_diff(before_map, after_map)
    return {
        "before": before["source"],
        "after": after["source"],
        "diff": diff,
        "policy": {
            "allowed_prefixes": list(ALLOWED_SOURCE_PREFIXES),
            "implementation_and_tests_only": diff["all_changes_allowed"],
        },
    }


def analyze(before_attempt: str, after_attempt: str, summary_path: Path, markdown_path: Path) -> dict[str, Any]:
    protocol, protocol_hash = _load_protocol()
    builds = _load_builds(protocol)
    source = _source_summary(builds)
    capture_gates = {
        "before": _validate_capture_gate("before", before_attempt, builds["before"]["normal"]["source"]),
        "after": _validate_capture_gate("after", after_attempt, builds["after"]["normal"]["source"]),
    }
    before_rows = _validate_formal_phase(protocol, builds["before"], "before", before_attempt)
    after_rows = _validate_formal_phase(protocol, builds["after"], "after", after_attempt)
    rows = before_rows + after_rows
    require(len(rows) == 144, "formal process cardinality differs from the 144-child comparison matrix")
    _check_content_identity(rows)
    candidate_archive_changes = _candidate_archive_changes(rows)
    comparisons = _comparisons(rows)
    summary: dict[str, Any] = {
        "schema": "docx-opc-splice-audit-consumed-prefix-comparison-summary-v1",
        "version": 1,
        "analyzer": {"path": Path(__file__).name, "sha256": sha(Path(__file__))},
        "protocol": {"path": PROTOCOL_FILE, "sha256": protocol_hash},
        "attempts": {"before": before_attempt, "after": after_attempt},
        "builds": {
            phase: {
                role: {
                    "record_path": str(builds[phase][role]["path"]),
                    "record_sha256": builds[phase][role]["sha256"],
                    "binary": builds[phase][role]["binary"],
                    "source": builds[phase][role]["source"],
                    "gate": {"path": str(builds[phase][role]["gate"]["path"]), "sha256": builds[phase][role]["gate"]["sha256"]},
                }
                for role in ROLE_NAMES
            }
            for phase in PHASES
        },
        "source_comparison": source,
        "capture_gates": capture_gates,
        "candidate_archive_changes": candidate_archive_changes,
        "inventory": {
            "arms": len(ARMS),
            "formal_processes": len(rows),
            "formal_samples": len(rows) * FORMAL_SAMPLES,
            "samples_per_process": FORMAL_SAMPLES,
            "warmups_per_process": FORMAL_WARMUPS,
            "roles": list(ROLE_NAMES),
            "repeats": list(REPEATS),
            "phases": list(PHASES),
            "pilots_included": False,
        },
        "processes": rows,
        "comparisons": comparisons,
        "claims": {
            "performance_claim": "none_until_reviewed",
            "causal_speedup": False,
            "review_threshold_percent": REVIEW_THRESHOLD_PERCENT,
            "rss_scope": "one GNU time maximum RSS observation per whole child; not a bounded-memory proof",
            "allocator_scope": "operation_global_system_allocator region peak minus live bytes before the operation",
            "uncertainty": "two independent process repeats are retained per phase; no confidence interval",
            "content_identity": "source/authored/candidate identities must match across all phase, role, and repeat rows for each arm",
        },
        "limitations": [
            "The matrix covers three exact workloads and selected input/store arms; it is not a full provider-by-input interaction grid.",
            "Two process repeats support a descriptive before/after range; they do not establish a strong confidence interval.",
            "A five-percent flag is a review trigger, not a causal speedup claim.",
            "GNU time observes the whole child once and is retained with n=1; allocator operation heap is a separate instrumented metric.",
        ],
    }
    write(summary_path, summary)
    _write_text_exclusive(markdown_path, _markdown(summary))
    return summary


def _fmt(value: Any) -> str:
    if value is None:
        return "—"
    if isinstance(value, float):
        return f"{value:.3f}"
    return str(value)


def _markdown(summary: Mapping[str, Any]) -> str:
    source = summary["source_comparison"]
    lines = [
        "# 0487 OPC consumed-prefix retention comparison",
        "",
        "This table retains every formal process row. Values are descriptive; a five-percent flag marks a review threshold and does not authorize a causal speedup claim.",
        "",
        f"Protocol `{summary['protocol']['sha256']}`; before attempt `{summary['attempts']['before']}`; after attempt `{summary['attempts']['after']}`.",
        "",
        f"The matrix contains {summary['inventory']['formal_processes']} formal children and {summary['inventory']['formal_samples']} measured samples across {summary['inventory']['arms']} arms.",
        "",
        "## Source custody",
        "",
        f"Before source manifest: `{source['before']['sha256']}` ({source['before']['files']} files). After source manifest: `{source['after']['sha256']}` ({source['after']['files']} files).",
        "",
        f"Changed source files: {source['diff']['changed_files']}; all changes in the configured OPC implementation/test allowlist: `{source['diff']['all_changes_allowed']}`.",
        "",
        f"Matched candidate ZIP archive identities changed for {len(summary['candidate_archive_changes'])} process pairs; decoded candidate content identity is checked independently and archive framing is retained per row.",
        "",
        "## Per-process rows",
        "",
        "| phase | arm | role | repeat | elapsed p50 (ns) | elapsed p95 (ns) | elapsed p99 (ns) | allocator operation heap p50 (B) | heap p95 (B) | heap p99 (B) | GNU time RSS (B, n=1) |",
        "|---|---|---|---:|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for row in summary["processes"]:
        elapsed = row["metrics"]["elapsed_ns"]
        heap = row["metrics"]["allocator_operation_peak_increment_bytes"]
        rss = row["metrics"]["time_max_rss_bytes"]["value"]
        lines.append(
            "| {phase} | {arm} | {role} | {repeat} | {e50} | {e95} | {e99} | {h50} | {h95} | {h99} | {rss} |".format(
                phase=row["phase"], arm=row["arm"], role=row["role"], repeat=row["repeat"],
                e50=_fmt(elapsed["p50"]), e95=_fmt(elapsed["p95"]), e99=_fmt(elapsed["p99"]),
                h50=_fmt(None if heap is None else heap["p50"]),
                h95=_fmt(None if heap is None else heap["p95"]),
                h99=_fmt(None if heap is None else heap["p99"]), rss=_fmt(rss),
            )
        )
    lines.extend([
        "",
        "## Matched before/after changes",
        "",
        "Each row compares one arm, role, repeat, and metric. Percentages use `(after - before) / before`; lower values are favorable for elapsed time and operation heap.",
        "",
        "| arm | role | repeat | metric | p50 change | p50 flag | p95 change | p95 flag | p99 change | p99 flag |",
        "|---|---|---:|---|---:|---|---:|---|---:|---|",
    ])
    for row in summary["comparisons"]:
        changes = row["percent_changes"]
        def pct(name: str) -> str:
            value = changes[name]["percent_change"]
            return "∞" if value is not None and not math.isfinite(value) else f"{value:.3f}%"
        lines.append(
            f"| {row['arm']} | {row['role']} | {row['repeat']} | {row['metric']} | {pct('p50')} | {changes['p50']['review_flag_5pct']} | {pct('p95')} | {changes['p95']['review_flag_5pct']} | {pct('p99')} | {changes['p99']['review_flag_5pct']} |"
        )
    lines.extend([
        "",
        "## Measurement limits",
        "",
    ])
    lines.extend(f"- {item}" for item in summary["limitations"])
    lines.append("")
    return "\n".join(lines)


def parser() -> argparse.ArgumentParser:
    command = argparse.ArgumentParser(description=__doc__)
    subcommands = command.add_subparsers(dest="command", required=True)
    subcommands.add_parser("plan", help="print the explicit matrix without writing it")
    subcommands.add_parser("freeze", help="write the immutable comparison protocol")
    for name, pilot in (("capture-all", False), ("pilot-all", True)):
        sub = subcommands.add_parser(name, help="run one phase of the selected comparison matrix")
        sub.add_argument("--phase", choices=PHASES, required=True)
        sub.add_argument("--attempt", required=True)
        sub.add_argument("--timeout-seconds", type=int, default=DEFAULT_TIMEOUT_SECONDS)
    sub = subcommands.add_parser("capture", help="run one formal arm/role/repeat child")
    sub.add_argument("--phase", choices=PHASES, required=True)
    sub.add_argument("--attempt", required=True)
    sub.add_argument("--arm", choices=tuple(ARM_BY_ID), required=True)
    sub.add_argument("--role", choices=ROLE_NAMES, required=True)
    sub.add_argument("--repeat", type=int, choices=REPEATS, required=True)
    sub.add_argument("--timeout-seconds", type=int, default=DEFAULT_TIMEOUT_SECONDS)
    sub = subcommands.add_parser("pilot", help="run one pilot arm/role child")
    sub.add_argument("--phase", choices=PHASES, required=True)
    sub.add_argument("--attempt", required=True)
    sub.add_argument("--arm", choices=tuple(ARM_BY_ID), required=True)
    sub.add_argument("--role", choices=ROLE_NAMES, required=True)
    sub.add_argument("--timeout-seconds", type=int, default=DEFAULT_TIMEOUT_SECONDS)
    sub = subcommands.add_parser("analyze", help="validate formal captures and write JSON/Markdown summaries")
    sub.add_argument("--before-attempt", required=True)
    sub.add_argument("--after-attempt", required=True)
    sub.add_argument("--summary", type=Path, default=ROOT / SUMMARY_JSON)
    sub.add_argument("--markdown", type=Path, default=ROOT / SUMMARY_MD)
    return command


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    try:
        if args.command == "plan":
            plan()
        elif args.command == "freeze":
            freeze()
        elif args.command in ("capture-all", "pilot-all"):
            require(args.timeout_seconds > 0, "timeout-seconds must be positive")
            capture_all(args.phase, _attempt(args.attempt), pilot=args.command == "pilot-all", timeout_seconds=args.timeout_seconds)
        elif args.command in ("capture", "pilot"):
            require(args.timeout_seconds > 0, "timeout-seconds must be positive")
            phase = args.phase
            attempt = _attempt(args.attempt)
            protocol, _ = _load_protocol()
            builds = _load_builds(protocol)[phase]
            wanted = [
                run for run in _runs(pilot=args.command == "pilot")
                if run["phase"] == phase and run["arm"] == args.arm and run["role"] == args.role
                and (args.command == "pilot" or run["repeat"] == args.repeat)
            ]
            require(len(wanted) == 1, "requested single-child identity is not unique")
            _capture_one(protocol, builds, wanted[0], attempt, pilot=args.command == "pilot", timeout_seconds=args.timeout_seconds)
        elif args.command == "analyze":
            analyze(_attempt(args.before_attempt), _attempt(args.after_attempt), args.summary.resolve(), args.markdown.resolve())
        else:  # pragma: no cover - argparse constrains commands
            fail(f"unknown command: {args.command}")
    except (CompareError, OSError, ValueError, KeyError, subprocess.SubprocessError) as error:
        print(f"compare.py: FAIL: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
