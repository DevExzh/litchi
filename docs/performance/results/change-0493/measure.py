#!/usr/bin/env python3
"""Capture and audit the 0493 managed DOCX read-ahead matrix.

The Rust target owns the managed operation.  This driver owns the process
boundary and makes the evidence replayable: a frozen protocol binds the
source manifest, successful build gates, retained executables, host
environment, and every helper; each child retains raw output, GNU time RSS,
and immutable terminal custody.  ``analyze`` and ``verify`` only consume
those retained files.

The Rust report contract is intentionally small and explicit.  A managed
report has one ``cache`` object per row with ``open`` and ``text`` diagnostic
snapshots.  The snapshots are taken inside the timed package lifetime, while
the report's ``budget`` snapshot is taken after the package is dropped.  The
physical trace is below the managed archive read policy; the pinned logical
trace is the 19-range 0188 DOCX oracle.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import fcntl
import hashlib
import json
import math
import os
from pathlib import Path
import re
import signal
import statistics
import subprocess
import sys
from typing import Any, Iterable, Callable

from support import ENV, ENV_KEYS, REPO, ROOT, TEMP, meta, now, read, sha, snapshot, write


VERSION = 1
CHANGE = 493
SCHEMA = "docx_managed_read_ahead_v1"
PROTOCOL_SCHEMA = "docx-managed-read-ahead-protocol-v1"
CAPTURE_SCHEMA = "docx-managed-read-ahead-capture-v1"
TERMINAL_SCHEMA = "docx-managed-read-ahead-terminal-v1"
ANALYSIS_SCHEMA = "docx-managed-read-ahead-analysis-v1"
VERIFICATION_SCHEMA = "docx-managed-read-ahead-verification-v1"
CASE = "docx_managed_source_open_document_extract_text"
CPU = 2
CPU_LOCK = "/home/zhuhe/.cache/litchi-goal-0484/cpu.lock"
ROLES = ("normal", "allocator")
REPEATS = (1, 2)
FORMAL_SAMPLES = 30
FORMAL_WARMUPS = 3
PILOT_SAMPLES = 3
PILOT_WARMUPS = 1
WINDOW_BYTES = 4096
MAX_RANGE_BYTES = 65_536
DEFAULT_TIMEOUT_SECONDS = 180
TERM_GRACE_SECONDS = 10
REQUEST_BUCKETS = 18
SAMPLE_ALLOCATION_SCOPE = "operation_global_system_allocator"
ATTEMPT_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]*$")
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
REVISION_RE = re.compile(r"^[0-9a-f]{40}$")

# This is the source-edit DOCX used by the preceding 0188 corpus lane.  The
# benchmark refuses a report whose semantic text or archive identity changes.
CORPUS = {
    "version": "0188-media-v1",
    "generator": "litchi-docx-source-edit-media-v1",
    "shape": "200 paragraphs + eight 2 MiB PNG-signature media members",
    "paragraph_count": 200,
    "media_member_count": 8,
    "media_member_bytes": 2 * 1024 * 1024,
    "archive_member_count": 20,
    "archive_bytes": 16_793_036,
    "archive_sha256": "a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4",
    "expected_text_bytes": 10_000,
    "expected_text_sha256": "ad4fe690f0ef2281ad8e64a78d1f4d64e7c8625d672b3ac2f7e9fbc28a82f4af",
}

MEDIA_RANGES = (
    (4027, 2101824), (2101891, 4199688), (4199755, 6297552),
    (6297619, 8395416), (8395483, 10493280), (10493347, 12591144),
    (12591211, 14689008), (14689075, 16786872),
)
LOGICAL_RANGES = (
    (16_793_014, 22), (16_791_709, 46), (16_791_709, 1_305),
    (0, 30), (49, 363), (412, 16), (412, 16), (428, 30), (469, 234),
    (703, 16), (703, 16), (2_907, 30), (2_965, 324), (3_289, 16),
    (3_289, 16), (1_420, 30), (1_467, 1_424), (2_891, 16), (2_891, 16),
)
PHYSICAL_CANDIDATE = ((16_793_014, 22), (16_791_709, 1_327), (0, 4_096))
BASELINE_PHYSICAL_CALLS = 19
BASELINE_PHYSICAL_BYTES = 3966
CANDIDATE_PHYSICAL_CALLS = 3
CANDIDATE_PHYSICAL_BYTES = 5445

TIMING_SCOPE = (
    "SourceBackedPackage managed open (including read-ahead window construction) + "
    "cache snapshot after open + document + extract_text + source-read and cache "
    "snapshots after text + package/document drop; returned text remains live after "
    "the clock; hashing, oracle comparison, source-version checks, physical traces, "
    "and post-drop budget checks are outside"
)
SETUP_SCOPE = (
    "deterministic corpus generation, source/provider construction, bounded physical-"
    "trace capacity reservation, range transport construction, finite limits and "
    "ExecutionContext construction are outside the operation clock; the production "
    "read-ahead window is constructed inside the clock by the managed DOCX owner"
)
ALLOCATION_SCOPE = (
    "operation-scoped global-system-allocator region surrounds the timing scope and "
    "includes managed package open, read-ahead window allocation, document "
    "materialization, text extraction, cache/source-read diagnostics, and "
    "package/document drop; normal binaries report unavailable and allocator binaries "
    "report the separate region sample"
)
PROVIDER_SCOPE = (
    "managed litchi-docx source-backed owner through litchi-opc; the source adapter "
    "is a deterministic in-memory range transport and is not disk or network I/O evidence"
)
PHYSICAL_SCOPE = (
    "CountingReadAt calls below PptxRangeSource and below the production SourceReader "
    "read-ahead policy; requested and returned ranges are physical transport "
    "observations for this synthetic adapter"
)
RANGE_SCOPE = (
    "PptxRangeSource logical adapter counters; fixed delay and transfer pacing "
    "describe the configured synthetic transport, not a filesystem or network "
    "service-level measurement"
)

# ReadLimits::default and the finite managed execution envelope.  Values are
# protocol inputs, not inferred from a result row.
READ_LIMITS = {
    "max_input_bytes": 512 * 1024 * 1024,
    "max_archive_members": 100_000,
    "max_archive_member_name_bytes": 4 * 1024,
    "max_archive_metadata_bytes": 64 * 1024 * 1024,
    "max_archive_compressed_bytes": 512 * 1024 * 1024,
    "max_archive_entry_bytes": 512 * 1024 * 1024,
    "max_archive_total_bytes": 2 * 1024 * 1024 * 1024,
    "max_parts": 100_000,
    "max_part_bytes": 512 * 1024 * 1024,
    "max_total_part_bytes": 512 * 1024 * 1024,
    "max_content_types_bytes": 8 * 1024 * 1024,
    "max_content_type_mappings": 100_000,
    "max_relationship_parts": 100_000,
    "max_relationship_xml_bytes": 8 * 1024 * 1024,
    "max_total_relationship_xml_bytes": 64 * 1024 * 1024,
    "max_relationships_per_part": 100_000,
    "max_total_relationships": 1_000_000,
    "max_relationship_graph_nodes": 100_000,
    "max_xml_events": 1_000_000,
    "max_total_relationship_xml_events": 8_000_000,
    "max_xml_depth": 256,
    "max_xml_attribute_bytes": 64 * 1024,
    "max_relationship_target_bytes": 4 * 1024,
}
LIMITS = {
    "read_limits": READ_LIMITS,
    "cache_max_bytes": 8 * 1024 * 1024,
    "cache_max_entries": 128,
    "budget_memory_bytes": 64 * 1024 * 1024,
    "budget_input_bytes": 64 * 1024 * 1024,
    "budget_output_bytes": 64 * 1024 * 1024,
    "budget_objects": 1_000_000,
    "budget_depth": 1_024,
    "budget_work": 2 * 1024 * 1024 * 1024,
    "execution_workers": 1,
    "execution_max_in_flight_tasks": 1,
    "execution_max_in_flight_bytes": 64 * 1024 * 1024,
    "execution_min_parallel_bytes": 0,
}

ARMS = (
    {"name": "exact-0us", "policy": "exact", "window_bytes": None,
     "max_range_bytes": MAX_RANGE_BYTES, "delay_us": 0,
     "transfer_bytes_per_second": None, "transfer_delay_policy": "separate-sleeps"},
    {"name": "managed-4096-0us", "policy": "forward_start", "window_bytes": WINDOW_BYTES,
     "max_range_bytes": MAX_RANGE_BYTES, "delay_us": 0,
     "transfer_bytes_per_second": None, "transfer_delay_policy": "separate-sleeps"},
    {"name": "exact-1000us-104857600bps", "policy": "exact", "window_bytes": None,
     "max_range_bytes": MAX_RANGE_BYTES, "delay_us": 1_000,
     "transfer_bytes_per_second": 104_857_600, "transfer_delay_policy": "minimum-service"},
    {"name": "managed-4096-1000us-104857600bps", "policy": "forward_start", "window_bytes": WINDOW_BYTES,
     "max_range_bytes": MAX_RANGE_BYTES, "delay_us": 1_000,
     "transfer_bytes_per_second": 104_857_600, "transfer_delay_policy": "minimum-service"},
)
ARM_BY_NAME = {arm["name"]: arm for arm in ARMS}
DRIVER_FILES = ("measure.py", "test_measure.py", "support.py", "gate.py", "retain_build.py")

CACHE_FIELDS = ("open_successful_loads", "successful_loads", "failed_loads",
                "retained_bytes", "retained_entries", "budget_managed")


class ProviderMatrixError(RuntimeError):
    """A fail-closed custody, schema, or recomputation error."""


def fail(message: str) -> None:
    raise ProviderMatrixError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _exact(value: Any, keys: Iterable[str], path: str) -> None:
    require(isinstance(value, dict), f"{path}: expected object")
    expected = set(keys)
    actual = set(value)
    require(actual == expected, f"{path}: fields differ; expected {sorted(expected)}, got {sorted(actual)}")


def _finite(value: Any, path: str = "json") -> None:
    if isinstance(value, float):
        require(math.isfinite(value), f"{path}: non-finite number")
    elif isinstance(value, dict):
        for key, child in value.items():
            require(isinstance(key, str), f"{path}: non-string key")
            _finite(child, f"{path}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            _finite(child, f"{path}[{index}]")


def _uint(value: Any, path: str, *, positive: bool = False) -> int:
    require(type(value) is int and value >= (1 if positive else 0),
            f"{path}: expected {'positive' if positive else 'non-negative'} integer")
    return value


def _text(value: Any, path: str, *, nonempty: bool = True) -> str:
    require(isinstance(value, str) and (bool(value) if nonempty else True),
            f"{path}: expected {'non-empty ' if nonempty else ''}string")
    return value


def _hash(value: Any, path: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
            f"{path}: expected lowercase SHA-256")
    return value


def _json(path: Path) -> Any:
    try:
        value = read(path)
    except (OSError, ValueError, json.JSONDecodeError) as error:
        fail(f"{path}: invalid JSON: {error}")
    _finite(value, str(path))
    return value


def _write(path: Path, value: Any) -> None:
    try:
        write(path, value)
    except FileExistsError:
        fail(f"refusing to replace immutable artifact: {path}")


def _write_or_match(path: Path, value: dict[str, Any], volatile: tuple[str, ...] = ()) -> None:
    if not path.exists():
        _write(path, value)
        return
    old = _json(path)
    require(isinstance(old, dict), f"{path}: existing receipt is not an object")
    expected = dict(value)
    actual = dict(old)
    for field in volatile:
        require(field in actual, f"{path}: volatile field {field} is missing")
        _timestamp(actual[field], f"{path}.{field}")
        actual.pop(field)
        expected.pop(field, None)
    require(actual == expected, f"{path}: immutable receipt differs on rerun")


def _path(value: Any, path: str, *, base: Path = ROOT) -> Path:
    text = _text(value, path)
    result = Path(text)
    if not result.is_absolute():
        require(".." not in result.parts, f"{path}: parent traversal is forbidden")
        result = base / result
    return result


def _file_meta(path: Path, *, executable: bool = False, allow_missing: bool = False) -> dict[str, Any]:
    require(not path.is_symlink(), f"{path}: symlink is forbidden")
    if allow_missing and not path.exists():
        return {"path": str(path)}
    require(path.is_file(), f"missing regular artifact: {path}")
    value = meta(path)
    _uint(value["bytes"], f"{path}.bytes")
    _hash(value["sha256"], f"{path}.sha256")
    if executable:
        require(os.access(path, os.X_OK), f"{path}: executable bit is absent")
    return {"path": str(path), **value}


def _json_hash(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing JSON artifact: {path}")
    return _hash(sha(path), str(path))


def _timestamp(value: Any, path: str) -> _datetime.datetime:
    text = _text(value, path)
    try:
        parsed = _datetime.datetime.fromisoformat(text)
    except ValueError as error:
        fail(f"{path}: invalid timestamp: {error}")
    require(parsed.tzinfo is not None, f"{path}: timezone is required")
    return parsed


def _normalized_snapshot() -> dict[str, Any]:
    value = snapshot()
    _exact(value, ("files", "path", "sha256"), "source snapshot")
    _uint(value["files"], "source snapshot.files", positive=True)
    _hash(value["sha256"], "source snapshot.sha256")
    path = _path(value["path"], "source snapshot.path")
    actual = _file_meta(path)
    require(actual["sha256"] == value["sha256"], "source snapshot digest changed")
    return {"files": value["files"], "path": str(path), "sha256": value["sha256"]}


def _source_binding(value: Any, owner: str) -> dict[str, Any]:
    _exact(value, ("files", "path", "sha256"), owner)
    path = _path(value["path"], f"{owner}.path")
    actual = _file_meta(path)
    _uint(value["files"], f"{owner}.files", positive=True)
    _hash(value["sha256"], f"{owner}.sha256")
    require(actual["sha256"] == value["sha256"], f"{owner}: manifest digest changed")
    manifest = _json(path)
    require(isinstance(manifest, dict) and len(manifest) == value["files"],
            f"{owner}: malformed source manifest")
    for name, digest in manifest.items():
        require(isinstance(name, str) and name and not Path(name).is_absolute()
                and ".." not in Path(name).parts, f"{owner}: unsafe manifest path")
        _hash(digest, f"{owner}.{name}")
    canonical = (json.dumps(manifest, indent=2, sort_keys=True) + "\n").encode()
    require(hashlib.sha256(canonical).hexdigest() == value["sha256"],
            f"{owner}: manifest content does not match binding")
    return {"files": value["files"], "path": str(path), "sha256": value["sha256"]}


def _artifact(value: Any, owner: str, *, retained: bool = True) -> dict[str, Any]:
    _exact(value, ("bytes", "executable", "path", "sha256"), owner)
    path = _path(value["path"], f"{owner}.path")
    require(value["executable"] is True, f"{owner}.executable is false")
    _uint(value["bytes"], f"{owner}.bytes", positive=True)
    _hash(value["sha256"], f"{owner}.sha256")
    if retained or path.exists():
        actual = _file_meta(path, executable=True)
        require(actual["bytes"] == value["bytes"] and actual["sha256"] == value["sha256"],
                f"{owner}: artifact identity changed")
    return {"path": str(path), "bytes": value["bytes"], "sha256": value["sha256"], "executable": True}


def _expected_build_command(role: str) -> list[str]:
    require(role in ROLES, f"unknown build role {role}")
    name = "litchi-perf-baseline" + ("-alloc" if role == "allocator" else "")
    command = ["cargo", "build", "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml"]
    if role == "allocator":
        command += ["--features", "allocator-metrics"]
    return command + ["--bin", name]


def _gate_binding(value: Any, owner: Path) -> dict[str, Any]:
    _exact(value, ("path", "sha256"), f"{owner}.gate")
    path = _path(value["path"], f"{owner}.gate.path")
    require(path.is_file() and not path.is_symlink(), f"{path}: gate receipt is missing")
    _hash(value["sha256"], f"{owner}.gate.sha256")
    require(_json_hash(path) == value["sha256"], f"{path}: gate receipt changed")
    receipt = _json(path)
    required = {
        "argv", "artifacts", "attempt", "common_sha256", "cwd", "driver_sha256",
        "environment", "exit_code", "finished_utc", "label", "schema",
        "source_after", "source_before", "source_unchanged", "started_utc",
        "termination", "timed_out", "timeout_seconds",
    }
    _exact(receipt, required, str(path))
    require(receipt["schema"] == "docx-managed-read-ahead-gate-v1", f"{path}: gate schema differs")
    require(receipt["exit_code"] == 0 and receipt["timed_out"] is False
            and receipt["termination"] is None and receipt["source_unchanged"] is True,
            f"{path}: gate did not pass unchanged")
    require(receipt["cwd"] == str(REPO), f"{path}: gate cwd differs")
    require(receipt["environment"] == {key: ENV[key] for key in ENV_KEYS}, f"{path}: gate environment differs")
    require(receipt["driver_sha256"] == _json_hash(ROOT / "gate.py"), f"{path}: gate helper changed")
    require(receipt["common_sha256"] == _json_hash(ROOT / "support.py"), f"{path}: support helper changed")
    require(_timestamp(receipt["finished_utc"], f"{path}.finished_utc") >
            _timestamp(receipt["started_utc"], f"{path}.started_utc"), f"{path}: chronology invalid")
    before = _source_binding(receipt["source_before"], f"{path}.source_before")
    after = _source_binding(receipt["source_after"], f"{path}.source_after")
    require(before == after, f"{path}: gate source changed")
    require(isinstance(receipt["argv"], list) and receipt["argv"] and
            all(isinstance(item, str) and item for item in receipt["argv"]), f"{path}: argv missing")
    names = (f"{path.stem}.stderr", f"{path.stem}.stdout")
    _exact(receipt["artifacts"], names, f"{path}.artifacts")
    for name in names:
        item = receipt["artifacts"][name]
        _exact(item, ("bytes", "sha256"), f"{path}.artifacts.{name}")
        _uint(item["bytes"], f"{path}.artifacts.{name}.bytes")
        _hash(item["sha256"], f"{path}.artifacts.{name}.sha256")
        actual = _file_meta(path.with_suffix("." + name.rsplit(".", 1)[1]))
        require(actual["bytes"] == item["bytes"] and actual["sha256"] == item["sha256"],
                f"{path}: gate artifact changed")
    return {"path": str(path), "sha256": value["sha256"], "source": after, "argv": list(receipt["argv"])}


def _build_from(value: Any, role: str, path: Path) -> dict[str, Any]:
    required = {
        "attempt", "binary", "command", "copied_utc", "environment", "gate",
        "original_binary", "role", "schema", "source_after", "source_before",
        "source_unchanged", "version", "git_revision", "retainer_sha256",
    }
    _exact(value, required, str(path))
    require(value["role"] == role and value["schema"] == "docx-provider-lifecycle-build-v1"
            and value["version"] == 1 and value["source_unchanged"] is True,
            f"{path}: build identity differs")
    require(value["retainer_sha256"] == _json_hash(ROOT / "retain_build.py"), f"{path}: retainer changed")
    _text(value["attempt"], f"{path}.attempt")
    _timestamp(value["copied_utc"], f"{path}.copied_utc")
    require(value["command"] == _expected_build_command(role), f"{path}: build argv differs")
    require(value["environment"] == {key: ENV[key] for key in ENV_KEYS}, f"{path}: build environment differs")
    revision = _text(value["git_revision"], f"{path}.git_revision")
    require(REVISION_RE.fullmatch(revision) is not None, f"{path}: malformed git revision")
    binary = _artifact(value["binary"], f"{path}.binary")
    original = _artifact(value["original_binary"], f"{path}.original_binary", retained=False)
    require(Path(binary["path"]).name == value["command"][-1] and
            Path(original["path"]).name == value["command"][-1], f"{path}: binary name differs")
    require(binary["bytes"] == original["bytes"] and binary["sha256"] == original["sha256"],
            f"{path}: binary copies differ")
    before = _source_binding(value["source_before"], f"{path}.source_before")
    after = _source_binding(value["source_after"], f"{path}.source_after")
    require(before == after, f"{path}: build source changed")
    gate = _gate_binding(value["gate"], path)
    require(gate["source"] == after and gate["argv"] == value["command"], f"{path}: gate binding differs")
    return {
        "path": str(path), "receipt_sha256": _json_hash(path), "role": role,
        "binary": binary, "original_binary": original, "source": after,
        "gate": gate, "command": list(value["command"]),
        "environment": dict(value["environment"]), "git_revision": revision,
    }


def load_builds(build_dir: Path | None = None) -> dict[str, dict[str, Any]]:
    directory = ROOT if build_dir is None else Path(build_dir)
    result = {
        role: _build_from(_json(directory / f"build-{role}.json"), role,
                          directory / f"build-{role}.json") for role in ROLES
    }
    require(result["normal"]["source"] == result["allocator"]["source"], "build source manifests differ")
    require(result["normal"]["git_revision"] == result["allocator"]["git_revision"], "build revisions differ")
    return result


def _environment_binding() -> dict[str, str]:
    path = ROOT / "environment.json"
    return {"path": "environment.json", "sha256": _json_hash(path)}


def formal_inventory(*, pilot: bool = False) -> list[dict[str, Any]]:
    samples = PILOT_SAMPLES if pilot else FORMAL_SAMPLES
    warmups = PILOT_WARMUPS if pilot else FORMAL_WARMUPS
    repeats = (1,) if pilot else REPEATS
    result = []
    for repeat in repeats:
        roles = ROLES if repeat == 1 else tuple(reversed(ROLES))
        arms = ARMS if repeat == 1 else tuple(reversed(ARMS))
        for role in roles:
            for arm in arms:
                result.append({
                    "kind": "pilot" if pilot else "formal", "repeat": repeat,
                    "role": role, "arm": arm["name"], "provider": "range",
                    "samples": samples, "warmups": warmups,
                    "label": f"{'pilot-' if pilot else ''}r{repeat}-{role}-{arm['name']}",
                })
    return result


def _build_binding(build: dict[str, Any]) -> dict[str, Any]:
    return {"receipt_sha256": build["receipt_sha256"], "binary": build["binary"],
            "source": build["source"], "gate": build["gate"],
            "command": build["command"], "environment": build["environment"],
            "git_revision": build["git_revision"]}


def protocol_value(builds: dict[str, dict[str, Any]] | None = None) -> dict[str, Any]:
    value: dict[str, Any] = {
        "schema": PROTOCOL_SCHEMA, "version": VERSION, "change": CHANGE,
        "claim_authorized": False,
        "performance_claim": "descriptive managed policy evidence; no cross-format or cross-arm optimization claim",
        "case": CASE, "cpu": CPU, "cpu_lock": CPU_LOCK,
        "roles": list(ROLES), "repeats": list(REPEATS),
        "samples": FORMAL_SAMPLES, "warmups": FORMAL_WARMUPS,
        "expected_formal_processes": 16, "expected_measured_samples": 480,
        "formal_runs": formal_inventory(), "pilot_runs": formal_inventory(pilot=True),
        "arms": [dict(arm) for arm in ARMS], "corpus": dict(CORPUS),
        "media_ranges": [{"start": start, "end": end} for start, end in MEDIA_RANGES],
        "logical_ranges": [{"offset": offset, "requested": size, "returned": size}
                           for offset, size in LOGICAL_RANGES],
        "physical_observation": {
            "baseline": {"calls": BASELINE_PHYSICAL_CALLS, "returned_bytes": BASELINE_PHYSICAL_BYTES},
            "candidate": {"calls": CANDIDATE_PHYSICAL_CALLS, "returned_bytes": CANDIDATE_PHYSICAL_BYTES},
        },
        "oracle": {"text_bytes": CORPUS["expected_text_bytes"],
                    "text_sha256": CORPUS["expected_text_sha256"]},
        "limits": LIMITS,
        "scopes": {"provider": PROVIDER_SCOPE, "timing": TIMING_SCOPE,
                    "setup": SETUP_SCOPE, "allocation": ALLOCATION_SCOPE,
                    "physical": PHYSICAL_SCOPE, "range": RANGE_SCOPE},
        "environment_artifact": _environment_binding(),
        "environment": {key: ENV[key] for key in ENV_KEYS},
        "driver": {name: _json_hash(ROOT / name) for name in DRIVER_FILES},
        "source": None, "builds": None,
    }
    if builds is not None:
        value["source"] = builds["normal"]["source"]
        value["builds"] = {role: _build_binding(builds[role]) for role in ROLES}
    return value


def create_freeze(plan: dict[str, Any] | None = None,
                  builds: dict[str, dict[str, Any]] | None = None) -> Path:
    """Create the immutable protocol after successful retained builds."""
    builds = load_builds() if builds is None else builds
    expected = protocol_value(builds)
    value = expected if plan is None else plan
    require(value == expected, "freeze plan does not equal current build-bound protocol")
    path = ROOT / "protocol.json"
    _write_or_match(path, value)
    return path


def _load_protocol(builds: dict[str, dict[str, Any]] | None = None) -> tuple[dict[str, Any], str]:
    path = ROOT / "protocol.json"
    require(path.is_file() and not path.is_symlink(), f"frozen protocol is missing: {path}")
    value = _json(path)
    require(value == protocol_value(builds) if builds is not None else
            value["schema"] == PROTOCOL_SCHEMA, f"{path}: frozen protocol differs")
    _exact(value, protocol_value(builds).keys() if builds is not None else value.keys(), str(path))
    require(value["schema"] == PROTOCOL_SCHEMA and value["version"] == VERSION
            and value["change"] == CHANGE and value["claim_authorized"] is False,
            f"{path}: protocol identity differs")
    require(value["formal_runs"] == formal_inventory() and value["pilot_runs"] == formal_inventory(pilot=True),
            f"{path}: run inventory differs")
    require(value["arms"] == [dict(arm) for arm in ARMS] and value["corpus"] == CORPUS,
            f"{path}: arms or corpus differ")
    require(value["driver"] == {name: _json_hash(ROOT / name) for name in DRIVER_FILES},
            f"{path}: helper binding differs")
    env_artifact = value["environment_artifact"]
    _exact(env_artifact, ("path", "sha256"), f"{path}.environment_artifact")
    require(_json_hash(_path(env_artifact["path"], f"{path}.environment_artifact.path")) == env_artifact["sha256"],
            f"{path}: environment artifact changed")
    if builds is None:
        require(value["source"] is None and value["builds"] is None, f"{path}: unbound plan is not capturable")
    return value, _json_hash(path)


def _arm(name: str) -> dict[str, Any]:
    try:
        return dict(ARM_BY_NAME[name])
    except KeyError:
        fail(f"unknown arm {name}")


def _command(spec: dict[str, Any], build: dict[str, Any], report: Path, resource: Path,
             revision: str) -> list[str]:
    arm = _arm(spec["arm"])
    command = ["/usr/bin/time", "-v", "-o", str(resource), "/usr/bin/taskset", "-c", str(CPU),
               str(build["binary"]["path"]), "docx-managed-read-ahead",
               "--policy", arm["policy"], "--max-range", str(arm["max_range_bytes"]),
               "--delay-us", str(arm["delay_us"])]
    if arm["window_bytes"] is not None:
        command += ["--window-bytes", str(arm["window_bytes"])]
    if arm["transfer_bytes_per_second"] is not None:
        command += ["--transfer-bytes-per-second", str(arm["transfer_bytes_per_second"]),
                    "--transfer-delay-policy", arm["transfer_delay_policy"]]
    command += ["--samples", str(spec["samples"]), "--warmup", str(spec["warmups"]),
                "--source-revision", revision, "--output", str(report)]
    return command


def _source_revision() -> str:
    try:
        revision = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=REPO, text=True).strip()
    except (OSError, subprocess.SubprocessError) as error:
        fail(f"cannot determine source revision: {error}")
    require(REVISION_RE.fullmatch(revision) is not None, "malformed source revision")
    return revision


def _run_root(attempt: str, label: str) -> Path:
    return TEMP / "managed" / attempt / label


def _capture_root(attempt: str, label: str) -> Path:
    return ROOT / "captures" / attempt / label


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


def _cleanup_private(run_root: Path, tmp_root: Path) -> dict[str, Any]:
    expected_parent = TEMP / "managed"
    require(tmp_root.parent == run_root and run_root.parent.parent == expected_parent,
            "private scratch escaped managed owner")
    removed: list[str] = []
    remaining: list[str] = []
    for path in (tmp_root, run_root):
        if path.exists():
            require(path.is_dir() and not path.is_symlink(), f"private path is not a directory: {path}")
            children = list(path.iterdir())
            if children:
                remaining.extend(str(child) for child in children)
                break
            path.rmdir()
            removed.append(str(path))
    attempt_parent = run_root.parent
    if not remaining and attempt_parent.exists():
        require(attempt_parent.is_dir() and not attempt_parent.is_symlink(), "attempt scratch parent is unsafe")
        if not list(attempt_parent.iterdir()):
            attempt_parent.rmdir()
            removed.append(str(attempt_parent))
        else:
            remaining.extend(str(child) for child in attempt_parent.iterdir())
    return {"schema": "docx-managed-private-cleanup-v1",
            "status": "pass" if not remaining else "failed",
            "root": str(run_root), "tmpdir": str(tmp_root),
            "removed": removed, "remaining": remaining}


def _with_cpu_lock(action: Callable[[], Any]) -> Any:
    lock_path = Path(CPU_LOCK)
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    with lock_path.open("a+") as lock:
        fcntl.flock(lock.fileno(), fcntl.LOCK_EX)
        try:
            return action()
        finally:
            fcntl.flock(lock.fileno(), fcntl.LOCK_UN)


def _resource_values(path: Path) -> dict[str, Any]:
    raw = path.read_bytes()
    require(raw, f"{path}: GNU time receipt is empty")
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
    _uint(result.get("maximum_resident_set_size_(kbytes)"), f"{path}: maximum RSS")
    return result


def _launch_unlocked(spec: dict[str, Any], build: dict[str, Any], protocol: dict[str, Any],
                     protocol_hash: str, timeout_seconds: int) -> Path:
    attempt = _attempt(spec["attempt"])
    label = _text(spec["label"], "capture label")
    capture = _capture_root(attempt, label)
    require(not capture.exists(), f"refusing to replace immutable capture: {capture}")
    capture.parent.mkdir(parents=True, exist_ok=True)
    capture.mkdir(exist_ok=False)
    run_root = _run_root(attempt, label)
    require(not run_root.exists(), f"refusing to replace private scratch: {run_root}")
    run_root.mkdir(parents=True, exist_ok=False)
    tmp_root = run_root / "tmp"
    tmp_root.mkdir()
    report, resource = capture / "report.json", capture / "resource.txt"
    stdout, stderr = capture / "stdout.txt", capture / "stderr.txt"
    stdout.touch(), stderr.touch()
    revision = _source_revision()
    require(revision == build["git_revision"], "capture HEAD differs from retained build")
    current = _normalized_snapshot()
    require(current == build["source"] == protocol["source"], "capture source differs from frozen build")
    argv = _command(spec, build, report, resource, revision)
    environment = {key: ENV[key] for key in ENV_KEYS}
    environment["TMPDIR"] = str(tmp_root)
    started = {
        "schema": CAPTURE_SCHEMA, "version": VERSION, "status": "running",
        "attempt": attempt, "run": {key: spec[key] for key in
                                      ("kind", "repeat", "role", "arm", "provider", "samples", "warmups", "label")},
        "protocol": {"path": "protocol.json", "sha256": protocol_hash},
        "build": {"path": build["path"], "sha256": build["receipt_sha256"], "source": build["source"]},
        "binary": build["binary"], "source": current, "argv": argv, "cwd": str(REPO),
        "environment": environment, "environment_artifact": protocol["environment_artifact"],
        "driver_bindings": protocol["driver"], "driver_sha256": _json_hash(ROOT / "measure.py"),
        "support_sha256": _json_hash(ROOT / "support.py"), "started_utc": now(),
        "timeout_seconds": timeout_seconds, "tmpdir": str(tmp_root),
    }
    _write(capture / "started.json", started)
    process: subprocess.Popen[bytes] | None = None
    timed_out = False
    termination: str | None = None
    launch_error: str | None = None
    validation_error: str | None = None
    exit_code: int | None = None
    try:
        run_environment = dict(ENV)
        run_environment["TMPDIR"] = str(tmp_root)
        with stdout.open("wb") as out, stderr.open("wb") as err:
            process = subprocess.Popen(argv, cwd=REPO, env=run_environment,
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
            validate_report(report, role=spec["role"], arm_name=spec["arm"],
                            samples=spec["samples"], warmups=spec["warmups"],
                            source_revision=revision)
        except (ProviderMatrixError, OSError, ValueError) as error:
            validation_error = str(error)
    try:
        after = _normalized_snapshot()
        unchanged = after == current
    except (OSError, ValueError, ProviderMatrixError) as error:
        after, unchanged = {"error": f"{type(error).__name__}: {error}"}, False
    try:
        cleanup = _cleanup_private(run_root, tmp_root)
    except (ProviderMatrixError, OSError) as error:
        cleanup = {"schema": "docx-managed-private-cleanup-v1", "status": "failed",
                   "root": str(run_root), "tmpdir": str(tmp_root), "removed": [],
                   "remaining": [str(error)]}
    _write(capture / "replay-cleanup.json", cleanup)
    expected_artifacts = (stdout, stderr, resource, report, capture / "replay-cleanup.json")
    artifacts = {path.name: _file_meta(path) for path in expected_artifacts if path.is_file()}
    missing = [path.name for path in expected_artifacts if not path.is_file()]
    passed = (exit_code == 0 and not timed_out and termination is None and launch_error is None
              and validation_error is None and unchanged and not missing and cleanup["status"] == "pass")
    terminal = dict(started)
    terminal.update({"schema": TERMINAL_SCHEMA, "status": "pass" if passed else "failed",
                     "exit_code": exit_code, "timed_out": timed_out, "termination": termination,
                     "finished_utc": now(), "artifacts": artifacts, "missing_artifacts": missing,
                     "tmpdir": str(tmp_root), "cleanup": cleanup,
                     "started_artifact": _file_meta(capture / "started.json"),
                     "source_before": current, "source_after": after, "source_unchanged": unchanged})
    if launch_error is not None:
        terminal["launch_error"] = launch_error
    if validation_error is not None:
        terminal["validation_error"] = validation_error
    _write(capture / "terminal.json", terminal)
    if not passed:
        fail(f"{label} failed; receipt retained at {capture / 'terminal.json'}")
    return capture / "terminal.json"


def _attempt(value: str) -> str:
    require(isinstance(value, str) and ATTEMPT_RE.fullmatch(value) is not None,
            "attempt must be a path-safe token")
    return value


def _capture_one(spec: dict[str, Any], build: dict[str, Any], protocol: dict[str, Any],
                 protocol_hash: str, timeout_seconds: int) -> Path:
    return _with_cpu_lock(lambda: _launch_unlocked(spec, build, protocol, protocol_hash, timeout_seconds))


def _check_corpus(value: Any, path: str) -> None:
    require(isinstance(value, dict), f"{path}: corpus manifest missing")
    require(value.get("generator") == CORPUS["generator"], f"{path}: corpus generator differs")
    require(value.get("archive_bytes") == CORPUS["archive_bytes"] and
            value.get("archive_sha256") == CORPUS["archive_sha256"] and
            value.get("archive_member_count") == CORPUS["archive_member_count"],
            f"{path}: archive identity differs")


def _check_limits(value: Any, path: str) -> None:
    _exact(value, LIMITS, path)
    require(value == LIMITS, f"{path}: finite limit contract differs")


def _check_provider(value: Any, arm: dict[str, Any], path: str) -> None:
    _exact(value, ("name", "max_range_bytes", "delay_us", "transfer_bytes_per_second",
                   "transfer_delay_policy", "physical_scope", "range_scope"), path)
    require(value["name"] == "PptxRangeSource" and value["max_range_bytes"] == arm["max_range_bytes"]
            and value["delay_us"] == arm["delay_us"]
            and value["transfer_bytes_per_second"] == arm["transfer_bytes_per_second"]
            and value["transfer_delay_policy"] == arm["transfer_delay_policy"], f"{path}: arm differs")
    require(value["physical_scope"] == PHYSICAL_SCOPE and value["range_scope"] == RANGE_SCOPE,
            f"{path}: scope differs")


def _check_version(value: Any, path: str) -> None:
    _exact(value, ("id", "revision"), path)
    _uint(value["id"], f"{path}.id", positive=True)
    _uint(value["revision"], f"{path}.revision")


def _check_range(value: Any, path: str) -> tuple[int, int, int]:
    _exact(value, ("offset", "requested", "returned"), path)
    offset = _uint(value["offset"], f"{path}.offset")
    requested = _uint(value["requested"], f"{path}.requested", positive=True)
    returned = _uint(value["returned"], f"{path}.returned")
    require(returned <= requested and offset + requested <= 2**64 - 1, f"{path}: invalid range conservation")
    return offset, requested, returned


def _bucket(size: int) -> int:
    return 0 if size <= 1 else min(REQUEST_BUCKETS - 1, (size - 1).bit_length())


def _check_physical(value: Any, arm: dict[str, Any], path: str) -> dict[str, Any]:
    _exact(value, ("scope", "calls", "requested_bytes", "returned_bytes", "short_reads", "ranges"), path)
    require(value["scope"] == PHYSICAL_SCOPE, f"{path}.scope differs")
    for key in ("calls", "requested_bytes", "returned_bytes", "short_reads"):
        _uint(value[key], f"{path}.{key}")
    require(value["returned_bytes"] <= value["requested_bytes"], f"{path}: returned bytes exceed requested")
    require(isinstance(value["ranges"], list) and len(value["ranges"]) == value["calls"], f"{path}: trace count differs")
    ranges = [_check_range(item, f"{path}.ranges[{index}]") for index, item in enumerate(value["ranges"])]
    require(all(item[2] == item[1] for item in ranges) and value["short_reads"] == 0, f"{path}: unexpected short read")
    expected = LOGICAL_RANGES if arm["policy"] == "exact" else PHYSICAL_CANDIDATE
    require([(offset, requested) for offset, requested, _ in ranges] == list(expected), f"{path}: physical trace differs")
    require(value["calls"] == len(expected) and value["requested_bytes"] == sum(size for _, size in expected)
            and value["returned_bytes"] == value["requested_bytes"], f"{path}: physical totals differ")
    return {"calls": value["calls"], "requested_bytes": value["requested_bytes"],
            "returned_bytes": value["returned_bytes"], "ranges": ranges}


def _check_transport(value: Any, arm: dict[str, Any], physical: dict[str, Any], path: str) -> None:
    fields = ("logical_calls", "requested_bytes", "returned_bytes", "min_request_bytes",
              "max_request_bytes", "short_reads", "delayed_calls", "transfer_paced_calls",
              "transfer_delay_ns", "request_size_counts")
    _exact(value, fields, path)
    for key in ("logical_calls", "requested_bytes", "returned_bytes", "short_reads",
                "delayed_calls", "transfer_paced_calls", "transfer_delay_ns"):
        _uint(value[key], f"{path}.{key}")
    require(value["logical_calls"] == physical["calls"] and value["requested_bytes"] == physical["requested_bytes"]
            and value["returned_bytes"] == physical["returned_bytes"] and value["short_reads"] == 0,
            f"{path}: transport and physical totals differ")
    for key in ("min_request_bytes", "max_request_bytes"):
        if value[key] is not None:
            _uint(value[key], f"{path}.{key}", positive=True)
    counts = value["request_size_counts"]
    require(isinstance(counts, list) and len(counts) == REQUEST_BUCKETS, f"{path}: request histogram shape differs")
    for item in counts:
        _uint(item, f"{path}.request_size_counts")
    require(sum(counts) == value["logical_calls"], f"{path}: request histogram does not conserve calls")
    expected_counts = [0] * REQUEST_BUCKETS
    for _, size, _ in physical["ranges"]:
        expected_counts[_bucket(size)] += 1
    require(counts == expected_counts, f"{path}: request histogram differs")
    require(value["delayed_calls"] == value["logical_calls"] and
            value["transfer_paced_calls"] == (value["logical_calls"] if arm["transfer_bytes_per_second"] else 0),
            f"{path}: delay accounting differs")


def _check_budget(value: Any, physical: dict[str, Any], path: str) -> None:
    fields = ("memory_before", "memory_live", "memory_after_drop", "memory_released",
              "input_bytes_before", "input_bytes_after_drop", "input_bytes_delta",
              "objects_before", "objects_live", "objects_after_drop", "objects_released",
              "managed", "memory_released_to_baseline", "input_bytes_match_physical_returned")
    _exact(value, fields, path)
    for key in fields:
        if key not in {"managed", "memory_released_to_baseline", "input_bytes_match_physical_returned"}:
            _uint(value[key], f"{path}.{key}")
    require(value["managed"] is True and value["memory_before"] == 0
            and value["memory_after_drop"] == 0 and value["objects_before"] == 0
            and value["objects_after_drop"] == 0
            and value["input_bytes_after_drop"] - value["input_bytes_before"] == value["input_bytes_delta"]
            and value["input_bytes_delta"] == physical["returned_bytes"]
            and value["input_bytes_match_physical_returned"] is True,
            f"{path}: managed resource accounting differs")
    require(value["memory_released"] == value["memory_live"] - value["memory_after_drop"]
            and value["objects_released"] == value["objects_live"] - value["objects_after_drop"],
            f"{path}: release accounting differs")
    require(value["memory_released_to_baseline"] is True and
            value["input_bytes_after_drop"] >= value["input_bytes_before"], f"{path}: resource regression")


def _check_cache(value: Any, path: str) -> dict[str, Any]:
    _exact(value, CACHE_FIELDS, path)
    for key in ("open_successful_loads", "successful_loads", "failed_loads",
                "retained_bytes", "retained_entries"):
        _uint(value[key], f"{path}.{key}")
    require(value["budget_managed"] is True, f"{path}.budget_managed: managed cache expected")
    require(value["open_successful_loads"] == 0 and value["successful_loads"] == 1
            and value["failed_loads"] == 0 and value["retained_entries"] == 1,
            f"{path}: expected zero open loads and one text load")
    return dict(value)


def _cache_snapshots(row: dict[str, Any], path: str) -> tuple[dict[str, Any], dict[str, Any]]:
    """Normalize the flat Rust CacheRecord into its two lifecycle points."""
    value = _check_cache(row.get("cache"), f"{path}.cache")
    opened = {"successful_loads": value["open_successful_loads"],
              "failed_loads": value["failed_loads"], "retained_entries": 0,
              "retained_bytes": 0, "budget_managed": value["budget_managed"]}
    text = {"successful_loads": value["successful_loads"],
            "failed_loads": value["failed_loads"], "retained_entries": value["retained_entries"],
            "retained_bytes": value["retained_bytes"], "budget_managed": value["budget_managed"]}
    return opened, text


def _check_allocation(value: Any, role: str, path: str) -> dict[str, Any] | None:
    require(isinstance(value, dict), f"{path}: allocation envelope is missing")
    if role == "normal":
        require(value == {}, f"{path}: normal allocation metrics must be absent")
        return None
    require(set(value) == {"sample", "region_peak_increment_bytes"},
            f"{path}: allocator envelope fields differ")
    sample = value["sample"]
    require(isinstance(sample, dict), f"{path}.sample: missing")
    status = sample.get("status")
    fields = ("status", "scope", "allocation_calls", "deallocation_calls", "reallocation_calls",
              "failed_allocation_calls", "allocated_bytes", "deallocated_bytes", "live_bytes_before",
              "live_bytes_after", "peak_live_bytes_before", "peak_live_bytes_after", "region_peak_live_bytes")
    _exact(sample, fields, f"{path}.sample")
    require(status == "measured" and sample["scope"] == SAMPLE_ALLOCATION_SCOPE, f"{path}: allocator sample unavailable")
    for key in fields[2:]:
        _uint(sample[key], f"{path}.sample.{key}")
    require(sample["live_bytes_after"] >= 0 and sample["peak_live_bytes_after"] >= sample["peak_live_bytes_before"]
            and sample["region_peak_live_bytes"] >= sample["live_bytes_before"], f"{path}: allocator counter order differs")
    if "region_peak_increment_bytes" in value:
        _uint(value["region_peak_increment_bytes"], f"{path}.region_peak_increment_bytes")
        require(value["region_peak_increment_bytes"] == sample["region_peak_live_bytes"] - sample["live_bytes_before"],
                f"{path}: region peak increment differs")
    else:
        fail(f"{path}: allocator region increment is missing")
    return {**sample, "region_peak_increment_bytes": value["region_peak_increment_bytes"]}


def _check_source_read(value: Any, arm: dict[str, Any], path: str) -> dict[str, Any] | None:
    if arm["policy"] == "exact":
        require(value is None, f"{path}: exact policy published read-ahead diagnostics")
        return None
    require(isinstance(value, dict), f"{path}: managed diagnostics are missing")
    fields = ("enabled", "configured_window_bytes", "retained_window_bytes", "requests", "hits",
              "misses", "fills", "requested_bytes", "returned_bytes")
    _exact(value, fields, path)
    require(value["enabled"] is True and value["configured_window_bytes"] == WINDOW_BYTES
            and value["retained_window_bytes"] == WINDOW_BYTES, f"{path}: managed window identity differs")
    for key in fields[3:]:
        _uint(value[key], f"{path}.{key}")
    require(value["requests"] == len(LOGICAL_RANGES) and value["hits"] == 16 and value["misses"] == 3
            and value["fills"] == 3 and value["requested_bytes"] == BASELINE_PHYSICAL_BYTES
            and value["returned_bytes"] == CANDIDATE_PHYSICAL_BYTES,
            f"{path}: managed read counters differ")
    return dict(value)


def validate_report(path: Path, *, role: str, arm_name: str, samples: int,
                    warmups: int, source_revision: str | None = None) -> dict[str, Any]:
    value = _json(path)
    required = ("schema", "version", "case_name", "provider_scope", "timing_scope", "setup_scope",
                "allocation_scope", "corpus_version", "corpus_generator", "corpus", "source_bytes",
                "source_sha256", "expected_text_bytes", "expected_text_sha256", "expected_archive_members",
                "media_ranges", "expected_exact_ranges", "expected_exact_physical_calls",
                "expected_exact_physical_bytes", "requested_source_revision", "limits", "provider", "policy",
                "warmup", "samples", "rows")
    _exact(value, required, str(path))
    arm = _arm(arm_name)
    require(value["schema"] == SCHEMA and value["version"] == 1 and value["case_name"] == CASE,
            f"{path}: report identity differs")
    require(value["provider_scope"] == PROVIDER_SCOPE and value["timing_scope"] == TIMING_SCOPE
            and value["setup_scope"] == SETUP_SCOPE and value["allocation_scope"] == ALLOCATION_SCOPE,
            f"{path}: report scope differs")
    require(value["corpus_version"] in ("source-edit-media-v1", "0188-media-v1")
            and value["corpus_generator"] == CORPUS["generator"], f"{path}: corpus identity differs")
    _check_corpus(value["corpus"], f"{path}.corpus")
    require(value["source_bytes"] == CORPUS["archive_bytes"] and value["source_sha256"] == CORPUS["archive_sha256"]
            and value["expected_text_bytes"] == CORPUS["expected_text_bytes"]
            and value["expected_text_sha256"] == CORPUS["expected_text_sha256"]
            and value["expected_archive_members"] == CORPUS["archive_member_count"], f"{path}: oracle differs")
    _check_limits(value["limits"], f"{path}.limits")
    _check_provider(value["provider"], arm, f"{path}.provider")
    policy = value["policy"]
    _exact(policy, ("name", "enabled", "configured_window_bytes"), f"{path}.policy")
    require(policy["name"] == arm["policy"] and policy["enabled"] is (arm["policy"] != "exact")
            and policy["configured_window_bytes"] == arm["window_bytes"], f"{path}: policy differs")
    require(value["warmup"] == warmups and value["samples"] == samples, f"{path}: sample counts differ")
    revision = _text(value["requested_source_revision"], f"{path}.requested_source_revision")
    require(REVISION_RE.fullmatch(revision) is not None, f"{path}: malformed source revision")
    if source_revision is not None:
        require(revision == source_revision, f"{path}: source revision differs")
    media = value["media_ranges"]
    require(isinstance(media, list) and len(media) == len(MEDIA_RANGES), f"{path}: media range inventory differs")
    require([(item.get("start"), item.get("end")) for item in media] == list(MEDIA_RANGES), f"{path}: media ranges differ")
    expected_ranges = [{"offset": o, "requested": n, "returned": n} for o, n in LOGICAL_RANGES]
    require(value["expected_exact_ranges"] == expected_ranges and value["expected_exact_physical_calls"] == BASELINE_PHYSICAL_CALLS
            and value["expected_exact_physical_bytes"] == BASELINE_PHYSICAL_BYTES, f"{path}: exact oracle differs")
    rows = value["rows"]
    require(isinstance(rows, list) and len(rows) == samples, f"{path}: row count differs")
    for index, row in enumerate(rows):
        _check_row(row, role, arm, f"{path}.rows[{index}]", warmups + index)
    return value


def _check_row(row: Any, role: str, arm: dict[str, Any], path: str, expected_index: int) -> None:
    fields = ("sample_index", "elapsed_ns", "actual_text_verified", "actual_text_bytes", "actual_text_sha256",
              "source_version_before", "source_version_after", "source_version_unchanged", "policy",
              "source_read", "physical", "transport", "budget", "allocation", "cache")
    _exact(row, fields, path)
    _uint(row["sample_index"], f"{path}.sample_index")
    require(row["sample_index"] == expected_index, f"{path}: sample index differs")
    _uint(row["elapsed_ns"], f"{path}.elapsed_ns", positive=True)
    require(row["actual_text_verified"] is True and row["actual_text_bytes"] == CORPUS["expected_text_bytes"]
            and row["actual_text_sha256"] == CORPUS["expected_text_sha256"]
            and row["source_version_unchanged"] is True, f"{path}: text/source oracle failed")
    _check_version(row["source_version_before"], f"{path}.source_version_before")
    _check_version(row["source_version_after"], f"{path}.source_version_after")
    require(row["source_version_before"] == row["source_version_after"], f"{path}: source version changed")
    policy = row["policy"]
    _exact(policy, ("name", "enabled", "configured_window_bytes"), f"{path}.policy")
    require(policy["name"] == arm["policy"] and policy["enabled"] is (arm["policy"] != "exact")
            and policy["configured_window_bytes"] == arm["window_bytes"], f"{path}: row policy differs")
    _check_source_read(row["source_read"], arm, f"{path}.source_read")
    physical = _check_physical(row["physical"], arm, f"{path}.physical")
    _check_transport(row["transport"], arm, physical, f"{path}.transport")
    _check_budget(row["budget"], physical, f"{path}.budget")
    _check_allocation(row["allocation"], role, f"{path}.allocation")
    opened, text = _cache_snapshots(row, path)
    require(opened["successful_loads"] == 0 and opened["failed_loads"] == 0
            and opened["retained_entries"] == 0 and text["successful_loads"] == 1
            and text["failed_loads"] == 0 and text["retained_entries"] == 1,
            f"{path}: cache open/text lifecycle differs")


def _check_artifacts(directory: Path, terminal: dict[str, Any]) -> None:
    names = ("stdout.txt", "stderr.txt", "resource.txt", "report.json", "replay-cleanup.json")
    artifacts = terminal.get("artifacts")
    require(isinstance(artifacts, dict) and set(artifacts) == set(names), f"{directory}: artifact inventory differs")
    for name in names:
        item = artifacts[name]
        _exact(item, ("bytes", "path", "sha256"), f"{directory}.artifacts.{name}")
        path = _path(item["path"], f"{directory}.artifacts.{name}.path")
        require(path == directory / name, f"{directory}: artifact escaped capture directory")
        actual = _file_meta(path)
        require(actual["bytes"] == item["bytes"] and actual["sha256"] == item["sha256"], f"{directory}: artifact changed")
        if name in ("resource.txt", "report.json", "replay-cleanup.json"):
            require(item["bytes"] > 0, f"{directory}/{name}: empty required artifact")


def _validate_terminal(started: dict[str, Any], terminal: dict[str, Any], spec: dict[str, Any],
                       build: dict[str, Any], protocol_hash: str, directory: Path) -> None:
    started_fields = {"schema", "version", "status", "attempt", "run", "protocol", "build", "binary",
                      "source", "argv", "cwd", "environment", "environment_artifact", "driver_bindings",
                      "driver_sha256", "support_sha256", "started_utc", "timeout_seconds", "tmpdir"}
    terminal_fields = started_fields | {"exit_code", "timed_out", "termination", "finished_utc", "artifacts",
                                       "missing_artifacts", "cleanup", "started_artifact", "source_before",
                                       "source_after", "source_unchanged"}
    _exact(started, started_fields, f"{directory}/started.json")
    _exact(terminal, terminal_fields, f"{directory}/terminal.json")
    require(started["schema"] == CAPTURE_SCHEMA and started["status"] == "running", f"{directory}: started identity differs")
    require(terminal["schema"] == TERMINAL_SCHEMA and terminal["status"] == "pass", f"{directory}: terminal did not pass")
    for field in ("version", "attempt", "run", "protocol", "build", "binary", "source", "cwd",
                  "environment_artifact", "driver_bindings", "driver_sha256", "support_sha256",
                  "timeout_seconds", "tmpdir"):
        require(terminal[field] == started[field], f"{directory}: terminal {field} changed")
    require(started["run"] == {key: spec[key] for key in
                                ("kind", "repeat", "role", "arm", "provider", "samples", "warmups", "label")},
            f"{directory}: run specification differs")
    require(started["attempt"] == spec["attempt"] and started["protocol"] == {"path": "protocol.json", "sha256": protocol_hash},
            f"{directory}: protocol binding differs")
    require(started["build"] == {"path": build["path"], "sha256": build["receipt_sha256"], "source": build["source"]}
            and started["binary"] == build["binary"] and started["source"] == build["source"], f"{directory}: build binding differs")
    require(started["cwd"] == str(REPO) and started["driver_bindings"] ==
            {name: _json_hash(ROOT / name) for name in DRIVER_FILES}
            and started["driver_sha256"] == _json_hash(ROOT / "measure.py")
            and started["support_sha256"] == _json_hash(ROOT / "support.py"), f"{directory}: helper custody differs")
    tmp_root = _path(started["tmpdir"], f"{directory}.tmpdir")
    expected_env = {key: ENV[key] for key in ENV_KEYS}
    expected_env["TMPDIR"] = str(tmp_root)
    require(started["environment"] == expected_env and terminal["environment"] == expected_env,
            f"{directory}: child environment differs")
    require(started["argv"] == terminal["argv"] and started["argv"] ==
            _command(spec, build, directory / "report.json", directory / "resource.txt", build["git_revision"]),
            f"{directory}: child argv differs")
    require(terminal["exit_code"] == 0 and terminal["timed_out"] is False and terminal["termination"] is None
            and terminal["source_before"] == started["source"] == terminal["source_after"]
            and terminal["source_unchanged"] is True and terminal["missing_artifacts"] == [], f"{directory}: terminal custody failed")
    require(_timestamp(terminal["finished_utc"], f"{directory}.finished_utc") >
            _timestamp(started["started_utc"], f"{directory}.started_utc"), f"{directory}: chronology invalid")
    cleanup = terminal["cleanup"]
    _exact(cleanup, ("schema", "status", "root", "tmpdir", "removed", "remaining"), f"{directory}.cleanup")
    require(cleanup["schema"] == "docx-managed-private-cleanup-v1" and cleanup["status"] == "pass"
            and cleanup["tmpdir"] == str(tmp_root) and cleanup["remaining"] == [], f"{directory}: private cleanup failed")
    require(terminal["started_artifact"] == _file_meta(directory / "started.json"), f"{directory}: started hash differs")
    _check_artifacts(directory, terminal)
    require(_json(directory / "replay-cleanup.json") == cleanup, f"{directory}: cleanup receipt differs")
    require({item.name for item in directory.iterdir()} ==
            {"started.json", "terminal.json", "report.json", "resource.txt", "stdout.txt", "stderr.txt", "replay-cleanup.json"},
            f"{directory}: child file inventory differs")


def _collect(attempt: str, builds: dict[str, dict[str, Any]], protocol: dict[str, Any], protocol_hash: str,
             *, pilot: bool = False) -> list[dict[str, Any]]:
    expected = formal_inventory(pilot=pilot)
    root = ROOT / "captures" / _attempt(attempt)
    require(root.is_dir() and not root.is_symlink(), f"capture attempt is missing: {root}")
    entries = []
    for item in expected:
        spec = dict(item, attempt=attempt)
        directory = root / spec["label"]
        started_path, terminal_path = directory / "started.json", directory / "terminal.json"
        require(started_path.is_file() and terminal_path.is_file(), f"{directory}: custody receipts missing")
        started, terminal = _json(started_path), _json(terminal_path)
        _validate_terminal(started, terminal, spec, builds[spec["role"]], protocol_hash, directory)
        report_path = directory / "report.json"
        report = validate_report(report_path, role=spec["role"], arm_name=spec["arm"],
                                 samples=spec["samples"], warmups=spec["warmups"],
                                 source_revision=builds[spec["role"]]["git_revision"])
        resource = _resource_values(directory / "resource.txt")
        entries.append({"spec": spec, "directory": directory, "started": started, "terminal": terminal,
                        "report": report, "resource": resource,
                        "terminal_sha256": _json_hash(terminal_path), "report_sha256": _json_hash(report_path)})
    labels = {item["label"] for item in expected}
    require({entry["spec"]["label"] for entry in entries} == labels, "capture labels differ")
    require({path.name for path in root.iterdir()} == labels, f"{root}: extra capture files remain")
    private_attempt = TEMP / "managed" / attempt
    require(not private_attempt.exists(), f"{private_attempt}: private attempt scratch remains")
    previous: _datetime.datetime | None = None
    for entry in entries:
        start = _timestamp(entry["started"]["started_utc"], "capture started_utc")
        end = _timestamp(entry["terminal"]["finished_utc"], "capture finished_utc")
        require(end > start and (previous is None or start >= previous), "capture chronology or order differs")
        previous = end
    return entries


def _percentiles(values: list[int]) -> dict[str, Any]:
    require(values, "cannot summarize an empty vector")
    ordered = sorted(values)
    middle = len(ordered) // 2
    p50 = ordered[middle] if len(ordered) % 2 else (ordered[middle - 1] + ordered[middle]) // 2
    return {"n": len(values), "min": ordered[0], "max": ordered[-1],
            "mean": statistics.fmean(values), "p50": p50,
            "p95": ordered[math.ceil(len(values) * .95) - 1],
            "p99": ordered[math.ceil(len(values) * .99) - 1]}


def _bootstrap_median(values: list[int], *, repetitions: int = 2000) -> dict[str, Any]:
    require(values, "cannot bootstrap an empty vector")
    seed = (0x0493_5EED ^ len(values) ^ sum(values)) & 0xFFFFFFFFFFFFFFFF
    medians: list[int] = []
    for _ in range(repetitions):
        sample = []
        for _ in values:
            seed = (seed * 6364136223846793005 + 1442695040888963407) & 0xFFFFFFFFFFFFFFFF
            sample.append(values[(seed >> 32) % len(values)])
        sample.sort()
        at = len(sample) // 2
        medians.append(sample[at] if len(sample) % 2 else (sample[at - 1] + sample[at]) // 2)
    medians.sort()
    return {"method": "deterministic_percentile_bootstrap_median", "seed": (0x0493_5EED ^ len(values) ^ sum(values)) & 0xFFFFFFFFFFFFFFFF,
            "resamples": repetitions, "confidence": 0.95, "median": _percentiles(values)["p50"],
            "ci_low": medians[(repetitions * 25) // 1000], "ci_high": medians[(repetitions * 975) // 1000 - 1]}


def _vectors(entry: dict[str, Any], role: str, arm: dict[str, Any]) -> dict[str, Any]:
    rows = entry["report"]["rows"]
    physical = [row["physical"] for row in rows]
    budgets = [row["budget"] for row in rows]
    source_reads = [row["source_read"] for row in rows]
    opened = [_cache_snapshots(row, entry["directory"])[0] for row in rows]
    text = [_cache_snapshots(row, entry["directory"])[1] for row in rows]
    values: dict[str, Any] = {
        "elapsed_ns": [row["elapsed_ns"] for row in rows],
        # GNU time observes one maximum for the child process, not one value
        # per in-process sample.  Preserve that distinction in the receipt.
        "whole_child_rss_kbytes": [entry["resource"]["maximum_resident_set_size_(kbytes)"]],
        "physical_read_calls": [row["calls"] for row in physical],
        "physical_read_requested_bytes": [row["requested_bytes"] for row in physical],
        "physical_read_returned_bytes": [row["returned_bytes"] for row in physical],
        "logical_read_calls": [len(LOGICAL_RANGES)] * len(rows),
        "logical_read_requested_bytes": [BASELINE_PHYSICAL_BYTES] * len(rows),
        "logical_read_returned_bytes": [BASELINE_PHYSICAL_BYTES] * len(rows),
        "budget_input_bytes": [row["input_bytes_delta"] for row in budgets],
        "budget_memory_after_drop": [row["memory_after_drop"] for row in budgets],
        "cache_open_successful_loads": [row["successful_loads"] for row in opened],
        "cache_text_successful_loads": [row["successful_loads"] for row in text],
        "cache_open_failed_loads": [row["failed_loads"] for row in opened],
        "cache_text_failed_loads": [row["failed_loads"] for row in text],
        "cache_open_retained_entries": [row["retained_entries"] for row in opened],
        "cache_text_retained_entries": [row["retained_entries"] for row in text],
    }
    if arm["policy"] == "exact":
        values.update({"read_ahead_requests": None, "read_ahead_hits": None, "read_ahead_misses": None,
                       "read_ahead_fills": None, "read_ahead_requested_bytes": None,
                       "read_ahead_returned_bytes": None})
    else:
        values.update({"read_ahead_requests": [item["requests"] for item in source_reads],
                       "read_ahead_hits": [item["hits"] for item in source_reads],
                       "read_ahead_misses": [item["misses"] for item in source_reads],
                       "read_ahead_fills": [item["fills"] for item in source_reads],
                       "read_ahead_requested_bytes": [item["requested_bytes"] for item in source_reads],
                       "read_ahead_returned_bytes": [item["returned_bytes"] for item in source_reads]})
    if role == "allocator":
        allocations = [row["allocation"]["sample"] for row in rows]
        values.update({"allocation_calls": [item["allocation_calls"] for item in allocations],
                       "deallocation_calls": [item["deallocation_calls"] for item in allocations],
                       "reallocation_calls": [item["reallocation_calls"] for item in allocations],
                       "allocated_bytes": [item["allocated_bytes"] for item in allocations],
                       "deallocated_bytes": [item["deallocated_bytes"] for item in allocations],
                       "allocation_region_peak_live_bytes": [item["region_peak_live_bytes"] for item in allocations],
                       "allocation_peak_increment_bytes": [row["allocation"]["region_peak_increment_bytes"] for row in rows]})
    else:
        values.update({"allocation_calls": None, "deallocation_calls": None, "reallocation_calls": None,
                       "allocated_bytes": None, "deallocated_bytes": None,
                       "allocation_region_peak_live_bytes": None, "allocation_peak_increment_bytes": None})
    return values


def _summaries(values: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    stats, bootstrap = {}, {}
    for name, value in values.items():
        if isinstance(value, list):
            stats[name] = _percentiles(value)
            bootstrap[name] = _bootstrap_median(value)
        else:
            stats[name], bootstrap[name] = None, None
    return stats, bootstrap


def _relative(first: int | float, second: int | float) -> float | None:
    if first == 0:
        return 0.0 if second == 0 else None
    return (float(second) - float(first)) / abs(float(first)) * 100.0


def _comparison(role: str, repeat: int, baseline: dict[str, Any], candidate: dict[str, Any]) -> dict[str, Any]:
    metrics = {}
    names = tuple(name for name, value in baseline["percentiles"].items()
                  if name in {"elapsed_ns", "whole_child_rss_kbytes", "physical_read_calls",
                              "physical_read_requested_bytes", "physical_read_returned_bytes",
                              "logical_read_calls", "read_ahead_fills", "read_ahead_requested_bytes",
                              "read_ahead_returned_bytes", "allocation_calls", "deallocation_calls",
                              "reallocation_calls", "allocated_bytes", "deallocated_bytes",
                              "allocation_region_peak_live_bytes", "allocation_peak_increment_bytes"}
                  and value is not None and candidate["percentiles"].get(name) is not None)
    for name in names:
        left, right = baseline["percentiles"][name], candidate["percentiles"][name]
        change = {percentile: _relative(left[percentile], right[percentile]) for percentile in ("p50", "p95", "p99")}
        flags = {percentile: item is not None and item > 5.0 for percentile, item in change.items()}
        metrics[name] = {"baseline": {p: left[p] for p in ("p50", "p95", "p99")},
                         "candidate": {p: right[p] for p in ("p50", "p95", "p99")},
                         "relative_percent": change, "adverse_over_5_percent": flags,
                         "flag_over_5_percent": any(flags.values())}
    return {"role": role, "repeat": repeat, "baseline_arm": baseline["arm"],
            "candidate_arm": candidate["arm"], "metrics": metrics,
            "scope": "descriptive same-role same-delay managed policy comparison; no optimization claim"}


def analyze_data(entries: list[dict[str, Any]], builds: dict[str, dict[str, Any]], *, pilot: bool = False,
                 protocol_hash: str = "") -> dict[str, Any]:
    rows = []
    for entry in entries:
        spec = entry["spec"]
        arm = _arm(spec["arm"])
        values = _vectors(entry, spec["role"], arm)
        stats, bootstrap = _summaries(values)
        rows.append({"label": spec["label"], "kind": spec["kind"], "repeat": spec["repeat"],
                     "role": spec["role"], "arm": spec["arm"], "provider": spec["provider"],
                     "samples": spec["samples"], "warmups": spec["warmups"],
                     "terminal_sha256": entry["terminal_sha256"], "report_sha256": entry["report_sha256"],
                     "raw_vectors": values, "percentiles": stats, "bootstrap_median_ci": bootstrap,
                     "resource": entry["resource"]})
    expected_repeats = (1,) if pilot else REPEATS
    comparisons = []
    for role in ROLES:
        for repeat in expected_repeats:
            for base, candidate in ((ARMS[0], ARMS[1]), (ARMS[2], ARMS[3])):
                left = [row for row in rows if row["role"] == role and row["repeat"] == repeat and row["arm"] == base["name"]]
                right = [row for row in rows if row["role"] == role and row["repeat"] == repeat and row["arm"] == candidate["name"]]
                require(len(left) == len(right) == 1, f"comparison inventory incomplete for {role}/r{repeat}")
                comparisons.append(_comparison(role, repeat, left[0], right[0]))
    repeat_variance = []
    if not pilot:
        for role in ROLES:
            for arm in ARMS:
                matched = [row for row in rows if row["role"] == role and row["arm"] == arm["name"]]
                require(len(matched) == 2 and {row["repeat"] for row in matched} == {1, 2}, f"repeat inventory incomplete for {role}/{arm['name']}")
                first, second = sorted(matched, key=lambda item: item["repeat"])
                metrics = {}
                for name in ("elapsed_ns", "whole_child_rss_kbytes"):
                    a, b = first["percentiles"][name], second["percentiles"][name]
                    changes = {p: _relative(a[p], b[p]) for p in ("p50", "p95", "p99")}
                    metrics[name] = {"repeat1": {p: a[p] for p in ("p50", "p95", "p99")},
                                     "repeat2": {p: b[p] for p in ("p50", "p95", "p99")},
                                     "relative_percent": changes,
                                     "flag_over_5_percent": any(x is not None and abs(x) > 5 for x in changes.values())}
                repeat_variance.append({"role": role, "arm": arm["name"], "metrics": metrics,
                                        "scope": "descriptive repeat variance; no optimization regression claim"})
    return {"schema": ANALYSIS_SCHEMA, "version": VERSION, "change": CHANGE,
            "claim_authorized": False,
            "performance_claim": "descriptive managed policy evidence; no cross-format or cross-arm optimization claim",
            "scope": "same-role managed exact-versus-forward-start evidence with source, budget, cache and allocation custody",
            "pilot": pilot, "protocol_sha256": protocol_hash, "corpus": dict(CORPUS), "case": CASE,
            "roles": list(ROLES), "repeats": list(expected_repeats), "arms": [dict(arm) for arm in ARMS],
            "samples": PILOT_SAMPLES if pilot else FORMAL_SAMPLES,
            "warmups": PILOT_WARMUPS if pilot else FORMAL_WARMUPS,
            "builds": {role: _build_binding(builds[role]) for role in ROLES},
            "rows": rows, "repeat_variance": repeat_variance, "comparisons": comparisons,
            "inventory": {"formal_processes": len(rows), "measured_samples": len(rows) * (PILOT_SAMPLES if pilot else FORMAL_SAMPLES),
                          "expected_formal_processes": len(formal_inventory(pilot=pilot)),
                          "expected_measured_samples": len(formal_inventory(pilot=pilot)) * (PILOT_SAMPLES if pilot else FORMAL_SAMPLES)}}


def capture_one(attempt: str, role: str, repeat: int, arm_name: str, *, pilot: bool = False,
                build_dir: Path | None = None, timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS) -> Path:
    require(role in ROLES and arm_name in ARM_BY_NAME, "unknown capture role or arm")
    require(repeat in ((1,) if pilot else REPEATS), "repeat is not valid for this capture")
    builds = load_builds(build_dir)
    protocol, protocol_hash = _load_protocol(builds)
    match = [item for item in formal_inventory(pilot=pilot)
             if item["role"] == role and item["repeat"] == repeat and item["arm"] == arm_name]
    require(len(match) == 1, "capture specification is not frozen")
    return _capture_one(dict(match[0], attempt=_attempt(attempt)), builds[role], protocol, protocol_hash, timeout_seconds)


def capture_all(attempt: str, build_dir: Path | None = None, *, pilot: bool = False,
                timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS) -> None:
    builds = load_builds(build_dir)
    protocol, protocol_hash = _load_protocol(builds)
    for item in formal_inventory(pilot=pilot):
        _capture_one(dict(item, attempt=_attempt(attempt)), builds[item["role"]], protocol, protocol_hash, timeout_seconds)


def analyze(attempt: str, build_dir: Path | None = None, *, pilot: bool = False) -> Path:
    builds = load_builds(build_dir)
    protocol, protocol_hash = _load_protocol(builds)
    entries = _collect(attempt, builds, protocol, protocol_hash, pilot=pilot)
    summary = analyze_data(entries, builds, pilot=pilot, protocol_hash=protocol_hash)
    path = ROOT / "analysis" / f"{attempt}{'-pilot' if pilot else ''}.json"
    _write_or_match(path, summary)
    return path


def verify(attempt: str, build_dir: Path | None = None, *, pilot: bool = False) -> Path:
    builds = load_builds(build_dir)
    protocol, protocol_hash = _load_protocol(builds)
    entries = _collect(attempt, builds, protocol, protocol_hash, pilot=pilot)
    expected = analyze_data(entries, builds, pilot=pilot, protocol_hash=protocol_hash)
    summary_path = ROOT / "analysis" / f"{attempt}{'-pilot' if pilot else ''}.json"
    require(summary_path.is_file(), f"analysis summary is missing: {summary_path}")
    require(_json(summary_path) == expected, "analysis summary differs from raw evidence")
    value = {"schema": VERIFICATION_SCHEMA, "version": VERSION, "status": "pass", "attempt": attempt,
             "pilot": pilot, "protocol_sha256": protocol_hash, "summary_sha256": _json_hash(summary_path),
             "raw_receipts": [{"label": item["spec"]["label"], "terminal_sha256": item["terminal_sha256"],
                               "report_sha256": item["report_sha256"]} for item in entries], "verified_utc": now()}
    path = ROOT / "verification" / f"{attempt}{'-pilot' if pilot else ''}.json"
    _write_or_match(path, value, volatile=("verified_utc",))
    return path


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("plan")
    sub.add_parser("freeze")
    for command in ("capture", "capture-all"):
        item = sub.add_parser(command)
        item.add_argument("--attempt", required=True)
        item.add_argument("--build-dir", type=Path)
        item.add_argument("--timeout-seconds", type=int, default=DEFAULT_TIMEOUT_SECONDS)
        item.add_argument("--pilot", action="store_true")
        if command == "capture":
            item.add_argument("--role", choices=ROLES, required=True)
            item.add_argument("--repeat", type=int, choices=REPEATS, required=True)
            item.add_argument("--arm", choices=tuple(ARM_BY_NAME), required=True)
    for command in ("analyze", "verify"):
        item = sub.add_parser(command)
        item.add_argument("--attempt", required=True)
        item.add_argument("--build-dir", type=Path)
        item.add_argument("--pilot", action="store_true")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        if args.command == "plan":
            builds = load_builds() if (ROOT / "build-normal.json").is_file() and (ROOT / "build-allocator.json").is_file() else None
            print(json.dumps(protocol_value(builds), indent=2, sort_keys=True))
        elif args.command == "freeze":
            print(create_freeze())
        elif args.command == "capture":
            capture_one(args.attempt, args.role, args.repeat, args.arm, pilot=args.pilot,
                        build_dir=args.build_dir, timeout_seconds=args.timeout_seconds)
        elif args.command == "capture-all":
            capture_all(args.attempt, args.build_dir, pilot=args.pilot, timeout_seconds=args.timeout_seconds)
        elif args.command == "analyze":
            print(analyze(args.attempt, args.build_dir, pilot=args.pilot))
        elif args.command == "verify":
            print(verify(args.attempt, args.build_dir, pilot=args.pilot))
    except (ProviderMatrixError, OSError, ValueError, subprocess.SubprocessError) as error:
        print(f"measure.py: FAIL: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
