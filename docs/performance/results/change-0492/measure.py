#!/usr/bin/env python3
"""Capture, analyze, and verify the 0492 DOCX read-ahead pilot/matrix.

The Rust target owns the typed DOCX operation.  This driver owns custody of
the process, build and source identities, the frozen order, and the retained
raw evidence.  Every validator is deliberately closed over the report schema:
missing fields, extra fields, changed artifacts, and stale source bindings
fail rather than being treated as an unavailable optional observation.
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
from typing import Any, Iterable

from support import ENV, ENV_KEYS, REPO, ROOT, TEMP, meta, now, read, sha, snapshot, write


SCHEMA = "docx_provider_lifecycle_v2"
CAPTURE_SCHEMA = "docx-provider-read-ahead-capture-v1"
TERMINAL_SCHEMA = "docx-provider-read-ahead-terminal-v1"
ANALYSIS_SCHEMA = "docx-read-ahead-analysis-v1"
VERIFICATION_SCHEMA = "docx-read-ahead-verification-v1"
VERSION = 1
CHANGE = 492
CPU = 2
CPU_LOCK = "/home/zhuhe/.cache/litchi-goal-0484/cpu.lock"
CASE = "docx_provider_open_full_text_lifecycle"
FORMAL_SAMPLES = 30
FORMAL_WARMUPS = 3
PILOT_SAMPLES = 3
PILOT_WARMUPS = 1
REPEATS = (1, 2)
ROLES = ("normal", "allocator")
DEFAULT_TIMEOUT_SECONDS = 1_800
TERM_GRACE_SECONDS = 10
ATTEMPT_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]*$")
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
REVISION_RE = re.compile(r"^[0-9a-f]{40}$")

# This is the accepted 0188 media corpus.  Keep all fields here because a
# hash and byte count alone do not prove that the semantic fixture is the one
# described by the protocol.
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
# The eight synthetic members are part of the physical-range oracle.  The
# report carries these ranges so overlap is recomputed from the captured trace,
# rather than inferred from a zero-valued summary counter.
MEDIA_RANGES = (
    (4027, 2101824),
    (2101891, 4199688),
    (4199755, 6297552),
    (6297619, 8395416),
    (8395483, 10493280),
    (10493347, 12591144),
    (12591211, 14689008),
    (14689075, 16786872),
)
EXPECTED_LOGICAL_RANGES = (
    (16_793_014, 22),
    (16_791_709, 46),
    (16_791_709, 1_305),
    (0, 30),
    (49, 363),
    (412, 16),
    (412, 16),
    (428, 30),
    (469, 234),
    (703, 16),
    (703, 16),
    (2_907, 30),
    (2_965, 324),
    (3_289, 16),
    (3_289, 16),
    (1_420, 30),
    (1_467, 1_424),
    (2_891, 16),
    (2_891, 16),
)
EXPECTED_TEXT_SCOPE = (
    "Package::from_read_at_with_limits_and_cache_limits + document + "
    "extract_text + two cache diagnostic snapshots + package/document drop; "
    "returned text remains live after the clock; text destruction, hashing, "
    "oracle comparison, range counters and traces are outside"
)
EXPECTED_PROVIDER_SCOPE = (
    "typed-owner source-backed DOCX provider baseline; direct comparison "
    "with the high-level filesystem facade is out of scope"
)
EXPECTED_TIMING_SCOPE = EXPECTED_TEXT_SCOPE
EXPECTED_SETUP_SCOPE = (
    "all v1 setup plus fresh bounded read-ahead buffer and wrapper construction "
    "outside the clock; fixed window capacity is reported separately; unmanaged "
    "benchmark pilot only"
)
EXPECTED_ALLOCATION_SCOPE = (
    "operation-scoped allocator region; allocation field omitted for normal "
    "binary; fresh read-ahead window allocated during setup and reported separately"
)
EXPECTED_FILE_SCOPE = (
    "FileSource is opened before each iteration and the staged file is "
    "recently written; this is a warm-cache/recent-file provider observation, "
    "not a filesystem-cold result"
)
EXPECTED_READ_SCOPE = "logical ReadAt calls observed by the named wrapper; counters are not physical filesystem or network I/O observations"
EXPECTED_PHYSICAL_SCOPE = "physical adapter calls below read-ahead; synthetic transport observations, not disk/network I/O"
EXPECTED_LOGICAL_SCOPE = "package logical ReadAt calls above read-ahead"
GENERAL_READ_SCOPE = (
    "logical ReadAt calls observed by the named wrapper; counters are not "
    "physical filesystem or network I/O observations"
)
EXPECTED_RANGE_SCOPE = "PptxRangeSource logical adapter counters; generic ReadAt behavior, not physical I/O"
EXPECTED_MEDIA_SCOPE = (
    "compressed ZIP data ranges for word/media members only; a zero overlap "
    "proves only that the observed logical wrapper did not request or return "
    "those ranges"
)
EXPECTED_SOURCE_CONSTRUCTION = "all provider adapters are constructed outside the operation clock"
EXPECTED_ALLOCATION_STATUS = "measured"
EXPECTED_ALLOCATION_SCOPE_NAME = "operation_global_system_allocator"
REQUEST_BUCKETS = 18
WINDOW_BYTES = 4096
BASELINE_PHYSICAL_CALLS = 19
BASELINE_PHYSICAL_BYTES = 3966

# ReadLimits::default and SourceCacheLimits::new(8 MiB, 128).  These values
# are part of the measured operation's safety envelope and are consequently
# frozen rather than accepted as arbitrary positive numbers.
LIMITS = {
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
    "cache_max_bytes": 8 * 1024 * 1024,
    "cache_max_entries": 128,
}

ARMS = (
    {
        "name": "baseline-65536-0us",
        "provider": "range",
        "max_range_bytes": 65_536,
        "delay_us": 0,
        "transfer_bytes_per_second": None,
        "transfer_delay_policy": "separate-sleeps",
        "read_ahead_window_bytes": None,
    },
    {
        "name": "candidate-4096-0us",
        "provider": "range",
        "max_range_bytes": 65_536,
        "delay_us": 0,
        "transfer_bytes_per_second": None,
        "transfer_delay_policy": "separate-sleeps",
        "read_ahead_window_bytes": 4096,
    },
    {
        "name": "baseline-65536-1000us-104857600bps-minimum-service",
        "provider": "range",
        "max_range_bytes": 65_536,
        "delay_us": 1_000,
        "transfer_bytes_per_second": 104_857_600,
        "transfer_delay_policy": "minimum-service",
        "read_ahead_window_bytes": None,
    },
    {
        "name": "candidate-4096-1000us-104857600bps-minimum-service",
        "provider": "range",
        "max_range_bytes": 65_536,
        "delay_us": 1_000,
        "transfer_bytes_per_second": 104_857_600,
        "transfer_delay_policy": "minimum-service",
        "read_ahead_window_bytes": 4096,
    },
)
ARM_BY_NAME = {item["name"]: item for item in ARMS}
DRIVER_FILES = ("measure.py", "test_measure.py", "support.py", "gate.py", "retain_build.py")


class ProviderMatrixError(RuntimeError):
    """A custody, schema, identity, or recomputation invariant failed."""


def fail(message: str) -> None:
    raise ProviderMatrixError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _attempt(value: str) -> str:
    require(ATTEMPT_RE.fullmatch(value) is not None, "attempt must be a path-safe token")
    return value


def _read_json(path: Path) -> Any:
    try:
        return read(path)
    except (OSError, ValueError, json.JSONDecodeError) as error:
        fail(f"{path}: invalid JSON: {error}")


def _write_json(path: Path, value: Any) -> None:
    try:
        write(path, value)
    except FileExistsError:
        fail(f"refusing to replace immutable artifact: {path}")


def _write_or_match_json(path: Path, value: dict[str, Any], *, volatile: tuple[str, ...] = ()) -> None:
    """Create a receipt once, or prove a later invocation is identical.

    Verification is intentionally rerunnable after cleanup and after a commit.
    Only explicitly listed fields may vary; the existing artifact remains
    byte immutable and its timestamp is checked when it is reused.
    """

    if not path.exists():
        _write_json(path, value)
        return
    existing = _read_json(path)
    require(isinstance(existing, dict), f"{path}: existing receipt is not an object")
    expected = dict(value)
    actual = dict(existing)
    for field in volatile:
        require(field in actual, f"{path}: volatile field {field} is missing")
        _timestamp(actual[field], f"{path}.{field}")
        expected.pop(field, None)
        actual.pop(field, None)
    require(actual == expected, f"{path}: immutable receipt differs on rerun")


def _finite(value: Any, path: str = "json") -> None:
    if isinstance(value, float):
        require(math.isfinite(value), f"{path}: non-finite number")
    elif isinstance(value, dict):
        for key, child in value.items():
            require(isinstance(key, str), f"{path}: non-string JSON key")
            _finite(child, f"{path}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            _finite(child, f"{path}[{index}]")


def _exact_keys(value: Any, keys: Iterable[str], path: str) -> None:
    require(isinstance(value, dict), f"{path}: expected object")
    expected = set(keys)
    actual = set(value)
    require(actual == expected, f"{path}: fields differ; expected {sorted(expected)}, got {sorted(actual)}")


def _hash(value: Any, path: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
            f"{path}: expected lowercase SHA-256")
    return value


def _uint(value: Any, path: str, *, positive: bool = False) -> int:
    require(type(value) is int and value >= (1 if positive else 0),
            f"{path}: expected {'positive' if positive else 'non-negative'} integer")
    return value


def _text(value: Any, path: str, *, nonempty: bool = True) -> str:
    require(isinstance(value, str) and (bool(value) if nonempty else True),
            f"{path}: expected {'non-empty ' if nonempty else ''}string")
    return value


def _file_meta(path: Path, *, executable: bool = False) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing regular artifact: {path}")
    value = meta(path)
    require(value["bytes"] >= 0 and SHA256_RE.fullmatch(value["sha256"]) is not None,
            f"{path}: malformed artifact metadata")
    if executable:
        require(os.access(path, os.X_OK), f"{path}: executable bit is absent")
    return {"path": str(path), **value}


def _resolve_rooted(value: str, *, base: Path = ROOT) -> Path:
    path = Path(value)
    return path if path.is_absolute() else base / path


def _json_hash(path: Path) -> str:
    return _hash(sha(path), str(path))


def _build_artifact(value: Any, path: Path, field: str, *, retained: bool = True) -> dict[str, Any]:
    require(isinstance(value, dict), f"{path}.{field}: artifact binding is missing")
    _exact_keys(value, ("bytes", "executable", "path", "sha256"), f"{path}.{field}")
    artifact_path = _resolve_rooted(_text(value["path"], f"{path}.{field}.path"))
    require(value["executable"] is True, f"{path}.{field}.executable is false")
    _uint(value["bytes"], f"{path}.{field}.bytes", positive=True)
    _hash(value["sha256"], f"{path}.{field}.sha256")
    if retained:
        actual = _file_meta(artifact_path, executable=True)
        require(value["bytes"] == actual["bytes"], f"{path}.{field}.bytes changed")
        require(value["sha256"] == actual["sha256"], f"{path}.{field}.sha256 changed")
    elif artifact_path.is_file():
        actual = _file_meta(artifact_path, executable=True)
        require(value["bytes"] == actual["bytes"], f"{path}.{field}.bytes changed")
        require(value["sha256"] == actual["sha256"], f"{path}.{field}.sha256 changed")
    return {"path": str(artifact_path), "bytes": value["bytes"], "sha256": value["sha256"], "executable": True}


def _source_binding(value: Any, path: Path, field: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{path}.{field}: source binding is missing")
    _exact_keys(value, ("files", "path", "sha256"), f"{path}.{field}")
    manifest_path = _resolve_rooted(_text(value["path"], f"{path}.{field}.path"))
    actual = _file_meta(manifest_path)
    require(value["sha256"] == actual["sha256"], f"{path}.{field}.sha256 changed")
    _uint(value["files"], f"{path}.{field}.files", positive=True)
    manifest = _read_json(manifest_path)
    require(isinstance(manifest, dict), f"{path}.{field}: source manifest is not an object")
    require(len(manifest) == value["files"], f"{path}.{field}: manifest file count differs")
    for name, digest in manifest.items():
        require(isinstance(name, str) and name and not Path(name).is_absolute()
                and ".." not in Path(name).parts,
                f"{path}.{field}: manifest path is unsafe")
        _hash(digest, f"{path}.{field}.{name}")
    canonical = (json.dumps(manifest, indent=2, sort_keys=True) + "\n").encode()
    require(hashlib.sha256(canonical).hexdigest() == value["sha256"],
            f"{path}.{field}: manifest content does not match its binding")
    return {"files": value["files"], "path": str(manifest_path), "sha256": value["sha256"]}


def _gate_binding(value: Any, path: Path) -> dict[str, Any]:
    require(isinstance(value, dict), f"{path}.gate: mandatory gate binding is missing")
    _exact_keys(value, ("path", "sha256"), f"{path}.gate")
    gate_path = _resolve_rooted(_text(value["path"], f"{path}.gate.path"))
    require(gate_path.is_file() and not gate_path.is_symlink(), f"{path}.gate.path is missing")
    _hash(value["sha256"], f"{path}.gate.sha256")
    require(_json_hash(gate_path) == value["sha256"], f"{path}.gate changed")
    receipt = _read_json(gate_path)
    require(isinstance(receipt, dict), f"{gate_path}: gate receipt is not an object")
    required = (
        "argv", "artifacts", "attempt", "common_sha256", "cwd", "driver_sha256",
        "environment", "exit_code", "finished_utc", "label", "schema",
        "source_after", "source_before", "source_unchanged", "started_utc",
    )
    _exact_keys(receipt, required, str(gate_path))
    require(receipt["schema"] == "docx-stream-append-gate-v1", f"{gate_path}: gate schema differs")
    require(receipt["exit_code"] == 0 and receipt["source_unchanged"] is True,
            f"{gate_path}: gate did not pass with unchanged source")
    require(receipt["cwd"] == str(REPO), f"{gate_path}: gate cwd differs")
    require(isinstance(receipt["argv"], list) and receipt["argv"]
            and all(isinstance(item, str) and item for item in receipt["argv"]),
            f"{gate_path}: gate argv is missing")
    require(receipt["environment"] == {key: ENV[key] for key in ENV_KEYS},
            f"{gate_path}: gate environment differs")
    require(receipt["driver_sha256"] == _json_hash(ROOT / "gate.py"),
            f"{gate_path}: gate driver changed")
    require(receipt["common_sha256"] == _json_hash(ROOT / "support.py"),
            f"{gate_path}: gate support binding differs")
    started_at = _timestamp(receipt["started_utc"], f"{gate_path}.started_utc")
    finished_at = _timestamp(receipt["finished_utc"], f"{gate_path}.finished_utc")
    require(finished_at > started_at, f"{gate_path}: gate chronology is invalid")
    source_before = _source_binding(receipt["source_before"], gate_path, "source_before")
    source_after = _source_binding(receipt["source_after"], gate_path, "source_after")
    require(source_before == source_after, f"{gate_path}: gate source bindings differ")
    _text(receipt["label"], f"{gate_path}.label")
    if receipt["attempt"] is not None:
        _text(receipt["attempt"], f"{gate_path}.attempt")
    artifact_names = (f"{gate_path.stem}.stderr", f"{gate_path.stem}.stdout")
    _exact_keys(receipt["artifacts"], artifact_names, f"{gate_path}.artifacts")
    for name in artifact_names:
        artifact = receipt["artifacts"][name]
        _exact_keys(artifact, ("bytes", "sha256"), f"{gate_path}.artifacts.{name}")
        _uint(artifact["bytes"], f"{gate_path}.artifacts.{name}.bytes")
        _hash(artifact["sha256"], f"{gate_path}.artifacts.{name}.sha256")
        artifact_path = gate_path.with_suffix("." + name.rsplit(".", 1)[1])
        actual = _file_meta(artifact_path)
        require(actual["bytes"] == artifact["bytes"] and actual["sha256"] == artifact["sha256"],
                f"{gate_path}: gate {name} artifact changed")
    return {
        "path": str(gate_path),
        "sha256": value["sha256"],
        "source": source_after,
        "argv": list(receipt["argv"]),
    }


def _expected_build_command(role: str) -> list[str]:
    require(role in ROLES, f"unknown build role {role}")
    name = "litchi-perf-baseline" + ("-alloc" if role == "allocator" else "")
    command = [
        "cargo", "build", "--release", "--locked",
        "--manifest-path", "tools/perf-baseline/Cargo.toml",
    ]
    if role == "allocator":
        command.extend(["--features", "allocator-metrics"])
    command.extend(["--bin", name])
    return command


def _binary_from_build(value: Any, role: str, path: Path) -> dict[str, Any]:
    require(isinstance(value, dict), f"{path}: build receipt is not an object")
    required = (
        "attempt", "binary", "command", "copied_utc", "environment", "gate",
        "original_binary", "role", "schema", "source_after", "source_before",
        "source_unchanged", "version", "git_revision", "retainer_sha256",
    )
    _exact_keys(value, required, str(path))
    require(value["role"] == role, f"{path}: build role differs from {role}")
    require(value["schema"] == "docx-provider-lifecycle-build-v1",
            f"{path}: build schema differs")
    require(value["version"] == 1, f"{path}: build version differs")
    require(value["retainer_sha256"] == _json_hash(ROOT / "retain_build.py"),
            f"{path}: build retainer helper changed")
    _text(value["attempt"], f"{path}.attempt")
    require(isinstance(value["command"], list) and value["command"] and
            all(isinstance(item, str) and item for item in value["command"]),
            f"{path}.command: non-empty argv is required")
    expected_command = _expected_build_command(role)
    require(value["command"] == expected_command,
            f"{path}.command: role-specific Cargo argv differs")
    git_revision = _text(value["git_revision"], f"{path}.git_revision")
    require(REVISION_RE.fullmatch(git_revision) is not None,
            f"{path}.git_revision: expected lowercase 40-character hash")
    _text(value["copied_utc"], f"{path}.copied_utc")
    require(isinstance(value["environment"], dict), f"{path}.environment: missing")
    for key in ENV_KEYS:
        require(value["environment"].get(key) == ENV[key],
                f"{path}.environment.{key}: build environment differs")
    binary = _build_artifact(value["binary"], path, "binary")
    # The copied executable is retained for historical verification.  The
    # original Cargo target may be removed during the bounded cleanup; its
    # immutable bytes/hash/path remain a mandatory receipt binding.
    original = _build_artifact(value["original_binary"], path, "original_binary", retained=False)
    expected_name = expected_command[-1]
    require(Path(binary["path"]).name == expected_name,
            f"{path}.binary.path: executable name does not match role")
    require(Path(original["path"]).name == expected_name,
            f"{path}.original_binary.path: executable name does not match role")
    require(binary["bytes"] == original["bytes"] and binary["sha256"] == original["sha256"],
            f"{path}: copied and original executable identities differ")
    source_before = _source_binding(value["source_before"], path, "source_before")
    source_after = _source_binding(value["source_after"], path, "source_after")
    require(value["source_unchanged"] is True and source_before == source_after,
            f"{path}: source custody failed")
    gate = _gate_binding(value["gate"], path)
    require(gate["source"] == source_after, f"{path}: gate source does not match build source")
    require(gate["argv"] == expected_command,
            f"{path}: gate argv does not match role-specific build command")
    return {
        "path": str(path),
        "receipt_sha256": _json_hash(path),
        "role": role,
        "binary": binary,
        "original_binary": original,
        "source": source_after,
        "gate": gate,
        "command": list(value["command"]),
        "environment": dict(value["environment"]),
        "git_revision": git_revision,
    }


def load_builds(build_dir: Path | None = None) -> dict[str, dict[str, Any]]:
    """Load immutable normal/allocator receipts without checking live sources."""

    directory = ROOT if build_dir is None else Path(build_dir)
    result = {
        role: _binary_from_build(
            _read_json(directory / f"build-{role}.json"),
            role,
            directory / f"build-{role}.json",
        )
        for role in ROLES
    }
    require(result["normal"]["source"] == result["allocator"]["source"],
            "normal and allocator build source identities differ")
    require(result["normal"]["git_revision"] == result["allocator"]["git_revision"],
            "normal and allocator build git revisions differ")
    return result


def _environment_binding() -> dict[str, Any]:
    """Bind the complete captured host/tool environment into the protocol."""

    path = ROOT / "environment.json"
    require(path.is_file() and not path.is_symlink(), f"environment artifact is missing: {path}")
    return {"path": "environment.json", "sha256": _json_hash(path)}


def _protocol_source(builds: dict[str, dict[str, Any]] | None) -> dict[str, Any] | None:
    if not builds:
        return None
    source = builds["normal"]["source"]
    return {"files": source["files"], "path": str(Path(source["path"]).relative_to(ROOT) if Path(source["path"]).is_relative_to(ROOT) else source["path"]), "sha256": source["sha256"]}


def formal_inventory(*, pilot: bool = False) -> list[dict[str, Any]]:
    samples = PILOT_SAMPLES if pilot else FORMAL_SAMPLES
    warmups = PILOT_WARMUPS if pilot else FORMAL_WARMUPS
    repeats = (1,) if pilot else REPEATS
    result: list[dict[str, Any]] = []
    for repeat in repeats:
        roles = ROLES if repeat == 1 else tuple(reversed(ROLES))
        arms = ARMS if repeat == 1 else tuple(reversed(ARMS))
        for role in roles:
            for arm in arms:
                result.append({
                    "kind": "pilot" if pilot else "formal",
                    "repeat": repeat,
                    "role": role,
                    "arm": arm["name"],
                    "provider": arm["provider"],
                    "samples": samples,
                    "warmups": warmups,
                    "label": f"{'pilot-' if pilot else ''}r{repeat}-{role}-{arm['name']}",
                })
    return result


def protocol_value(builds: dict[str, dict[str, Any]] | None = None) -> dict[str, Any]:
    """Return the frozen protocol; build/machine bindings are added when present."""

    runs = formal_inventory()
    return {
        "schema": "docx-read-ahead-protocol-v1",
        "version": VERSION,
        "change": CHANGE,
        "claim_authorized": False,
        "performance_claim": "descriptive provider evidence; no cross-arm optimization claim",
        "case": CASE,
        "cpu": CPU,
        "cpu_lock": CPU_LOCK,
        "roles": list(ROLES),
        "repeats": list(REPEATS),
        "arms": [dict(arm) for arm in ARMS],
        "samples": FORMAL_SAMPLES,
        "warmups": FORMAL_WARMUPS,
        "expected_formal_processes": 16,
        "expected_measured_samples": 480,
        "formal_runs": runs,
        "corpus": dict(CORPUS),
        "media_ranges": [{"start": start, "end": end} for start, end in MEDIA_RANGES],
        "read_ahead_window_bytes": WINDOW_BYTES,
        "baseline_physical_observation": {
            "logical_calls": BASELINE_PHYSICAL_CALLS,
            "requested_bytes": BASELINE_PHYSICAL_BYTES,
        },
        "expected_logical_ranges": [
            {"offset": offset, "requested": length, "returned": length}
            for offset, length in EXPECTED_LOGICAL_RANGES
        ],
        "actual_text_oracle": {
            "sha256": CORPUS["expected_text_sha256"],
            "bytes": CORPUS["expected_text_bytes"],
            "timing_scope": EXPECTED_TEXT_SCOPE,
        },
        "limits": dict(LIMITS),
        "scopes": {
            "provider": EXPECTED_PROVIDER_SCOPE,
            "timing": EXPECTED_TIMING_SCOPE,
            "setup": EXPECTED_SETUP_SCOPE,
            "allocation": EXPECTED_ALLOCATION_SCOPE,
            "file": EXPECTED_FILE_SCOPE,
            "read": EXPECTED_READ_SCOPE,
            "range": EXPECTED_RANGE_SCOPE,
            "media": EXPECTED_MEDIA_SCOPE,
        },
        "source": _protocol_source(builds),
        "environment_artifact": _environment_binding(),
        "environment": {key: ENV[key] for key in ENV_KEYS},
        "driver": {
            name: _json_hash(ROOT / name) for name in DRIVER_FILES if (ROOT / name).is_file()
        },
        "limitations": [
            "provider construction is outside the operation clock",
            "the read-ahead adapter is an unmanaged bounded pilot; managed OPC budgets remain a production prerequisite",
            "logical ReadAt counters are not physical filesystem or network I/O",
            "two repeats provide descriptive variance only and no strong confidence claim",
            "media overlap is measured and reported; compressed-byte overlap does not prove media decompression",
            "borrowed non-static input, publication atomicity, parallel scaling, and native producer coverage remain outside this matrix",
        ],
    }


def _protocol_path() -> Path:
    return ROOT / "protocol.json"


def _load_protocol() -> tuple[dict[str, Any], str]:
    path = _protocol_path()
    require(path.is_file(), f"frozen protocol is missing: {path}")
    value = _read_json(path)
    _finite(value)
    require(value.get("schema") == "docx-read-ahead-protocol-v1",
            f"{path}: protocol schema differs")
    require(value.get("version") == VERSION and value.get("change") == CHANGE,
            f"{path}: protocol version/change differs")
    require(value.get("formal_runs") == formal_inventory(), f"{path}: frozen run order differs")
    require(value.get("arms") == [dict(arm) for arm in ARMS], f"{path}: frozen arms differ")
    require(value.get("corpus") == CORPUS, f"{path}: frozen corpus differs")
    require(value.get("media_ranges") == [{"start": start, "end": end} for start, end in MEDIA_RANGES],
            f"{path}: frozen media ranges differ")
    require(value.get("read_ahead_window_bytes") == WINDOW_BYTES,
            f"{path}: frozen read-ahead window differs")
    require(value.get("baseline_physical_observation") == {
        "logical_calls": BASELINE_PHYSICAL_CALLS,
        "requested_bytes": BASELINE_PHYSICAL_BYTES,
    }, f"{path}: frozen baseline observation differs")
    require(value.get("expected_logical_ranges") == [
        {"offset": offset, "requested": length, "returned": length}
        for offset, length in EXPECTED_LOGICAL_RANGES
    ], f"{path}: frozen logical range trace differs")
    require(value.get("limits") == LIMITS, f"{path}: frozen limits differ")
    require(value.get("environment") == {key: ENV[key] for key in ENV_KEYS},
            f"{path}: frozen environment differs")
    driver = value.get("driver")
    require(isinstance(driver, dict), f"{path}: helper bindings are missing")
    for name in DRIVER_FILES:
        require(driver.get(name) == _json_hash(ROOT / name), f"{path}: helper {name} changed")
    environment_artifact = value.get("environment_artifact")
    require(isinstance(environment_artifact, dict), f"{path}: environment binding is missing")
    _exact_keys(environment_artifact, ("path", "sha256"), f"{path}.environment_artifact")
    environment_path = _resolve_rooted(
        _text(environment_artifact.get("path"), f"{path}.environment_artifact.path")
    )
    require(_json_hash(environment_path) == environment_artifact.get("sha256"),
            f"{path}: environment artifact changed")
    source = value.get("source")
    require(isinstance(source, dict), f"{path}: source binding is missing")
    _hash(source.get("sha256"), f"{path}.source.sha256")
    source_path = _resolve_rooted(_text(source.get("path"), f"{path}.source.path"))
    require(_json_hash(source_path) == source["sha256"], f"{path}: source artifact changed")
    _uint(source.get("files"), f"{path}.source.files", positive=True)
    return value, _json_hash(path)


def _run_root(attempt: str, label: str) -> Path:
    return TEMP / "provider" / attempt / label


def _capture_root(attempt: str, label: str) -> Path:
    return ROOT / "captures" / attempt / label


def _source_revision() -> str:
    try:
        value = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=REPO, text=True).strip()
    except (OSError, subprocess.SubprocessError) as error:
        fail(f"cannot determine source revision: {error}")
    require(REVISION_RE.fullmatch(value) is not None, "git revision is not a lowercase 40-character hash")
    return value


def _normalized_snapshot() -> dict[str, Any]:
    """Normalize support.py's relative manifest path for receipt comparisons."""

    value = snapshot()
    _exact_keys(value, ("files", "path", "sha256"), "source snapshot")
    _uint(value["files"], "source snapshot.files", positive=True)
    _hash(value["sha256"], "source snapshot.sha256")
    path = _resolve_rooted(_text(value["path"], "source snapshot.path"))
    actual = _file_meta(path)
    require(actual["sha256"] == value["sha256"], "source snapshot digest changed")
    return {"files": value["files"], "path": str(path), "sha256": value["sha256"]}


def _arm(name: str) -> dict[str, Any]:
    try:
        return dict(ARM_BY_NAME[name])
    except KeyError:
        fail(f"unknown provider arm {name}")


def _command(spec: dict[str, Any], binary: dict[str, Any], report: Path, resource: Path,
             source_revision: str) -> list[str]:
    arm = _arm(spec["arm"])
    command = [
        "/usr/bin/time", "-v", "-o", str(resource),
        "/usr/bin/taskset", "-c", str(CPU),
        str(binary["path"]), "docx-provider-lifecycle",
        "--provider", arm["provider"],
        "--trace-ranges",
    ]
    if arm["provider"] == "range":
        command.extend(["--max-range", str(arm["max_range_bytes"]), "--delay-us", str(arm["delay_us"])])
        if arm["transfer_bytes_per_second"] is not None:
            command.extend([
                "--transfer-bytes-per-second", str(arm["transfer_bytes_per_second"]),
                "--transfer-delay-policy", str(arm["transfer_delay_policy"]),
            ])
    if arm.get("read_ahead_window_bytes") is not None:
        command.extend(["--read-ahead", str(arm["read_ahead_window_bytes"])])
    command.extend([
        "--samples", str(spec["samples"]),
        "--warmup", str(spec["warmups"]),
        "--source-revision", source_revision,
        "--output", str(report),
    ])
    return command


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


def _timestamp(value: Any, path: str) -> _datetime.datetime:
    text = _text(value, path)
    try:
        parsed = _datetime.datetime.fromisoformat(text)
    except ValueError as error:
        fail(f"{path}: invalid timestamp: {error}")
    require(parsed.tzinfo is not None, f"{path}: timestamp lacks timezone")
    return parsed


def _resource_values(path: Path) -> dict[str, Any]:
    require(path.is_file(), f"missing process resource receipt: {path}")
    raw = path.read_bytes()
    require(raw, f"empty process resource receipt: {path}")
    text = raw.decode("utf-8", errors="replace")
    result: dict[str, Any] = {"raw_sha256": hashlib.sha256(raw).hexdigest(), "raw_bytes": len(raw)}
    for line in text.splitlines():
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        key = key.strip().lower().replace(" ", "_")
        value = value.strip()
        if value.isdigit():
            result[key] = int(value)
        else:
            try:
                result[key] = float(value.split()[0])
            except (ValueError, IndexError):
                result[key] = value
    require("maximum_resident_set_size_(kbytes)" in result,
            f"{path}: /usr/bin/time did not retain maximum RSS")
    _uint(result["maximum_resident_set_size_(kbytes)"], f"{path}: maximum RSS", positive=False)
    return result


def _cleanup_private(run_root: Path, tmp_root: Path) -> dict[str, Any]:
    """Remove only the explicitly owned empty TMPDIR and its empty parents."""

    removed: list[str] = []
    remaining: list[str] = []
    require(tmp_root.parent == run_root and run_root.parent.parent == TEMP / "provider", "scratch root escaped owner")
    require(not run_root.is_symlink() and not run_root.parent.is_symlink(), "scratch parent is a symlink")
    if tmp_root.exists():
        require(tmp_root.is_dir() and not tmp_root.is_symlink(), f"private TMPDIR is not a directory: {tmp_root}")
        children = list(tmp_root.iterdir())
        if children:
            remaining.extend(str(child) for child in children)
        else:
            tmp_root.rmdir()
            removed.append(str(tmp_root))
    else:
        remaining.append(str(tmp_root))
    if not remaining and run_root.exists():
        require(run_root.is_dir() and not run_root.is_symlink(), f"private run root is not a directory: {run_root}")
        if not list(run_root.iterdir()):
            run_root.rmdir()
            removed.append(str(run_root))
        else:
            remaining.extend(str(child) for child in run_root.iterdir())
    return {
        "schema": "docx-provider-private-cleanup-v1",
        "status": "pass" if not remaining else "failed",
        "root": str(run_root),
        "tmpdir": str(tmp_root),
        "removed": removed,
        "remaining": remaining,
    }


def _report_keys() -> tuple[str, ...]:
    return (
        "allocation_scope", "corpus", "limits", "provider", "provider_scope", "schema",
        "media_ranges", "requested_source_revision", "rows", "samples", "setup_scope",
        "source_bytes", "source_sha256", "timing_scope", "warmup",
    )


def _corpus(value: Any, path: str) -> dict[str, Any]:
    _exact_keys(value, CORPUS.keys(), path)
    require(value == CORPUS, f"{path}: pinned corpus manifest differs")
    return dict(value)


def _check_limits(value: Any, path: str) -> dict[str, Any]:
    _exact_keys(value, LIMITS.keys(), path)
    require(value == LIMITS, f"{path}: configured limits differ")
    return dict(value)


def _check_provider(value: Any, arm: dict[str, Any], path: str) -> dict[str, Any]:
    fields = [
        "delay_us", "file_scope", "max_range_bytes", "provider",
        "read_counter_scope", "source_construction", "transfer_bytes_per_second",
        "transfer_delay_policy",
    ]
    if arm.get("read_ahead_window_bytes") is not None:
        fields.append("read_ahead_window_bytes")
    _exact_keys(value, fields, path)
    for field in ("max_range_bytes", "delay_us", "transfer_bytes_per_second"):
        item = value[field]
        if item is not None:
            _uint(item, f"{path}.{field}", positive=field != "delay_us")
    require(value["provider"] == arm["provider"], f"{path}.provider differs")
    for field in ("max_range_bytes", "delay_us", "transfer_bytes_per_second", "transfer_delay_policy"):
        require(value[field] == arm[field], f"{path}.{field} differs from frozen arm")
    if arm.get("read_ahead_window_bytes") is not None:
        _uint(value["read_ahead_window_bytes"], f"{path}.read_ahead_window_bytes", positive=True)
        require(value["read_ahead_window_bytes"] == arm["read_ahead_window_bytes"],
                f"{path}.read_ahead_window_bytes differs from frozen arm")
    require(value["source_construction"] == EXPECTED_SOURCE_CONSTRUCTION,
            f"{path}.source_construction differs")
    require(value["file_scope"] == EXPECTED_FILE_SCOPE, f"{path}.file_scope differs")
    require(value["read_counter_scope"] == GENERAL_READ_SCOPE, f"{path}.read_counter_scope differs")
    return dict(value)


def _check_source_version(value: Any, path: str) -> dict[str, int]:
    _exact_keys(value, ("id", "revision"), path)
    _uint(value["id"], f"{path}.id", positive=True)
    _uint(value["revision"], f"{path}.revision")
    return {"id": value["id"], "revision": value["revision"]}


def _check_counter(value: Any, path: str, *, available: bool, range_counter: bool = False,
                   scope: str | None = None) -> dict[str, Any]:
    fields = (
        "availability", "logical_calls", "max_request_bytes", "min_request_bytes",
        "request_size_counts", "requested_bytes", "returned_bytes", "scope",
        "short_reads", "delayed_calls", "transfer_paced_calls", "transfer_delay_ns",
    )
    _exact_keys(value, fields, path)
    require(value["availability"] == ("available" if available else "unavailable"),
            f"{path}.availability differs")
    expected_scope = EXPECTED_RANGE_SCOPE if range_counter else (
        scope if scope is not None else (EXPECTED_READ_SCOPE if available else GENERAL_READ_SCOPE)
    )
    require(value["scope"] == expected_scope,
            f"{path}.scope differs")
    numeric = ("logical_calls", "requested_bytes", "returned_bytes", "short_reads",
               "delayed_calls", "transfer_paced_calls", "transfer_delay_ns")
    extrema = ("min_request_bytes", "max_request_bytes")
    if available:
        for field in numeric:
            _uint(value[field], f"{path}.{field}")
        for field in extrema:
            if value[field] is not None:
                _uint(value[field], f"{path}.{field}", positive=True)
        for field in ("short_reads", "delayed_calls", "transfer_paced_calls"):
            require(value[field] <= value["logical_calls"], f"{path}: {field} exceeds call count")
        if value["logical_calls"]:
            require(value["min_request_bytes"] is not None and value["max_request_bytes"] is not None
                    and value["min_request_bytes"] <= value["max_request_bytes"], f"{path}: request extrema invalid")
        counts = value["request_size_counts"]
        require(isinstance(counts, list) and len(counts) == REQUEST_BUCKETS,
                f"{path}.request_size_counts differs")
        for index, item in enumerate(counts):
            _uint(item, f"{path}.request_size_counts[{index}]")
        require(sum(counts) == value["logical_calls"], f"{path}: histogram does not reconcile calls")
        require(value["returned_bytes"] <= value["requested_bytes"],
                f"{path}: returned bytes exceed requested bytes")
    else:
        for field in numeric + extrema + ("request_size_counts",):
            require(value[field] is None, f"{path}.{field}: unavailable counter has a value")
    return dict(value)


def _check_media(value: Any, path: str, *, available: bool, allow_overlap: bool = False) -> dict[str, Any]:
    fields = (
        "availability", "media_range_count", "observed_call_count",
        "reason", "returned_overlap_bytes", "requested_overlap_bytes", "scope", "status",
    )
    _exact_keys(value, fields, path)
    require(value["scope"] == EXPECTED_MEDIA_SCOPE, f"{path}.scope differs")
    _uint(value["media_range_count"], f"{path}.media_range_count", positive=True)
    require(value["media_range_count"] == len(MEDIA_RANGES),
            f"{path}.media_range_count differs from the frozen corpus media map")
    require(value["availability"] == ("available" if available else "unavailable"),
            f"{path}.availability differs")
    if available:
        require(value["status"] in ("proved_no_media_overlap", "media_overlap_observed"),
                f"{path}.status differs")
        require(value["reason"] is None, f"{path}.reason unexpectedly present")
        for field in ("observed_call_count", "requested_overlap_bytes", "returned_overlap_bytes"):
            _uint(value[field], f"{path}.{field}")
        if not allow_overlap:
            require(value["requested_overlap_bytes"] == 0 and value["returned_overlap_bytes"] == 0,
                    f"{path}: media overlap was observed")
        if value["requested_overlap_bytes"] == 0 and value["returned_overlap_bytes"] == 0:
            require(value["status"] == "proved_no_media_overlap", f"{path}: zero overlap status is wrong")
        else:
            require(value["status"] == "media_overlap_observed", f"{path}: overlap status is wrong")
    else:
        require(value["status"] == "unavailable" and value["reason"] == "provider has no offset instrumentation",
                f"{path}: unavailable media proof reason differs")
        for field in ("observed_call_count", "requested_overlap_bytes", "returned_overlap_bytes"):
            require(value[field] is None, f"{path}.{field}: unavailable proof has a value")
    return dict(value)


def _check_allocation(value: Any, path: str) -> dict[str, Any]:
    fields = (
        "allocated_bytes", "allocation_calls", "deallocated_bytes", "deallocation_calls",
        "failed_allocation_calls", "live_bytes_after", "live_bytes_before",
        "peak_live_bytes_after", "peak_live_bytes_before", "reallocation_calls",
        "region_peak_live_bytes", "scope", "status",
    )
    _exact_keys(value, fields, path)
    require(value["status"] == EXPECTED_ALLOCATION_STATUS and value["scope"] == EXPECTED_ALLOCATION_SCOPE_NAME,
            f"{path}: allocator identity differs")
    for field in fields:
        if field not in ("scope", "status"):
            _uint(value[field], f"{path}.{field}")
    require(value["region_peak_live_bytes"] >= value["live_bytes_before"],
            f"{path}: region peak precedes live bytes")
    require(value["region_peak_live_bytes"] >= value["live_bytes_after"],
            f"{path}: region peak below after live bytes")
    return dict(value)


def _bucket(requested: int) -> int:
    """Mirror the fixed Rust request-size histogram without floating point."""

    if requested <= 1:
        return 0
    return min((requested - 1).bit_length(), REQUEST_BUCKETS - 1)


def _check_ranges(value: Any, path: str, counter: dict[str, Any]) -> list[dict[str, int]]:
    require(isinstance(value, list), f"{path}: range trace is not a list")
    require(len(value) == counter["logical_calls"], f"{path}: trace/counter call count differs")
    total_requested = 0
    total_returned = 0
    histogram = [0] * REQUEST_BUCKETS
    result: list[dict[str, int]] = []
    for index, item in enumerate(value):
        _exact_keys(item, ("offset", "requested", "returned"), f"{path}[{index}]")
        offset = _uint(item["offset"], f"{path}[{index}].offset")
        requested = _uint(item["requested"], f"{path}[{index}].requested")
        returned = _uint(item["returned"], f"{path}[{index}].returned")
        require(requested > 0, f"{path}[{index}]: empty range was recorded")
        require(returned <= requested, f"{path}[{index}]: returned bytes exceed request")
        require(offset + requested <= (1 << 64) - 1,
                f"{path}[{index}]: range end overflows u64")
        total_requested += requested
        total_returned += returned
        histogram[_bucket(requested)] += 1
        result.append({"offset": offset, "requested": requested, "returned": returned})
    require(total_requested == counter["requested_bytes"], f"{path}: requested sum differs")
    require(total_returned == counter["returned_bytes"], f"{path}: returned sum differs")
    require(histogram == counter["request_size_counts"], f"{path}: request histogram differs")
    require(sum(item["returned"] < item["requested"] for item in result) == counter["short_reads"],
            f"{path}: short-read count differs")
    require(counter["min_request_bytes"] == min(item["requested"] for item in result),
            f"{path}: minimum request differs from trace")
    require(counter["max_request_bytes"] == max(item["requested"] for item in result),
            f"{path}: maximum request differs from trace")
    return result


def _media_overlap(ranges: list[dict[str, int]], *, returned: bool) -> int:
    total = 0
    for item in ranges:
        length = item["returned"] if returned else item["requested"]
        start = item["offset"]
        end = start + length
        for media_start, media_end in MEDIA_RANGES:
            total += max(0, min(end, media_end) - max(start, media_start))
    return total


def _check_read_ahead(value: Any, path: str, *, candidate: bool) -> dict[str, Any] | None:
    if not candidate:
        require(value is None, f"{path}: baseline unexpectedly contains read-ahead evidence")
        return None
    fields = (
        "window_capacity", "requests", "hits", "misses", "fills",
        "fill_requested_bytes", "fill_returned_bytes", "short_fills",
        "max_fill_bytes", "failures",
    )
    _exact_keys(value, fields, path)
    for field in fields:
        _uint(value[field], f"{path}.{field}")
    require(value["window_capacity"] == WINDOW_BYTES, f"{path}.window_capacity differs")
    require(value["hits"] + value["misses"] == value["requests"],
            f"{path}: requests are not classified exactly once")
    require(value["fills"] <= value["misses"], f"{path}: fills exceed misses")
    require(value["fill_returned_bytes"] <= value["fill_requested_bytes"],
            f"{path}: returned fill bytes exceed requested fill bytes")
    require(value["short_fills"] <= value["fills"], f"{path}: short fills exceed fills")
    require(value["max_fill_bytes"] <= WINDOW_BYTES, f"{path}: fill exceeds window")
    if value["fills"] == 0:
        require(all(value[field] == 0 for field in (
            "fill_requested_bytes", "fill_returned_bytes", "short_fills", "max_fill_bytes")),
                f"{path}: nonzero fill counters without a fill")
    return dict(value)


def _check_sample(value: Any, index: int, arm: dict[str, Any], role: str, path: str) -> dict[str, Any]:
    fields = (
        "cache", "actual_text_bytes", "actual_text_sha256", "actual_text_verified", "allocation",
        "latency_ns", "reads", "sample_index", "source_version_after",
        "source_version_before", "source_version_unchanged",
    )
    expected_fields = set(fields)
    _exact_keys(value, expected_fields if role == "allocator" else expected_fields - {"allocation"}, path)
    require(value["sample_index"] == index, f"{path}.sample_index differs")
    _uint(value["latency_ns"], f"{path}.latency_ns", positive=True)
    require(value["actual_text_verified"] is True, f"{path}.actual_text_verified is false")
    require(value["actual_text_bytes"] == CORPUS["expected_text_bytes"], f"{path}.actual_text_bytes differs")
    require(value["actual_text_sha256"] == CORPUS["expected_text_sha256"], f"{path}.actual_text_sha256 differs")
    before = _check_source_version(value["source_version_before"], f"{path}.source_version_before")
    after = _check_source_version(value["source_version_after"], f"{path}.source_version_after")
    require(value["source_version_unchanged"] is True and before == after,
            f"{path}: source version changed")
    cache = value["cache"]
    _exact_keys(cache, ("open_successful_loads", "successful_loads", "failed_loads",
                        "retained_bytes", "retained_entries", "budget_managed"), f"{path}.cache")
    for field in ("open_successful_loads", "successful_loads", "failed_loads", "retained_bytes", "retained_entries"):
        _uint(cache[field], f"{path}.cache.{field}")
    require(cache["open_successful_loads"] == 0, f"{path}: open cache loads were not zero")
    require(cache["successful_loads"] == 1, f"{path}: final cache load count differs")
    require(cache["failed_loads"] == 0 and cache["budget_managed"] is False,
            f"{path}: cache diagnostic is invalid")

    reads = value["reads"]
    candidate = arm.get("read_ahead_window_bytes") is not None
    read_fields = ["logical_wrapper", "logical_ranges", "physical_ranges",
                   "logical_media_range_proof", "wrapper", "range_adapter", "media_range_proof"]
    if candidate:
        read_fields.insert(0, "read_ahead")
    _exact_keys(reads, read_fields, f"{path}.reads")
    physical = _check_counter(reads["wrapper"], f"{path}.reads.wrapper", available=True,
                              scope=EXPECTED_PHYSICAL_SCOPE)
    logical = _check_counter(reads["logical_wrapper"], f"{path}.reads.logical_wrapper", available=True,
                             scope=EXPECTED_LOGICAL_SCOPE)
    physical_ranges = _check_ranges(reads["physical_ranges"], f"{path}.reads.physical_ranges", physical)
    logical_ranges = _check_ranges(reads["logical_ranges"], f"{path}.reads.logical_ranges", logical)
    expected_ranges = [
        {"offset": offset, "requested": length, "returned": length}
        for offset, length in EXPECTED_LOGICAL_RANGES
    ]
    require(logical_ranges == expected_ranges,
            f"{path}: logical range trace differs from the pinned successful replay")
    physical_media = _check_media(reads["media_range_proof"], f"{path}.reads.media_range_proof", available=True,
                 allow_overlap=True)
    logical_media = _check_media(reads["logical_media_range_proof"], f"{path}.reads.logical_media_range_proof", available=True,
                 allow_overlap=True)
    require(physical_media["observed_call_count"] == len(physical_ranges),
            f"{path}: physical media call count differs from trace")
    require(logical_media["observed_call_count"] == len(logical_ranges),
            f"{path}: logical media call count differs from trace")
    require(physical_media["requested_overlap_bytes"] == _media_overlap(physical_ranges, returned=False),
            f"{path}: physical requested media overlap differs from trace")
    require(physical_media["returned_overlap_bytes"] == _media_overlap(physical_ranges, returned=True),
            f"{path}: physical returned media overlap differs from trace")
    require(logical_media["requested_overlap_bytes"] == _media_overlap(logical_ranges, returned=False),
            f"{path}: logical requested media overlap differs from trace")
    require(logical_media["returned_overlap_bytes"] == _media_overlap(logical_ranges, returned=True),
            f"{path}: logical returned media overlap differs from trace")
    ahead = _check_read_ahead(reads.get("read_ahead"), f"{path}.reads.read_ahead", candidate=candidate)
    range_counter = _check_counter(reads["range_adapter"], f"{path}.reads.range_adapter", available=True,
                                   range_counter=True)
    require(range_counter["logical_calls"] == physical["logical_calls"] and
            range_counter["requested_bytes"] == physical["requested_bytes"] and
            range_counter["returned_bytes"] == physical["returned_bytes"],
            f"{path}: range-adapter counters differ from physical trace")
    require(range_counter["short_reads"] == 0 and
            range_counter["returned_bytes"] == range_counter["requested_bytes"],
            f"{path}: fixed transport returned a short read")
    expected_transport_calls = physical["logical_calls"] if arm["delay_us"] is not None else 0
    expected_paced_calls = physical["logical_calls"] if arm["transfer_bytes_per_second"] else 0
    require(range_counter["delayed_calls"] == expected_transport_calls and
            range_counter["transfer_paced_calls"] == expected_paced_calls,
            f"{path}: transport delay/pacing counters differ from the frozen arm")
    if arm["delay_us"] and arm["transfer_bytes_per_second"]:
        require(range_counter["delayed_calls"] > 0 and range_counter["transfer_paced_calls"] > 0
                and range_counter["transfer_delay_ns"] > 0,
                f"{path}: configured delay/pacing was not observed")
    if candidate:
        require(ahead is not None, f"{path}: candidate read-ahead evidence is missing")
        require(ahead["requests"] == logical["logical_calls"],
                f"{path}: read-ahead requests differ from package logical calls")
        require(all(item["requested"] <= WINDOW_BYTES for item in physical_ranges),
                f"{path}: physical fill exceeded candidate window")
        require(ahead["fills"] == physical["logical_calls"] == range_counter["logical_calls"],
                f"{path}: fill/physical/adapter call counts differ")
        require(ahead["fill_requested_bytes"] == physical["requested_bytes"] == range_counter["requested_bytes"],
                f"{path}: fill/physical/adapter requested bytes differ")
        require(ahead["fill_returned_bytes"] == physical["returned_bytes"] == range_counter["returned_bytes"],
                f"{path}: fill/physical/adapter returned bytes differ")
        require(ahead["max_fill_bytes"] <= WINDOW_BYTES, f"{path}: max fill exceeds candidate window")
        require(ahead["max_fill_bytes"] == max(item["requested"] for item in physical_ranges),
                f"{path}: maximum fill does not match physical trace")
        require(ahead["short_fills"] == physical["short_reads"],
                f"{path}: read-ahead short-fill count differs from physical short reads")
        require(ahead["failures"] == 0, f"{path}: read-ahead recorded a failure")
    else:
        require(physical_ranges == logical_ranges,
                f"{path}: baseline physical trace differs from logical trace")
        require(physical["logical_calls"] == BASELINE_PHYSICAL_CALLS and
                physical["requested_bytes"] == BASELINE_PHYSICAL_BYTES,
                f"{path}: baseline physical observation differs from 0491 trace")
    if role == "allocator":
        _check_allocation(value["allocation"], f"{path}.allocation")
    return dict(value)


def validate_report(report_path: Path, *, role: str, arm_name: str,
                    binary: dict[str, Any], samples: int = FORMAL_SAMPLES,
                    warmups: int = FORMAL_WARMUPS,
                    source_revision: str | None = None) -> dict[str, Any]:
    """Validate one raw Rust report against the frozen provider contract."""

    require(role in ROLES, f"unknown role {role}")
    arm = _arm(arm_name)
    value = _read_json(report_path)
    _finite(value)
    _exact_keys(value, _report_keys(), str(report_path))
    require(value["schema"] == SCHEMA, f"{report_path}: report schema differs")
    require(value["provider_scope"] == EXPECTED_PROVIDER_SCOPE, f"{report_path}: provider scope differs")
    require(value["timing_scope"] == EXPECTED_TIMING_SCOPE, f"{report_path}: timing scope differs")
    require(value["setup_scope"] == EXPECTED_SETUP_SCOPE, f"{report_path}: setup scope differs")
    require(value["allocation_scope"] == EXPECTED_ALLOCATION_SCOPE, f"{report_path}: allocation scope differs")
    _corpus(value["corpus"], f"{report_path}.corpus")
    require(value["media_ranges"] == [{"start": start, "end": end} for start, end in MEDIA_RANGES],
            f"{report_path}.media_ranges differs from the pinned compressed-member oracle")
    require(value["source_bytes"] == CORPUS["archive_bytes"], f"{report_path}.source_bytes differs")
    require(value["source_sha256"] == CORPUS["archive_sha256"], f"{report_path}.source_sha256 differs")
    revision = _text(value["requested_source_revision"], f"{report_path}.requested_source_revision")
    require(REVISION_RE.fullmatch(revision) is not None, f"{report_path}: source revision is malformed")
    if source_revision is not None:
        require(revision == source_revision, f"{report_path}: source revision differs from argv")
    _check_limits(value["limits"], f"{report_path}.limits")
    _check_provider(value["provider"], arm, f"{report_path}.provider")
    require(value["warmup"] == warmups and value["samples"] == samples,
            f"{report_path}: sample contract differs")
    rows = value["rows"]
    require(isinstance(rows, list) and len(rows) == samples, f"{report_path}: row count differs")
    for index, row in enumerate(rows):
        _check_sample(row, index + warmups, arm, role, f"{report_path}.rows[{index}]")
    # The executable identity is bound by the process receipt.  The report's
    # request revision and corpus are separately checked above; no report field
    # is allowed to silently substitute another binary.
    require(binary["sha256"] and binary["bytes"] > 0, "binary binding is incomplete")
    return value


def _file_artifacts(directory: Path) -> dict[str, dict[str, Any]]:
    names = ("report.json", "resource.txt", "stdout.txt", "stderr.txt", "replay-cleanup.json")
    return {name: _file_meta(directory / name) for name in names}


def _validate_terminal(started: dict[str, Any], terminal: dict[str, Any], spec: dict[str, Any],
                       build: dict[str, Any], protocol_hash: str, directory: Path,
                       tmp_root: Path, environment_artifact: dict[str, Any]) -> None:
    started_fields = {
        "schema", "version", "status", "attempt", "run", "protocol", "build", "binary",
        "source", "argv", "cwd", "environment", "environment_artifact", "driver_bindings",
        "driver_sha256", "support_sha256", "started_utc", "timeout_seconds", "tmpdir",
    }
    terminal_fields = started_fields | {
        "exit_code", "timed_out", "termination", "finished_utc", "artifacts", "missing_artifacts",
        "cleanup", "started_artifact", "source_before", "source_after", "source_unchanged",
    }
    _exact_keys(started, started_fields, f"{directory}/started.json")
    _exact_keys(terminal, terminal_fields, f"{directory}/terminal.json")
    require(started.get("schema") == CAPTURE_SCHEMA and started.get("version") == VERSION,
            f"{directory}: started receipt schema differs")
    require(terminal.get("schema") == TERMINAL_SCHEMA and terminal.get("version") == VERSION,
            f"{directory}: terminal receipt schema differs")
    require(started.get("status") == "running", f"{directory}: started status differs")
    require(terminal.get("status") == "pass", f"{directory}: terminal status is not pass")
    for item, label in ((started, "started"), (terminal, "terminal")):
        require(item.get("attempt") == spec["attempt"], f"{directory}: {label} attempt differs")
        require(item.get("run") == {key: spec[key] for key in ("kind", "repeat", "role", "arm", "provider", "samples", "warmups", "label")},
                f"{directory}: {label} run identity differs")
        require(item.get("protocol") == {"path": "protocol.json", "sha256": protocol_hash},
                f"{directory}: {label} protocol binding differs")
        require(item.get("build") == {
            "path": build["path"], "sha256": build["receipt_sha256"], "source": build["source"],
        }, f"{directory}: {label} build binding differs")
        require(item.get("binary") == build["binary"], f"{directory}: {label} binary binding differs")
        require(item.get("cwd") == str(REPO), f"{directory}: {label} cwd differs")
        require(item.get("environment_artifact") == environment_artifact,
                f"{directory}: {label} environment artifact binding differs")
        require(item.get("driver_bindings") == {
            name: _json_hash(ROOT / name) for name in DRIVER_FILES
        }, f"{directory}: {label} helper bindings differ")
        require(item.get("source") == build["source"], f"{directory}: {label} source binding differs")
        require(item.get("driver_sha256") == _json_hash(ROOT / "measure.py"),
                f"{directory}: {label} driver binding differs")
        require(item.get("support_sha256") == _json_hash(ROOT / "support.py"),
                f"{directory}: {label} support binding differs")
    require(terminal.get("source_before") == started.get("source"),
            f"{directory}: terminal source-before binding differs")
    require(terminal.get("source_after") == terminal.get("source_before") and
            terminal.get("source_unchanged") is True,
            f"{directory}: source changed during child capture")
    started_environment = started.get("environment")
    require(isinstance(started_environment, dict), f"{directory}: started environment missing")
    expected_environment = {key: ENV[key] for key in ENV_KEYS}
    expected_environment["TMPDIR"] = str(tmp_root)
    require(started_environment == expected_environment, f"{directory}: started environment differs")
    require(terminal.get("environment") == expected_environment, f"{directory}: terminal environment differs")
    require(started.get("argv") == terminal.get("argv"), f"{directory}: argv changed at terminal")
    source_revision = _argv_revision(started["argv"])
    require(source_revision == build["git_revision"], f"{directory}: source revision differs from build")
    expected_argv = _command(
        spec,
        build["binary"],
        directory / "report.json",
        directory / "resource.txt",
        source_revision,
    )
    require(started["argv"] == expected_argv, f"{directory}: argv does not match frozen run configuration")
    _timestamp(started.get("started_utc"), f"{directory}.started.started_utc")
    started_at = _timestamp(started.get("started_utc"), f"{directory}.started.started_utc")
    finished_at = _timestamp(terminal.get("finished_utc"), f"{directory}.terminal.finished_utc")
    require(finished_at > started_at, f"{directory}: terminal chronology is invalid")
    require(terminal.get("exit_code") == 0 and terminal.get("timed_out") is False,
            f"{directory}: process did not terminate successfully")
    require(terminal.get("termination") is None, f"{directory}: process was terminated")
    require("launch_error" not in terminal and "validation_error" not in terminal,
            f"{directory}: terminal contains a failure detail")
    require(terminal.get("missing_artifacts") == [], f"{directory}: terminal reports missing artifacts")
    require(terminal.get("tmpdir") == str(tmp_root), f"{directory}: terminal TMPDIR differs")
    expected_tmpdir = _run_root(spec["attempt"], spec["label"]) / "tmp"
    require(tmp_root == expected_tmpdir, f"{directory}: private TMPDIR escaped its owned run root")
    cleanup = terminal.get("cleanup")
    require(isinstance(cleanup, dict), f"{directory}: private scratch cleanup is missing")
    _exact_keys(cleanup, ("schema", "status", "root", "tmpdir", "removed", "remaining"),
                f"{directory}/replay-cleanup.json")
    require(cleanup.get("status") == "pass", f"{directory}: private scratch cleanup did not pass")
    require(cleanup.get("root") == str(_run_root(spec["attempt"], spec["label"])),
            f"{directory}: cleanup root differs")
    require(cleanup.get("tmpdir") == str(tmp_root), f"{directory}: cleanup TMPDIR differs")
    require(isinstance(cleanup.get("removed"), list) and isinstance(cleanup.get("remaining"), list),
            f"{directory}: cleanup lists are malformed")
    cleanup_value = _read_json(directory / "replay-cleanup.json")
    require(cleanup == cleanup_value, f"{directory}: cleanup receipt differs from terminal")
    require(cleanup_value.get("remaining") == [], f"{directory}: private scratch remains")
    require(terminal.get("artifacts") == _file_artifacts(directory), f"{directory}: artifact inventory differs")
    require(terminal.get("started_artifact") == _file_meta(directory / "started.json"),
            f"{directory}: started artifact binding differs")
    expected_files = {"started.json", "terminal.json", "report.json", "resource.txt",
                      "stdout.txt", "stderr.txt", "replay-cleanup.json"}
    require({path.name for path in directory.iterdir()} == expected_files,
            f"{directory}: capture file inventory differs")


def _with_cpu_lock(action: Any) -> Any:
    """Serialize the full child capture, including custody snapshots."""

    lock_path = Path(CPU_LOCK)
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    with lock_path.open("a+") as lock:
        fcntl.flock(lock.fileno(), fcntl.LOCK_EX)
        try:
            return action()
        finally:
            fcntl.flock(lock.fileno(), fcntl.LOCK_UN)


def _launch_unlocked(spec: dict[str, Any], build: dict[str, Any], protocol: dict[str, Any],
                     protocol_hash: str, timeout_seconds: int) -> Path:
    attempt = spec["attempt"]
    label = spec["label"]
    capture = _capture_root(attempt, label)
    require(not capture.exists(), f"refusing to replace immutable capture: {capture}")
    capture.mkdir(parents=True, exist_ok=False)
    run_root = _run_root(attempt, label)
    require(not run_root.exists(), f"refusing to replace private scratch: {run_root}")
    run_root.mkdir(parents=True, exist_ok=False)
    tmp_root = run_root / "tmp"
    tmp_root.mkdir(exist_ok=False)
    report = capture / "report.json"
    resource = capture / "resource.txt"
    stdout = capture / "stdout.txt"
    stderr = capture / "stderr.txt"
    source_revision = _source_revision()
    require(source_revision == build["git_revision"], "capture HEAD differs from build revision")
    argv = _command(spec, build["binary"], report, resource, source_revision)
    current = _normalized_snapshot()
    require(current["sha256"] == build["source"]["sha256"],
            "capture source differs from the immutable build source manifest")
    require(protocol.get("source", {}).get("sha256") == current["sha256"],
            "capture source differs from the frozen protocol source")
    environment = {key: ENV[key] for key in ENV_KEYS}
    environment["TMPDIR"] = str(tmp_root)
    run_identity = {key: spec[key] for key in ("kind", "repeat", "role", "arm", "provider", "samples", "warmups", "label")}
    started = {
        "schema": CAPTURE_SCHEMA,
        "version": VERSION,
        "status": "running",
        "attempt": attempt,
        "run": run_identity,
        "protocol": {"path": "protocol.json", "sha256": protocol_hash},
        "build": {"path": build["path"], "sha256": build["receipt_sha256"], "source": build["source"]},
        "binary": build["binary"],
        "source": build["source"],
        "argv": argv,
        "cwd": str(REPO),
        "environment": environment,
        "environment_artifact": protocol["environment_artifact"],
        "driver_bindings": protocol["driver"],
        "driver_sha256": _json_hash(ROOT / "measure.py"),
        "support_sha256": _json_hash(ROOT / "support.py"),
        "started_utc": now(),
        "timeout_seconds": timeout_seconds,
        "tmpdir": str(tmp_root),
    }
    _write_json(capture / "started.json", started)
    stdout.touch(mode=0o664, exist_ok=False)
    stderr.touch(mode=0o664, exist_ok=False)
    timed_out = False
    termination: str | None = None
    launch_error: str | None = None
    exit_code: int | None = None
    process: subprocess.Popen[bytes] | None = None
    run_environment = dict(os.environ)
    run_environment.update(environment)
    try:
        with stdout.open("wb") as out, stderr.open("wb") as err:
            process = subprocess.Popen(
                argv,
                cwd=REPO,
                env=run_environment,
                stdin=subprocess.DEVNULL,
                stdout=out,
                stderr=err,
                start_new_session=True,
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
    validation_error: str | None = None
    try:
        if exit_code == 0 and not timed_out and launch_error is None:
            validate_report(report, role=spec["role"], arm_name=spec["arm"], binary=build["binary"],
                            samples=spec["samples"], warmups=spec["warmups"], source_revision=source_revision)
    except (ProviderMatrixError, OSError, ValueError) as error:
        validation_error = str(error)
    try:
        source_after = _normalized_snapshot()
        source_unchanged = source_after == current
    except (OSError, ValueError) as error:
        source_after = {"error": f"{type(error).__name__}: {error}"}
        source_unchanged = False
    cleanup: dict[str, Any] | None = None
    try:
        cleanup = _cleanup_private(run_root, tmp_root)
    except (ProviderMatrixError, OSError) as error:
        cleanup = {
            "schema": "docx-provider-private-cleanup-v1", "status": "failed",
            "root": str(run_root), "tmpdir": str(tmp_root), "removed": [],
            "remaining": [str(error)],
        }
    _write_json(capture / "replay-cleanup.json", cleanup)
    artifacts: dict[str, dict[str, Any]] = {}
    for path in (stdout, stderr, resource, report, capture / "replay-cleanup.json"):
        if path.is_file():
            artifacts[path.name] = _file_meta(path)
    missing = [path.name for path in (stdout, stderr, resource, report, capture / "replay-cleanup.json") if not path.is_file()]
    passed = (
        exit_code == 0 and not timed_out and termination is None and launch_error is None
        and validation_error is None and source_unchanged and not missing and cleanup.get("status") == "pass"
    )
    terminal = dict(started)
    terminal.update({
        "schema": TERMINAL_SCHEMA,
        "status": "pass" if passed else "failed",
        "exit_code": exit_code,
        "timed_out": timed_out,
        "termination": termination,
        "finished_utc": now(),
        "artifacts": artifacts,
        "missing_artifacts": missing,
        "tmpdir": str(tmp_root),
        "cleanup": cleanup,
        "started_artifact": _file_meta(capture / "started.json"),
        "source_before": current,
        "source_after": source_after,
        "source_unchanged": source_unchanged,
    })
    if launch_error is not None:
        terminal["launch_error"] = launch_error
    if validation_error is not None:
        terminal["validation_error"] = validation_error
    _write_json(capture / "terminal.json", terminal)
    if not passed:
        fail(f"{label} failed; receipt retained: {capture / 'terminal.json'}")
    return capture / "terminal.json"


def _launch(spec: dict[str, Any], build: dict[str, Any], protocol: dict[str, Any],
            protocol_hash: str, timeout_seconds: int) -> Path:
    return _with_cpu_lock(lambda: _launch_unlocked(spec, build, protocol, protocol_hash, timeout_seconds))


def _check_report_artifacts(directory: Path, terminal: dict[str, Any]) -> None:
    artifacts = terminal.get("artifacts")
    require(isinstance(artifacts, dict), f"{directory}: terminal artifact inventory is missing")
    for name in ("stdout.txt", "stderr.txt", "resource.txt", "report.json", "replay-cleanup.json"):
        item = artifacts.get(name)
        require(isinstance(item, dict), f"{directory}: {name} artifact binding is missing")
        path = _resolve_rooted(_text(item.get("path"), f"{directory}.artifacts.{name}.path"))
        require(path == directory / name and path.parent == directory,
                f"{directory}: {name} artifact escaped its capture directory")
        actual = _file_meta(path)
        require(actual["bytes"] == item.get("bytes") and actual["sha256"] == item.get("sha256"),
                f"{directory}: {name} changed after capture")
        if name in {"resource.txt", "report.json", "replay-cleanup.json"}:
            require(actual["bytes"] > 0, f"{directory}: {name} is empty")


def _collect(attempt: str, builds: dict[str, dict[str, Any]], protocol: dict[str, Any],
             protocol_hash: str, *, pilot: bool = False) -> list[dict[str, Any]]:
    require(protocol == protocol_value(builds), "protocol differs from retained builds or current frozen contract")
    expected = formal_inventory(pilot=pilot)
    root = ROOT / "captures" / attempt
    require(root.is_dir(), f"capture attempt is missing: {root}")
    entries: list[dict[str, Any]] = []
    for item in expected:
        spec = dict(item)
        spec["attempt"] = attempt
        directory = root / spec["label"]
        terminal_path = directory / "terminal.json"
        started_path = directory / "started.json"
        require(terminal_path.is_file() and started_path.is_file(), f"{directory}: custody receipts are missing")
        started = _read_json(started_path)
        terminal = _read_json(terminal_path)
        _validate_terminal(
            started,
            terminal,
            spec,
            builds[spec["role"]],
            protocol_hash,
            directory,
            Path(str(started["tmpdir"])),
            protocol["environment_artifact"],
        )
        _check_report_artifacts(directory, terminal)
        report_path = _resolve_rooted(terminal["artifacts"]["report.json"]["path"])
        report = validate_report(report_path, role=spec["role"], arm_name=spec["arm"],
                                binary=builds[spec["role"]]["binary"], samples=spec["samples"],
                                warmups=spec["warmups"], source_revision=_argv_revision(terminal["argv"]))
        resource = _resource_values(_resolve_rooted(terminal["artifacts"]["resource.txt"]["path"]))
        entries.append({
            "spec": spec,
            "directory": directory,
            "started": started,
            "terminal": terminal,
            "terminal_sha256": _json_hash(terminal_path),
            "report": report,
            "report_sha256": _json_hash(report_path),
            "resource": resource,
        })
    actual = {entry["spec"]["label"] for entry in entries}
    expected_labels = {item["label"] for item in expected}
    require(actual == expected_labels, "capture inventory has missing or extra runs")
    actual_dirs = {path.name for path in root.iterdir() if path.is_dir()}
    require(actual_dirs == expected_labels, f"{root}: unexpected capture directories remain")
    require({path.name for path in root.iterdir()} == expected_labels,
            f"{root}: attempt file inventory differs")
    run_attempt = TEMP / "provider" / attempt
    require(not run_attempt.exists(), f"{run_attempt}: private run/tmp files remain after capture")
    _check_chronology(entries)
    return entries


def _check_chronology(entries: list[dict[str, Any]]) -> None:
    previous_end = None
    for entry in entries:
        start = _timestamp(entry["started"]["started_utc"], "capture start")
        end = _timestamp(entry["terminal"]["finished_utc"], "capture end")
        require(end > start, "capture interval is empty or backwards")
        require(previous_end is None or start >= previous_end, "capture order changed or intervals overlap")
        previous_end = end


def _argv_revision(argv: Any) -> str:
    require(isinstance(argv, list), "terminal argv is not a list")
    try:
        index = argv.index("--source-revision")
        value = argv[index + 1]
    except (ValueError, IndexError):
        fail("terminal argv has no source revision")
    require(isinstance(value, str) and REVISION_RE.fullmatch(value) is not None,
            "terminal argv source revision is malformed")
    return value


def _percentiles(values: list[int]) -> dict[str, Any]:
    require(values, "cannot calculate percentiles of an empty vector")
    ordered = sorted(values)
    p50 = (ordered[len(values) // 2 - 1] + ordered[len(values) // 2]) // 2 if len(values) % 2 == 0 else ordered[len(values) // 2]
    return {
        "n": len(values),
        "min": ordered[0],
        "max": ordered[-1],
        "mean": statistics.fmean(values),
        "p50": p50,
        "p95": ordered[math.ceil(len(values) * 0.95) - 1],
        "p99": ordered[math.ceil(len(values) * 0.99) - 1],
    }


def _raw_vectors(report: dict[str, Any], role: str, arm: dict[str, Any], resource: dict[str, Any]) -> dict[str, Any]:
    rows = report["rows"]
    latency = [row["latency_ns"] for row in rows]
    wrapper = [row["reads"]["wrapper"] for row in rows]
    logical_wrapper = [row["reads"]["logical_wrapper"] for row in rows]
    vectors: dict[str, Any] = {
        "latency_ns": latency,
        "actual_text_bytes": [row["actual_text_bytes"] for row in rows],
        "source_version_id": [row["source_version_before"]["id"] for row in rows],
        "source_version_revision": [row["source_version_before"]["revision"] for row in rows],
        "physical_read_calls": [item["logical_calls"] for item in wrapper],
        "physical_read_requested_bytes": [item["requested_bytes"] for item in wrapper],
        "physical_read_returned_bytes": [item["returned_bytes"] for item in wrapper],
        "physical_read_short_reads": [item["short_reads"] for item in wrapper],
        "logical_read_calls": [item["logical_calls"] for item in logical_wrapper],
        "logical_read_requested_bytes": [item["requested_bytes"] for item in logical_wrapper],
        "logical_read_returned_bytes": [item["returned_bytes"] for item in logical_wrapper],
        "logical_read_short_reads": [item["short_reads"] for item in logical_wrapper],
        "physical_media_requested_overlap_bytes": [
            row["reads"]["media_range_proof"]["requested_overlap_bytes"] for row in rows
        ],
        "physical_media_returned_overlap_bytes": [
            row["reads"]["media_range_proof"]["returned_overlap_bytes"] for row in rows
        ],
        "logical_media_requested_overlap_bytes": [
            row["reads"]["logical_media_range_proof"]["requested_overlap_bytes"] for row in rows
        ],
        "logical_media_returned_overlap_bytes": [
            row["reads"]["logical_media_range_proof"]["returned_overlap_bytes"] for row in rows
        ],
        "whole_child_rss_kbytes": resource["maximum_resident_set_size_(kbytes)"],
    }
    range_rows = [row["reads"]["range_adapter"] for row in rows]
    vectors.update({
        "range_logical_calls": [item["logical_calls"] for item in range_rows],
        "range_requested_bytes": [item["requested_bytes"] for item in range_rows],
        "range_returned_bytes": [item["returned_bytes"] for item in range_rows],
        "range_short_reads": [item["short_reads"] for item in range_rows],
        "range_delayed_calls": [item["delayed_calls"] for item in range_rows],
        "range_transfer_paced_calls": [item["transfer_paced_calls"] for item in range_rows],
        "range_transfer_delay_ns": [item["transfer_delay_ns"] for item in range_rows],
    })
    if arm.get("read_ahead_window_bytes") is not None:
        ahead = [row["reads"]["read_ahead"] for row in rows]
        vectors.update({
            "read_ahead_requests": [item["requests"] for item in ahead],
            "read_ahead_hits": [item["hits"] for item in ahead],
            "read_ahead_misses": [item["misses"] for item in ahead],
            "read_ahead_fills": [item["fills"] for item in ahead],
            "read_ahead_fill_requested_bytes": [item["fill_requested_bytes"] for item in ahead],
            "read_ahead_fill_returned_bytes": [item["fill_returned_bytes"] for item in ahead],
            "read_ahead_short_fills": [item["short_fills"] for item in ahead],
            "read_ahead_max_fill_bytes": [item["max_fill_bytes"] for item in ahead],
            "read_ahead_failures": [item["failures"] for item in ahead],
        })
    else:
        for name in ("requests", "hits", "misses", "fills", "fill_requested_bytes",
                     "fill_returned_bytes", "short_fills", "max_fill_bytes", "failures"):
            vectors[f"read_ahead_{name}"] = None
    if role == "allocator":
        allocation_values: list[int] = []
        for index, row in enumerate(rows):
            allocation = row.get("allocation")
            require(isinstance(allocation, dict),
                    f"report.rows[{index}].allocation: allocator evidence is missing")
            _uint(allocation.get("region_peak_live_bytes"),
                  f"report.rows[{index}].allocation.region_peak_live_bytes")
            allocation_values.append(allocation["region_peak_live_bytes"])
        vectors["allocation_region_peak_live_bytes"] = allocation_values
        for field in ("allocation_calls", "reallocation_calls", "allocated_bytes", "deallocated_bytes"):
            vectors["allocation_" + field] = [row["allocation"][field] for row in rows]
        vectors["allocation_peak_increment_bytes"] = [row["allocation"]["region_peak_live_bytes"] - row["allocation"]["live_bytes_before"] for row in rows]
    else:
        require(all("allocation" not in row for row in rows),
                "normal report contains an explicit null allocator record")
        vectors["allocation_region_peak_live_bytes"] = None
    return vectors


def _stats_for_vectors(vectors: dict[str, Any]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for name, values in vectors.items():
        if isinstance(values, list) and values and all(type(value) is int for value in values):
            result[name] = _percentiles(values)
        elif isinstance(values, int):
            result[name] = _percentiles([values])
        else:
            result[name] = None
    return result


def _bootstrap_for_vectors(vectors: dict[str, Any]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for name, values in vectors.items():
        if isinstance(values, list) and values and all(type(item) is int for item in values):
            result[name] = _bootstrap_median_ci(values)
        elif isinstance(values, int):
            result[name] = _bootstrap_median_ci([values])
    return result


def _bootstrap_median_ci(values: list[int], *, repetitions: int = 2_000) -> dict[str, Any]:
    """Return a reproducible percentile bootstrap interval for the median."""

    require(values, "cannot bootstrap an empty vector")
    # A tiny local LCG keeps the receipt independent of Python's versioned
    # ``random`` implementation while remaining deterministic across hosts.
    seed = (0x0492_5EED ^ len(values) ^ sum(values)) & 0xFFFFFFFFFFFFFFFF
    medians: list[int] = []
    count = len(values)
    for _ in range(repetitions):
        sample: list[int] = []
        for _ in range(count):
            seed = (seed * 6364136223846793005 + 1442695040888963407) & 0xFFFFFFFFFFFFFFFF
            sample.append(values[(seed >> 32) % count])
        sample.sort()
        middle = len(sample) // 2
        medians.append(sample[middle] if len(sample) % 2 else
                       (sample[middle - 1] + sample[middle]) // 2)
    medians.sort()
    low = medians[(len(medians) * 25) // 1000]
    high = medians[(len(medians) * 975) // 1000 - 1]
    return {
        "method": "deterministic_percentile_bootstrap_median",
        "seed": (0x0492_5EED ^ len(values) ^ sum(values)) & 0xFFFFFFFFFFFFFFFF,
        "resamples": repetitions,
        "confidence": 0.95,
        "median": _percentiles(values)["p50"],
        "ci_low": low,
        "ci_high": high,
    }


def _relative_percent(first: int | float, second: int | float) -> float | None:
    if first == 0:
        return 0.0 if second == 0 else None
    return (float(second) - float(first)) / abs(float(first)) * 100.0


def analyze_data(entries: list[dict[str, Any]], builds: dict[str, dict[str, Any]], *, pilot: bool = False) -> dict[str, Any]:
    rows: list[dict[str, Any]] = []
    for entry in entries:
        spec = entry["spec"]
        arm = _arm(spec["arm"])
        vectors = _raw_vectors(entry["report"], spec["role"], arm, entry["resource"])
        rows.append({
            "label": spec["label"], "kind": spec["kind"], "repeat": spec["repeat"],
            "role": spec["role"], "arm": spec["arm"], "provider": spec["provider"],
            "samples": spec["samples"], "warmups": spec["warmups"],
            "terminal_sha256": entry["terminal_sha256"], "report_sha256": entry["report_sha256"],
            "raw_vectors": vectors, "percentiles": _stats_for_vectors(vectors),
            "bootstrap_median_ci": _bootstrap_for_vectors(vectors),
            "resource": entry["resource"],
        })
    pairs: list[dict[str, Any]] = []
    for role in ROLES:
        for arm in ARMS:
            matching = [row for row in rows if row["role"] == role and row["arm"] == arm["name"]]
            require(len(matching) == (1 if pilot else 2), f"repeat inventory incomplete for {role}/{arm['name']}")
            if pilot:
                continue
            by_repeat = {row["repeat"]: row for row in matching}
            require(set(by_repeat) == {1, 2}, f"repeat labels incomplete for {role}/{arm['name']}")
            metrics: dict[str, Any] = {}
            for metric in ("latency_ns", "whole_child_rss_kbytes", "allocation_region_peak_live_bytes"):
                first = by_repeat[1]["percentiles"][metric]
                second = by_repeat[2]["percentiles"][metric]
                if first is None or second is None:
                    metrics[metric] = None
                    continue
                changes = {
                    percentile: _relative_percent(first[percentile], second[percentile])
                    for percentile in ("p50", "p95", "p99")
                }
                metrics[metric] = {
                    "repeat1": {key: first[key] for key in ("p50", "p95", "p99")},
                    "repeat2": {key: second[key] for key in ("p50", "p95", "p99")},
                    "relative_percent": changes,
                    "flag_over_5_percent": any(value is not None and abs(value) > 5.0 for value in changes.values()),
                }
            pairs.append({"role": role, "arm": arm["name"], "metrics": metrics,
                          "scope": "descriptive repeat-to-repeat variance; no optimization regression claim"})
    comparisons: list[dict[str, Any]] = []
    adverse_rows: list[dict[str, Any]] = []
    for role in ROLES:
        for repeat in ((1,) if pilot else REPEATS):
            for baseline_arm, candidate_arm in ((ARMS[0], ARMS[1]), (ARMS[2], ARMS[3])):
                baseline_matches = [row for row in rows if row["role"] == role and row["repeat"] == repeat
                                    and row["arm"] == baseline_arm["name"]]
                candidate_matches = [row for row in rows if row["role"] == role and row["repeat"] == repeat
                                     and row["arm"] == candidate_arm["name"]]
                require(len(baseline_matches) == 1 and len(candidate_matches) == 1,
                        f"comparison inventory incomplete for {role}/r{repeat}/{baseline_arm['name']}")
                baseline_row = baseline_matches[0]
                candidate_row = candidate_matches[0]
                metric_names = (
                    "latency_ns", "whole_child_rss_kbytes", "physical_read_calls",
                    "physical_read_requested_bytes", "physical_read_returned_bytes",
                    "logical_read_calls", "logical_read_requested_bytes",
                    "logical_read_returned_bytes", "read_ahead_fills",
                    "read_ahead_fill_requested_bytes", "read_ahead_fill_returned_bytes",
                    "allocation_region_peak_live_bytes", "allocation_peak_increment_bytes",
                    "allocation_allocated_bytes", "allocation_deallocated_bytes",
                    "allocation_calls", "allocation_reallocation_calls",
                )
                metrics: dict[str, Any] = {}
                for metric in metric_names:
                    first = baseline_row["percentiles"].get(metric)
                    second = candidate_row["percentiles"].get(metric)
                    if first is None or second is None:
                        metrics[metric] = None
                        continue
                    changes = {
                        percentile: _relative_percent(first[percentile], second[percentile])
                        for percentile in ("p50", "p95", "p99")
                    }
                    flags = {
                        percentile: change is not None and change > 5.0
                        for percentile, change in changes.items()
                    }
                    metrics[metric] = {
                        "baseline": {key: first[key] for key in ("p50", "p95", "p99")},
                        "candidate": {key: second[key] for key in ("p50", "p95", "p99")},
                        "relative_percent": changes,
                        "adverse_over_5_percent": flags,
                        "flag_over_5_percent": any(flags.values()),
                    }
                    for percentile, flagged in flags.items():
                        if flagged:
                            adverse_rows.append({
                                "role": role,
                                "repeat": repeat,
                                "baseline_arm": baseline_arm["name"],
                                "candidate_arm": candidate_arm["name"],
                                "metric": metric,
                                "percentile": percentile,
                                "relative_percent": changes[percentile],
                            })
                comparisons.append({
                    "role": role,
                    "repeat": repeat,
                    "baseline_arm": baseline_arm["name"],
                    "candidate_arm": candidate_arm["name"],
                    "metrics": metrics,
                    "scope": "descriptive same-role same-delay comparison; positive change is candidate cost; no optimization claim",
                })
    return {
        "schema": ANALYSIS_SCHEMA,
        "version": VERSION,
        "change": CHANGE,
        "claim_authorized": False,
        "performance_claim": "descriptive provider evidence; no cross-arm optimization claim",
        "scope": "same-role/provider repeat variance, bounded read-ahead comparison, and raw lifecycle evidence",
        "pilot": pilot,
        "corpus": dict(CORPUS),
        "case": CASE,
        "roles": list(ROLES),
        "repeats": list((1,) if pilot else REPEATS),
        "arms": [dict(arm) for arm in ARMS],
        "samples": PILOT_SAMPLES if pilot else FORMAL_SAMPLES,
        "warmups": PILOT_WARMUPS if pilot else FORMAL_WARMUPS,
        "builds": {
            role: {
                "receipt_sha256": builds[role]["receipt_sha256"],
                "binary": builds[role]["binary"],
                "source": builds[role]["source"],
                "gate": builds[role]["gate"],
            }
            for role in ROLES
        },
        "rows": rows,
        "repeat_variance": pairs,
        "comparisons": comparisons,
        "adverse_rows_over_5_percent": adverse_rows,
        "inventory": {
            "formal_processes": len(rows),
            "measured_samples": sum(row["samples"] for row in rows),
            "expected_formal_processes": len(formal_inventory(pilot=pilot)),
            "expected_measured_samples": len(formal_inventory(pilot=pilot)) * (PILOT_SAMPLES if pilot else FORMAL_SAMPLES),
        },
    }


def analyze(attempt: str, build_dir: Path | None = None, *, pilot: bool = False) -> Path:
    protocol, protocol_hash = _load_protocol()
    builds = load_builds(build_dir)
    entries = _collect(_attempt(attempt), builds, protocol, protocol_hash, pilot=pilot)
    summary = analyze_data(entries, builds, pilot=pilot)
    directory = ROOT / "analysis"
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / f"{attempt}{'-pilot' if pilot else ''}.json"
    _write_or_match_json(path, summary)
    return path


def verify(attempt: str, build_dir: Path | None = None, *, pilot: bool = False) -> Path:
    """Recompute the summary from retained raw reports without live-source checks."""

    protocol, protocol_hash = _load_protocol()
    builds = load_builds(build_dir)
    entries = _collect(_attempt(attempt), builds, protocol, protocol_hash, pilot=pilot)
    recomputed = analyze_data(entries, builds, pilot=pilot)
    summary_path = ROOT / "analysis" / f"{attempt}{'-pilot' if pilot else ''}.json"
    require(summary_path.is_file(), f"analysis summary is missing: {summary_path}")
    require(_read_json(summary_path) == recomputed, "analysis summary differs from retained raw evidence")
    verification = {
        "schema": VERIFICATION_SCHEMA,
        "version": VERSION,
        "status": "pass",
        "attempt": attempt,
        "pilot": pilot,
        "protocol_sha256": protocol_hash,
        "summary_sha256": _json_hash(summary_path),
        "raw_receipts": [
            {"label": entry["spec"]["label"], "terminal_sha256": entry["terminal_sha256"],
             "report_sha256": entry["report_sha256"]}
            for entry in entries
        ],
        "verified_utc": now(),
    }
    directory = ROOT / "verification"
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / f"{attempt}{'-pilot' if pilot else ''}.json"
    _write_or_match_json(path, verification, volatile=("verified_utc",))
    return path


def capture_one(attempt: str, role: str, repeat: int, arm_name: str, *, pilot: bool = False,
                build_dir: Path | None = None, timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS) -> Path:
    require(role in ROLES, f"unknown role {role}")
    require(repeat in ((1,) if pilot else REPEATS), f"invalid repeat {repeat}")
    require(arm_name in ARM_BY_NAME, f"unknown provider arm {arm_name}")
    protocol, protocol_hash = _load_protocol()
    builds = load_builds(build_dir)
    require(protocol == protocol_value(builds), "protocol/build contract differs")
    match = [item for item in formal_inventory(pilot=pilot)
             if item["role"] == role and item["repeat"] == repeat and item["arm"] == arm_name]
    require(len(match) == 1, "capture specification is not in frozen inventory")
    return _launch({**match[0], "attempt": _attempt(attempt)}, builds[role], protocol, protocol_hash, timeout_seconds)


def capture_all(attempt: str, build_dir: Path | None = None, *, pilot: bool = False,
                timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS) -> None:
    protocol, protocol_hash = _load_protocol()
    builds = load_builds(build_dir)
    require(protocol == protocol_value(builds), "protocol/build contract differs")
    for spec in formal_inventory(pilot=pilot):
        _launch({**spec, "attempt": _attempt(attempt)}, builds[spec["role"]], protocol, protocol_hash, timeout_seconds)


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("plan", help="print the frozen provider matrix")
    for command in ("capture", "capture-all"):
        item = sub.add_parser(command, help="capture immutable provider process evidence")
        item.add_argument("--attempt", required=True)
        item.add_argument("--build-dir", type=Path)
        item.add_argument("--timeout-seconds", type=int, default=DEFAULT_TIMEOUT_SECONDS)
        item.add_argument("--pilot", action="store_true")
        if command == "capture":
            item.add_argument("--role", choices=ROLES, required=True)
            item.add_argument("--repeat", type=int, choices=REPEATS, required=True)
            item.add_argument("--arm", choices=tuple(ARM_BY_NAME), required=True)
    for command in ("analyze", "verify"):
        item = sub.add_parser(command, help="recompute retained provider evidence")
        item.add_argument("--attempt", required=True)
        item.add_argument("--build-dir", type=Path)
        item.add_argument("--pilot", action="store_true")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        if args.command == "plan":
            builds = None
            if (ROOT / "build-normal.json").is_file() and (ROOT / "build-allocator.json").is_file():
                builds = load_builds()
            print(json.dumps(protocol_value(builds), indent=2, sort_keys=True))
        elif args.command == "capture":
            capture_one(args.attempt, args.role, args.repeat, args.arm, pilot=args.pilot,
                        build_dir=args.build_dir, timeout_seconds=args.timeout_seconds)
        elif args.command == "capture-all":
            capture_all(args.attempt, args.build_dir, pilot=args.pilot, timeout_seconds=args.timeout_seconds)
        elif args.command == "analyze":
            print(analyze(args.attempt, args.build_dir, pilot=args.pilot))
        elif args.command == "verify":
            print(verify(args.attempt, args.build_dir, pilot=args.pilot))
        else:  # pragma: no cover
            fail(f"unknown command {args.command}")
    except (ProviderMatrixError, OSError, ValueError, subprocess.SubprocessError) as error:
        print(f"measure.py: FAIL: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
