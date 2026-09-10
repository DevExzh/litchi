#!/usr/bin/env python3
"""Capture and audit the 0491 DOCX filesystem cache-state matrix.

This helper is deliberately a custody and validation layer around the existing
``litchi-perf-baseline`` filesystem selectors.  Importing it never builds or
executes a benchmark.  A formal child selects one cache state and one role;
the Rust harness then owns the fresh child per sample and the timed lifecycle.

The matrix has no before/after arm and therefore makes no optimization claim.
Its output is standalone warm/cold-requested/cold-verified evidence plus the
explicit prepared-query cold-ineligible control.
"""

from __future__ import annotations

import argparse
import copy
import datetime
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

from support import ENV, ENV_KEYS, REPO, ROOT, TEMP, meta, now, read, sha, write


SCHEMA = "docx-filesystem-cold-matrix-v1"
REPORT_SCHEMA_VERSION = 1
CASE = "docx_file_source_open_full_text_lifecycle"
CONTROL_CASE = "docx_file_source_full_text"
ROLES = ("normal", "allocator")
REPEATS = (1, 2)
CACHE_STATES = ("warm", "cold-requested", "cold-verified")
CPU = 2
FORMAL_SAMPLES = 30
FORMAL_WARMUPS = 3
PILOT_SAMPLES = 3
PILOT_WARMUPS = 1
DEFAULT_TIMEOUT_SECONDS = 1_800
GRACE_SECONDS = 10
ATTEMPT_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]*$")
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")

# This is the fixed corpus bound by change 0188.  The driver intentionally
# does not replace these values with a hash observed from a failed pilot.
CORPUS = {
    "name": "docx-source-backed-media",
    "generator": "litchi-docx-source-edit-media-v1",
    "package_format": "DOCX/OPC/ZIP",
    "shape": "media-rich",
    "payload_kind": "deterministic-incompressible-media",
    "compression": "deflate",
    "archive_member_count": 20,
    "archive_bytes": 16_793_036,
    "archive_sha256": "a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4",
}
EXPECTED_PARAGRAPHS = 200
EXPECTED_TEXT_BYTES = 10_000
EXPECTED_TEXT_SHA256 = "ad4fe690f0ef2281ad8e64a78d1f4d64e7c8625d672b3ac2f7e9fbc28a82f4af"
# Historical replay hashes the finalized SHA-256 bytes a second time.
EXPECTED_REPLAY_SHA256 = hashlib.sha256(bytes.fromhex(EXPECTED_TEXT_SHA256)).hexdigest()
EXPECTED_MEDIA_MEMBERS = 8
EXPECTED_REPLAY_CLASSIFICATION = (
    "semantic-query:one-complete-main-range-preparation-zero-query-"
    "unselected-media-core"
)
EXPECTED_ALIGNED_REPLAY_CLASSIFICATION = (
    "semantic-query:aligned-eocd-tail-metadata-probe;"
    "one-complete-main-range-preparation-zero-query-unselected-media-core"
)
EXPECTED_COLD_SCOPE = (
    "external fincore page-cache residency/dirty/writeback proof plus "
    "positive process read_bytes; no physical-media claim"
)
EXPECTED_FINCORE_COMMAND = (
    "fincore --json --bytes --output FILE,SIZE,RES,DIRTY,WRITEBACK --"
)
EXPECTED_FINCORE_METHOD = "external_fincore_json_columns"
EXPECTED_FINCORE_FALLBACK = "none"
EXPECTED_COLD_ADVICE = "posix_fadvise_dontneed_accepted"
EXPECTED_TEXT_SCOPE = (
    "v2:open_plus_full_text;document_drop_inside_timer;text_digest_after_timer"
)
DRIVER_FILES = ("cold_matrix.py", "support.py")
HELPER_FILES = ("cold_matrix.py", "support.py")
CAPTURE_ARTIFACTS = (
    "stdout.txt",
    "stderr.txt",
    "resource.txt",
    "report.json",
    "corpus-manifest.json",
)
MACHINE_FILE = ROOT / "machine.json"
ALIGNMENT_FILE = ROOT / "aligned-corpus-oracle.json"
PROTOCOL_FILE = ROOT / "cold-protocol.json"
ALIGNMENT_BYTES = 419
ALIGNMENT_SHA256 = "82e2a8595746b96f8c6e93e90678337a5d4e982524e07da45967c33f46903169"
ALIGNMENT_ORACLE = {
    "source_sha256": CORPUS["archive_sha256"],
    "source_bytes": CORPUS["archive_bytes"],
    "page_size_bytes": 4096,
    "padding_bytes": 564,
    "aligned_source_bytes": 16_793_600,
    "aligned_source_sha256": "d1e6f59e6c6c698aa91463a2ed341351da3b549b818122bc0a2b9b9b8449dc88",
    "algorithm": "EOCD comment length and zero comment padding to next page",
    "all_members_crc_and_bytes_equal": True,
}
EXPECTED_ALIGNED_TAIL_PROBE = {
    "aligned_source_sha256": ALIGNMENT_ORACLE["aligned_source_sha256"],
    "aligned_source_bytes": ALIGNMENT_ORACLE["aligned_source_bytes"],
    "unaligned_source_bytes": CORPUS["archive_bytes"],
    "page_size_bytes": ALIGNMENT_ORACLE["page_size_bytes"],
    "alignment_padding_bytes": ALIGNMENT_ORACLE["padding_bytes"],
    "eocd_offset": 16_793_014,
    "eocd_comment_bytes": 564,
    "eocd_tail_probe_offset": 16_728_064,
    "eocd_tail_probe_bytes": 65_536,
    "eocd_tail_probe_read_count": 1,
    "eocd_tail_probe_main_payload_overlap_bytes": 0,
    "eocd_tail_probe_media_payload_overlap_bytes": 58_808,
    "eocd_tail_probe_unselected_payload_overlap_bytes": 4_500,
    "eocd_tail_probe_core_payload_overlap_bytes": 0,
    "eocd_tail_probe_payload_overlap_bytes": 63_308,
    "cache_successful_loads": 0,
}
EXPECTED_ENVIRONMENT = {key: ENV.get(key) for key in ENV_KEYS}


class ColdMatrixError(RuntimeError):
    """A protocol, custody, report, or proof invariant failed closed."""


def fail(message: str) -> None:
    raise ColdMatrixError(message)


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


def _file_meta(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing regular artifact: {path}")
    value = meta(path)
    require(value["bytes"] >= 0 and SHA256_RE.fullmatch(value["sha256"]) is not None,
            f"{path}: malformed artifact metadata")
    return {"path": str(path), **value}


def _finite(value: Any, path: str = "json") -> None:
    """Reject NaN/infinity and non-string dictionary keys recursively."""

    if isinstance(value, float):
        require(math.isfinite(value), f"{path}: non-finite number")
    elif isinstance(value, dict):
        for key, child in value.items():
            require(isinstance(key, str), f"{path}: non-string JSON key")
            _finite(child, f"{path}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            _finite(child, f"{path}[{index}]")


def _positive_int(value: Any, path: str) -> int:
    require(type(value) is int and value > 0, f"{path}: expected positive integer")
    return value


def _nonnegative_int(value: Any, path: str) -> int:
    require(type(value) is int and value >= 0, f"{path}: expected non-negative integer")
    return value


def _hash(value: Any, path: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
            f"{path}: expected lowercase SHA-256")
    return value


def _json_hash(path: Path) -> str:
    return _hash(sha(path), str(path))


def _resolve_path(value: Any, base: Path, label: str) -> Path:
    require(isinstance(value, str) and value, f"{label}: path is missing")
    candidate = Path(value)
    if candidate.is_absolute():
        return candidate
    local = base / candidate
    if local.exists():
        return local
    # support.snapshot() records paths relative to the change directory.  A
    # gate receipt lives one directory below it, so retain that established
    # root-relative spelling while still requiring the resolved file below.
    return ROOT / candidate


def _timestamp(value: Any, path: str) -> datetime.datetime:
    require(isinstance(value, str) and value, f"{path}: timestamp is missing")
    try:
        parsed = datetime.datetime.fromisoformat(value)
    except ValueError as error:
        fail(f"{path}: invalid timestamp: {error}")
    require(parsed.tzinfo is not None, f"{path}: timestamp must include timezone")
    return parsed


def _environment() -> dict[str, Any]:
    return dict(EXPECTED_ENVIRONMENT)


def _machine_binding() -> dict[str, Any]:
    path = MACHINE_FILE
    value = _read_json(path)
    _finite(value, str(path))
    require(isinstance(value, dict), f"{path}: machine record is not an object")
    require(value.get("schema") == "docx-stream-route-machine-v1",
            f"{path}: machine schema differs")
    require(value.get("selected_cpu") == CPU, f"{path}: selected CPU differs")
    actual = _file_meta(path)
    return {"path": str(path), **{key: actual[key] for key in ("bytes", "sha256")}}


def _alignment_binding() -> dict[str, Any]:
    actual = _file_meta(ALIGNMENT_FILE)
    require(actual["bytes"] == ALIGNMENT_BYTES and actual["sha256"] == ALIGNMENT_SHA256,
            f"{ALIGNMENT_FILE}: alignment oracle changed")
    value = _read_json(ALIGNMENT_FILE)
    _finite(value, str(ALIGNMENT_FILE))
    require(value == ALIGNMENT_ORACLE, f"{ALIGNMENT_FILE}: alignment oracle values changed")
    return {"path": str(ALIGNMENT_FILE), **{key: actual[key] for key in ("bytes", "sha256")}}


def _source_binding(value: Any, base: Path, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: source identity is missing")
    source_path = _resolve_path(value.get("path"), base, f"{label}.path")
    actual = _file_meta(source_path)
    source_hash = _hash(value.get("sha256"), f"{label}.sha256")
    files = _positive_int(value.get("files"), f"{label}.files")
    require(actual["sha256"] == source_hash, f"{label}: source manifest hash changed")
    manifest = _read_json(source_path)
    _finite(manifest, str(source_path))
    require(isinstance(manifest, dict), f"{label}: source manifest is not an object")
    require(len(manifest) == files, f"{label}: source manifest file count differs")
    for name, digest in manifest.items():
        require(isinstance(name, str) and name, f"{label}: source manifest name is invalid")
        _hash(digest, f"{label}.{name}")
    return {"path": str(source_path), "files": files, "sha256": source_hash}


def _tool_identity(role: str) -> dict[str, str]:
    require(role in ROLES, f"unknown role {role}")
    return {
        "binary": "litchi-perf-baseline" if role == "normal" else "litchi-perf-baseline-alloc",
        "instrumentation": "none" if role == "normal" else "system_allocator_operation_scoped",
    }


def _command_contract(command: Any, binary_name: str, label: str) -> list[str]:
    expected = ["cargo", "build", "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml"]
    if binary_name.endswith("-alloc"):
        expected += ["--features", "allocator-metrics"]
    expected += ["--bin", binary_name]
    require(command == expected, f"{label}: exact locked release build command differs")
    return list(command)



def _gate_binding(value: dict[str, Any], path: Path, source_before: dict[str, Any],
                  source_after: dict[str, Any], command: list[str], role: str) -> dict[str, Any]:
    gate = value.get("gate")
    require(isinstance(gate, dict), f"{path}: bound gate receipt is missing")
    gate_path = _resolve_path(gate.get("path"), path.parent, f"{path}.gate")
    gate_hash = _hash(gate.get("sha256"), f"{path}.gate.sha256")
    gate_actual = _file_meta(gate_path)
    require(gate_actual["sha256"] == gate_hash, f"{path}: gate receipt changed")
    gate_value = _read_json(gate_path)
    _finite(gate_value, str(gate_path))
    require(isinstance(gate_value, dict), f"{gate_path}: gate receipt is not an object")
    require(gate_value.get("schema") == "docx-stream-append-gate-v1",
            f"{gate_path}: gate schema differs")
    require(gate_value.get("exit_code") == 0, f"{gate_path}: gate command did not pass")
    require(gate_value.get("source_unchanged") is True,
            f"{gate_path}: gate source custody failed")
    gate_before = _source_binding(gate_value.get("source_before"), gate_path.parent,
                                  f"{gate_path}.source_before")
    gate_after = _source_binding(gate_value.get("source_after"), gate_path.parent,
                                 f"{gate_path}.source_after")
    require(gate_before == source_before and gate_after == source_after,
            f"{gate_path}: gate source binding differs")
    require(gate_before == gate_after, f"{gate_path}: gate source changed")
    require(gate_value.get("argv") == command, f"{gate_path}: gate command differs")
    require(gate_value.get("cwd") == str(REPO), f"{gate_path}: gate cwd differs")
    require(gate_value.get("environment") == _environment(),
            f"{gate_path}: gate environment differs")
    require(gate_value.get("driver_sha256") == _json_hash(ROOT / "gate.py"),
            f"{gate_path}: gate helper hash differs")
    require(gate_value.get("common_sha256") == _json_hash(ROOT / "support.py"),
            f"{gate_path}: support helper hash differs")
    started = _timestamp(gate_value.get("started_utc"), f"{gate_path}.started_utc")
    finished = _timestamp(gate_value.get("finished_utc"), f"{gate_path}.finished_utc")
    require(finished > started, f"{gate_path}: gate finished before it started")
    label = gate_value.get("label")
    require(isinstance(label, str) and label, f"{gate_path}: gate label is missing")
    artifacts = gate_value.get("artifacts")
    require(isinstance(artifacts, dict), f"{gate_path}: gate artifacts are missing")
    expected_artifacts = {f"{label}.stdout", f"{label}.stderr"}
    require(set(artifacts) == expected_artifacts, f"{gate_path}: gate artifact inventory differs")
    for name, item in artifacts.items():
        require(isinstance(item, dict), f"{gate_path}: malformed gate artifact {name}")
        artifact_path = gate_path.parent / name
        actual = _file_meta(artifact_path)
        require(item.get("bytes") == actual["bytes"] and item.get("sha256") == actual["sha256"],
                f"{gate_path}: gate artifact {name} changed")
    return {"path": str(gate_path), "bytes": gate_actual["bytes"], "sha256": gate_hash}


def _canonical_hash(value: Any) -> str:
    encoded = _canonical_bytes(value)
    return hashlib.sha256(encoded).hexdigest()


def _canonical_bytes(value: Any) -> bytes:
    """Return the protocol/catalog canonical JSON representation."""

    return json.dumps(value, ensure_ascii=False, sort_keys=True,
                      separators=(",", ":"), allow_nan=False).encode("utf-8")


def _freeze_protocol() -> Path:
    """Write the immutable capture protocol once, after all helpers are final."""

    require(not PROTOCOL_FILE.exists(), f"refusing to replace immutable protocol: {PROTOCOL_FILE}")
    value = protocol_value()
    encoded = _canonical_bytes(value) + b"\n"
    try:
        with PROTOCOL_FILE.open("xb") as stream:
            stream.write(encoded)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError:
        fail(f"refusing to replace immutable protocol: {PROTOCOL_FILE}")
    return PROTOCOL_FILE


def _protocol_binding() -> dict[str, Any]:
    """Load the frozen protocol and reject any byte or value drift."""

    actual = _file_meta(PROTOCOL_FILE)
    expected = _canonical_bytes(protocol_value()) + b"\n"
    require(PROTOCOL_FILE.read_bytes() == expected,
            f"{PROTOCOL_FILE}: retained bytes differ from canonical protocol")
    value = _read_json(PROTOCOL_FILE)
    _finite(value, str(PROTOCOL_FILE))
    require(value == protocol_value(),
            f"{PROTOCOL_FILE}: retained values differ from canonical protocol")
    return {
        "path": str(PROTOCOL_FILE),
        "bytes": actual["bytes"],
        "sha256": actual["sha256"],
    }


def _binary_from_build(value: dict[str, Any], role: str, path: Path) -> dict[str, Any]:
    _finite(value, str(path))
    require(value.get("schema") == "docx-provider-lifecycle-build-v1" and value.get("version") == 1, f"{path}: build schema differs")
    revision = value.get("git_revision")
    require(isinstance(revision, str) and re.fullmatch(r"[0-9a-f]{40}", revision) is not None, f"{path}: build revision missing")
    require(value.get("role") == role, f"{path}: build role differs from {role}")
    binary = value.get("binary")
    require(isinstance(binary, dict), f"{path}: binary metadata is missing")
    binary_path = _resolve_path(binary.get("path"), path.parent, f"{path}.binary")
    require(binary_path.is_absolute(), f"{path}: binary path must be absolute")
    actual = _file_meta(binary_path)
    executable = os.access(binary_path, os.X_OK)
    require(executable, f"{path}: binary is not executable")
    for field in ("bytes", "sha256"):
        require(binary.get(field) == actual[field], f"{path}: binary.{field} changed")
    require(binary.get("executable") == executable, f"{path}: binary executable bit changed")
    require(binary.get("executable") is True, f"{path}: binary is not marked executable")
    binary_name = _tool_identity(role)["binary"]
    require(binary_path.name == binary_name, f"{path}: binary filename differs")
    command = _command_contract(value.get("command"), binary_name, str(path))
    require(value.get("environment") == _environment(), f"{path}: build environment differs")
    source_before = _source_binding(value.get("source_before"), path.parent,
                                    f"{path}.source_before")
    source_after = _source_binding(value.get("source_after"), path.parent,
                                   f"{path}.source_after")
    require(value.get("source_unchanged") is True and source_before == source_after,
            f"{path}: source custody failed")
    gate = _gate_binding(value, path, source_before, source_after, command, role)
    path = path.resolve()
    receipt_hash = _json_hash(path)
    normalized_binary = {
        "path": str(binary_path),
        "bytes": actual["bytes"],
        "sha256": actual["sha256"],
        "executable": True,
    }
    binding = {
        "path": str(path),
        "sha256": receipt_hash,
        "role": role,
        "git_revision": revision,
        "source_before": source_before,
        "source_after": source_after,
        "source_unchanged": True,
        "command": command,
        "environment": _environment(),
        "gate": gate,
    }
    return {
        "role": role,
        "path": path,
        "receipt_sha256": receipt_hash,
        "binary": normalized_binary,
        "source": source_after,
        "source_before": source_before,
        "source_after": source_after,
        "environment": _environment(),
        "command": command,
        "gate": gate,
        "binding": binding,
        "git_revision": revision,
    }


def load_builds(build_dir: Path | None = None) -> dict[str, dict[str, Any]]:
    """Load and rehash the coordinator's immutable normal/allocator receipts."""

    directory = (ROOT if build_dir is None else Path(build_dir)).resolve()
    require(directory.is_dir() and not directory.is_symlink(),
            f"build directory is missing: {directory}")
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
    return result


def formal_inventory(*, pilot: bool = False) -> list[dict[str, Any]]:
    """Return the immutable run order; repeat two reverses role order."""

    samples = PILOT_SAMPLES if pilot else FORMAL_SAMPLES
    warmups = PILOT_WARMUPS if pilot else FORMAL_WARMUPS
    runs: list[dict[str, Any]] = []
    repeats = (1,) if pilot else REPEATS
    for repeat in repeats:
        roles = ROLES if repeat == 1 else tuple(reversed(ROLES))
        for role in roles:
            for state in CACHE_STATES:
                runs.append({
                    "kind": "formal" if not pilot else "pilot",
                    "repeat": repeat,
                    "role": role,
                    "cache_state": state,
                    "case": CASE,
                    "samples": samples,
                    "warmups": warmups,
                    "label": f"{'pilot-' if pilot else ''}r{repeat}-{role}-{state}",
                })
            runs.append({
                "kind": "control",
                "repeat": repeat,
                "role": role,
                "cache_state": "cold-verified",
                "case": CONTROL_CASE,
                "samples": samples,
                "warmups": warmups,
                "label": f"{'pilot-' if pilot else ''}r{repeat}-{role}-prepared-control",
            })
    return runs


def protocol_value() -> dict[str, Any]:
    runs = formal_inventory()
    return {
        "schema": SCHEMA,
        "version": 1,
        "change": 491,
        "claim_authorized": False,
        "performance_claim": "none",
        "scope": "fixed 0188 DOCX source-backed full-text lifecycle cache-state evidence",
        "corpus": dict(CORPUS),
        "case": CASE,
        "prepared_query_control": CONTROL_CASE,
        "cpu": CPU,
        "cache_states": list(CACHE_STATES),
        "roles": list(ROLES),
        "repeats": list(REPEATS),
        "samples": FORMAL_SAMPLES,
        "warmups": FORMAL_WARMUPS,
        "expected_formal_processes": len(runs),
        "formal_runs": runs,
        "driver": {name: _json_hash(ROOT / name) for name in DRIVER_FILES if (ROOT / name).is_file()},
        "machine": _machine_binding(),
        "aligned_corpus_oracle": _alignment_binding(),
        "environment": {key: ENV.get(key) for key in ENV_KEYS},
        "fincore_tool": _read_json(ROOT / "fincore-tool.json"),
        "cold_proof": {
            "scope": EXPECTED_COLD_SCOPE,
            "fincore_command": EXPECTED_FINCORE_COMMAND,
            "fincore_method": EXPECTED_FINCORE_METHOD,
            "fincore_fallback": EXPECTED_FINCORE_FALLBACK,
            "advice": EXPECTED_COLD_ADVICE,
        },
        "limitations": [
            "cold-verified is page-cache/procfs evidence and makes no physical-media claim",
            "no before/after, speedup, optimization, or global cache claim is authorized",
            "the prepared full-text selector is an explicit cold-ineligible control",
        ],
    }


def _run_root(attempt: str, label: str) -> Path:
    return TEMP / "cold-matrix" / attempt / label


def _capture_root(attempt: str, label: str) -> Path:
    return ROOT / "captures" / attempt / label


def _command(spec: dict[str, Any], binary: dict[str, Any], report: Path,
             catalog: Path, filesystem_root: Path) -> list[str]:
    return [
        "/usr/bin/time", "-v", "-o", str(_capture_root(spec["attempt"], spec["label"]) / "resource.txt"),
        "/usr/bin/taskset", "-c", str(CPU),
        str(binary["path"]),
        "--case", spec["case"],
        "--filesystem-cache", spec["cache_state"],
        "--filesystem-root", str(filesystem_root),
        "--samples", str(spec["samples"]),
        "--warmup", str(spec["warmups"]),
        "--json", str(report),
        "--corpus-manifest", str(catalog),
    ]


def _kill_group(process: subprocess.Popen[bytes], timed_out: bool) -> str | None:
    if process.poll() is not None:
        return None
    try:
        os.killpg(process.pid, signal.SIGTERM)
    except ProcessLookupError:
        return None
    try:
        process.communicate(timeout=GRACE_SECONDS)
    except subprocess.TimeoutExpired:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
        process.communicate()
        return "SIGKILL"
    return "SIGTERM"


def _check_report_build(report: dict[str, Any], build: dict[str, Any]) -> None:
    identity = report.get("binary_identity")
    require(isinstance(identity, dict), "report binary identity missing")
    for actual, expected in (("path", "path"), ("binary_bytes", "bytes"), ("binary_sha256", "sha256"), ("executable", "executable")):
        require(identity.get(actual) == build["binary"][expected], "report executable identity differs from build")
    require(identity.get("profile") == "release", "report is not a release build")
    require(report.get("environment", {}).get("git_revision") == build["git_revision"], "report revision differs from build")


def _launch(spec: dict[str, Any], build: dict[str, Any], timeout_seconds: int) -> Path:
    require(type(timeout_seconds) is int and timeout_seconds > 0,
            "timeout must be a positive integer")
    attempt = spec["attempt"]
    label = spec["label"]
    capture = _capture_root(attempt, label)
    require(not capture.exists(), f"refusing to replace immutable capture: {capture}")
    capture.mkdir(parents=True, exist_ok=False)
    replay = _run_root(attempt, label)
    replay.mkdir(parents=True, exist_ok=False)
    report = capture / "report.json"
    catalog = capture / "corpus-manifest.json"
    resource = capture / "resource.txt"
    stdout = capture / "stdout.txt"
    stderr = capture / "stderr.txt"
    filesystem_root = replay / "filesystem"
    filesystem_root.mkdir(parents=True, exist_ok=False)
    binary = build["binary"]
    argv = _command(spec, binary, report, catalog, filesystem_root)
    protocol = _protocol_binding()
    machine = _machine_binding()
    tool = _tool_identity(spec["role"])
    source = copy.deepcopy(build["source"])
    started = {
        "schema": SCHEMA,
        "version": 1,
        "status": "running",
        "attempt": attempt,
        "run": {key: spec[key] for key in ("kind", "repeat", "role", "cache_state", "case", "samples", "warmups", "label")},
        "build": {
            **copy.deepcopy(build["binding"]),
        },
        "binary": binary,
        "source": source,
        "tool": tool,
        "machine": machine,
        "protocol": protocol,
        "argv": argv,
        "cwd": str(REPO),
        "environment": _environment(),
        "driver_sha256": _json_hash(Path(__file__)),
        "support_sha256": _json_hash(ROOT / "support.py"),
        "started_utc": now(),
        "timeout_seconds": timeout_seconds,
        "process_group": {"start_new_session": True},
    }
    _write_json(capture / "started.json", started)
    stdout.touch(mode=0o664, exist_ok=False)
    stderr.touch(mode=0o664, exist_ok=False)
    timed_out = False
    termination: str | None = None
    launch_error: str | None = None
    exit_code: int | None = None
    process: subprocess.Popen[bytes] | None = None
    try:
        with stdout.open("wb") as out, stderr.open("wb") as err:
            process = subprocess.Popen(
                argv,
                cwd=REPO,
                env=ENV,
                stdin=subprocess.DEVNULL,
                stdout=out,
                stderr=err,
                start_new_session=True,
            )
            try:
                process.communicate(timeout=timeout_seconds)
            except subprocess.TimeoutExpired:
                timed_out = True
                termination = _kill_group(process, timed_out)
            exit_code = process.returncode
    except (OSError, subprocess.SubprocessError) as error:
        launch_error = f"{type(error).__name__}: {error}"
        if process is not None and process.poll() is None:
            termination = _kill_group(process, timed_out)
            exit_code = process.returncode

    validation_error: str | None = None
    report_value: dict[str, Any] | None = None
    try:
        if exit_code == 0:
            report_value = validate_report(
                report,
                role=spec["role"],
                cache_state=spec["cache_state"],
                control=spec["kind"] == "control",
                samples=spec["samples"],
                warmups=spec["warmups"],
            )
            _check_report_build(report_value, build)
            checked_catalog = _check_catalog(catalog, report_value, control=spec["kind"] == "control")
            require(checked_catalog["build"]["git_revision"] == build["git_revision"], "catalog revision differs from build")
    except (ColdMatrixError, OSError, ValueError) as error:
        validation_error = str(error)

    artifacts: dict[str, dict[str, Any]] = {}
    for path in (stdout, stderr, resource, report, catalog):
        if path.is_file():
            artifacts[path.name] = _file_meta(path)
    missing = [path.name for path in (stdout, stderr, resource, report, catalog) if not path.is_file()]
    passed = (
        exit_code == 0
        and not timed_out
        and launch_error is None
        and validation_error is None
        and not missing
    )
    terminal = dict(started)
    terminal.update({
        "status": "pass" if passed else "failed",
        "exit_code": exit_code,
        "timed_out": timed_out,
        "termination": termination,
        "finished_utc": now(),
        "artifacts": artifacts,
        "missing_artifacts": missing,
        "replay_root": str(replay),
        "process_id": None if process is None else process.pid,
        "process_group_id": None if process is None else process.pid,
    })
    if launch_error is not None:
        terminal["launch_error"] = launch_error
    if validation_error is not None:
        terminal["validation_error"] = validation_error
    terminal["started_artifact"] = _file_meta(capture / "started.json")
    _write_json(capture / "terminal.json", terminal)
    # The terminal receipt itself is the final custody record.  It is written
    # once above; a hash of it is recorded by the analyzer after completion.
    if not passed:
        fail(f"{label} failed; receipt retained: {capture / 'terminal.json'}")
    return capture / "terminal.json"


def _manifest_matches(actual: Any, path: str) -> dict[str, Any]:
    require(isinstance(actual, dict), f"{path}: corpus manifest is missing")
    for key, expected in CORPUS.items():
        require(actual.get(key) == expected, f"{path}.{key}: fixed 0188 corpus drifted")
    return actual


def _catalog_hash(value: dict[str, Any]) -> str:
    without_hash = copy.deepcopy(value)
    without_hash.pop("catalog_sha256", None)
    return _canonical_hash(without_hash)


def _catalog_content_hash(value: dict[str, Any]) -> str:
    corpora = []
    for corpus in value["corpora"]:
        require(isinstance(corpus, dict), "corpus catalog corpus entry is malformed")
        bytes_value = corpus.get("bytes")
        require(isinstance(bytes_value, dict), "corpus catalog corpus bytes are malformed")
        members_value = corpus.get("members")
        require(isinstance(members_value, dict) and isinstance(members_value.get("items"), list),
                "corpus catalog member inventory is malformed")
        members = members_value["items"]
        normalized_members = []
        for item in members:
            require(isinstance(item, dict), "corpus catalog member is malformed")
            require(set(item) >= {"ordinal", "name", "sha256"},
                    "corpus catalog member identity is incomplete")
            normalized_members.append({
                "ordinal": item["ordinal"],
                "name": item["name"],
                "sha256": item["sha256"],
            })
        corpora.append({
            "id": corpus["id"],
            "archive_sha256": bytes_value["archive_sha256"],
            "members": normalized_members,
        })
    bindings = []
    for item in value["case_bindings"]:
        require(isinstance(item, dict) and set(item) >= {"case", "corpus_id", "role"},
                "corpus catalog case binding is malformed")
        bindings.append({"case": item["case"], "corpus_id": item["corpus_id"], "role": item["role"]})
    return _canonical_hash({"corpora": corpora, "case_bindings": bindings})


def _check_catalog(path: Path, report: dict[str, Any], *, control: bool) -> dict[str, Any]:
    value = _read_json(path)
    _finite(value, str(path))
    require(isinstance(value, dict), f"{path}: corpus catalog is not an object")
    required = {
        "manifest_version", "manifest_kind", "catalog_id", "canonicalization",
        "catalog_sha256", "content_set_sha256", "build", "corpora", "case_bindings",
    }
    require(set(value) == required, f"{path}: corpus catalog fields differ")
    require(value["manifest_version"] == 2 and value["manifest_kind"] == "corpus-catalog",
            f"{path}: corpus catalog identity differs")
    require(value["catalog_id"] == "litchi-perf-corpus-v2", f"{path}: catalog id differs")
    require(value["canonicalization"] == {
        "algorithm": "sorted-json-utf8-compact-v1", "hash": "sha256"
    }, f"{path}: catalog canonicalization differs")
    _hash(value["catalog_sha256"], f"{path}.catalog_sha256")
    _hash(value["content_set_sha256"], f"{path}.content_set_sha256")
    require(value["catalog_sha256"] == _catalog_hash(value),
            f"{path}: catalog hash differs from retained content")
    require(value["content_set_sha256"] == _catalog_content_hash(value),
            f"{path}: content-set hash differs from retained content")
    build = value["build"]
    require(isinstance(build, dict), f"{path}.build: build identity is missing")
    require(set(build) == {"tool", "tool_version", "git_revision", "git_worktree_dirty", "source_files"},
            f"{path}.build: build fields differ")
    require(isinstance(build.get("tool"), str) and build["tool"] == "litchi-perf-baseline",
            f"{path}.build.tool: tool identity differs")
    require(isinstance(build.get("tool_version"), str) and build["tool_version"],
            f"{path}.build.tool_version: version is missing")
    require(isinstance(build.get("source_files"), list), f"{path}.build.source_files: source list is malformed")
    for index, source in enumerate(build["source_files"]):
        require(isinstance(source, dict) and set(source) == {"path", "sha256"},
                f"{path}.build.source_files[{index}]: source identity is malformed")
        require(isinstance(source["path"], str) and source["path"],
                f"{path}.build.source_files[{index}].path: source path is missing")
        _hash(source["sha256"], f"{path}.build.source_files[{index}].sha256")
    corpora = value["corpora"]
    bindings = value["case_bindings"]
    require(isinstance(corpora, list) and isinstance(bindings, list),
            f"{path}: catalog arrays are malformed")
    expected_count = 0 if control else 1
    require(len(corpora) == expected_count and len(bindings) == expected_count,
            f"{path}: catalog corpus/binding count differs")
    if not control:
        corpus = corpora[0]
        require(isinstance(corpus, dict), f"{path}.corpora[0]: corpus is malformed")
        legacy = corpus.get("legacy_v1")
        _manifest_matches(legacy, f"{path}.corpora[0].legacy_v1")
        require(corpus.get("id") == f"docx-opc-zip:sha256:{CORPUS['archive_sha256']}",
                f"{path}.corpora[0].id: corpus identity differs")
        require(corpus.get("name") == CORPUS["name"] and corpus.get("format") == CORPUS["package_format"],
                f"{path}.corpora[0]: corpus naming differs")
        bytes_value = corpus.get("bytes")
        require(isinstance(bytes_value, dict), f"{path}.corpora[0].bytes: byte identity is missing")
        require(bytes_value.get("archive_bytes") == CORPUS["archive_bytes"] and
                bytes_value.get("archive_sha256") == CORPUS["archive_sha256"],
                f"{path}.corpora[0].bytes: archive identity differs")
        generator = corpus.get("generator")
        require(isinstance(generator, dict) and generator.get("id") == CORPUS["generator"],
                f"{path}.corpora[0].generator: generator identity differs")
        coverage = corpus.get("coverage")
        require(isinstance(coverage, dict) and CASE in coverage.get("timed_cases", []),
                f"{path}.corpora[0].coverage: timed case is missing")
        binding = bindings[0]
        require(binding == {
            "case": CASE,
            "corpus_id": corpus["id"],
            "legacy_name": CORPUS["name"],
            "legacy_archive_sha256": CORPUS["archive_sha256"],
            "role": "timed",
        }, f"{path}.case_bindings[0]: case binding differs")
    reference = report.get("corpus_catalog")
    require(isinstance(reference, dict), f"{path}: report corpus catalog reference is missing")
    require(reference == {
        "manifest_version": value["manifest_version"],
        "catalog_id": value["catalog_id"],
        "catalog_sha256": value["catalog_sha256"],
        "content_set_sha256": value["content_set_sha256"],
    }, f"{path}: report corpus catalog reference differs")
    return value


def _check_stats(stats: Any, raw: list[int], path: str) -> None:
    require(isinstance(stats, dict), f"{path}: elapsed statistics are missing")
    require(stats.get("unit") == "ns", f"{path}.unit: expected ns")
    require(stats.get("samples") == raw, f"{path}.samples: raw vector differs")
    ordered = sorted(raw)
    p50 = (ordered[len(raw) // 2 - 1] + ordered[len(raw) // 2]) // 2 if len(raw) % 2 == 0 else ordered[len(raw) // 2]
    p95 = ordered[math.ceil(0.95 * len(raw)) - 1]
    p99 = ordered[math.ceil(0.99 * len(raw)) - 1]
    require(stats.get("p50") == p50, f"{path}.p50: does not match retained raw vector")
    require(stats.get("p95") == p95, f"{path}.p95: does not match retained raw vector")
    require(stats.get("p99") == p99, f"{path}.p99: does not match retained raw vector")
    require(stats.get("min") == ordered[0] and stats.get("max") == ordered[-1],
            f"{path}: min/max do not match retained raw vector")
    mean = stats.get("mean")
    require(type(mean) in (int, float) and math.isfinite(float(mean)), f"{path}.mean: invalid")
    require(math.isclose(float(mean), statistics.fmean(raw), rel_tol=1e-9, abs_tol=1e-6),
            f"{path}.mean: does not match retained raw vector")


def _check_replay(replay: Any, path: str, *, source_bytes: int = CORPUS["archive_bytes"],
                  source_sha256: str = CORPUS["archive_sha256"],
                  aligned_tail_probe: bool = False) -> dict[str, Any]:
    require(isinstance(replay, dict), f"{path}: source replay oracle is missing")
    require(replay.get("source_bytes") == source_bytes, f"{path}.source_bytes: source drifted")
    _hash(replay.get("source_sha256"), f"{path}.source_sha256")
    require(replay.get("source_sha256") == source_sha256, f"{path}.source_sha256: source drifted")
    require(replay.get("paragraph_count") == EXPECTED_PARAGRAPHS, f"{path}.paragraph_count: oracle drifted")
    require(replay.get("operation") == "open_full_text_lifecycle", f"{path}.operation: full-text oracle missing")
    if aligned_tail_probe:
        require(replay.get("classification") == EXPECTED_ALIGNED_REPLAY_CLASSIFICATION,
                f"{path}.classification: aligned source-range oracle failed")
        probe = replay.get("aligned_eocd_tail_probe")
        require(isinstance(probe, dict), f"{path}.aligned_eocd_tail_probe: proof is missing")
        require(set(probe) == set(EXPECTED_ALIGNED_TAIL_PROBE),
                f"{path}.aligned_eocd_tail_probe: fields differ")
        for field, expected in EXPECTED_ALIGNED_TAIL_PROBE.items():
            actual = probe.get(field)
            if field == "aligned_source_sha256":
                _hash(actual, f"{path}.aligned_eocd_tail_probe.{field}")
            elif type(expected) is int:
                _nonnegative_int(actual, f"{path}.aligned_eocd_tail_probe.{field}")
            require(actual == expected,
                    f"{path}.aligned_eocd_tail_probe.{field}: value differs")
        require(probe["aligned_source_sha256"] == source_sha256 and
                probe["aligned_source_bytes"] == source_bytes,
                f"{path}.aligned_eocd_tail_probe: aligned source identity differs")
    else:
        require(replay.get("classification") == EXPECTED_REPLAY_CLASSIFICATION,
                f"{path}.classification: source-range oracle failed")
        require("aligned_eocd_tail_probe" not in replay,
                f"{path}.aligned_eocd_tail_probe: unexpected cold-only proof")
    _hash(replay.get("semantic_sha256"), f"{path}.semantic_sha256")
    require(replay["semantic_sha256"] == EXPECTED_REPLAY_SHA256, f"{path}: independent text oracle differs")
    for phase in ("open", "preparation", "query"):
        sizes = replay.get(f"{phase}_read_return_sizes")
        require(isinstance(sizes, list), f"{path}: {phase} read sizes missing")
        for size in sizes:
            _nonnegative_int(size, f"{path}: {phase} read size")
        require(replay.get(f"{phase}_read_calls") == len(sizes), f"{path}: {phase} call count differs")
        require(replay.get(f"{phase}_read_bytes") == sum(sizes), f"{path}: {phase} byte count differs")
    require(replay["query_read_calls"] == replay["query_read_bytes"] == 0,
            f"{path}: prepared full-text query performed I/O")
    require(replay.get("materializations") == 1, f"{path}: expected one main-document cache materialization")
    require(replay.get("preparation_main_payload_fully_covered") is True,
            f"{path}: preparation did not cover main payload")
    for category in ("main", "media", "unselected", "core"):
        expected_open = EXPECTED_ALIGNED_TAIL_PROBE[f"eocd_tail_probe_{category}_payload_overlap_bytes"] if aligned_tail_probe else 0
        expected_preparation = 1424 if category == "main" else 0
        for suffix in ("overlap_bytes", "covered_bytes"):
            require(replay.get(f"open_{category}_payload_{suffix}") == expected_open,
                    f"{path}: raw open {category} payload differs")
            require(replay.get(f"preparation_{category}_payload_{suffix}") == expected_preparation,
                    f"{path}: preparation {category} payload differs")
            require(replay.get(f"query_{category}_payload_{suffix}") == 0,
                    f"{path}: query {category} payload differs")
    return replay


def _text_observation(sample: dict[str, Any], path: str) -> dict[str, Any]:
    """Validate the frozen additive timed DOCX text oracle fields."""

    fields = (
        "docx_timed_full_text_sha256",
        "docx_timed_full_text_bytes",
        "docx_timed_full_text_timing_scope",
    )
    require(all(field in sample for field in fields),
            f"{path}: frozen timed full-text oracle fields are incomplete")
    digest = sample["docx_timed_full_text_sha256"]
    length = sample["docx_timed_full_text_bytes"]
    scope = sample["docx_timed_full_text_timing_scope"]
    _hash(digest, f"{path}.docx_timed_full_text_sha256")
    _positive_int(length, f"{path}.docx_timed_full_text_bytes")
    require(digest == EXPECTED_TEXT_SHA256 and length == EXPECTED_TEXT_BYTES,
            f"{path}: observed text differs from deterministic oracle")
    require(scope == EXPECTED_TEXT_SCOPE,
            f"{path}.docx_timed_full_text_timing_scope: scope drifted")
    return {"sha256": digest, "utf8_bytes": length, "scope": scope}


def _check_cold_sample(proof: Any, path: str) -> dict[str, Any]:
    require(isinstance(proof, dict), f"{path}: cold proof is missing")
    require(proof.get("status") == "eligible", f"{path}.status: cold proof is not eligible")
    tool = _read_json(ROOT / "fincore-tool.json")
    require(proof.get("filesystem_magic") == tool["filesystem_magic"], f"{path}: filesystem differs from machine")
    require(proof.get("fincore_sha256") == tool["sha256"] and proof.get("fincore_version") == tool["version"], f"{path}: fincore executable differs")
    empty_hash = hashlib.sha256(b"").hexdigest()
    for prefix in ("fincore_stderr", "fincore_version_stderr"):
        require(proof.get(prefix + "_bytes") == 0 and proof.get(prefix + "_sha256") == empty_hash, f"{path}: fincore stderr differs")
    require(proof.get("fsync_completed") is True, f"{path}.fsync_completed: sync was not proven")
    require(proof.get("advice") == EXPECTED_COLD_ADVICE, f"{path}.advice: DONTNEED was not accepted")
    for field in ("filesystem_magic", "page_size_bytes", "source_bytes", "source_pages",
                  "aligned_source_bytes", "fincore_size_bytes", "resident_bytes",
                  "dirty_bytes", "writeback_bytes", "fincore_stderr_bytes",
                  "fincore_version_stderr_bytes"):
        _nonnegative_int(proof.get(field), f"{path}.{field}")
    require(proof["page_size_bytes"] > 0 and proof["source_bytes"] > 0,
            f"{path}: source/page size must be positive")
    require(proof["source_bytes"] % proof["page_size_bytes"] == 0,
            f"{path}: source is not page aligned")
    require(proof["page_size_bytes"] == ALIGNMENT_ORACLE["page_size_bytes"],
            f"{path}: page size differs from frozen alignment oracle")
    require(proof["source_bytes"] == ALIGNMENT_ORACLE["aligned_source_bytes"],
            f"{path}: aligned source size differs from frozen alignment oracle")
    require(proof["source_pages"] == ALIGNMENT_ORACLE["aligned_source_bytes"] // ALIGNMENT_ORACLE["page_size_bytes"],
            f"{path}: aligned source page count differs")
    require(proof["aligned_source_bytes"] == proof["fincore_size_bytes"],
            f"{path}: aligned source and fincore sizes differ")
    require(proof["aligned_source_bytes"] == ALIGNMENT_ORACLE["aligned_source_bytes"],
            f"{path}: fincore source size differs from frozen alignment oracle")
    require(proof["resident_bytes"] == proof["dirty_bytes"] == proof["writeback_bytes"] == 0,
            f"{path}: final fincore residency was not zero")
    before = _nonnegative_int(proof.get("read_bytes_before"), f"{path}.read_bytes_before")
    after = _nonnegative_int(proof.get("read_bytes_after"), f"{path}.read_bytes_after")
    delta = _positive_int(proof.get("read_bytes_delta"), f"{path}.read_bytes_delta")
    require(after >= before and after - before == delta, f"{path}: read_bytes delta is inconsistent")
    _hash(proof.get("aligned_source_sha256"), f"{path}.aligned_source_sha256")
    require(proof["aligned_source_sha256"] == ALIGNMENT_ORACLE["aligned_source_sha256"],
            f"{path}: aligned source hash differs from frozen alignment oracle")
    for field in ("fincore_sha256", "fincore_stderr_sha256", "fincore_version_stderr_sha256"):
        _hash(proof.get(field), f"{path}.{field}")
    require(proof.get("fincore_tool") == "fincore", f"{path}.fincore_tool: noncanonical tool")
    require(isinstance(proof.get("fincore_version"), str) and proof["fincore_version"],
            f"{path}.fincore_version: missing")
    require(proof.get("fincore_method") == EXPECTED_FINCORE_METHOD,
            f"{path}.fincore_method: method drifted")
    require(proof.get("fincore_fallback") == EXPECTED_FINCORE_FALLBACK,
            f"{path}.fincore_fallback: fallback is not explicit none")
    return proof


def _check_allocation_sample(sample: dict[str, Any], role: str, path: str) -> None:
    allocation = sample.get("allocation_metrics")
    if role == "normal":
        require("allocation_metrics" not in sample, f"{path}: normal child emitted allocator evidence")
        return
    fields = {"allocated_bytes", "allocation_calls", "deallocated_bytes", "deallocation_calls",
              "failed_allocation_calls", "live_bytes_after", "live_bytes_before",
              "peak_live_bytes_after", "peak_live_bytes_before", "reallocation_calls",
              "region_peak_live_bytes", "scope", "status"}
    require(isinstance(allocation, dict) and set(allocation) == fields,
            f"{path}: allocator evidence missing or fields differ")
    require(allocation["status"] == "measured" and allocation["scope"] == "operation_global_system_allocator",
            f"{path}: allocator scope differs")
    for field in fields - {"status", "scope"}:
        _nonnegative_int(allocation[field], f"{path}.{field}")
    require(allocation["failed_allocation_calls"] == 0, f"{path}: allocation failed")
    require(allocation["live_bytes_before"] + allocation["allocated_bytes"] ==
            allocation["live_bytes_after"] + allocation["deallocated_bytes"], f"{path}: live bytes do not reconcile")
    require(allocation["region_peak_live_bytes"] >= max(allocation["live_bytes_before"], allocation["live_bytes_after"]),
            f"{path}: region peak below live bytes")


def _check_sample(sample: Any, index: int, state: str, path: str,
                  replay_digest: str | None = None,
                  source_bytes: int = CORPUS["archive_bytes"],
                  source_sha256: str = CORPUS["archive_sha256"],
                  aligned_tail_probe: bool = False) -> dict[str, Any]:
    require(isinstance(sample, dict), f"{path}: sample is not an object")
    require(sample.get("sample_index") == index, f"{path}.sample_index: order drifted")
    require(sample.get("cache_state") == state, f"{path}.cache_state: state drifted")
    _positive_int(sample.get("elapsed_ns"), f"{path}.elapsed_ns")
    _positive_int(sample.get("parent_wall_ns"), f"{path}.parent_wall_ns")
    _check_replay(sample.get("docx_source_replay"), f"{path}.docx_source_replay",
                  source_bytes=source_bytes, source_sha256=source_sha256,
                  aligned_tail_probe=aligned_tail_probe)
    if replay_digest is not None:
        require(sample["docx_source_replay"].get("semantic_sha256") == replay_digest,
                f"{path}: semantic oracle differs within row")
    return sample


def validate_report(report_path: Path, *, role: str, cache_state: str,
                    control: bool, samples: int = FORMAL_SAMPLES,
                    warmups: int = FORMAL_WARMUPS) -> dict[str, Any]:
    """Validate one benchmark report and return its parsed JSON object."""

    require(role in ROLES, f"unknown role {role}")
    require(cache_state in CACHE_STATES, f"unknown cache state {cache_state}")
    value = _read_json(report_path)
    _finite(value)
    require(isinstance(value, dict) and value.get("schema_version") == REPORT_SCHEMA_VERSION,
            f"{report_path}: report schema differs")
    configuration = value.get("configuration")
    require(isinstance(configuration, dict), f"{report_path}: configuration is missing")
    require(configuration.get("cases") == [CONTROL_CASE if control else CASE],
            f"{report_path}: selected case differs")
    require(configuration.get("samples_per_case") == samples and configuration.get("warmup_iterations_per_case") == warmups,
            f"{report_path}: sample contract differs")
    require(configuration.get("filesystem_cache_states") == [cache_state],
            f"{report_path}: cache-state contract differs")
    require(configuration.get("filesystem_fresh_child_per_sample") is True,
            f"{report_path}: fresh-child contract is absent")
    require(configuration.get("filesystem_process_isolated") is True,
            f"{report_path}: process-isolation contract is absent")
    require(configuration.get("filesystem_root_selected") is True,
            f"{report_path}: selected filesystem root is absent")
    tool = value.get("tool")
    require(isinstance(tool, dict), f"{report_path}: tool identity is missing")
    expected_binary = "litchi-perf-baseline" if role == "normal" else "litchi-perf-baseline-alloc"
    require(tool.get("binary") == expected_binary, f"{report_path}: tool binary identity differs")
    expected_instrumentation = "none" if role == "normal" else "system_allocator_operation_scoped"
    require(tool.get("instrumentation") == expected_instrumentation,
            f"{report_path}: allocator instrumentation identity differs")
    evidence_list = value.get("filesystem_evidence")
    require(isinstance(evidence_list, list) and len(evidence_list) == 1,
            f"{report_path}: exactly one filesystem evidence record is required")
    evidence = evidence_list[0]
    require(isinstance(evidence, dict), f"{report_path}: filesystem evidence is malformed")
    expected_case = CONTROL_CASE if control else CASE
    require(evidence.get("case") == expected_case, f"{report_path}: evidence case differs")
    _manifest_matches(evidence.get("corpus"), f"{report_path}.filesystem_evidence[0].corpus")
    require(evidence.get("warmup_iterations") == warmups and evidence.get("sample_count") == samples,
            f"{report_path}: evidence sample contract differs")
    require(evidence.get("cache_states") == [cache_state] and evidence.get("fresh_child_per_sample") is True,
            f"{report_path}: evidence cache-state contract differs")
    results = value.get("results")
    require(isinstance(results, list), f"{report_path}: results are malformed")
    if control:
        require(results == [], f"{report_path}: prepared query control emitted timed result")
        require(evidence.get("cold_verified_status") == "ineligible_prepared_query_control",
                f"{report_path}: prepared-query ineligibility is not explicit")
        require(not evidence.get("samples"), f"{report_path}: ineligible control emitted samples")
        cold_samples = evidence.get("cold_verified_samples")
        require(cold_samples is None or cold_samples == [],
                f"{report_path}: ineligible control emitted cold proof samples")
        return value
    require(len(results) == 1, f"{report_path}: exactly one selected result is required")
    result = results[0]
    require(isinstance(result, dict) and result.get("case") == CASE and result.get("cache_state") == cache_state,
            f"{report_path}: selected result identity differs")
    _manifest_matches(result.get("corpus"), f"{report_path}.results[0].corpus")
    elapsed = result.get("elapsed_ns")
    require(isinstance(elapsed, dict) and isinstance(elapsed.get("samples"), list),
            f"{report_path}: elapsed vector is missing")
    raw = [_positive_int(item, f"{report_path}.elapsed_ns.samples[{i}]") for i, item in enumerate(elapsed["samples"])]
    require(len(raw) == samples, f"{report_path}: elapsed sample count differs")
    _check_stats(elapsed, raw, f"{report_path}.results[0].elapsed_ns")
    order = elapsed.get("sample_order")
    require(isinstance(order, list) and all(type(i) is int for i in order)
            and sorted(order) == list(range(samples)),
            f"{report_path}: elapsed sample order is not a permutation")
    chronological_elapsed = dict(zip(order, raw, strict=True))
    observed = evidence.get("samples")
    require(isinstance(observed, list) and len(observed) == samples,
            f"{report_path}: evidence sample count differs")
    checked_proofs: list[dict[str, Any]] | None = None
    aligned_identity: tuple[str, int] | None = None
    if cache_state == "cold-verified":
        require(evidence.get("cold_verified_status") == "eligible",
                f"{report_path}: verified-cold eligibility status is missing")
        proofs = evidence.get("cold_verified_samples")
        require(isinstance(proofs, list) and len(proofs) == samples,
                f"{report_path}: verified-cold proof count differs")
        checked_proofs = []
        for index, proof in enumerate(proofs):
            checked = _check_cold_sample(proof, f"{report_path}.cold_verified_samples[{index}]")
            checked_proofs.append(checked)
            current = (checked["aligned_source_sha256"], checked["aligned_source_bytes"])
            if aligned_identity is None:
                aligned_identity = current
            require(current == aligned_identity, f"{report_path}: aligned cold source changed")
        require(aligned_identity == (
            ALIGNMENT_ORACLE["aligned_source_sha256"],
            ALIGNMENT_ORACLE["aligned_source_bytes"],
        ), f"{report_path}: aligned cold identity is not the frozen oracle")
        require(evidence.get("cold_verified_claim_scope") == EXPECTED_COLD_SCOPE,
                f"{report_path}: cold claim scope differs")
        require(evidence.get("cold_verified_fincore_command") == EXPECTED_FINCORE_COMMAND,
                f"{report_path}: fincore command differs")
    replay_digest: str | None = None
    text_identity: tuple[str, int] | None = None
    for index, sample in enumerate(observed):
        source_bytes = CORPUS["archive_bytes"]
        source_sha256 = CORPUS["archive_sha256"]
        if aligned_identity is not None:
            source_bytes, source_sha256 = aligned_identity[1], aligned_identity[0]
        item = _check_sample(sample, index, cache_state,
                             f"{report_path}.filesystem_evidence[0].samples[{index}]",
                             replay_digest, source_bytes, source_sha256,
                             aligned_tail_probe=cache_state == "cold-verified")
        _check_allocation_sample(item, role, f"{report_path}.sample[{index}]")
        replay = item["docx_source_replay"]
        digest = replay["semantic_sha256"]
        if replay_digest is None:
            replay_digest = digest
        text = _text_observation(item,
                                 f"{report_path}.filesystem_evidence[0].samples[{index}].observed_text")
        identity = (text["sha256"], text["utf8_bytes"])
        if text_identity is None:
            text_identity = identity
        require(identity == text_identity, f"{report_path}: timed text oracle changed between samples")
        require(hashlib.sha256(bytes.fromhex(text["sha256"])).hexdigest() == digest,
                f"{report_path}: timed text digest differs from independent semantic oracle")
        require(item["elapsed_ns"] == chronological_elapsed[index], f"{report_path}: evidence/result elapsed vectors differ")
    require(replay_digest is not None and text_identity is not None,
            f"{report_path}: text oracle is absent")
    if cache_state == "warm":
        require(evidence.get("cold_verified_status") is None and not evidence.get("cold_verified_samples"),
                f"{report_path}: warm report contains verified-cold evidence")
        for index, sample in enumerate(observed):
            require(sample.get("cold_advice") == "not_requested",
                    f"{report_path}: warm sample {index} requested cache eviction")
            require(sample.get("cold_verified") is None,
                    f"{report_path}: warm sample contains cold proof")
    elif cache_state == "cold-requested":
        require(evidence.get("cold_verified_status") is None and not evidence.get("cold_verified_samples"),
                f"{report_path}: cold-requested report contains verified-cold evidence")
        for index, sample in enumerate(observed):
            require(sample.get("cold_advice") == "requested",
                    f"{report_path}: cold-requested sample {index} did not record an accepted request")
            require(sample.get("cold_verified") is None,
                    f"{report_path}: cold-requested sample contains verified proof")
    else:
        require(checked_proofs is not None, f"{report_path}: cold proof validation was skipped")
        for index, (sample, checked) in enumerate(zip(observed, checked_proofs)):
            sample_proof = sample.get("cold_verified")
            require(sample_proof == checked, f"{report_path}: sample/top-level cold proof differs")
    return value


def _resource_values(path: Path) -> dict[str, Any]:
    if not path.is_file():
        return {}
    text = path.read_text(encoding="utf-8", errors="replace")
    result: dict[str, Any] = {"raw_sha256": sha(path), "raw_bytes": len(text.encode())}
    for line in text.splitlines():
        if ":" not in line:
            continue
        key, raw = line.split(":", 1)
        key = key.strip().lower().replace(" ", "_")
        raw = raw.strip()
        if raw.isdigit():
            result[key] = int(raw)
        else:
            try:
                result[key] = float(raw.split()[0])
            except (ValueError, IndexError):
                result[key] = raw
    return result


def _vectors(evidence: dict[str, Any]) -> dict[str, list[int]]:
    samples = evidence.get("samples") or []
    keys = ("elapsed_ns", "parent_wall_ns", "logical_read_calls", "logical_read_bytes")
    values = {key: [sample[key] for sample in samples] for key in keys}
    for key in ("read_bytes_delta", "rss_bytes", "peak_rss_bytes", "region_peak_live_bytes"):
        vector: list[int] = []
        present = True
        for sample in samples:
            if key == "read_bytes_delta":
                proof = sample.get("cold_verified")
                item = proof.get(key) if isinstance(proof, dict) else None
            else:
                process = sample.get("process_metrics") or {}
                allocation = sample.get("allocation_metrics") or {}
                item = process.get(key) if key in process else allocation.get(key)
            if type(item) is not int:
                present = False
                break
            vector.append(item)
        if present and vector:
            values[key] = vector
    if samples and all(isinstance(sample.get("allocation_metrics"), dict) for sample in samples):
        for field in ("allocation_calls", "reallocation_calls", "allocated_bytes", "deallocated_bytes"):
            values["allocation_" + field] = [sample["allocation_metrics"][field] for sample in samples]
        values["allocation_peak_increment_bytes"] = [sample["allocation_metrics"]["region_peak_live_bytes"] - sample["allocation_metrics"]["live_bytes_before"] for sample in samples]
    return values


def _percentiles(raw: list[int]) -> dict[str, Any]:
    ordered = sorted(raw)
    p50 = (ordered[len(raw) // 2 - 1] + ordered[len(raw) // 2]) // 2 if len(raw) % 2 == 0 else ordered[len(raw) // 2]
    return {
        "n": len(raw),
        "min": ordered[0],
        "max": ordered[-1],
        "mean": statistics.fmean(raw),
        "p50": p50,
        "p95": ordered[math.ceil(0.95 * len(raw)) - 1],
        "p99": ordered[math.ceil(0.99 * len(raw)) - 1],
    }


def _confined_artifact(directory: Path, name: str, value: Any, *, required_bytes: bool = False) -> dict[str, Any]:
    require(name in CAPTURE_ARTIFACTS, f"unknown capture artifact: {name}")
    require(isinstance(value, dict), f"{directory}/terminal.json: {name} hash is missing")
    expected = directory / name
    artifact_path = Path(str(value.get("path", "")))
    require(artifact_path == expected, f"{directory}: {name} path is outside its capture directory")
    actual = _file_meta(expected)
    require(actual["sha256"] == value.get("sha256") and actual["bytes"] == value.get("bytes"),
            f"{directory}: {name} changed after capture")
    if required_bytes:
        require(actual["bytes"] > 0, f"{directory}: {name} is empty")
    return actual


def _check_label_set(root: Path, expected_labels: Iterable[str]) -> None:
    require(root.is_dir() and not root.is_symlink(), f"capture attempt is missing: {root}")
    actual = sorted(item.name for item in root.iterdir())
    expected = sorted(expected_labels)
    require(actual == expected, f"{root}: on-disk capture labels differ from frozen inventory")
    for item in root.iterdir():
        require(item.is_dir() and not item.is_symlink(),
                f"{root}: capture label is not a regular directory: {item.name}")


def _capture_header(spec: dict[str, Any], directory: Path, started: dict[str, Any],
                    terminal: dict[str, Any], build: dict[str, Any]) -> tuple[datetime.datetime, datetime.datetime]:
    expected_binary = build["binary"]
    expected_argv = _command(
        spec,
        expected_binary,
        directory / "report.json",
        directory / "corpus-manifest.json",
        _run_root(spec["attempt"], spec["label"]) / "filesystem",
    )
    expected = {
        "run": {key: value for key, value in spec.items() if key != "attempt"},
        "attempt": spec["attempt"],
        "build": build["binding"],
        "binary": expected_binary,
        "source": build["source"],
        "tool": _tool_identity(spec["role"]),
        "machine": _machine_binding(),
        "protocol": _protocol_binding(),
        "argv": expected_argv,
        "cwd": str(REPO),
        "environment": _environment(),
        "driver_sha256": _json_hash(Path(__file__)),
        "support_sha256": _json_hash(ROOT / "support.py"),
    }
    require(started.get("schema") == SCHEMA and started.get("version") == 1,
            f"{directory}/started.json: receipt schema differs")
    require(terminal.get("schema") == SCHEMA and terminal.get("version") == 1,
            f"{directory}/terminal.json: receipt schema differs")
    require(started.get("status") == "running", f"{directory}/started.json: status differs")
    require(terminal.get("status") == "pass", f"{directory}/terminal.json: status is not pass")
    for receipt, name in ((started, "started"), (terminal, "terminal")):
        for key, expected_value in expected.items():
            require(receipt.get(key) == expected_value,
                    f"{directory}/{name}.json: {key} binding differs")
    require(terminal.get("exit_code") == 0, f"{directory}/terminal.json: process exit was not zero")
    require(terminal.get("timed_out") is False, f"{directory}/terminal.json: process timed out")
    require(terminal.get("termination") is None, f"{directory}/terminal.json: process was terminated")
    require(started.get("process_group") == {"start_new_session": True},
            f"{directory}/started.json: process-group isolation differs")
    require(type(terminal.get("process_id")) is int and terminal["process_id"] > 0,
            f"{directory}/terminal.json: process id is missing")
    require(terminal.get("process_group_id") == terminal.get("process_id"),
            f"{directory}/terminal.json: process-group id differs")
    require(terminal.get("missing_artifacts") == [],
            f"{directory}/terminal.json: missing artifact list is non-empty")
    require("launch_error" not in terminal and "validation_error" not in terminal,
            f"{directory}/terminal.json: failure detail is present")
    require(terminal.get("started_utc") == started.get("started_utc"),
            f"{directory}/terminal.json: start timestamp differs")
    started_at = _timestamp(started.get("started_utc"), f"{directory}/started.json.started_utc")
    finished_at = _timestamp(terminal.get("finished_utc"), f"{directory}/terminal.json.finished_utc")
    require(finished_at > started_at, f"{directory}: terminal finished before capture started")
    require(terminal.get("replay_root") == str(_run_root(spec["attempt"], spec["label"])),
            f"{directory}/terminal.json: replay root differs")
    return started_at, finished_at


def _collect(attempt: str, builds: dict[str, dict[str, Any]], *, pilot: bool = False) -> list[dict[str, Any]]:
    _attempt(attempt)
    expected = [{**spec, "attempt": attempt} for spec in formal_inventory(pilot=pilot)]
    root = ROOT / "captures" / attempt
    require(root.is_dir() and not root.is_symlink(), f"capture attempt is missing: {root}")
    expected_labels = [item["label"] for item in expected]
    _check_label_set(root, expected_labels)
    entries: list[dict[str, Any]] = []
    previous_finished: datetime.datetime | None = None
    for spec in expected:
        directory = root / spec["label"]
        require(directory.is_dir() and not directory.is_symlink(),
                f"missing capture directory: {directory}")
        actual_names = {item.name for item in directory.iterdir()}
        required_names = set(CAPTURE_ARTIFACTS) | {"started.json", "terminal.json"}
        require(actual_names == required_names,
                f"{directory}: retained artifact inventory differs")
        for item in directory.iterdir():
            require(item.is_file() and not item.is_symlink(),
                    f"{directory}: unexpected non-regular artifact {item.name}")
        started_path = directory / "started.json"
        terminal_path = directory / "terminal.json"
        started = _read_json(started_path)
        require(terminal_path.is_file(), f"missing terminal receipt: {terminal_path}")
        terminal = _read_json(terminal_path)
        _finite(started, str(started_path))
        _finite(terminal, str(terminal_path))
        build = builds[spec["role"]]
        started_at, finished_at = _capture_header(spec, directory, started, terminal, build)
        if previous_finished is not None:
            require(started_at >= previous_finished,
                    f"{directory}: declared capture chronology overlaps or is out of order")
        previous_finished = finished_at
        started_meta = _file_meta(started_path)
        require(terminal.get("started_artifact") == started_meta,
                f"{terminal_path}: started receipt hash differs")
        artifacts = terminal.get("artifacts")
        require(isinstance(artifacts, dict), f"{terminal_path}: artifact hashes are missing")
        require(set(artifacts) == set(CAPTURE_ARTIFACTS),
                f"{terminal_path}: artifact inventory differs")
        for name in CAPTURE_ARTIFACTS:
            _confined_artifact(directory, name, artifacts[name],
                               required_bytes=name in {"resource.txt", "report.json", "corpus-manifest.json"})
        report_path = Path(str(artifacts["report.json"]["path"]))
        report = validate_report(report_path, role=spec["role"], cache_state=spec["cache_state"],
                                 control=spec["kind"] == "control", samples=spec["samples"], warmups=spec["warmups"])
        _check_report_build(report, build)
        checked_catalog = _check_catalog(directory / "corpus-manifest.json", report, control=spec["kind"] == "control")
        require(checked_catalog["build"]["git_revision"] == build["git_revision"], "catalog revision differs from build")
        entries.append({"spec": spec, "directory": directory, "terminal": terminal,
                        "terminal_sha256": _json_hash(terminal_path), "report": report,
                        "report_sha256": _json_hash(report_path),
                        "resource": _resource_values(directory / "resource.txt"),
                        "started_at": started_at.isoformat(),
                        "finished_at": finished_at.isoformat()})
    return entries


def analyze_data(entries: list[dict[str, Any]], builds: dict[str, dict[str, Any]], *, pilot: bool = False) -> dict[str, Any]:
    rows: list[dict[str, Any]] = []
    adverse: list[dict[str, Any]] = []
    controls: list[dict[str, Any]] = []
    for entry in entries:
        spec = entry["spec"]
        evidence = entry["report"]["filesystem_evidence"][0]
        base = {
            "label": spec["label"],
            "kind": spec["kind"],
            "role": spec["role"],
            "repeat": spec["repeat"],
            "cache_state": spec["cache_state"],
            "case": spec["case"],
            "samples": spec["samples"],
            "warmups": spec["warmups"],
            "terminal_sha256": entry["terminal_sha256"],
            "report_sha256": entry["report_sha256"],
            "resource": entry["resource"],
        }
        if spec["kind"] == "control":
            base["status"] = evidence.get("cold_verified_status")
            controls.append(base)
            continue
        vectors = _vectors(evidence)
        percentiles = {name: _percentiles(values) for name, values in vectors.items()}
        base["raw_vectors"] = vectors
        base["percentiles"] = percentiles
        rows.append(base)
        elapsed = vectors["elapsed_ns"]
        p50 = percentiles["elapsed_ns"]["p50"]
        if p50 > 0:
            for index, value in enumerate(elapsed):
                if value > p50 * 2:
                    adverse.append({
                        "label": spec["label"],
                        "kind": "tail_spread_diagnostic",
                        "metric": "elapsed_ns",
                        "sample_index": index,
                        "value": value,
                        "row_p50": p50,
                        "ratio": value / p50,
                    })
    return {
        "schema": SCHEMA + "-analysis",
        "version": 1,
        "claim_authorized": False,
        "performance_claim": "none",
        "scope": "standalone cache-state/provider evidence; comparison unavailable",
        "pilot": pilot,
        "corpus": dict(CORPUS),
        "case": CASE,
        "prepared_query_control": CONTROL_CASE,
        "roles": list(ROLES),
        "repeats": list((1,) if pilot else REPEATS),
        "cache_states": list(CACHE_STATES),
        "samples": PILOT_SAMPLES if pilot else FORMAL_SAMPLES,
        "warmups": PILOT_WARMUPS if pilot else FORMAL_WARMUPS,
        "builds": {
            role: {
                "receipt_sha256": builds[role]["receipt_sha256"],
                "binary": builds[role]["binary"],
                "source": builds[role]["source"],
            }
            for role in ROLES
        },
        "rows": rows,
        "controls": controls,
        "adverse_rows": adverse,
        "inventory": {
            "formal_processes": len(entries),
            "measured_rows": len(rows),
            "control_processes": len(controls),
            "measured_samples": sum(len(row["raw_vectors"]["elapsed_ns"]) for row in rows),
        },
    }


def analyze(attempt: str, build_dir: Path | None = None, *, pilot: bool = False) -> Path:
    builds = load_builds(build_dir)
    entries = _collect(attempt, builds, pilot=pilot)
    summary = analyze_data(entries, builds, pilot=pilot)
    directory = ROOT / "analysis"
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / f"{attempt}{'-pilot' if pilot else ''}.json"
    _write_json(path, summary)
    return path


def verify(attempt: str, build_dir: Path | None = None, *, pilot: bool = False) -> Path:
    """Recompute analysis solely from hashed retained raw receipts/reports."""

    builds = load_builds(build_dir)
    entries = _collect(attempt, builds, pilot=pilot)
    recomputed = analyze_data(entries, builds, pilot=pilot)
    summary_path = ROOT / "analysis" / f"{attempt}{'-pilot' if pilot else ''}.json"
    require(summary_path.is_file(), f"analysis summary is missing: {summary_path}")
    retained = _read_json(summary_path)
    require(retained == recomputed, "analysis summary differs from recomputation over retained raw files")
    verification = {
        "schema": SCHEMA + "-verification",
        "version": 1,
        "status": "pass",
        "attempt": attempt,
        "pilot": pilot,
        "summary_sha256": _json_hash(summary_path),
        "raw_receipts": [
            {"label": entry["spec"]["label"], "terminal_sha256": entry["terminal_sha256"], "report_sha256": entry["report_sha256"]}
            for entry in entries
        ],
        "verified_utc": now(),
    }
    directory = ROOT / "verification"
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / f"{attempt}{'-pilot' if pilot else ''}.json"
    _write_json(path, verification)
    return path


def capture_one(attempt: str, role: str, repeat: int, cache_state: str,
                *, control: bool = False, pilot: bool = False,
                build_dir: Path | None = None,
                timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS) -> Path:
    require(role in ROLES, f"unknown role {role}")
    require(repeat in ((1,) if pilot else REPEATS), f"invalid repeat {repeat}")
    require(cache_state in CACHE_STATES, f"unknown cache state {cache_state}")
    if control:
        require(cache_state == "cold-verified", "prepared control must select cold-verified")
        case = CONTROL_CASE
        label = f"{'pilot-' if pilot else ''}r{repeat}-{role}-prepared-control"
        kind = "control"
    else:
        case = CASE
        label = f"{'pilot-' if pilot else ''}r{repeat}-{role}-{cache_state}"
        kind = "pilot" if pilot else "formal"
    spec = {
        "attempt": _attempt(attempt),
        "kind": kind,
        "repeat": repeat,
        "role": role,
        "cache_state": cache_state,
        "case": case,
        "samples": PILOT_SAMPLES if pilot else FORMAL_SAMPLES,
        "warmups": PILOT_WARMUPS if pilot else FORMAL_WARMUPS,
        "label": label,
    }
    builds = load_builds(build_dir)
    return _launch(spec, builds[role], timeout_seconds)


def capture_all(attempt: str, build_dir: Path | None = None, *, pilot: bool = False,
                timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS) -> None:
    builds = load_builds(build_dir)
    for spec in formal_inventory(pilot=pilot):
        # The inventory is generated in role order; launch serially while the
        # coordinator's shared CPU lock is held by the caller/gate.
        _launch({**spec, "attempt": _attempt(attempt)}, builds[spec["role"]], timeout_seconds)


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("plan", help="print the frozen matrix without writing it")
    sub.add_parser("freeze-protocol", help="write cold-protocol.json exactly once")
    for command, help_text in (("capture", "capture one immutable process"), ("capture-all", "capture the complete matrix")):
        item = sub.add_parser(command, help=help_text)
        item.add_argument("--attempt", required=True)
        item.add_argument("--build-dir", type=Path)
        item.add_argument("--timeout-seconds", type=int, default=DEFAULT_TIMEOUT_SECONDS)
        item.add_argument("--pilot", action="store_true")
        if command == "capture":
            item.add_argument("--role", choices=ROLES, required=True)
            item.add_argument("--repeat", type=int, choices=REPEATS, required=True)
            item.add_argument("--cache-state", choices=CACHE_STATES, required=True)
            item.add_argument("--control", action="store_true")
    for command, help_text in (("analyze", "analyze retained raw reports"), ("verify", "recompute and verify retained raw reports")):
        item = sub.add_parser(command, help=help_text)
        item.add_argument("--attempt", required=True)
        item.add_argument("--build-dir", type=Path)
        item.add_argument("--pilot", action="store_true")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        if args.command == "plan":
            print(json.dumps(protocol_value(), indent=2, sort_keys=True))
        elif args.command == "freeze-protocol":
            print(_freeze_protocol())
        elif args.command == "capture":
            capture_one(args.attempt, args.role, args.repeat, args.cache_state,
                        control=args.control, pilot=args.pilot, build_dir=args.build_dir,
                        timeout_seconds=args.timeout_seconds)
        elif args.command == "capture-all":
            capture_all(args.attempt, args.build_dir, pilot=args.pilot,
                        timeout_seconds=args.timeout_seconds)
        elif args.command == "analyze":
            print(analyze(args.attempt, args.build_dir, pilot=args.pilot))
        elif args.command == "verify":
            print(verify(args.attempt, args.build_dir, pilot=args.pilot))
        else:  # pragma: no cover - argparse enforces this
            fail(f"unknown command {args.command}")
    except (ColdMatrixError, OSError, ValueError, subprocess.SubprocessError) as error:
        print(f"cold_matrix.py: FAIL: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
