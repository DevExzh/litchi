#!/usr/bin/env python3
"""Capture and audit the 0495 DOCX opened-document edit/provider matrix.

The Rust target owns one opened-document paragraph edit, commit, and
sequential publication.  This driver owns the process boundary and makes the
evidence replayable: a frozen protocol binds the source manifest, successful
build gates, retained executables, host environment, and every helper; each
child retains raw output, GNU time RSS, and immutable terminal custody.
``analyze`` and ``verify`` only consume those retained files.

The Rust report contract is intentionally small and explicit.  Each report
has one row per measured edit/save iteration.  The row keeps the edit/output
oracle, source and sink summaries, cache snapshots, and optional
allocator sample together.  The timed lifecycle is the opened source,
paragraph edit, commit, and sequential sink publication; setup, oracle
hashing, and report serialization remain outside it.
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
CHANGE = 495
SCHEMA = "docx_edit_provider_managed_v1"
PROTOCOL_SCHEMA = "docx-edit-provider-managed-protocol-v1"
CAPTURE_SCHEMA = "docx-edit-provider-managed-capture-v1"
TERMINAL_SCHEMA = "docx-edit-provider-terminal-v1"
ANALYSIS_SCHEMA = "docx-edit-provider-managed-analysis-v1"
VERIFICATION_SCHEMA = "docx-edit-provider-managed-verification-v1"
CASE = "docx_opened_document_managed_vs_unmanaged_one_paragraph_edit_save"
CPU = 2
CPU_LOCK = "/home/zhuhe/.cache/litchi-goal-0484/cpu.lock"
ROLES = ("normal", "allocator")
PHASES = ("before", "after")
APIS = ("unmanaged-api", "managed-api")
REPEATS = (1, 2)
FORMAL_SAMPLES = 30
FORMAL_WARMUPS = 3
PILOT_SAMPLES = 3
PILOT_WARMUPS = 1
WINDOW_BYTES = 4096
MAX_RANGE_BYTES = 65_536
DEFAULT_TIMEOUT_SECONDS = 180
MAX_TIMEOUT_SECONDS = 3600
TERM_GRACE_SECONDS = 10
REQUEST_BUCKETS = 18
SAMPLE_ALLOCATION_SCOPE = "operation_global_system_allocator"
PATCH_ORACLE_SCOPE = "untimed_preflight_commit_patch_oracles"
UNMANAGED_API = "unmanaged-api"
MANAGED_API = "managed-api"
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

TIMING_SCOPE = (
    "fresh source-backed DOCX open + one paragraph edit staging/commit + sequential "
    "publication + commit/package/document drops; provider construction, sink reservation "
    "and retained output are outside; commit diagnostics and source/candidate XML identity "
    "comparisons are inside, while output hashing, semantic/media verification and preflight "
    "patch oracles are outside"
)
SETUP_SCOPE = (
    "deterministic corpus construction, expected publication, source adapter "
    "construction, warm file staging, sink reservation, and oracle preparation are "
    "outside the clock"
)
ALLOCATION_SCOPE = (
    "operation-scoped global-system-allocator region surrounds the timed opened "
    "document edit, commit, publication, diagnostics, and package/document drop; "
    "normal binaries report unavailable and allocator binaries report the separate "
    "region sample"
)
PROVIDER_SCOPE = (
    "explicit caller-owned positional provider matrix for one DOCX paragraph "
    "replacement; no ambient filesystem or network behavior"
)
FILE_SCOPE = (
    "recently-written FileSource reopened for each iteration; warm-cache observation "
    "only, with no cold-filesystem claim"
)
PHYSICAL_SCOPE = (
    "caller-visible ReadAt calls at the selected provider/source boundary; requested "
    "and returned ranges are transport observations, not filesystem or network I/O"
)
RANGE_SCOPE = (
    "logical ReadAt calls and adapter calls observed by harness wrappers; zero-length "
    "caller calls are counted separately; counters are transport-model evidence, not "
    "physical network or filesystem observations"
)
ZERO_LENGTH_SCOPE = (
    "zero-length caller ReadAt calls are delegated to the wrapped provider, preserve its "
    "return/error behavior, and are counted separately from nonempty range totals"
)
BUDGET_SCOPE = (
    "managed budget gauges are sampled before package open, after publication while the "
    "commit remains live, and after package/document/commit drops; cumulative "
    "input/output/work counters are not release gauges"
)

# These are the limits emitted by the Rust provider report.  They are kept in
# the frozen protocol as well so a report cannot silently move to a different
# resource envelope between the pilot and formal matrix.
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

UNMANAGED_REASON = "unmanaged-api uses the compatibility constructor without an ExecutionContext"
MANAGED_REASON = "managed-api uses the explicit finite ExecutionContext"
BUDGET_LIMITS = {
    "memory_bytes": 64 * 1024 * 1024,
    "input_bytes": 64 * 1024 * 1024,
    "output_bytes": 64 * 1024 * 1024,
    "objects": 1_000_000,
    "depth": 1_024,
    "work": 2 * 1024 * 1024 * 1024,
}

BASE_LIMITS = {
    "read_limits": READ_LIMITS,
    "cache_max_bytes": 8 * 1024 * 1024,
    "cache_max_entries": 128,
    "max_tracked_ranges": 131_072,
    "sink_max_write_bytes": 1_048_576,
}


def _limits(api: str) -> dict[str, Any]:
    managed = api == MANAGED_API
    value = dict(BASE_LIMITS)
    value["resource_budget"] = {
        "managed": managed,
        "mode_reason": MANAGED_REASON if managed else UNMANAGED_REASON,
        **({key: BUDGET_LIMITS[key] for key in BUDGET_LIMITS} if managed
           else {key: None for key in BUDGET_LIMITS}),
    }
    return value


LIMITS = _limits(UNMANAGED_API)

# The protocol also bounds the driver itself.  These values describe accepted
# CLI request sizes rather than the provider's report-level safety limits.
DRIVER_LIMITS = {
    "max_samples": 10_000,
    "max_warmup": 1_000,
    "max_traced_ranges": LIMITS["max_tracked_ranges"],
}

EXPECTED_OUTPUT_BYTES = 16_793_048
EXPECTED_OUTPUT_SHA256 = "9af99bf7f63aac1ffc13ff59a5038703c96229b906b3db1602d5af5590171795"

CORPUS_MANIFEST = {
    "name": "docx-source-backed-media",
    "generator": CORPUS["generator"],
    "package_format": "DOCX/OPC/ZIP",
    "shape": "media-rich",
    "payload_kind": "deterministic-incompressible-media",
    "compression": "deflate",
    "entry_count": 17,
    "archive_member_count": CORPUS["archive_member_count"],
    "entry_bytes": CORPUS["media_member_bytes"],
    "uncompressed_payload_bytes": 16_833_643,
    "archive_bytes": CORPUS["archive_bytes"],
    "archive_sha256": CORPUS["archive_sha256"],
    "target_entry": "paragraph:100",
    "target_payload_bytes": 50,
    "target_payload_sha256": "08ea83c6e3287e74125f5a70c3ce3a4bdb6de13b7930312a5059060e24e56c14",
    "xlsx": None,
}

REPORT_FIELDS = (
    "schema", "version", "case_name", "benchmark", "api", "api_name",
    "provider_scope", "timing_scope",
    "setup_scope", "allocation_scope", "physical_scope", "range_scope", "zero_length_scope",
    "file_scope", "budget_scope",
    "corpus_version", "corpus_generator", "source_archive_sha256", "source_archive_bytes",
    "source_bytes", "source_sha256", "expected_text_bytes", "expected_text_sha256",
    "expected_text_scope", "expected_archive_members", "expected_output_sha256",
    "expected_output_bytes", "source_revision", "requested_source_revision", "limits",
    "binary_sha256", "binary_bytes", "current_exe", "provider", "corpus", "preflight",
    "warmup", "samples", "allocator", "instrumentation", "rows",
)

READ_COUNTER_FIELDS = (
    "availability", "scope", "calls", "empty_calls", "requested_bytes", "returned_bytes",
    "short_reads", "min_request_bytes", "max_request_bytes", "traced_ranges",
    "ranges", "requested_media_overlap_bytes", "returned_media_overlap_bytes",
)

RESOURCE_FIELDS = (
    "availability", "managed", "memory_used", "input_bytes_used", "output_bytes_used",
    "objects_used", "depth_used", "work_used",
)

CACHE_BUDGET_FIELDS = (
    "availability", "managed", "budget_reservation_failures", "budget_memory_used",
    "budget_cache_reserved_bytes", "budget_input_bytes_used", "budget_output_bytes_used",
    "budget_work_used", "budget_objects_used", "budget_catalog_reserved_objects",
    "budget_cache_reserved_objects", "retained_entries", "retained_bytes", "in_flight_loads",
)

BUDGET_FIELDS = (
    "scope", "before", "live", "after_drop", "cache_before", "cache_live",
    "memory_released_to_baseline", "objects_released_to_baseline", "reservation_failures",
)

PREFLIGHT_FIELDS = (
    "expected_materializations", "source_document_xml_sha256",
    "candidate_document_xml_sha256", "output_exact_source_changed",
    "semantic_reopen_verified", "unchanged_media_preserved",
    "replay_forward_verified", "inverse_restores_source_verified",
    "stale_target_refusal_verified", "foreign_source_refusal_verified",
)

RANGE_ADAPTER_FIELDS = (
    "availability", "logical_calls", "requested_bytes", "returned_bytes", "short_reads",
    "delayed_calls", "transfer_paced_calls", "transfer_delay_ns",
)

PROVIDER_FIELDS = (
    "name", "provider", "kind", "trace_ranges", "short_read_bytes", "max_range_bytes",
    "zero_length_scope", "delay_us", "transfer_bytes_per_second", "transfer_delay_policy",
    "source_construction", "file_scope", "range_scope",
)

ARMS = (
    {"name": "owned", "provider": "owned", "kind": "owned", "trace_ranges": False,
     "max_range_bytes": None, "delay_us": None, "transfer_bytes_per_second": None,
     "transfer_delay_policy": "separate-sleeps", "short_read_bytes": None},
    {"name": "instrumented", "provider": "instrumented", "kind": "instrumented", "trace_ranges": True,
     "max_range_bytes": None, "delay_us": None, "transfer_bytes_per_second": None,
     "transfer_delay_policy": "separate-sleeps", "short_read_bytes": None},
    {"name": "file-warm", "provider": "file", "kind": "file-warm", "trace_ranges": False,
     "max_range_bytes": None, "delay_us": None, "transfer_bytes_per_second": None,
     "transfer_delay_policy": "separate-sleeps", "short_read_bytes": None},
    {"name": "short", "provider": "short", "kind": "short-read", "trace_ranges": True,
     "max_range_bytes": 4_096, "delay_us": None, "transfer_bytes_per_second": None,
     "transfer_delay_policy": "separate-sleeps", "short_read_bytes": 4_096},
    {"name": "delayed", "provider": "delayed", "kind": "range-delay", "trace_ranges": True,
     "max_range_bytes": 65_536, "delay_us": 1_000, "transfer_bytes_per_second": 104_857_600,
     "transfer_delay_policy": "minimum-service", "short_read_bytes": None},
    {"name": "range-zero", "provider": "delayed", "kind": "range-control", "trace_ranges": True,
     "max_range_bytes": 65_536, "delay_us": 0, "transfer_bytes_per_second": 104_857_600,
     "transfer_delay_policy": "minimum-service", "short_read_bytes": None},
)
ARM_BY_NAME = {arm["name"]: arm for arm in ARMS}
DRIVER_FILES = ("measure.py", "test_measure.py", "support.py", "gate.py", "retain_build.py")
HARNESS_FILES = (
    "tools/perf-baseline/src/docx_managed_edit.rs",
    "tools/perf-baseline/src/lib.rs",
    "tools/perf-baseline/src/main.rs",
    "tools/perf-baseline/src/bin/litchi-perf-baseline-alloc.rs",
)

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
    # The retained copy is immutable custody.  The original Cargo target is
    # historical metadata only and may be overwritten by the next phase's
    # build, so never re-stat or re-hash it during later analysis.
    if retained:
        actual = _file_meta(path, executable=True)
        require(actual["bytes"] == value["bytes"] and actual["sha256"] == value["sha256"],
                f"{owner}: artifact identity changed")
    elif path.exists():
        require(path.is_file() and not path.is_symlink(),
                f"{owner}: historical original path is not a regular file")
    return {"path": str(path), "bytes": value["bytes"], "sha256": value["sha256"], "executable": True}


def _expected_build_command(role: str) -> list[str]:
    require(role in ROLES, f"unknown build role {role}")
    name = "litchi-perf-baseline" + ("-alloc" if role == "allocator" else "")
    command = ["cargo", "build", "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml"]
    if role == "allocator":
        command += ["--features", "allocator-metrics"]
    return command + ["--bin", name]


def _authenticated_archived_helpers(driver_sha256: str, common_sha256: str,
                                    owner: Path) -> None:
    """Require an exact archived helper pair for a historical gate receipt.

    The gate helper remains shared across attempts, so an archived pair may
    differ only in ``support.py``.  Both files are still re-hashed from the
    archive; a receipt cannot authorize an arbitrary historical digest merely
    because it has the right development label.
    """

    current_driver = _json_hash(ROOT / "gate.py")
    require(driver_sha256 == current_driver,
            f"{owner}: archived gate helper differs")
    for archive in sorted(ROOT.glob("development-*")):
        if not archive.is_dir() or archive.is_symlink():
            continue
        gate = archive / "gate.py"
        support = archive / "support.py"
        if (not gate.is_file() or gate.is_symlink()
                or not support.is_file() or support.is_symlink()):
            continue
        if (_json_hash(gate) == driver_sha256
                and _json_hash(support) == common_sha256):
            return
    require(False, f"{owner}: archived helper pair is not authenticated")


def _gate_binding(value: Any, owner: Path, *, allow_archived_helpers: bool = False) -> dict[str, Any]:
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
    require(receipt["schema"] == "docx-edit-provider-gate-v1", f"{path}: gate schema differs")
    require(receipt["exit_code"] == 0 and receipt["timed_out"] is False
            and receipt["termination"] is None and receipt["source_unchanged"] is True,
            f"{path}: gate did not pass unchanged")
    require(receipt["cwd"] == str(REPO), f"{path}: gate cwd differs")
    require(receipt["environment"] == {key: ENV[key] for key in ENV_KEYS}, f"{path}: gate environment differs")
    current_driver = _json_hash(ROOT / "gate.py")
    current_common = _json_hash(ROOT / "support.py")
    if receipt["driver_sha256"] == current_driver and receipt["common_sha256"] == current_common:
        pass
    elif allow_archived_helpers:
        _authenticated_archived_helpers(receipt["driver_sha256"],
                                         receipt["common_sha256"], path)
    else:
        require(receipt["driver_sha256"] == current_driver,
                f"{path}: gate helper changed")
        require(receipt["common_sha256"] == current_common,
                f"{path}: support helper changed")
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


def _build_from(value: Any, phase: str, role: str, path: Path) -> dict[str, Any]:
    required = {
        "attempt", "binary", "command", "copied_utc", "environment", "gate", "phase",
        "original_binary", "role", "schema", "source_after", "source_before",
        "source_unchanged", "version", "git_revision", "retainer_sha256",
    }
    _exact(value, required, str(path))
    require(value["phase"] == phase and value["role"] == role
            and value["schema"] == "docx-provider-managed-build-v1"
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
    gate = _gate_binding(value["gate"], path, allow_archived_helpers=phase == "before")
    require(gate["source"] == after and gate["argv"] == value["command"], f"{path}: gate binding differs")
    return {
        "path": str(path), "receipt_sha256": _json_hash(path), "phase": phase, "role": role,
        "binary": binary, "original_binary": original, "source": after,
        "gate": gate, "command": list(value["command"]),
        "environment": dict(value["environment"]), "git_revision": revision,
    }


def _build_key(phase: str, role: str) -> str:
    require(phase in PHASES and role in ROLES, f"unknown build phase/role {phase}/{role}")
    return f"{phase}-{role}"


def load_builds(build_dir: Path | None = None) -> dict[str, dict[str, Any]]:
    directory = ROOT if build_dir is None else Path(build_dir)
    result: dict[str, dict[str, Any]] = {}
    for phase in PHASES:
        for role in ROLES:
            path = directory / f"build-{phase}-{role}.json"
            result[_build_key(phase, role)] = _build_from(
                _json(path), phase, role, path)
        normal = result[_build_key(phase, "normal")]
        allocator = result[_build_key(phase, "allocator")]
        require(normal["source"] == allocator["source"], f"{phase}: build source manifests differ")
        require(normal["git_revision"] == allocator["git_revision"], f"{phase}: build revisions differ")
    require(result[_build_key("before", "normal")]["source"] ==
            result[_build_key("before", "allocator")]["source"],
            "before build source manifests differ")
    require(result[_build_key("after", "normal")]["source"] ==
            result[_build_key("after", "allocator")]["source"],
            "after build source manifests differ")
    require(result[_build_key("before", "normal")]["source"] !=
            result[_build_key("after", "normal")]["source"],
            "before and after builds must bind distinct source manifests")
    harness = _shared_harness_binding()
    for phase in PHASES:
        source = result[_build_key(phase, "normal")]["source"]
        manifest = _json(_path(source["path"], f"{phase} source manifest.path"))
        require(all(manifest.get(name) == digest for name, digest in harness["files"].items()),
                f"{phase}: source does not bind the frozen common harness")
    return result


def _environment_binding() -> dict[str, str]:
    path = ROOT / "environment.json"
    return {"path": "environment.json", "sha256": _json_hash(path)}


def _shared_harness_binding() -> dict[str, Any]:
    path = ROOT / "baseline-harness-manifest.json"
    require(path.is_file() and not path.is_symlink(),
            f"missing common harness manifest: {path}")
    value = _json(path)
    _exact(value, HARNESS_FILES, str(path))
    for name, digest in value.items():
        _hash(digest, f"{path}.{name}")
    return {"path": "baseline-harness-manifest.json", "sha256": _json_hash(path),
            "files": dict(value)}


def formal_inventory(*, pilot: bool = False) -> list[dict[str, Any]]:
    samples = PILOT_SAMPLES if pilot else FORMAL_SAMPLES
    warmups = PILOT_WARMUPS if pilot else FORMAL_WARMUPS
    repeats = (1,) if pilot else REPEATS
    result = []
    for repeat in repeats:
        roles = ROLES if repeat == 1 else tuple(reversed(ROLES))
        arms = ARMS if repeat == 1 else tuple(reversed(ARMS))
        phases = PHASES if repeat == 1 else tuple(reversed(PHASES))
        for phase in phases:
            phase_apis = (UNMANAGED_API,) if phase == "before" else APIS
            apis = phase_apis if repeat == 1 else tuple(reversed(phase_apis))
            for api in apis:
                for role in roles:
                    for arm in arms:
                        result.append({
                            "kind": "pilot" if pilot else "formal", "repeat": repeat,
                            "phase": phase, "api": api, "role": role,
                            "arm": arm["name"], "provider": arm["provider"],
                            "samples": samples, "warmups": warmups,
                            "label": f"{'pilot-' if pilot else ''}r{repeat}-{phase}-{api}-{role}-{arm['name']}",
                        })
    return result


def _build_binding(build: dict[str, Any]) -> dict[str, Any]:
    return {"phase": build["phase"], "role": build["role"],
            "receipt_sha256": build["receipt_sha256"], "binary": build["binary"],
            "source": build["source"], "gate": build["gate"],
            "command": build["command"], "environment": build["environment"],
            "git_revision": build["git_revision"]}


def protocol_value(builds: dict[str, dict[str, Any]] | None = None) -> dict[str, Any]:
    inventory = formal_inventory()
    build_bindings = None
    source_bindings = None
    if builds is not None:
        build_bindings = {key: _build_binding(builds[key]) for key in sorted(builds)}
        source_bindings = {phase: builds[_build_key(phase, "normal")]["source"]
                           for phase in PHASES}
    value: dict[str, Any] = {
        "schema": PROTOCOL_SCHEMA, "version": VERSION, "change": CHANGE,
        "claim_authorized": False,
        "comparison_policy": (
            "paired unmanaged before/after latency and RSS regression evidence with "
            "greater-than-5-percent review flags; managed API rows are capability and "
            "budget evidence without a before-managed baseline"
        ),
        "case": CASE, "cpu": CPU, "cpu_lock": CPU_LOCK,
        "roles": list(ROLES), "phases": list(PHASES), "edit_apis": list(APIS),
        "repeats": list(REPEATS),
        "samples": FORMAL_SAMPLES, "warmups": FORMAL_WARMUPS,
        "expected_formal_processes": len(inventory),
        "expected_measured_samples": len(inventory) * FORMAL_SAMPLES,
        "formal_runs": inventory, "pilot_runs": formal_inventory(pilot=True),
        "arms": [dict(arm) for arm in ARMS], "corpus": dict(CORPUS),
        "budget_profile": {"managed": True, "finite": True,
                           "limits": dict(BUDGET_LIMITS)},
        "historical_0494_control": {"managed": False, "comparison_only": True,
                                     "schema": "docx_edit_provider_v1"},
        "oracle": {"text_bytes": CORPUS["expected_text_bytes"],
            "text_sha256": CORPUS["expected_text_sha256"]},
        "limits": {"unmanaged": _limits(UNMANAGED_API), "managed": _limits(MANAGED_API)},
        "driver_limits": DRIVER_LIMITS,
        "scopes": {"provider": PROVIDER_SCOPE, "timing": TIMING_SCOPE,
                    "setup": SETUP_SCOPE, "allocation": ALLOCATION_SCOPE,
                    "physical": PHYSICAL_SCOPE, "range": RANGE_SCOPE,
                    "zero_length": ZERO_LENGTH_SCOPE},
        "environment_artifact": _environment_binding(),
        "shared_harness": _shared_harness_binding(),
        "environment": {key: ENV[key] for key in ENV_KEYS},
        "driver": {name: _json_hash(ROOT / name) for name in DRIVER_FILES},
        "source": source_bindings, "builds": build_bindings,
    }
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
    require(value["shared_harness"] == _shared_harness_binding(),
            f"{path}: common harness binding changed")
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
               str(build["binary"]["path"]), "docx-managed-edit",
               "--edit-api", spec["api"],
               "--provider", arm["provider"]]
    if arm["max_range_bytes"] is not None:
        command += ["--max-range", str(arm["max_range_bytes"])]
    if arm["delay_us"] is not None:
        command += ["--delay-us", str(arm["delay_us"])]
    if arm["transfer_bytes_per_second"] is not None:
        command += ["--transfer-bytes-per-second", str(arm["transfer_bytes_per_second"]),
                    "--transfer-delay-policy", arm["transfer_delay_policy"]]
    if arm["trace_ranges"]:
        command += ["--trace-ranges"]
    # `--max-range` is the provider's single range-bound option.  For the
    # short arm it also determines the reported short-read bound, so do not
    # pass the alias twice and trigger the CLI's duplicate-option guard.
    if arm["short_read_bytes"] is not None and arm["max_range_bytes"] is None:
        command += ["--short-read-bytes", str(arm["short_read_bytes"])]
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
    return {"schema": "docx-edit-provider-private-cleanup-v1",
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
    require(type(timeout_seconds) is int and 0 < timeout_seconds <= MAX_TIMEOUT_SECONDS,
            "capture timeout must be in 1..3600 seconds")
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
    source_revision_at_run = _source_revision()
    current = _normalized_snapshot()
    expected_source = protocol["source"][spec["phase"]]
    require(build["source"] == expected_source,
            "retained build source differs from frozen phase binding")
    require(any(current == source for source in protocol["source"].values()),
            "capture source-at-run is outside the frozen phase bindings")
    argv = _command(spec, build, report, resource, build["git_revision"])
    environment = {key: ENV[key] for key in ENV_KEYS}
    environment["TMPDIR"] = str(tmp_root)
    started = {
        "schema": CAPTURE_SCHEMA, "version": VERSION, "status": "running",
        "attempt": attempt, "run": {key: spec[key] for key in
                                      ("kind", "repeat", "phase", "api", "role", "arm",
                                       "provider", "samples", "warmups", "label")},
        "protocol": {"path": "protocol.json", "sha256": protocol_hash},
        "build": {"path": build["path"], "sha256": build["receipt_sha256"],
                  "phase": build["phase"], "role": build["role"], "source": build["source"]},
        "binary": build["binary"], "source": build["source"],
        "source_at_run": current, "source_revision_at_run": source_revision_at_run,
        "argv": argv, "cwd": str(REPO),
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
            validate_report(report, role=spec["role"], api=spec["api"], arm_name=spec["arm"],
                            samples=spec["samples"], warmups=spec["warmups"],
                            source_revision=build["git_revision"],
                            binary_sha256=build["binary"]["sha256"],
                            binary_bytes=build["binary"]["bytes"])
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
        cleanup = {"schema": "docx-edit-provider-private-cleanup-v1", "status": "failed",
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
    _exact(value, CORPUS_MANIFEST.keys(), path)
    require(value == CORPUS_MANIFEST, f"{path}: corpus manifest differs")


def _check_resource_snapshot(value: Any, path: str, *, managed: bool) -> dict[str, Any]:
    require(isinstance(value, dict), f"{path}: resource snapshot is missing")
    _exact(value, RESOURCE_FIELDS, path)
    if not managed:
        require(value["availability"] == "unavailable" and value["managed"] is False,
                f"{path}: unmanaged resource snapshot is available")
        require(all(value[key] is None for key in RESOURCE_FIELDS[2:]),
                f"{path}: unmanaged resource snapshot contains invented values")
        return dict(value)
    require(value["availability"] == "available" and value["managed"] is True,
            f"{path}: managed resource snapshot is unavailable")
    for key in RESOURCE_FIELDS[2:]:
        _uint(value[key], f"{path}.{key}")
    for limit_key, usage_key in (
        ("memory_bytes", "memory_used"), ("input_bytes", "input_bytes_used"),
        ("output_bytes", "output_bytes_used"), ("objects", "objects_used"),
        ("depth", "depth_used"), ("work", "work_used"),
    ):
        require(value[usage_key] <= BUDGET_LIMITS[limit_key],
                f"{path}.{usage_key}: managed usage exceeds configured limit")
    return dict(value)


def _check_cache_budget(value: Any, path: str, *, managed: bool) -> dict[str, Any]:
    require(isinstance(value, dict), f"{path}: cache budget snapshot is missing")
    _exact(value, CACHE_BUDGET_FIELDS, path)
    if not managed:
        require(value["availability"] == "unavailable" and value["managed"] is False,
                f"{path}: unmanaged cache budget snapshot is available")
        require(all(value[key] is None for key in CACHE_BUDGET_FIELDS[2:]),
                f"{path}: unmanaged cache budget contains invented values")
        return dict(value)
    require(value["availability"] == "available" and value["managed"] is True,
            f"{path}: managed cache budget snapshot is unavailable")
    for key in CACHE_BUDGET_FIELDS[2:]:
        _uint(value[key], f"{path}.{key}")
    for limit_key, usage_key in (
        ("memory_bytes", "budget_memory_used"),
        ("input_bytes", "budget_input_bytes_used"),
        ("output_bytes", "budget_output_bytes_used"),
        ("objects", "budget_objects_used"),
        ("work", "budget_work_used"),
    ):
        require(value[usage_key] <= BUDGET_LIMITS[limit_key],
                f"{path}.{usage_key}: managed usage exceeds configured limit")
    require(value["retained_entries"] <= BASE_LIMITS["cache_max_entries"],
            f"{path}: retained cache entries exceed configured bound")
    require(value["retained_bytes"] <= BASE_LIMITS["cache_max_bytes"],
            f"{path}: retained cache bytes exceed configured bound")
    require(value["budget_cache_reserved_bytes"] <= BASE_LIMITS["cache_max_bytes"],
            f"{path}: cache reservation exceeds configured byte bound")
    require(value["budget_cache_reserved_bytes"] <= value["budget_memory_used"],
            f"{path}: cache reservation exceeds same-phase memory usage")
    require(value["budget_cache_reserved_objects"] <= value["budget_objects_used"],
            f"{path}: cache reservation exceeds object usage")
    require(value["budget_catalog_reserved_objects"] <= value["budget_objects_used"],
            f"{path}: catalog reservation exceeds object usage")
    return dict(value)


def _check_budget(value: Any, api: str, path: str, *, reads: dict[str, Any] | None = None,
                  sink_bytes: int | None = None) -> dict[str, Any]:
    require(isinstance(value, dict), f"{path}: budget evidence is missing")
    _exact(value, BUDGET_FIELDS, path)
    require(value["scope"] == BUDGET_SCOPE, f"{path}.scope: scope differs")
    managed = api == MANAGED_API
    before = _check_resource_snapshot(value["before"], f"{path}.before", managed=managed)
    live = _check_resource_snapshot(value["live"], f"{path}.live", managed=managed)
    after = _check_resource_snapshot(value["after_drop"], f"{path}.after_drop", managed=managed)
    cache_before = _check_cache_budget(value["cache_before"], f"{path}.cache_before", managed=managed)
    cache_live = _check_cache_budget(value["cache_live"], f"{path}.cache_live", managed=managed)
    if not managed:
        require(value["memory_released_to_baseline"] is None
                and value["objects_released_to_baseline"] is None
                and value["reservation_failures"] is None,
                f"{path}: unmanaged budget contains managed ownership claims")
        return dict(value)

    for resource_key, cache_key in (
        ("input_bytes_used", "budget_input_bytes_used"),
        ("output_bytes_used", "budget_output_bytes_used"),
        ("work_used", "budget_work_used"),
    ):
        phases = (before[resource_key], cache_before[cache_key],
                  cache_live[cache_key], live[resource_key], after[resource_key])
        require(all(left <= right for left, right in zip(phases, phases[1:])),
                f"{path}: {resource_key} phase budget moved backwards")
        require(live[resource_key] == after[resource_key],
                f"{path}: {resource_key} changed after publication")
    require(cache_before["budget_output_bytes_used"] == before["output_bytes_used"]
            and cache_live["budget_output_bytes_used"] == before["output_bytes_used"],
            f"{path}: pre-publication output budget is not at baseline")
    for key in ("memory_used", "objects_used", "depth_used"):
        require(after[key] == before[key], f"{path}: {key} did not return to baseline")
    require(value["memory_released_to_baseline"] is True
            and value["objects_released_to_baseline"] is True,
            f"{path}: release booleans do not prove baseline restoration")
    require(type(value["reservation_failures"]) is int and value["reservation_failures"] == 0,
            f"{path}: reservation failures are nonzero")
    require(cache_live["budget_reservation_failures"] == value["reservation_failures"],
            f"{path}: reservation failure receipt diverges from cache diagnostics")
    if sink_bytes is not None:
        output_delta = live["output_bytes_used"] - before["output_bytes_used"]
        require(output_delta == sink_bytes,
                f"{path}: managed output charge differs from accepted sink bytes")
    if reads is not None:
        physical = reads.get("physical")
        logical = reads.get("logical")
        counter = (physical if isinstance(physical, dict)
                   and physical.get("availability") == "available" else logical)
        if isinstance(counter, dict) and counter.get("availability") == "available":
            input_delta = live["input_bytes_used"] - before["input_bytes_used"]
            require(input_delta == counter["returned_bytes"],
                    f"{path}: managed input charge differs from authenticated returned bytes")
    return dict(value)


def _check_limits(value: Any, api: str, path: str) -> None:
    require(isinstance(value, dict), f"{path}: limits object is missing")
    _finite(value, path)
    expected = _limits(api)
    _exact(value, expected.keys(), path)
    require(value == expected, f"{path}: provider limits differ for {api}")


def _check_provider(value: Any, arm: dict[str, Any], path: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{path}: provider object is missing")
    _exact(value, PROVIDER_FIELDS, path)
    require(value["name"] == arm["provider"] and value["provider"] == arm["provider"],
            f"{path}: provider differs from arm")
    require(value["kind"] == arm["kind"], f"{path}: provider kind differs from arm")
    require(value["trace_ranges"] is arm["trace_ranges"], f"{path}: trace identity differs")
    require(value["short_read_bytes"] == arm["short_read_bytes"],
            f"{path}: short-read bound differs")
    require(value["max_range_bytes"] == arm["max_range_bytes"], f"{path}: max range differs")
    require(value["delay_us"] == arm["delay_us"], f"{path}: delay differs")
    require(value["transfer_bytes_per_second"] == arm["transfer_bytes_per_second"],
            f"{path}: transfer rate differs")
    require(value["transfer_delay_policy"] == arm["transfer_delay_policy"],
            f"{path}: transfer policy differs")
    require(value["zero_length_scope"] ==
            "zero-length caller ReadAt calls are delegated to the wrapped provider, preserve its return/error behavior, and are counted separately from nonempty range totals",
            f"{path}: zero-length scope differs")
    require(value["source_construction"] ==
            "source adapters and FileSource handles are constructed before the operation clock",
            f"{path}: source construction scope differs")
    require(value["file_scope"] == FILE_SCOPE and value["range_scope"] == RANGE_SCOPE,
            f"{path}: provider scope differs")
    return dict(value)


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


def _request_histogram(counters: list[dict[str, Any]]) -> list[int] | None:
    if not counters or not all(item.get("availability") == "available" for item in counters):
        return None
    result = [0] * REQUEST_BUCKETS
    for counter in counters:
        for item in counter["ranges"]:
            result[_bucket(item["requested"])] += 1
    return result


def _counter_values(counters: list[dict[str, Any]], key: str) -> list[int] | None:
    if not counters or not all(item.get("availability") == "available" for item in counters):
        return None
    return [item[key] or 0 for item in counters]


def _budget_values(rows: list[dict[str, Any]], location: str, key: str) -> list[int] | None:
    values = [row.get("budget", {}).get(location) for row in rows]
    if not values or not all(isinstance(item, dict) and item.get("availability") == "available"
                             for item in values):
        return None
    result = [item.get(key) for item in values]
    if not all(type(item) is int for item in result):
        return None
    return result


def _analysis_read_summary(reads: dict[str, Any]) -> dict[str, Any]:
    """Keep per-row counters while replacing large raw traces with derived tails."""
    result = dict(reads)
    for name in ("logical", "physical"):
        counter = result.get(name)
        if not isinstance(counter, dict):
            continue
        counter = dict(counter)
        ranges = counter.pop("ranges", None)
        if ranges is not None:
            counter["range_count"] = len(ranges)
            counter["request_size_histogram"] = _request_histogram([{
                **counter, "ranges": ranges, "availability": "available"
            }])
        result[name] = counter
    return result


def _bool(value: Any, path: str) -> None:
    require(type(value) is bool, f"{path}: expected boolean")


def _find(value: dict[str, Any], names: tuple[str, ...], path: str, *, required: bool = True) -> Any:
    for name in names:
        if name in value:
            return value[name]
    if required:
        fail(f"{path}: none of {names} is present")
    return None


def _check_read_counter(value: Any, path: str, *, expected_scope: str,
                        expected_available: bool) -> dict[str, Any]:
    require(isinstance(value, dict), f"{path}: read counter is missing")
    _exact(value, READ_COUNTER_FIELDS, path)
    require(value["scope"] == expected_scope, f"{path}: scope differs")
    available = value["availability"] == "available"
    require(available is expected_available, f"{path}: availability differs")
    numeric = ("calls", "empty_calls", "requested_bytes", "returned_bytes", "short_reads",
               "traced_ranges", "requested_media_overlap_bytes", "returned_media_overlap_bytes")
    optional_numeric = ("min_request_bytes", "max_request_bytes")
    if not available:
        require(value["availability"] == "unavailable", f"{path}: unknown availability")
        require(all(value[key] is None for key in numeric + optional_numeric) and value["ranges"] is None,
                f"{path}: unavailable counter contains measurements")
        return dict(value)
    for key in numeric:
        _uint(value[key], f"{path}.{key}")
    for key in optional_numeric:
        if value[key] is not None:
            _uint(value[key], f"{path}.{key}", positive=True)
    require(value["returned_bytes"] <= value["requested_bytes"],
            f"{path}: returned bytes exceed requested")
    require(value["empty_calls"] >= 0 and value["short_reads"] <= value["calls"]
            and value["traced_ranges"] == value["calls"],
            f"{path}: call counters do not conserve")
    require(value["returned_media_overlap_bytes"] <= value["requested_media_overlap_bytes"]
            <= value["requested_bytes"], f"{path}: media overlap exceeds request")
    ranges = value["ranges"]
    require(isinstance(ranges, list) and len(ranges) == value["calls"] == value["traced_ranges"],
            f"{path}: retained range trace does not conserve calls")
    checked = [_check_range(item, f"{path}.ranges[{index}]")
               for index, item in enumerate(ranges)]
    require(sum(item[1] for item in checked) == value["requested_bytes"]
            and sum(item[2] for item in checked) == value["returned_bytes"],
            f"{path}: retained range trace does not conserve bytes")
    require(sum(1 for item in checked if item[2] < item[1]) == value["short_reads"],
            f"{path}: retained range trace does not conserve short reads")
    if value["calls"] == 0:
        require(value["min_request_bytes"] is None and value["max_request_bytes"] is None,
                f"{path}: empty counter has request bounds")
    else:
        require(min(item[1] for item in checked) == value["min_request_bytes"]
                and max(item[1] for item in checked) == value["max_request_bytes"],
                f"{path}: retained range trace does not conserve request bounds")
    requested_media = 0
    returned_media = 0
    for offset, requested, returned in checked:
        require(offset + requested <= CORPUS["archive_bytes"],
                f"{path}: retained range exceeds source length")
        requested_end = offset + requested
        returned_end = offset + returned
        for media_start, media_end in MEDIA_RANGES:
            requested_media += max(0, min(requested_end, media_end) - max(offset, media_start))
            returned_media += max(0, min(returned_end, media_end) - max(offset, media_start))
    require(requested_media == value["requested_media_overlap_bytes"]
            and returned_media == value["returned_media_overlap_bytes"],
            f"{path}: media overlap does not conserve")
    if value["calls"] > 0:
        require(value["min_request_bytes"] is not None and value["max_request_bytes"] is not None
                and value["min_request_bytes"] <= value["max_request_bytes"],
                f"{path}: request bounds do not conserve")
    require(value["traced_ranges"] <= LIMITS["max_tracked_ranges"],
            f"{path}: range trace exceeds configured capacity")
    return dict(value)


def _check_range_adapter(value: Any, path: str, *, arm: dict[str, Any],
                         logical: dict[str, Any], physical: dict[str, Any],
                         expected_available: bool) -> dict[str, Any]:
    require(isinstance(value, dict), f"{path}: range adapter evidence is missing")
    _exact(value, RANGE_ADAPTER_FIELDS, path)
    available = value["availability"] == "available"
    require(available is expected_available, f"{path}: availability differs")
    numeric = ("logical_calls", "requested_bytes", "returned_bytes", "short_reads",
               "delayed_calls", "transfer_paced_calls", "transfer_delay_ns")
    if not available:
        require(value["availability"] == "unavailable" and all(value[key] is None for key in numeric),
                f"{path}: unavailable adapter contains measurements")
        return dict(value)
    for key in numeric:
        _uint(value[key], f"{path}.{key}")
    require(value["returned_bytes"] <= value["requested_bytes"]
            and value["short_reads"] <= value["logical_calls"]
            and value["delayed_calls"] <= value["logical_calls"]
            and value["transfer_paced_calls"] <= value["logical_calls"],
            f"{path}: adapter counters do not conserve")

    # The adapter is the transport model between the caller-facing logical
    # counter and the underlying physical source counter.  Reconcile each
    # retained range so aggregate counters cannot hide a reordered or
    # substituted trace.
    require(logical["availability"] == "available"
            and physical["availability"] == "available",
            f"{path}: available adapter lacks available source counters")
    logical_ranges = logical["ranges"]
    physical_ranges = physical["ranges"]
    require(isinstance(logical_ranges, list) and isinstance(physical_ranges, list)
            and len(logical_ranges) == len(physical_ranges) == value["logical_calls"],
            f"{path}: adapter range count does not conserve")
    require(value["logical_calls"] == logical["calls"] == physical["calls"]
            and value["requested_bytes"] == logical["requested_bytes"]
            and value["returned_bytes"] == logical["returned_bytes"] == physical["returned_bytes"]
            and value["short_reads"] == logical["short_reads"],
            f"{path}: logical/physical adapter totals do not conserve")
    maximum = arm["max_range_bytes"]
    physical_requested = 0
    expected_transfer_calls = 0
    expected_transfer_delay = 0
    rate = arm["transfer_bytes_per_second"]
    for index, (logical_range, physical_range) in enumerate(zip(logical_ranges, physical_ranges)):
        logical_offset = logical_range["offset"]
        logical_requested = logical_range["requested"]
        logical_returned = logical_range["returned"]
        require(physical_range["offset"] == logical_offset,
                f"{path}: physical range {index} offset differs from logical range")
        expected_requested = logical_requested if maximum is None else min(logical_requested, maximum)
        require(physical_range["requested"] == expected_requested
                and physical_range["returned"] == logical_returned,
                f"{path}: physical range {index} does not match adapter delegation")
        physical_requested += expected_requested
        if logical_returned:
            expected_transfer_calls += 1
            if rate is not None:
                expected_transfer_delay += (logical_returned * 1_000_000_000 + rate - 1) // rate
    require(physical_requested == physical["requested_bytes"],
            f"{path}: physical request total does not conserve the configured range cap")

    delay_us = arm["delay_us"]
    if delay_us is None:
        require(arm["transfer_bytes_per_second"] is None
                and arm["transfer_delay_policy"] == "separate-sleeps"
                and value["delayed_calls"] == 0
                and value["transfer_paced_calls"] == 0
                and value["transfer_delay_ns"] == 0,
                f"{path}: unpaced arm reported service counters")
    else:
        require(value["delayed_calls"] == value["logical_calls"],
                f"{path}: configured fixed delay was not recorded for every delegated call")
        require(arm["transfer_bytes_per_second"] is not None
                and arm["transfer_delay_policy"] == "minimum-service"
                and value["transfer_paced_calls"] == expected_transfer_calls
                and value["transfer_delay_ns"] == expected_transfer_delay,
                f"{path}: transfer pacing counters do not match the configured service")
    return dict(value)


def _check_row(row: Any, role: str, api: str, arm: dict[str, Any], path: str,
               expected_index: int, source_sha: str) -> dict[str, Any]:
    require(isinstance(row, dict), f"{path}: row is missing")
    row_fields = {"sample_index", "api", "latency_ns", "output_bytes", "output_sha256",
                  "output_exact_bytes", "materializations", "commit_changed",
                  "commit_operations", "source_version_before", "source_version_after",
                  "source_version_unchanged", "reads", "cache", "sink", "oracles", "budget"}
    if role == "allocator":
        row_fields.add("allocation")
    _exact(row, row_fields, path)
    require(row["api"] == api, f"{path}.api: row API differs")
    index = _find(row, ("sample_index", "iteration"), path)
    _uint(index, f"{path}.sample_index")
    require(index == expected_index, f"{path}: sample index differs")
    elapsed = _find(row, ("latency_ns", "elapsed_ns"), path)
    _uint(elapsed, f"{path}.elapsed_ns", positive=True)
    before = _find(row, ("source_version_before", "source_before"), path)
    after = _find(row, ("source_version_after", "source_after"), path)
    _check_version(before, f"{path}.source_version_before")
    _check_version(after, f"{path}.source_version_after")
    unchanged = _find(row, ("source_version_unchanged", "source_unchanged"), path)
    _bool(unchanged, f"{path}.source_version_unchanged")
    require(unchanged is True and before == after, f"{path}: source changed during edit")

    reads = _find(row, ("reads",), path)
    require(isinstance(reads, dict), f"{path}.reads: evidence is missing")
    _exact(reads, ("logical", "physical", "range_adapter", "media_scope"), f"{path}.reads")
    expected_logical = arm["provider"] in {"instrumented", "file", "short", "delayed"}
    expected_physical = arm["provider"] in {"short", "delayed"}
    logical = _check_read_counter(reads["logical"], f"{path}.reads.logical",
                                  expected_scope="caller-visible logical ReadAt calls",
                                  expected_available=expected_logical)
    physical = _check_read_counter(reads["physical"], f"{path}.reads.physical",
                                   expected_scope="underlying adapter ReadAt calls; transport model only",
                                   expected_available=expected_physical)
    adapter = _check_range_adapter(reads["range_adapter"], f"{path}.reads.range_adapter",
                                   arm=arm, logical=logical, physical=physical,
                                   expected_available=expected_physical)
    require(reads["media_scope"] ==
            "source compressed ranges for word/media members; exact output and OPC semantic checks additionally prove unchanged media payloads",
            f"{path}.reads.media_scope: scope differs")

    sink = _find(row, ("sink",), path)
    require(isinstance(sink, dict), f"{path}.sink: summary is missing")
    _exact(sink, ("accepted_bytes", "write_calls", "largest_write"), f"{path}.sink")
    sink_bytes = _find(sink, ("accepted_bytes", "bytes"), f"{path}.sink")
    sink_calls = _find(sink, ("write_calls", "calls"), f"{path}.sink")
    _uint(sink_bytes, f"{path}.sink.bytes")
    _uint(sink_calls, f"{path}.sink.write_calls")
    largest = sink.get("largest_write", 0)
    _uint(largest, f"{path}.sink.largest_write")
    require(0 < largest <= min(sink_bytes, LIMITS["sink_max_write_bytes"]) and sink_calls > 0,
            f"{path}.sink: write bounds or count differ")

    output_bytes = _find(row, ("output_bytes",), path)
    output_sha = _find(row, ("output_sha256",), path)
    _uint(output_bytes, f"{path}.output_bytes")
    _hash(output_sha, f"{path}.output_sha256")
    output_exact = _find(row, ("output_exact_bytes",), path)
    _bool(output_exact, f"{path}.output_exact_bytes")
    require(output_bytes == sink_bytes and output_exact is True,
            f"{path}: output oracle failed")
    require(output_sha != source_sha, f"{path}: output digest equals source")
    materializations = _find(row, ("materializations",), path)
    _uint(materializations, f"{path}.materializations", positive=True)
    commit_changed = _find(row, ("commit_changed",), path)
    _bool(commit_changed, f"{path}.commit_changed")
    commit_operations = _find(row, ("commit_operations",), path)
    _uint(commit_operations, f"{path}.commit_operations")
    require(commit_changed is True and commit_operations == 1,
            f"{path}: commit oracle failed")

    cache = _find(row, ("cache",), path)
    require(isinstance(cache, dict), f"{path}.cache: evidence is missing")
    _exact(cache, ("successful_loads", "expected_successful_loads",
                   "exactly_one_main_part_materialization"), f"{path}.cache")
    successful = _find(cache, ("successful_loads",), f"{path}.cache")
    expected = _find(cache, ("expected_successful_loads",), f"{path}.cache")
    _uint(successful, f"{path}.cache.successful_loads", positive=True)
    _uint(expected, f"{path}.cache.expected_successful_loads", positive=True)
    exactly_one = _find(cache, ("exactly_one_main_part_materialization",), f"{path}.cache")
    _bool(exactly_one, f"{path}.cache.exactly_one_main_part_materialization")
    require(successful == expected == materializations == 1 and exactly_one is True,
            f"{path}: cache materialization oracle failed")

    oracles = _find(row, ("oracles",), path)
    require(isinstance(oracles, dict), f"{path}.oracles: evidence is missing")
    _exact(oracles, ("commit_identity_verified", "output_exact_bytes", "semantic_reopen",
                     "unchanged_media_preserved", "source_version_unchanged",
                     "cache_load_count", "patch_oracles"), f"{path}.oracles")
    for key in ("commit_identity_verified", "output_exact_bytes", "semantic_reopen",
                "unchanged_media_preserved", "source_version_unchanged", "cache_load_count"):
        _bool(oracles.get(key), f"{path}.oracles.{key}")
        require(oracles[key] is True, f"{path}.oracles.{key}: failed")
    patch = oracles.get("patch_oracles")
    require(isinstance(patch, dict), f"{path}.oracles.patch_oracles: missing")
    _exact(patch, ("scope", "replay_forward", "inverse_restores_source",
                   "stale_target_refused", "foreign_source_refused"),
           f"{path}.oracles.patch_oracles")
    require(patch["scope"] == PATCH_ORACLE_SCOPE,
            f"{path}.oracles.patch_oracles.scope: scope differs")
    for key in ("replay_forward", "inverse_restores_source", "stale_target_refused", "foreign_source_refused"):
        _bool(patch.get(key), f"{path}.oracles.patch_oracles.{key}")
        require(patch[key] is True, f"{path}.oracles.patch_oracles.{key}: failed")

    budget = _check_budget(row["budget"], api, f"{path}.budget", reads=reads,
                           sink_bytes=sink_bytes)

    allocation = row.get("allocation")
    if role == "normal":
        require(allocation is None, f"{path}.allocation: normal metrics present")
    else:
        require(isinstance(allocation, dict), f"{path}.allocation: allocator metrics missing")
        fields = ("status", "scope", "allocation_calls", "deallocation_calls", "reallocation_calls",
                  "failed_allocation_calls", "allocated_bytes", "deallocated_bytes", "live_bytes_before",
                  "live_bytes_after", "peak_live_bytes_before", "peak_live_bytes_after", "region_peak_live_bytes")
        require(set(allocation) == set(fields), f"{path}.allocation: fields differ")
        require(allocation["status"] == "measured" and allocation["scope"] == SAMPLE_ALLOCATION_SCOPE,
                f"{path}.allocation: sample unavailable")
        for key in fields[2:]:
            _uint(allocation[key], f"{path}.allocation.{key}")
        require(allocation["allocation_calls"] >= allocation["reallocation_calls"],
                f"{path}.allocation: operation counts do not conserve reallocations")
        before_live = allocation["live_bytes_before"]
        after_live = allocation["live_bytes_after"]
        allocated = allocation["allocated_bytes"]
        deallocated = allocation["deallocated_bytes"]
        require(before_live + allocated >= deallocated
                and after_live == before_live + allocated - deallocated,
                f"{path}.allocation: live-byte conservation failed")
        require(allocation["peak_live_bytes_before"] >= before_live
                and allocation["peak_live_bytes_after"] >= allocation["peak_live_bytes_before"]
                and allocation["peak_live_bytes_after"] >= after_live
                and allocation["region_peak_live_bytes"] >= max(before_live, after_live)
                and allocation["region_peak_live_bytes"] <= allocation["peak_live_bytes_after"],
                f"{path}.allocation: peak-live conservation failed")
    return {"sample_index": index, "elapsed_ns": elapsed, "logical": logical,
            "physical": physical, "sink": {"bytes": sink_bytes, "write_calls": sink_calls,
                                            "largest_write": largest},
            "output_bytes": output_bytes, "output_sha256": output_sha,
            "materializations": materializations, "allocation": allocation,
            "cache": cache, "range_adapter": adapter, "budget": budget}


def validate_report(path: Path, *, role: str, api: str, arm_name: str, samples: int,
                    warmups: int, source_revision: str | None = None,
                    binary_sha256: str | None = None, binary_bytes: int | None = None) -> dict[str, Any]:
    value = _json(path)
    arm = _arm(arm_name)
    _exact(value, REPORT_FIELDS, str(path))
    require(api in APIS, f"{path}: unknown API mode")
    require(value["schema"] == SCHEMA and value["version"] == VERSION
            and value["case_name"] == CASE and isinstance(value["benchmark"], str),
            f"{path}: report identity differs")
    require(value["api"] == api and value["api_name"] == api,
            f"{path}: report API identity differs")
    require(value["provider_scope"] == PROVIDER_SCOPE and value["timing_scope"] == TIMING_SCOPE
            and value["setup_scope"] == SETUP_SCOPE and value["allocation_scope"] == ALLOCATION_SCOPE
            and value["physical_scope"] == PHYSICAL_SCOPE and value["range_scope"] == RANGE_SCOPE
            and value["zero_length_scope"] == ZERO_LENGTH_SCOPE
            and value["file_scope"] == FILE_SCOPE
            and value["budget_scope"] == BUDGET_SCOPE,
            f"{path}: report scope differs")
    require(value["corpus_version"] == CORPUS["version"]
            and value["corpus_generator"] == CORPUS["generator"],
            f"{path}: top-level corpus identity differs")
    _check_corpus(value["corpus"], f"{path}.corpus")
    require(value["source_archive_bytes"] == CORPUS["archive_bytes"]
            and value["source_bytes"] == CORPUS["archive_bytes"]
            and value["source_archive_sha256"] == CORPUS["archive_sha256"]
            and value["source_sha256"] == CORPUS["archive_sha256"]
            and value["expected_text_bytes"] == CORPUS["expected_text_bytes"]
            and value["expected_text_sha256"] == CORPUS["expected_text_sha256"]
            and value["expected_text_scope"] ==
            "original deterministic corpus text identity before the selected paragraph replacement; it is not the changed output text identity"
            and value["expected_archive_members"] == CORPUS["archive_member_count"],
            f"{path}: source oracle differs")
    _hash(value["source_archive_sha256"], f"{path}.source_archive_sha256")
    _uint(value["source_archive_bytes"], f"{path}.source_archive_bytes", positive=True)
    _uint(value["source_bytes"], f"{path}.source_bytes", positive=True)
    _hash(value["source_sha256"], f"{path}.source_sha256")
    _uint(value["expected_text_bytes"], f"{path}.expected_text_bytes", positive=True)
    _hash(value["expected_text_sha256"], f"{path}.expected_text_sha256")
    _uint(value["expected_archive_members"], f"{path}.expected_archive_members", positive=True)
    _hash(value["expected_output_sha256"], f"{path}.expected_output_sha256")
    _uint(value["expected_output_bytes"], f"{path}.expected_output_bytes", positive=True)
    require(value["expected_output_sha256"] == EXPECTED_OUTPUT_SHA256
            and value["expected_output_bytes"] == EXPECTED_OUTPUT_BYTES,
            f"{path}: expected output identity differs")
    require(value["requested_source_revision"] == value["source_revision"],
            f"{path}: requested/source revision differs")
    _check_limits(value["limits"], api, f"{path}.limits")
    _hash(value["binary_sha256"], f"{path}.binary_sha256")
    _uint(value["binary_bytes"], f"{path}.binary_bytes", positive=True)
    if binary_sha256 is not None:
        require(value["binary_sha256"] == binary_sha256, f"{path}: binary digest differs from build")
    if binary_bytes is not None:
        require(value["binary_bytes"] == binary_bytes, f"{path}: binary size differs from build")
    _text(value["current_exe"], f"{path}.current_exe")
    _check_provider(value["provider"], arm, f"{path}.provider")
    if role == "normal":
        require(value["allocator"] == "Rust system allocator"
                and value["instrumentation"] == "none",
                f"{path}: normal allocator identity differs")
    else:
        require(value["allocator"] == "CountingSystemAllocator(std::alloc::System)"
                and value["instrumentation"] == "system_allocator_operation_scoped",
                f"{path}: allocator identity differs")
    require(value["warmup"] == warmups and value["samples"] == samples, f"{path}: sample counts differ")
    revision = _text(value["source_revision"], f"{path}.source_revision")
    require(REVISION_RE.fullmatch(revision) is not None, f"{path}: malformed source revision")
    if source_revision is not None:
        require(revision == source_revision, f"{path}: source revision differs")
    preflight = value["preflight"]
    require(isinstance(preflight, dict), f"{path}.preflight: missing")
    _exact(preflight, PREFLIGHT_FIELDS, f"{path}.preflight")
    expected_materializations = preflight.get("expected_materializations")
    _uint(expected_materializations, f"{path}.preflight.expected_materializations", positive=True)
    for key in ("source_document_xml_sha256", "candidate_document_xml_sha256"):
        _hash(preflight.get(key), f"{path}.preflight.{key}")
    require(preflight["source_document_xml_sha256"] != preflight["candidate_document_xml_sha256"],
            f"{path}.preflight: source and candidate XML identities are equal")
    for key in ("output_exact_source_changed", "semantic_reopen_verified", "unchanged_media_preserved",
                "replay_forward_verified", "inverse_restores_source_verified",
                "stale_target_refusal_verified", "foreign_source_refusal_verified"):
        require(preflight.get(key) is True, f"{path}.preflight.{key}: failed")
    rows = value["rows"]
    require(isinstance(rows, list) and len(rows) == samples, f"{path}: row count differs")
    for index, row in enumerate(rows):
        _check_row(row, role, api, arm, f"{path}.rows[{index}]", index,
                   value["source_archive_sha256"])
        require(row["output_sha256"] == value["expected_output_sha256"]
                and row["output_bytes"] == value["expected_output_bytes"]
                and row["materializations"] == expected_materializations,
                f"{path}.rows[{index}]: output identity differs")
    return value


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
                       build: dict[str, Any], protocol: dict[str, Any], protocol_hash: str,
                       directory: Path) -> None:
    started_fields = {"schema", "version", "status", "attempt", "run", "protocol", "build", "binary",
                      "source", "source_at_run", "source_revision_at_run", "argv", "cwd",
                      "environment", "environment_artifact", "driver_bindings",
                      "driver_sha256", "support_sha256", "started_utc", "timeout_seconds", "tmpdir"}
    terminal_fields = started_fields | {"exit_code", "timed_out", "termination", "finished_utc", "artifacts",
                                       "missing_artifacts", "cleanup", "started_artifact", "source_before",
                                       "source_after", "source_unchanged"}
    _exact(started, started_fields, f"{directory}/started.json")
    _exact(terminal, terminal_fields, f"{directory}/terminal.json")
    require(started["schema"] == CAPTURE_SCHEMA and started["status"] == "running", f"{directory}: started identity differs")
    require(terminal["schema"] == TERMINAL_SCHEMA and terminal["status"] == "pass", f"{directory}: terminal did not pass")
    for field in ("version", "attempt", "run", "protocol", "build", "binary", "source",
                  "source_at_run", "source_revision_at_run", "cwd",
                  "environment_artifact", "driver_bindings", "driver_sha256", "support_sha256",
                  "timeout_seconds", "tmpdir"):
        require(terminal[field] == started[field], f"{directory}: terminal {field} changed")
    require(started["run"] == {key: spec[key] for key in
                                ("kind", "repeat", "phase", "api", "role", "arm", "provider",
                                 "samples", "warmups", "label")},
            f"{directory}: run specification differs")
    require(started["attempt"] == spec["attempt"] and started["protocol"] == {"path": "protocol.json", "sha256": protocol_hash},
            f"{directory}: protocol binding differs")
    require(started["build"] == {"path": build["path"], "sha256": build["receipt_sha256"],
                                  "phase": build["phase"], "role": build["role"],
                                  "source": build["source"]}
            and started["binary"] == build["binary"] and started["source"] == build["source"]
            and any(started["source_at_run"] == source
                    for source in protocol["source"].values()),
            f"{directory}: build binding differs")
    _source_binding(started["source_at_run"], f"{directory}.source_at_run")
    source_revision_at_run = _text(started["source_revision_at_run"],
                                    f"{directory}.source_revision_at_run")
    require(REVISION_RE.fullmatch(source_revision_at_run) is not None,
            f"{directory}: source-at-run revision is malformed")
    require(started["cwd"] == str(REPO) and started["driver_bindings"] ==
            {name: _json_hash(ROOT / name) for name in DRIVER_FILES}
            and started["driver_sha256"] == _json_hash(ROOT / "measure.py")
            and started["support_sha256"] == _json_hash(ROOT / "support.py"), f"{directory}: helper custody differs")
    tmp_root = _path(started["tmpdir"], f"{directory}.tmpdir")
    expected_run_root = _run_root(spec["attempt"], spec["label"])
    require(tmp_root == expected_run_root / "tmp", f"{directory}: private scratch path differs")
    expected_env = {key: ENV[key] for key in ENV_KEYS}
    expected_env["TMPDIR"] = str(tmp_root)
    require(started["environment"] == expected_env and terminal["environment"] == expected_env,
            f"{directory}: child environment differs")
    require(started["argv"] == terminal["argv"] and started["argv"] ==
            _command(spec, build, directory / "report.json", directory / "resource.txt", build["git_revision"]),
            f"{directory}: child argv differs")
    require(terminal["exit_code"] == 0 and terminal["timed_out"] is False and terminal["termination"] is None
            and terminal["source_before"] == started["source_at_run"] == terminal["source_after"]
            and terminal["source_unchanged"] is True and terminal["missing_artifacts"] == [], f"{directory}: terminal custody failed")
    require(_timestamp(terminal["finished_utc"], f"{directory}.finished_utc") >
            _timestamp(started["started_utc"], f"{directory}.started_utc"), f"{directory}: chronology invalid")
    cleanup = terminal["cleanup"]
    _exact(cleanup, ("schema", "status", "root", "tmpdir", "removed", "remaining"), f"{directory}.cleanup")
    require(cleanup["schema"] == "docx-edit-provider-private-cleanup-v1" and cleanup["status"] == "pass"
            and cleanup["tmpdir"] == str(tmp_root) and cleanup["root"] == str(expected_run_root)
            and not expected_run_root.exists() and cleanup["remaining"] == [], f"{directory}: private cleanup failed")
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
        build = builds[_build_key(spec["phase"], spec["role"])]
        _validate_terminal(started, terminal, spec, build, protocol, protocol_hash, directory)
        report_path = directory / "report.json"
        report = validate_report(report_path, role=spec["role"], api=spec["api"], arm_name=spec["arm"],
                                 samples=spec["samples"], warmups=spec["warmups"],
                                 source_revision=build["git_revision"],
                                 binary_sha256=build["binary"]["sha256"],
                                 binary_bytes=build["binary"]["bytes"])
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
    seed = (0x0495_5EED ^ len(values) ^ sum(values)) & 0xFFFFFFFFFFFFFFFF
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
    return {"method": "deterministic_percentile_bootstrap_median", "seed": (0x0495_5EED ^ len(values) ^ sum(values)) & 0xFFFFFFFFFFFFFFFF,
            "resamples": repetitions, "confidence": 0.95, "median": _percentiles(values)["p50"],
            "ci_low": medians[(repetitions * 25) // 1000], "ci_high": medians[(repetitions * 975) // 1000 - 1]}


def _vectors(entry: dict[str, Any], role: str, arm: dict[str, Any]) -> dict[str, Any]:
    rows = entry["report"]["rows"]
    reads = [row["reads"] for row in rows]
    logical = [item["logical"] for item in reads]
    physical = [item["physical"] for item in reads
                if isinstance(item.get("physical"), dict)
                and item["physical"].get("availability") == "available"]
    adapters = [item["range_adapter"] for item in reads
                if isinstance(item.get("range_adapter"), dict)
                and item["range_adapter"].get("availability") == "available"]
    sinks = [row["sink"] for row in rows]
    outputs = rows
    values: dict[str, Any] = {
        "elapsed_ns": [row["latency_ns"] for row in rows],
        # GNU time observes one maximum for the child process, not one value
        # per in-process sample.  Preserve that distinction in the receipt.
        "whole_child_rss_kbytes": [entry["resource"]["maximum_resident_set_size_(kbytes)"]],
        "source_logical_read_calls": _counter_values(logical, "calls"),
        "source_logical_read_requested_bytes": _counter_values(logical, "requested_bytes"),
        "source_logical_read_returned_bytes": _counter_values(logical, "returned_bytes"),
        "source_logical_empty_calls": _counter_values(logical, "empty_calls"),
        "source_logical_short_reads": _counter_values(logical, "short_reads"),
        "source_logical_requested_media_overlap_bytes":
            _counter_values(logical, "requested_media_overlap_bytes"),
        "source_logical_returned_media_overlap_bytes":
            _counter_values(logical, "returned_media_overlap_bytes"),
        "source_logical_request_size_histogram": _request_histogram(logical),
        "source_physical_read_calls": _counter_values(physical, "calls"),
        "source_physical_read_requested_bytes": _counter_values(physical, "requested_bytes"),
        "source_physical_read_returned_bytes": _counter_values(physical, "returned_bytes"),
        "source_physical_empty_calls": _counter_values(physical, "empty_calls"),
        "source_physical_short_reads": _counter_values(physical, "short_reads"),
        "source_physical_requested_media_overlap_bytes":
            _counter_values(physical, "requested_media_overlap_bytes"),
        "source_physical_returned_media_overlap_bytes":
            _counter_values(physical, "returned_media_overlap_bytes"),
        "source_physical_request_size_histogram": _request_histogram(physical),
        "range_adapter_logical_calls": _counter_values(adapters, "logical_calls"),
        "range_adapter_requested_bytes": _counter_values(adapters, "requested_bytes"),
        "range_adapter_returned_bytes": _counter_values(adapters, "returned_bytes"),
        "range_adapter_short_reads": _counter_values(adapters, "short_reads"),
        "range_adapter_delayed_calls": _counter_values(adapters, "delayed_calls"),
        "range_adapter_transfer_paced_calls": _counter_values(adapters, "transfer_paced_calls"),
        "range_adapter_transfer_delay_ns": _counter_values(adapters, "transfer_delay_ns"),
        "sink_write_calls": [item["write_calls"] for item in sinks],
        "sink_bytes": [item["accepted_bytes"] for item in sinks],
        "sink_largest_write": [item["largest_write"] for item in sinks],
        "output_bytes": [item["output_bytes"] for item in outputs],
        "materializations": [item["materializations"] for item in rows],
    }
    for location in ("before", "live", "after_drop"):
        for key in ("memory_used", "input_bytes_used", "output_bytes_used",
                    "objects_used", "depth_used", "work_used"):
            values[f"budget_{location}_{key}"] = _budget_values(rows, location, key)
    for location in ("cache_before", "cache_live"):
        for key in ("budget_reservation_failures", "budget_memory_used",
                    "budget_cache_reserved_bytes", "budget_input_bytes_used",
                    "budget_output_bytes_used", "budget_work_used",
                    "budget_objects_used", "budget_catalog_reserved_objects",
                    "budget_cache_reserved_objects", "retained_entries",
                    "retained_bytes", "in_flight_loads"):
            values[f"{location}_{key}"] = _budget_values(rows, location, key)
    release = [row.get("budget", {}).get("memory_released_to_baseline") for row in rows]
    values["budget_memory_released_to_baseline"] = release if all(type(item) is bool for item in release) else None
    release = [row.get("budget", {}).get("objects_released_to_baseline") for row in rows]
    values["budget_objects_released_to_baseline"] = release if all(type(item) is bool for item in release) else None
    cache_values = [row["cache"] for row in rows]
    if all(isinstance(item, dict) for item in cache_values):
        for key in ("successful_loads", "expected_successful_loads"):
            values[f"cache_{key}"] = [item[key] for item in cache_values]
        # Keep this retained oracle typed as a boolean.  It is a validation
        # result, not a synthetic numeric zero, so it has no percentile.
        values["cache_exactly_one_main_part_materialization"] = [
            item["exactly_one_main_part_materialization"] for item in cache_values
        ]
    else:
        values.update({"cache_successful_loads": None,
                       "cache_expected_successful_loads": None,
                       "cache_exactly_one_main_part_materialization": None})
    if role == "allocator":
        samples = [row["allocation"] for row in rows]
        values.update({"allocation_calls": [item["allocation_calls"] for item in samples],
                       "deallocation_calls": [item["deallocation_calls"] for item in samples],
                       "reallocation_calls": [item["reallocation_calls"] for item in samples],
                       "allocated_bytes": [item["allocated_bytes"] for item in samples],
                       "deallocated_bytes": [item["deallocated_bytes"] for item in samples],
                       "allocation_region_peak_live_bytes": [item["region_peak_live_bytes"] for item in samples],
                       "allocation_peak_increment_bytes": [item["region_peak_live_bytes"] - item["live_bytes_before"]
                                                           for item in samples]})
    else:
        values.update({"allocation_calls": None, "deallocation_calls": None, "reallocation_calls": None,
                       "allocated_bytes": None, "deallocated_bytes": None,
                       "allocation_region_peak_live_bytes": None, "allocation_peak_increment_bytes": None})
    return values


def _summaries(values: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    stats, bootstrap = {}, {}
    for name, value in values.items():
        if isinstance(value, list) and all(type(item) is int for item in value):
            stats[name] = _percentiles(value)
            bootstrap[name] = _bootstrap_median(value)
        else:
            stats[name], bootstrap[name] = None, None
    return stats, bootstrap


def _relative(first: int | float, second: int | float) -> float | None:
    if first == 0:
        return 0.0 if second == 0 else None
    return (float(second) - float(first)) / abs(float(first)) * 100.0


def _regression_metric(before: dict[str, Any], after: dict[str, Any], name: str) -> dict[str, Any] | None:
    before_stats = before["percentiles"].get(name)
    after_stats = after["percentiles"].get(name)
    if before_stats is None or after_stats is None:
        return None
    percent = {point: _relative(before_stats[point], after_stats[point])
               for point in ("p50", "p95", "p99")}
    return {
        "before": {point: before_stats[point] for point in ("p50", "p95", "p99")},
        "after": {point: after_stats[point] for point in ("p50", "p95", "p99")},
        "after_minus_before_percent": percent,
        "regression_over_5_percent": any(value is not None and value > 5
                                          for value in percent.values()),
    }


def analyze_data(entries: list[dict[str, Any]], builds: dict[str, dict[str, Any]], *, pilot: bool = False,
                 protocol_hash: str = "") -> dict[str, Any]:
    rows = []
    for entry in entries:
        spec = entry["spec"]
        arm = _arm(spec["arm"])
        values = _vectors(entry, spec["role"], arm)
        stats, bootstrap = _summaries(values)
        rows.append({"label": spec["label"], "kind": spec["kind"], "repeat": spec["repeat"],
                     "phase": spec["phase"], "api": spec["api"], "role": spec["role"],
                     "arm": spec["arm"], "provider": spec["provider"],
                     "samples": spec["samples"], "warmups": spec["warmups"],
                     "terminal_sha256": entry["terminal_sha256"], "report_sha256": entry["report_sha256"],
                     "build_source": entry["started"]["source"],
                     "source_at_run": entry["started"]["source_at_run"],
                     "source_revision_at_run": entry["started"]["source_revision_at_run"],
                     "source_summary": [_analysis_read_summary(item["reads"])
                                        for item in entry["report"]["rows"]],
                     "sink_summary": [item["sink"] for item in entry["report"]["rows"]],
                     "budget_summary": [item["budget"] for item in entry["report"]["rows"]],
                     "output_sha256": sorted({item["output_sha256"] for item in entry["report"]["rows"]}),
                     "output_bytes": sorted({item["output_bytes"] for item in entry["report"]["rows"]}),
                     "raw_vectors": values, "percentiles": stats, "bootstrap_median_ci": bootstrap,
                     "resource": entry["resource"]})
    expected_repeats = (1,) if pilot else REPEATS
    # Every provider has a standalone row with its own tails and bootstrap CI.
    # The only paired comparison is the explicitly matched unmanaged path
    # across the retained before/after binaries. Managed rows have no before
    # baseline and therefore never enter this list.
    comparisons = []
    for repeat in expected_repeats:
        for role in ROLES:
            for arm in ARMS:
                before_matches = [row for row in rows
                                  if row["repeat"] == repeat and row["phase"] == "before"
                                  and row["api"] == UNMANAGED_API and row["role"] == role
                                  and row["arm"] == arm["name"]]
                after_matches = [row for row in rows
                                 if row["repeat"] == repeat and row["phase"] == "after"
                                 and row["api"] == UNMANAGED_API and row["role"] == role
                                 and row["arm"] == arm["name"]]
                require(len(before_matches) == len(after_matches) == 1,
                        f"unmanaged before/after pair is incomplete for repeat {repeat}/{role}/{arm['name']}")
                before, after = before_matches[0], after_matches[0]
                metrics = {}
                for name in ("elapsed_ns", "whole_child_rss_kbytes"):
                    metric = _regression_metric(before, after, name)
                    if metric is not None:
                        metrics[name] = metric
                comparisons.append({
                    "repeat": repeat, "role": role, "arm": arm["name"],
                    "provider": arm["provider"], "before_api": UNMANAGED_API,
                    "after_api": UNMANAGED_API, "before_phase": "before",
                    "after_phase": "after", "before_label": before["label"],
                    "after_label": after["label"], "metrics": metrics,
                    "scope": "paired unmanaged before/after regression review; managed API excluded",
                })
    repeat_variance = []
    if not pilot:
        for phase in PHASES:
            for api in ((UNMANAGED_API,) if phase == "before" else APIS):
                for role in ROLES:
                    for arm in ARMS:
                        matched = [row for row in rows if row["phase"] == phase
                                   and row["api"] == api and row["role"] == role
                                   and row["arm"] == arm["name"]]
                        require(len(matched) == 2 and {row["repeat"] for row in matched} == {1, 2},
                                f"repeat inventory incomplete for {phase}/{api}/{role}/{arm['name']}")
                        first, second = sorted(matched, key=lambda item: item["repeat"])
                        metrics = {}
                        for name in ("elapsed_ns", "whole_child_rss_kbytes", "source_logical_read_calls",
                                     "source_logical_read_returned_bytes", "sink_bytes", "output_bytes",
                                     "allocation_calls", "allocated_bytes", "allocation_peak_increment_bytes"):
                            if first["percentiles"].get(name) is None or second["percentiles"].get(name) is None:
                                continue
                            a, b = first["percentiles"][name], second["percentiles"][name]
                            changes = {p: _relative(a[p], b[p]) for p in ("p50", "p95", "p99")}
                            metrics[name] = {"repeat1": {p: a[p] for p in ("p50", "p95", "p99")},
                                             "repeat2": {p: b[p] for p in ("p50", "p95", "p99")},
                                             "relative_percent": changes,
                                             "flag_over_5_percent": any(x is not None and abs(x) > 5 for x in changes.values())}
                        repeat_variance.append({"phase": phase, "api": api, "role": role,
                                                "arm": arm["name"], "metrics": metrics,
                                                "scope": "descriptive repeat variance within one phase/API; no optimization regression claim"})
    return {"schema": ANALYSIS_SCHEMA, "version": VERSION, "change": CHANGE,
            "claim_authorized": False,
            "performance_claim": (
                "descriptive provider evidence with paired unmanaged before/after "
                "latency and RSS regression review; managed API has no before baseline"
            ),
            "scope": "per-provider opened-document edit/save evidence with source, sink, output, limits, cache and allocation custody",
            "pilot": pilot, "protocol_sha256": protocol_hash, "corpus": dict(CORPUS), "case": CASE,
            "roles": list(ROLES), "phases": list(PHASES), "edit_apis": list(APIS),
            "repeats": list(expected_repeats), "arms": [dict(arm) for arm in ARMS],
            "samples": PILOT_SAMPLES if pilot else FORMAL_SAMPLES,
            "warmups": PILOT_WARMUPS if pilot else FORMAL_WARMUPS,
            "builds": {key: _build_binding(builds[key]) for key in sorted(builds)},
            "rows": rows, "repeat_variance": repeat_variance, "comparisons": comparisons,
            "inventory": {"formal_processes": len(rows), "measured_samples": len(rows) * (PILOT_SAMPLES if pilot else FORMAL_SAMPLES),
                          "expected_formal_processes": len(formal_inventory(pilot=pilot)),
                          "expected_measured_samples": len(formal_inventory(pilot=pilot)) * (PILOT_SAMPLES if pilot else FORMAL_SAMPLES)}}


def capture_one(attempt: str, phase: str, api: str, role: str, repeat: int, arm_name: str, *, pilot: bool = False,
                build_dir: Path | None = None, timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS) -> Path:
    require(phase in PHASES and api in APIS and role in ROLES and arm_name in ARM_BY_NAME,
            "unknown capture phase, API, role, or arm")
    require(not (phase == "before" and api == MANAGED_API),
            "before phase has no successful managed timing row")
    require(repeat in ((1,) if pilot else REPEATS), "repeat is not valid for this capture")
    builds = load_builds(build_dir)
    protocol, protocol_hash = _load_protocol(builds)
    match = [item for item in formal_inventory(pilot=pilot)
             if item["phase"] == phase and item["api"] == api and item["role"] == role
             and item["repeat"] == repeat and item["arm"] == arm_name]
    require(len(match) == 1, "capture specification is not frozen")
    return _capture_one(dict(match[0], attempt=_attempt(attempt)),
                        builds[_build_key(phase, role)], protocol, protocol_hash, timeout_seconds)


def capture_all(attempt: str, build_dir: Path | None = None, *, pilot: bool = False,
                phase: str | None = None, api: str | None = None,
                repeat: int | None = None,
                timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS) -> None:
    require(phase is None or phase in PHASES, "unknown capture phase")
    require(api is None or api in APIS, "unknown capture API")
    require(repeat is None or repeat in ((1,) if pilot else REPEATS),
            "unknown capture repeat")
    if phase == "before":
        require(api in (None, UNMANAGED_API),
                "before phase has no successful managed timing row")
    builds = load_builds(build_dir)
    protocol, protocol_hash = _load_protocol(builds)
    items = [item for item in formal_inventory(pilot=pilot)
             if (phase is None or item["phase"] == phase)
             and (api is None or item["api"] == api)
             and (repeat is None or item["repeat"] == repeat)]
    require(items, "capture filter selected no frozen runs")
    for item in items:
        _capture_one(dict(item, attempt=_attempt(attempt)),
                     builds[_build_key(item["phase"], item["role"])], protocol,
                     protocol_hash, timeout_seconds)


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
            item.add_argument("--phase", choices=PHASES, required=True)
            item.add_argument("--api", choices=APIS, required=True)
            item.add_argument("--role", choices=ROLES, required=True)
            item.add_argument("--repeat", type=int, choices=REPEATS, required=True)
            item.add_argument("--arm", choices=tuple(ARM_BY_NAME), required=True)
        else:
            item.add_argument("--phase", choices=PHASES)
            item.add_argument("--api", choices=APIS)
            item.add_argument("--repeat", type=int, choices=REPEATS)
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
            build_files = [ROOT / f"build-{phase}-{role}.json"
                           for phase in PHASES for role in ROLES]
            builds = load_builds() if all(path.is_file() for path in build_files) else None
            print(json.dumps(protocol_value(builds), indent=2, sort_keys=True))
        elif args.command == "freeze":
            print(create_freeze())
        elif args.command == "capture":
            capture_one(args.attempt, args.phase, args.api, args.role, args.repeat, args.arm,
                        pilot=args.pilot,
                        build_dir=args.build_dir, timeout_seconds=args.timeout_seconds)
        elif args.command == "capture-all":
            capture_all(args.attempt, args.build_dir, pilot=args.pilot, phase=args.phase,
                        api=args.api, repeat=args.repeat, timeout_seconds=args.timeout_seconds)
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
