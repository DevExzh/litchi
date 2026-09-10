#!/usr/bin/env python3
"""Capture and verify the 0494 verified-cold DOCX edit lane.

The ordinary provider matrix in :mod:`measure` deliberately measures warm and
caller-controlled providers.  This module is a separate custody boundary for
``docx-edit-provider-cold``.  Each requested sample is a fresh invocation of
the Rust parent; that parent starts one fresh measured child after its final
residency probe.  A cold result is admitted only when the report proves an
aligned source, a successful ``fincore`` observation, zero resident/dirty/
writeback pages, and a positive process ``read_bytes`` delta.

This lane is descriptive.  Its analysis reports same-cold-cell repeat
variance and never treats a warm row and a cold row as a before/after timing
comparison.
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
import shutil
import signal
import statistics
import subprocess
import sys
from typing import Any, Callable, Iterable

from support import ENV, ENV_KEYS, REPO, ROOT, TEMP, meta, now, read, sha, snapshot, write

import measure as warm


VERSION = 1
CHANGE = 494
COLD_PROTOCOL_SCHEMA = "docx-edit-provider-cold-protocol-v1"
COLD_CAPTURE_SCHEMA = "docx-edit-provider-cold-capture-v1"
COLD_SAMPLE_SCHEMA = "docx-edit-provider-cold-sample-v1"
COLD_TERMINAL_SCHEMA = "docx-edit-provider-cold-terminal-v1"
COLD_ANALYSIS_SCHEMA = "docx-edit-provider-cold-analysis-v1"
COLD_VERIFICATION_SCHEMA = "docx-edit-provider-cold-verification-v1"
COLD_REPORT_SCHEMA = "docx_edit_provider_cold_v1"
COLD_CHILD_STATUS = "eligible"
BENCHMARK = "one DOCX paragraph replacement through verified cold FileSource open/commit/sequential publication"
COLD_REPORT_PATH = ROOT / "cold-protocol.json"
# Use the bundle's standard evidence roots so verify_bundle.py can discover
# cold aggregate reports and bind their analysis/verification receipts beside
# the warm lane.  Cold attempts use distinct attempt tokens (for example
# ``cold-formal1``) and cold-prefixed cell labels.
COLD_CAPTURE_ROOT = ROOT / "captures"
COLD_ANALYSIS_ROOT = ROOT / "analysis"
COLD_VERIFICATION_ROOT = ROOT / "verification"
COLD_MANAGED_ROOT = TEMP / "cold-managed"
CPU = warm.CPU
CPU_LOCK = warm.CPU_LOCK
ROLES = warm.ROLES
REPEATS = warm.REPEATS
FORMAL_SAMPLES = 30
FORMAL_WARMUPS = 0
PILOT_SAMPLES = 3
PILOT_WARMUPS = 0
DEFAULT_TIMEOUT_SECONDS = 180
TERM_GRACE_SECONDS = 10
PAGE_SIZE = 4096
SUPPORTED_FILESYSTEM_MAGICS = (0xEF53, 0x5846_5342, 0x9123_683E, 0xF2F5_2010, 0x2FC1_2FC1)
INELIGIBLE_STATUSES = frozenset({
    "ineligible_non_linux", "ineligible_linux_non64_bit",
    "ineligible_filesystem_unknown", "ineligible_filesystem_unsupported",
    "ineligible_fincore_unavailable", "ineligible_fincore_failed",
    "ineligible_fincore_invalid_json", "ineligible_fincore_multiple_records",
    "ineligible_fincore_path_mismatch", "ineligible_fincore_size_mismatch",
    "ineligible_fincore_metadata_unavailable", "ineligible_fincore_unrecognized_fallback",
    "ineligible_source_not_regular", "ineligible_source_empty",
    "ineligible_source_read_write_unavailable", "ineligible_source_hash_failed",
    "ineligible_source_page_size_unavailable", "ineligible_source_not_page_aligned",
    "ineligible_source_fsync_failed", "ineligible_source_advice_failed",
    "ineligible_source_resident", "ineligible_source_dirty", "ineligible_source_writeback",
    "ineligible_proc_io_unavailable", "ineligible_read_bytes_backwards",
    "ineligible_read_bytes_zero", "ineligible_prepared_query_control",
    "ineligible_source_alignment_unavailable", "ineligible_source_write_failed",
})
FINFORE_COMMAND = "fincore --json --bytes --output FILE,SIZE,RES,DIRTY,WRITEBACK --"
FINFORE_METHOD = "external_fincore_json_columns"
FINFORE_FALLBACK = "none"
FINFORE_ADVICE = "posix_fadvise_dontneed_accepted"
# Keep the correctly-spelled names as the public contract.  The short aliases
# above are retained only because the initial implementation used that local
# spelling internally.
FINCORE_COMMAND = FINFORE_COMMAND
FINCORE_METHOD = FINFORE_METHOD
FINCORE_FALLBACK = FINFORE_FALLBACK
FINCORE_ADVICE = FINFORE_ADVICE
PROCESS_METRICS_SCOPE = (
    "fresh measured child process /proc interval only; the parent and child "
    "supervisor process tree are excluded; the after snapshot includes procfs "
    "probe overhead"
)
RSS_SCOPE = (
    "fresh measured child /proc/self/status RSS and VmHWM only; this is "
    "process-local evidence and excludes the parent process tree"
)
READ_SCOPE = "FileSource positional ReadAt calls; logical source evidence"
READ_EVIDENCE_SCOPE = (
    "FileSource positional ReadAt calls; counters are logical ranges and process "
    "read_bytes is the cold admission proof"
)
TIMING_SCOPE = (
    "FileSource open + source-backed DOCX open + one paragraph edit "
    "staging/commit + sequential publication + commit/package/document drops; "
    "source version fences, commit diagnostics and source/candidate XML identity "
    "comparisons are inside; cold preparation, expected-value construction, sink "
    "reservation, output hashing, semantic/media verification and preflight patch "
    "oracles are outside"
)
SETUP_SCOPE = (
    "aligned source construction, source staging, expected publication, all "
    "semantic and patch oracles, counter reservation, and cold verifier setup "
    "are outside the operation clock"
)
COLD_CLAIM_SCOPE = (
    "external fincore page-cache residency/dirty/writeback proof plus positive "
    "process read_bytes; no physical-media claim"
)
PROVIDER_SCOPE = (
    "fresh-child verified cold regular-file FileSource for one source-backed "
    "DOCX paragraph replacement"
)
UNMANAGED_REASON = (
    "publish_docx_source_edit uses the compatibility constructor without an "
    "ExecutionContext"
)
SHA256_RE = warm.SHA256_RE
REVISION_RE = warm.REVISION_RE
ATTEMPT_RE = warm.ATTEMPT_RE


class ColdMeasureError(RuntimeError):
    """A fail-closed cold custody, schema, or recomputation error."""


def fail(message: str) -> None:
    raise ColdMeasureError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _exact(value: Any, keys: Iterable[str], path: str) -> None:
    require(isinstance(value, dict), f"{path}: expected object")
    expected = set(keys)
    actual = set(value)
    require(actual == expected,
            f"{path}: fields differ; expected {sorted(expected)}, got {sorted(actual)}")


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


def _timestamp(value: Any, path: str) -> _datetime.datetime:
    text = _text(value, path)
    try:
        parsed = _datetime.datetime.fromisoformat(text)
    except ValueError as error:
        fail(f"{path}: invalid timestamp: {error}")
    require(parsed.tzinfo is not None, f"{path}: timezone is required")
    return parsed


def _path(value: Any, path: str, *, base: Path = ROOT) -> Path:
    text = _text(value, path)
    result = Path(text)
    if not result.is_absolute():
        require(".." not in result.parts, f"{path}: parent traversal is forbidden")
        result = base / result
    return result


def _file_meta(path: Path, *, executable: bool = False) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing regular artifact: {path}")
    value = meta(path)
    _uint(value["bytes"], f"{path}.bytes")
    _hash(value["sha256"], f"{path}.sha256")
    if executable:
        require(os.access(path, os.X_OK), f"{path}: executable bit is absent")
    return {"path": str(path), **value}


def _json_hash(path: Path) -> str:
    return _hash(sha(path), str(path))


def _source_binding(value: Any, owner: str) -> dict[str, Any]:
    _exact(value, ("files", "path", "sha256"), owner)
    path = _path(value["path"], f"{owner}.path")
    _uint(value["files"], f"{owner}.files", positive=True)
    digest = _hash(value["sha256"], f"{owner}.sha256")
    actual = _file_meta(path)
    require(actual["sha256"] == digest, f"{owner}: manifest digest changed")
    manifest = _json(path)
    require(isinstance(manifest, dict) and len(manifest) == value["files"],
            f"{owner}: malformed source manifest")
    for name, item in manifest.items():
        require(isinstance(name, str) and name and not Path(name).is_absolute()
                and ".." not in Path(name).parts, f"{owner}: unsafe manifest path")
        _hash(item, f"{owner}.{name}")
    canonical = (json.dumps(manifest, indent=2, sort_keys=True) + "\n").encode()
    require(hashlib.sha256(canonical).hexdigest() == digest,
            f"{owner}: manifest content does not match binding")
    return {"files": value["files"], "path": str(path), "sha256": digest}


def _artifact(value: Any, owner: str, *, retained: bool = True) -> dict[str, Any]:
    _exact(value, ("bytes", "executable", "path", "sha256"), owner)
    path = _path(value["path"], f"{owner}.path")
    require(value["executable"] is True, f"{owner}.executable is false")
    _uint(value["bytes"], f"{owner}.bytes", positive=True)
    digest = _hash(value["sha256"], f"{owner}.sha256")
    if retained:
        actual = _file_meta(path, executable=True)
        require(actual["bytes"] == value["bytes"] and actual["sha256"] == digest,
                f"{owner}: artifact identity changed")
    return {"path": str(path), "bytes": value["bytes"], "sha256": digest,
            "executable": True}


def _build_binding(build: dict[str, Any]) -> dict[str, Any]:
    return {"receipt_sha256": build["receipt_sha256"], "binary": build["binary"],
            "source": build["source"], "gate": build["gate"],
            "command": build["command"], "environment": build["environment"],
            "git_revision": build["git_revision"]}


def _environment_binding() -> dict[str, str]:
    path = ROOT / "environment.json"
    return {"path": "environment.json", "sha256": _json_hash(path)}


def formal_inventory(*, pilot: bool = False) -> list[dict[str, Any]]:
    """Return the immutable cold cell inventory.

    A cell has one provider (verified cold file), while ``samples`` is the
    number of fresh parent invocations.  The Rust parent then creates one fresh
    measured child for every invocation.
    """

    samples = PILOT_SAMPLES if pilot else FORMAL_SAMPLES
    repeats = (1,) if pilot else REPEATS
    result: list[dict[str, Any]] = []
    for repeat in repeats:
        roles = ROLES if repeat == 1 else tuple(reversed(ROLES))
        for role in roles:
            result.append({
                "kind": "pilot" if pilot else "formal", "repeat": repeat,
                "role": role, "provider": "file-cold-verified",
                "samples": samples, "warmups": 0,
                "label": f"{'pilot-' if pilot else ''}r{repeat}-{role}-cold",
            })
    return result


def protocol_value(builds: dict[str, dict[str, Any]] | None = None,
                   warm_protocol: dict[str, Any] | None = None,
                   warm_protocol_hash: str | None = None) -> dict[str, Any]:
    """Return the cold protocol, optionally bound to retained warm builds."""

    if warm_protocol is None and builds is not None:
        warm_protocol, warm_protocol_hash = warm._load_protocol(builds)
    if warm_protocol is not None and warm_protocol_hash is None:
        warm_protocol_hash = _json_hash(ROOT / "protocol.json")
    value: dict[str, Any] = {
        "schema": COLD_PROTOCOL_SCHEMA, "version": VERSION, "change": CHANGE,
        "claim_authorized": False,
        "performance_claim": (
            "descriptive verified-cold source-backed edit evidence; repeat variance only; "
            "no warm-versus-cold timing claim"
        ),
        "report_schema": COLD_REPORT_SCHEMA,
        "child_status": COLD_CHILD_STATUS,
        "case": "docx_opened_document_one_paragraph_edit_save",
        "provider": "file-cold-verified", "cache_state": "cold-verified",
        "cpu": CPU, "cpu_lock": CPU_LOCK,
        "roles": list(ROLES), "repeats": list(REPEATS),
        "samples": FORMAL_SAMPLES, "warmups": FORMAL_WARMUPS,
        "pilot_samples": PILOT_SAMPLES, "pilot_warmups": PILOT_WARMUPS,
        "expected_formal_processes": len(ROLES) * len(REPEATS),
        "expected_formal_invocations": len(ROLES) * len(REPEATS) * FORMAL_SAMPLES,
        "expected_pilot_invocations": len(ROLES) * PILOT_SAMPLES,
        "formal_runs": formal_inventory(), "pilot_runs": formal_inventory(pilot=True),
        "corpus": dict(warm.CORPUS),
        "limits": {"page_size_bytes": PAGE_SIZE, "timeout_seconds": DEFAULT_TIMEOUT_SECONDS,
                   "term_grace_seconds": TERM_GRACE_SECONDS},
        "scopes": {"provider": PROVIDER_SCOPE, "timing": TIMING_SCOPE,
                    "setup": SETUP_SCOPE, "cold_claim": COLD_CLAIM_SCOPE,
                    "process_metrics": PROCESS_METRICS_SCOPE, "rss": RSS_SCOPE},
        "environment_artifact": _environment_binding(),
        "environment": {key: ENV[key] for key in ENV_KEYS},
        "source": None, "builds": None,
        "warm_protocol": None,
        "warm_protocol_sha256": None,
        "drivers": {"cold_measure.py": _json_hash(ROOT / "cold_measure.py"),
                     "test_cold_measure.py": _json_hash(ROOT / "test_cold_measure.py")},
    }
    if builds is not None:
        require(warm_protocol is not None and warm_protocol_hash is not None,
                "warm protocol is required for a build-bound cold protocol")
        value["source"] = builds["normal"]["source"]
        value["builds"] = {role: _build_binding(builds[role]) for role in ROLES}
        value["warm_protocol"] = {"path": "protocol.json", "sha256": warm_protocol_hash,
                                  "schema": warm_protocol["schema"],
                                  "version": warm_protocol["version"]}
        value["warm_protocol_sha256"] = warm_protocol_hash
    return value


def create_freeze() -> Path:
    builds = warm.load_builds()
    warm_protocol, warm_hash = warm._load_protocol(builds)
    expected = protocol_value(builds, warm_protocol, warm_hash)
    _write_or_match(COLD_REPORT_PATH, expected)
    return COLD_REPORT_PATH


def _load_protocol(builds: dict[str, dict[str, Any]] | None = None) -> tuple[dict[str, Any], str]:
    require(COLD_REPORT_PATH.is_file() and not COLD_REPORT_PATH.is_symlink(),
            f"cold protocol is missing: {COLD_REPORT_PATH}")
    builds = warm.load_builds() if builds is None else builds
    warm_protocol, warm_hash = warm._load_protocol(builds)
    value = _json(COLD_REPORT_PATH)
    expected = protocol_value(builds, warm_protocol, warm_hash)
    require(value == expected, f"{COLD_REPORT_PATH}: cold protocol differs")
    return value, _json_hash(COLD_REPORT_PATH)


def _attempt(value: str) -> str:
    require(isinstance(value, str) and ATTEMPT_RE.fullmatch(value) is not None,
            "attempt must be a path-safe token")
    return value


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


def _resource_values(path: Path) -> dict[str, Any]:
    raw = path.read_bytes()
    require(raw, f"{path}: GNU time receipt is empty")
    result: dict[str, Any] = {"raw_bytes": len(raw),
                              "raw_sha256": hashlib.sha256(raw).hexdigest()}
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
    _uint(result.get("maximum_resident_set_size_(kbytes)"),
          f"{path}: maximum RSS")
    result["scope"] = (
        "GNU time maximum resident set size for the parent invocation and its "
        "waited child; it is not operation-local RSS"
    )
    return result


def _command(spec: dict[str, Any], build: dict[str, Any], report: Path,
             resource: Path, revision: str) -> list[str]:
    require(spec["provider"] == "file-cold-verified", "cold cell provider differs")
    return [
        "/usr/bin/time", "-v", "-o", str(resource), "/usr/bin/taskset", "-c", str(CPU),
        str(build["binary"]["path"]), "docx-edit-provider-cold",
        "--provider", "file-cold-verified", "--filesystem-cache", "cold-verified",
        "--samples", "1", "--warmup", "0", "--source-revision", revision,
        "--output", str(report),
    ]


def _source_revision() -> str:
    try:
        revision = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=REPO,
                                           text=True).strip()
    except (OSError, subprocess.SubprocessError) as error:
        fail(f"cannot determine source revision: {error}")
    require(REVISION_RE.fullmatch(revision) is not None, "malformed source revision")
    return revision


def _run_root(attempt: str, label: str, index: int) -> Path:
    return COLD_MANAGED_ROOT / attempt / label / f"{index:03d}"


def _capture_root(attempt: str, label: str) -> Path:
    return COLD_CAPTURE_ROOT / attempt / label


def _overlap(offset: int, length: int, ranges: tuple[tuple[int, int], ...]) -> int:
    end = offset + length
    return sum(max(0, min(end, right) - max(offset, left)) for left, right in ranges)


def _read_counter(value: Any, path: str, *, aligned_bytes: int,
                  source_ranges: tuple[tuple[int, int], ...]) -> dict[str, Any]:
    fields = ("availability", "scope", "calls", "empty_calls", "requested_bytes",
              "returned_bytes", "short_reads", "min_request_bytes", "max_request_bytes",
              "traced_ranges", "requested_media_overlap_bytes",
              "returned_media_overlap_bytes", "ranges")
    _exact(value, fields, path)
    require(value["availability"] == "available", f"{path}: cold source counter unavailable")
    require(value["scope"] == READ_SCOPE, f"{path}.scope: scope differs")
    for key in ("calls", "empty_calls", "requested_bytes", "returned_bytes", "short_reads",
                "traced_ranges", "requested_media_overlap_bytes",
                "returned_media_overlap_bytes"):
        _uint(value[key], f"{path}.{key}")
    for key in ("min_request_bytes", "max_request_bytes"):
        if value[key] is not None:
            _uint(value[key], f"{path}.{key}", positive=True)
    require(value["returned_bytes"] <= value["requested_bytes"],
            f"{path}: returned bytes exceed requested")
    require(value["calls"] > 0 and value["requested_bytes"] > 0
            and value["returned_bytes"] > 0,
            f"{path}: logical source reads are not positive")
    ranges = value["ranges"]
    require(isinstance(ranges, list) and len(ranges) == value["calls"],
            f"{path}.ranges: trace does not conserve calls")
    requested = returned = short = requested_media = returned_media = 0
    minimum: int | None = None
    maximum: int | None = None
    for index, item in enumerate(ranges):
        _exact(item, ("offset", "requested", "returned"), f"{path}.ranges[{index}]")
        offset = _uint(item["offset"], f"{path}.ranges[{index}].offset")
        length = _uint(item["requested"], f"{path}.ranges[{index}].requested", positive=True)
        got = _uint(item["returned"], f"{path}.ranges[{index}].returned")
        require(got <= length and offset + length <= 2**64 - 1
                and offset + length <= aligned_bytes,
                f"{path}.ranges[{index}]: invalid range conservation")
        requested += length
        returned += got
        short += got < length
        requested_media += _overlap(offset, length, source_ranges)
        returned_media += min(_overlap(offset, got, source_ranges),
                              _overlap(offset, length, source_ranges))
        minimum = length if minimum is None else min(minimum, length)
        maximum = length if maximum is None else max(maximum, length)
    require(requested == value["requested_bytes"] and returned == value["returned_bytes"],
            f"{path}: range byte totals do not conserve")
    require(short == value["short_reads"] and value["traced_ranges"] == value["calls"],
            f"{path}: range short/read trace totals do not conserve")
    require(requested_media == value["requested_media_overlap_bytes"]
            and returned_media == value["returned_media_overlap_bytes"],
            f"{path}: media overlap totals do not conserve")
    require(returned_media <= requested_media <= requested,
            f"{path}: media overlap exceeds request")
    if value["calls"]:
        require(value["min_request_bytes"] == minimum and value["max_request_bytes"] == maximum,
                f"{path}: request bounds do not conserve")
    else:
        require(value["min_request_bytes"] is None and value["max_request_bytes"] is None,
                f"{path}: empty counter has request bounds")
    require(value["empty_calls"] >= 0 and value["calls"] <= warm.LIMITS["max_tracked_ranges"],
            f"{path}: trace exceeds configured capacity")
    require(aligned_bytes > 0, f"{path}: aligned source is empty")
    return dict(value)


def _check_limits(value: Any, path: str) -> dict[str, Any]:
    fields = {"read_limits", "cache_max_bytes", "cache_max_entries", "resource_budget",
              "max_tracked_ranges", "sink_max_write_bytes"}
    _exact(value, fields, path)
    read_fields = {
        "max_input_bytes": 512 * 1024 * 1024, "max_archive_members": 100_000,
        "max_archive_member_name_bytes": 4 * 1024, "max_archive_metadata_bytes": 64 * 1024 * 1024,
        "max_archive_compressed_bytes": 512 * 1024 * 1024, "max_archive_entry_bytes": 512 * 1024 * 1024,
        "max_archive_total_bytes": 2 * 1024 * 1024 * 1024, "max_parts": 100_000,
        "max_part_bytes": 512 * 1024 * 1024, "max_total_part_bytes": 512 * 1024 * 1024,
        "max_content_types_bytes": 8 * 1024 * 1024, "max_content_type_mappings": 100_000,
        "max_relationship_parts": 100_000, "max_relationship_xml_bytes": 8 * 1024 * 1024,
        "max_total_relationship_xml_bytes": 64 * 1024 * 1024, "max_relationships_per_part": 100_000,
        "max_total_relationships": 1_000_000, "max_relationship_graph_nodes": 100_000,
        "max_xml_events": 1_000_000, "max_total_relationship_xml_events": 8_000_000,
        "max_xml_depth": 256, "max_xml_attribute_bytes": 64 * 1024,
        "max_relationship_target_bytes": 4 * 1024,
    }
    _exact(value["read_limits"], read_fields, f"{path}.read_limits")
    require(value["read_limits"] == read_fields, f"{path}.read_limits: default policy differs")
    _uint(value["cache_max_bytes"], f"{path}.cache_max_bytes", positive=True)
    _uint(value["cache_max_entries"], f"{path}.cache_max_entries", positive=True)
    _uint(value["max_tracked_ranges"], f"{path}.max_tracked_ranges", positive=True)
    _uint(value["sink_max_write_bytes"], f"{path}.sink_max_write_bytes", positive=True)
    require(value["cache_max_bytes"] == warm.LIMITS["cache_max_bytes"]
            and value["cache_max_entries"] == warm.LIMITS["cache_max_entries"]
            and value["max_tracked_ranges"] == warm.LIMITS["max_tracked_ranges"]
            and value["sink_max_write_bytes"] == warm.LIMITS["sink_max_write_bytes"],
            f"{path}: package limits differ")
    budget = value["resource_budget"]
    _exact(budget, {"managed", "unmanaged_reason", "memory_bytes", "input_bytes",
                    "output_bytes", "objects", "depth", "work"}, f"{path}.resource_budget")
    require(budget["managed"] is False and budget["unmanaged_reason"] == UNMANAGED_REASON,
            f"{path}.resource_budget: unmanaged boundary differs")
    require(all(budget[key] is None for key in
                ("memory_bytes", "input_bytes", "output_bytes", "objects", "depth", "work")),
            f"{path}.resource_budget: unmanaged limits must remain absent")
    return dict(value)


def _validate_proof(value: Any, aligned_sha: str, aligned_bytes: int, path: str) -> bool:
    fields = {"status", "filesystem_magic", "page_size_bytes", "source_bytes", "source_pages",
              "aligned_source_bytes", "aligned_source_sha256", "fsync_completed", "advice",
              "fincore_size_bytes", "resident_bytes", "dirty_bytes", "writeback_bytes",
              "fincore_tool", "fincore_sha256", "fincore_version", "fincore_stderr_sha256",
              "fincore_stderr_bytes", "fincore_version_stderr_sha256",
              "fincore_version_stderr_bytes", "fincore_method", "fincore_fallback",
              "read_bytes_before", "read_bytes_after", "read_bytes_delta"}
    require(isinstance(value, dict), f"{path}: expected object")
    require(set(value) <= fields, f"{path}: proof contains unknown fields")
    require("status" in value, f"{path}.status: missing")
    status = _text(value["status"], f"{path}.status")
    require(status == "eligible" or status in INELIGIBLE_STATUSES,
            f"{path}: unknown cold status")
    if status != "eligible":
        # `cold_verified::Sample` omits every unavailable optional field on
        # early failures.  A later failure may retain a partial observation,
        # so validate each retained field without requiring omitted fields.
        optional_uints = {
            "filesystem_magic", "page_size_bytes", "source_bytes", "source_pages",
            "aligned_source_bytes", "fincore_size_bytes", "resident_bytes",
            "dirty_bytes", "writeback_bytes", "fincore_stderr_bytes",
            "fincore_version_stderr_bytes", "read_bytes_before", "read_bytes_after",
            "read_bytes_delta",
        }
        optional_hashes = {
            "aligned_source_sha256", "fincore_sha256", "fincore_stderr_sha256",
            "fincore_version_stderr_sha256",
        }
        optional_text = {"advice", "fincore_tool", "fincore_version", "fincore_method",
                         "fincore_fallback"}
        for key, item in value.items():
            if key == "status" or item is None:
                continue
            if key in optional_uints:
                _uint(item, f"{path}.{key}")
            elif key in optional_hashes:
                _hash(item, f"{path}.{key}")
            elif key in optional_text:
                _text(item, f"{path}.{key}")
            elif key == "fsync_completed":
                require(isinstance(item, bool), f"{path}.{key}: expected boolean")
            else:
                fail(f"{path}.{key}: unsupported retained proof field")
        return False
    _exact(value, fields, path)
    for key in ("filesystem_magic", "page_size_bytes", "source_bytes", "source_pages",
                "aligned_source_bytes", "fincore_size_bytes", "resident_bytes",
                "dirty_bytes", "writeback_bytes", "fincore_stderr_bytes",
                "fincore_version_stderr_bytes", "read_bytes_before", "read_bytes_after",
                "read_bytes_delta"):
        _uint(value[key], f"{path}.{key}")
    require(value["filesystem_magic"] in SUPPORTED_FILESYSTEM_MAGICS,
            f"{path}: filesystem is outside the admitted block-filesystem set")
    require(value["page_size_bytes"] == PAGE_SIZE
            and value["aligned_source_bytes"] == aligned_bytes
            and value["source_bytes"] == aligned_bytes
            and value["source_pages"] == aligned_bytes // PAGE_SIZE
            and value["fincore_size_bytes"] == aligned_bytes,
            f"{path}: alignment/size proof differs")
    require(_hash(value["aligned_source_sha256"], f"{path}.aligned_source_sha256") == aligned_sha,
            f"{path}: aligned source hash differs")
    require(value["fsync_completed"] is True and value["advice"] == FINFORE_ADVICE,
            f"{path}: fsync/advice proof failed")
    require(value["resident_bytes"] == 0 and value["dirty_bytes"] == 0
            and value["writeback_bytes"] == 0,
            f"{path}: source residency proof is not cold")
    require(value["fincore_tool"] == "fincore"
            and _hash(value["fincore_sha256"], f"{path}.fincore_sha256")
            and _text(value["fincore_version"], f"{path}.fincore_version")
            and _hash(value["fincore_stderr_sha256"], f"{path}.fincore_stderr_sha256")
            and _hash(value["fincore_version_stderr_sha256"], f"{path}.fincore_version_stderr_sha256")
            and value["fincore_method"] == FINFORE_METHOD
            and value["fincore_fallback"] == FINFORE_FALLBACK,
            f"{path}: fincore provenance differs")
    require(value["read_bytes_after"] >= value["read_bytes_before"]
            and value["read_bytes_delta"] == value["read_bytes_after"] - value["read_bytes_before"]
            and value["read_bytes_delta"] > 0,
            f"{path}: process read_bytes proof is not positive")
    return True


def _check_version(value: Any, path: str) -> None:
    _exact(value, ("id", "revision"), path)
    _uint(value["id"], f"{path}.id", positive=True)
    _uint(value["revision"], f"{path}.revision")


def _check_process_metrics(value: Any, path: str) -> dict[str, Any]:
    fields = {"rchar", "wchar", "read_bytes", "write_bytes", "cancelled_write_bytes",
              "syscr", "syscw", "minor_faults", "major_faults", "user_cpu_ticks",
              "system_cpu_ticks", "clock_ticks_per_second", "voluntary_context_switches",
              "nonvoluntary_context_switches", "rss_bytes", "peak_rss_bytes"}
    _exact(value, fields, path)
    for key in fields:
        _uint(value[key], f"{path}.{key}")
    require(value["clock_ticks_per_second"] > 0 and value["read_bytes"] > 0,
            f"{path}: process-I/O proof is not positive")
    return dict(value)


def _check_row(row: Any, role: str, *, aligned_sha: str, aligned_bytes: int,
               expected_output_sha: str, expected_output_bytes: int,
               expected_materializations: int, source_ranges: tuple[tuple[int, int], ...],
               path: str) -> dict[str, Any]:
    require(isinstance(row, dict), f"{path}: row is missing")
    fields = {"sample_index", "child_process_id", "latency_ns", "output_bytes",
              "output_sha256", "materializations", "commit_changed", "commit_operations",
              "source_version_before", "source_version_after", "source_version_unchanged",
              "reads", "sink", "cache", "process_metrics", "process_metrics_scope",
              "rss_scope", "oracles"}
    if role == "allocator":
        fields.add("allocation")
    _exact(row, fields, path)
    _uint(row["sample_index"], f"{path}.sample_index")
    require(row["sample_index"] == 0, f"{path}.sample_index: fresh invocation must have index zero")
    _uint(row["child_process_id"], f"{path}.child_process_id", positive=True)
    _uint(row["latency_ns"], f"{path}.latency_ns", positive=True)
    _uint(row["output_bytes"], f"{path}.output_bytes", positive=True)
    require(_hash(row["output_sha256"], f"{path}.output_sha256") == expected_output_sha
            and row["output_bytes"] == expected_output_bytes,
            f"{path}: output identity differs")
    _uint(row["materializations"], f"{path}.materializations", positive=True)
    _uint(row["commit_operations"], f"{path}.commit_operations", positive=True)
    require(row["commit_changed"] is True and row["commit_operations"] == 1,
            f"{path}: commit oracle failed")
    _check_version(row["source_version_before"], f"{path}.source_version_before")
    _check_version(row["source_version_after"], f"{path}.source_version_after")
    require(row["source_version_unchanged"] is True
            and row["source_version_before"] == row["source_version_after"],
            f"{path}: source version changed")
    reads = row["reads"]
    _exact(reads, {"source", "scope"}, f"{path}.reads")
    _read_counter(reads["source"], f"{path}.reads.source", aligned_bytes=aligned_bytes,
                  source_ranges=source_ranges)
    require(reads["scope"] == READ_EVIDENCE_SCOPE, f"{path}.reads.scope: scope differs")
    sink = row["sink"]
    _exact(sink, {"accepted_bytes", "write_calls", "largest_write"}, f"{path}.sink")
    accepted = _uint(sink["accepted_bytes"], f"{path}.sink.accepted_bytes", positive=True)
    calls = _uint(sink["write_calls"], f"{path}.sink.write_calls", positive=True)
    largest = _uint(sink["largest_write"], f"{path}.sink.largest_write", positive=True)
    require(accepted == expected_output_bytes and row["output_bytes"] == accepted
            and largest <= accepted and largest <= warm.LIMITS["sink_max_write_bytes"]
            and calls > 0, f"{path}: sink accounting differs")
    cache = row["cache"]
    _exact(cache, {"successful_loads", "expected_successful_loads",
                   "exactly_one_main_part_materialization"}, f"{path}.cache")
    for key in ("successful_loads", "expected_successful_loads"):
        _uint(cache[key], f"{path}.cache.{key}", positive=True)
    require(cache["successful_loads"] == cache["expected_successful_loads"]
            == row["materializations"] == expected_materializations == 1
            and cache["exactly_one_main_part_materialization"] is True,
            f"{path}: cache materialization oracle failed")
    _check_process_metrics(row["process_metrics"], f"{path}.process_metrics")
    require(row["process_metrics_scope"] == PROCESS_METRICS_SCOPE
            and row["rss_scope"] == RSS_SCOPE, f"{path}: process scope differs")
    oracles = row["oracles"]
    _exact(oracles, {"commit_identity_verified", "output_exact_bytes", "semantic_reopen",
                     "unchanged_media_preserved", "source_version_unchanged",
                     "cache_load_count", "logical_source_reads_positive",
                     "patch_oracles"}, f"{path}.oracles")
    for key in ("commit_identity_verified", "output_exact_bytes", "semantic_reopen",
                "unchanged_media_preserved", "source_version_unchanged", "cache_load_count",
                "logical_source_reads_positive"):
        require(oracles[key] is True, f"{path}.oracles.{key}: failed")
    patch = oracles["patch_oracles"]
    _exact(patch, {"scope", "replay_forward", "inverse_restores_source",
                   "stale_target_refused", "foreign_source_refused"}, f"{path}.oracles.patch_oracles")
    for key in ("replay_forward", "inverse_restores_source", "stale_target_refused",
                "foreign_source_refused"):
        require(patch[key] is True, f"{path}.oracles.patch_oracles.{key}: failed")
    require(patch["scope"] == "untimed cold-child preflight commit patch oracles",
            f"{path}.oracles.patch_oracles.scope: scope differs")
    if role == "allocator":
        allocation = row["allocation"]
        fields = {"status", "scope", "allocation_calls", "deallocation_calls",
                  "reallocation_calls", "failed_allocation_calls", "allocated_bytes",
                  "deallocated_bytes", "live_bytes_before", "live_bytes_after",
                  "peak_live_bytes_before", "peak_live_bytes_after", "region_peak_live_bytes"}
        _exact(allocation, fields, f"{path}.allocation")
        require(allocation["status"] == "measured"
                and allocation["scope"] == warm.SAMPLE_ALLOCATION_SCOPE,
                f"{path}.allocation: allocator evidence unavailable")
        for key in fields - {"status", "scope"}:
            _uint(allocation[key], f"{path}.allocation.{key}")
        require(allocation["region_peak_live_bytes"] >= allocation["live_bytes_before"],
                f"{path}.allocation: peak ordering differs")
    return {"latency_ns": row["latency_ns"], "child_process_id": row["child_process_id"],
            "process_rss_bytes": row["process_metrics"]["rss_bytes"],
            "process_peak_rss_bytes": row["process_metrics"]["peak_rss_bytes"],
            "process_read_bytes": row["process_metrics"]["read_bytes"],
            "source_calls": row["reads"]["source"]["calls"],
            "source_returned_bytes": row["reads"]["source"]["returned_bytes"],
            "output_bytes": row["output_bytes"],
            "allocation": row.get("allocation")}


def validate_report(path: Path, *, role: str, source_revision: str | None = None,
                    allow_ineligible: bool = True) -> dict[str, Any]:
    """Strictly validate one Rust parent report.

    The report is one fresh parent invocation, hence it must have
    ``samples=1`` and ``warmup=0``.  An ineligible cold proof is valid custody
    evidence but has zero rows and can never count as a formal measurement.
    The report schema has no executable identity field; the retained
    started/terminal receipts bind the outer launched role binary instead.
    """

    require(role in ROLES, f"unknown cold role {role}")
    value = _json(path)
    fields = {"limits", "schema", "benchmark", "provider_scope", "timing_scope",
              "setup_scope", "cold_claim_scope", "process_metrics_scope", "rss_scope",
              "provider", "cache_state", "filesystem_root_selected", "source_archive_sha256",
              "source_archive_bytes", "aligned_source_sha256", "aligned_source_bytes",
              "expected_output_sha256", "expected_output_bytes", "source_revision", "corpus",
              "preflight", "cold_verified_status", "cold_verified_samples",
              "cold_verified_fincore_command", "warmup", "samples", "rows"}
    _exact(value, fields, str(path))
    require(value["schema"] == COLD_REPORT_SCHEMA and value["benchmark"] == BENCHMARK
            and value["provider"] == "file-cold-verified"
            and value["cache_state"] == "cold-verified"
            and value["filesystem_root_selected"] is False,
            f"{path}: cold report identity differs")
    require(value["provider_scope"] == PROVIDER_SCOPE and value["timing_scope"] == TIMING_SCOPE
            and value["setup_scope"] == SETUP_SCOPE and value["cold_claim_scope"] == COLD_CLAIM_SCOPE
            and value["process_metrics_scope"] == PROCESS_METRICS_SCOPE
            and value["rss_scope"] == RSS_SCOPE,
            f"{path}: cold report scope differs")
    _check_limits(value["limits"], f"{path}.limits")
    require(value["corpus"] == warm.CORPUS_MANIFEST, f"{path}: corpus identity differs")
    require(value["source_archive_bytes"] == warm.CORPUS["archive_bytes"]
            and value["source_archive_sha256"] == warm.CORPUS["archive_sha256"],
            f"{path}: source archive identity differs")
    _hash(value["source_archive_sha256"], f"{path}.source_archive_sha256")
    _uint(value["source_archive_bytes"], f"{path}.source_archive_bytes", positive=True)
    aligned_value = value["aligned_source_sha256"]
    aligned_size = value["aligned_source_bytes"]
    if aligned_value is None or aligned_size is None:
        require(aligned_value is None and aligned_size is None,
                f"{path}: aligned source identity is only partially present")
        aligned_sha, aligned_bytes = None, None
    else:
        aligned_sha = _hash(aligned_value, f"{path}.aligned_source_sha256")
        aligned_bytes = _uint(aligned_size, f"{path}.aligned_source_bytes", positive=True)
        require(aligned_bytes % PAGE_SIZE == 0 and aligned_bytes >= value["source_archive_bytes"],
                f"{path}: aligned source is not page-aligned")
    output_value = value["expected_output_sha256"]
    output_size = value["expected_output_bytes"]
    if output_value is None or output_size is None:
        require(output_value is None and output_size is None,
                f"{path}: expected output identity is only partially present")
        expected_output_sha, expected_output_bytes = None, None
    else:
        expected_output_sha = _hash(output_value, f"{path}.expected_output_sha256")
        expected_output_bytes = _uint(output_size, f"{path}.expected_output_bytes", positive=True)
        if aligned_sha is not None and aligned_bytes is not None:
            require(expected_output_sha != aligned_sha and expected_output_bytes >= aligned_bytes,
                    f"{path}: expected output does not describe a changed archive")
    revision = _text(value["source_revision"], f"{path}.source_revision")
    require(REVISION_RE.fullmatch(revision) is not None, f"{path}: malformed source revision")
    if source_revision is not None:
        require(revision == source_revision, f"{path}: source revision differs")
    _uint(value["warmup"], f"{path}.warmup")
    _uint(value["samples"], f"{path}.samples", positive=True)
    require(value["warmup"] == 0 and value["samples"] == 1,
            f"{path}: cold invocation must have one sample and no warmup")
    _text(value["cold_verified_fincore_command"], f"{path}.cold_verified_fincore_command")
    require(value["cold_verified_fincore_command"] == FINFORE_COMMAND,
            f"{path}: fincore command differs")
    proofs = value["cold_verified_samples"]
    require(isinstance(proofs, list) and len(proofs) == 1, f"{path}: proof sample count differs")
    eligible = _validate_proof(proofs[0], aligned_sha or "", aligned_bytes or 0,
                               f"{path}.cold_verified_samples[0]")
    if not eligible:
        require(value["cold_verified_status"] == proofs[0]["status"] and value["rows"] == [],
                f"{path}: ineligible cold report must contain zero rows")
        require(value["preflight"] is None or isinstance(value["preflight"], dict),
                f"{path}: ineligible preflight must be absent or an object")
        require(allow_ineligible, f"{path}: ineligible cold sample is not admissible")
        return {"report": value, "eligible": False, "rows": [], "proof": proofs[0]}
    require(aligned_sha is not None and aligned_bytes is not None
            and expected_output_sha is not None and expected_output_bytes is not None,
            f"{path}: eligible cold report is missing aligned/output identity")
    preflight = value["preflight"]
    preflight_fields = {"expected_materializations", "source_document_xml_sha256",
                        "candidate_document_xml_sha256", "output_exact_source_changed",
                        "semantic_reopen_verified", "unchanged_media_preserved",
                        "replay_forward_verified", "inverse_restores_source_verified",
                        "stale_target_refusal_verified", "foreign_source_refusal_verified"}
    _exact(preflight, preflight_fields, f"{path}.preflight")
    expected_materializations = _uint(preflight["expected_materializations"],
                                      f"{path}.preflight.expected_materializations", positive=True)
    for key in ("source_document_xml_sha256", "candidate_document_xml_sha256"):
        _hash(preflight[key], f"{path}.preflight.{key}")
    for key in preflight_fields - {"expected_materializations", "source_document_xml_sha256",
                                   "candidate_document_xml_sha256"}:
        require(preflight[key] is True, f"{path}.preflight.{key}: failed")
    require(value["cold_verified_status"] == "eligible", f"{path}: status differs from proof")
    rows = value["rows"]
    require(isinstance(rows, list) and len(rows) == 1, f"{path}: eligible report must contain one row")
    media = tuple((left, right) for left, right in warm.MEDIA_RANGES)
    checked = _check_row(rows[0], role, aligned_sha=aligned_sha, aligned_bytes=aligned_bytes,
                         expected_output_sha=expected_output_sha,
                         expected_output_bytes=expected_output_bytes,
                         expected_materializations=expected_materializations,
                         source_ranges=media, path=f"{path}.rows[0]")
    return {"report": value, "eligible": True, "rows": [checked], "proof": proofs[0]}


def _check_sample_terminal(started: dict[str, Any], terminal: dict[str, Any],
                           spec: dict[str, Any], build: dict[str, Any],
                           protocol_hash: str, directory: Path) -> dict[str, Any]:
    started_fields = {"schema", "version", "status", "attempt", "run", "protocol", "build",
                      "binary", "source", "argv", "cwd", "environment", "environment_artifact",
                      "driver_bindings", "driver_sha256", "test_driver_sha256", "started_utc",
                      "timeout_seconds", "tmpdir", "parent_tree_rss_scope"}
    terminal_fields = started_fields | {"exit_code", "timed_out", "termination", "finished_utc",
                                       "artifacts", "missing_artifacts", "source_before", "source_after",
                                       "source_unchanged", "resource", "cold_status", "eligible",
                                       "validation_error", "started_artifact", "cleanup"}
    _exact(started, started_fields, f"{directory}/started.json")
    _exact(terminal, terminal_fields, f"{directory}/terminal.json")
    require(started["schema"] == COLD_SAMPLE_SCHEMA and terminal["schema"] == COLD_TERMINAL_SCHEMA
            and started["version"] == VERSION and terminal["version"] == VERSION,
            f"{directory}: custody schema differs")
    require(terminal["status"] == "pass", f"{directory}: terminal did not pass")
    require(started["status"] == "running" and started["attempt"] == spec["attempt"],
            f"{directory}: started status/spec differs")
    for field in ("version", "attempt", "run", "protocol", "build", "binary", "source", "argv",
                  "cwd", "environment", "environment_artifact", "driver_bindings", "driver_sha256",
                  "test_driver_sha256", "started_utc", "timeout_seconds", "tmpdir",
                  "parent_tree_rss_scope"):
        require(terminal[field] == started[field],
                f"{directory}: terminal {field} changed")
    require(started["run"] == {key: spec[key] for key in
                                ("kind", "repeat", "role", "provider", "samples", "warmups", "label")},
            f"{directory}: run specification differs")
    require(started["protocol"] == {"path": "cold-protocol.json", "sha256": protocol_hash},
            f"{directory}: cold protocol binding differs")
    require(started["build"] == {"path": build["path"], "sha256": build["receipt_sha256"],
                                  "source": build["source"]}
            and started["binary"] == build["binary"] and started["source"] == build["source"],
            f"{directory}: build/source binding differs")
    require(started["cwd"] == str(REPO) and started["driver_bindings"] ==
            {"cold_measure.py": _json_hash(ROOT / "cold_measure.py"),
             "test_cold_measure.py": _json_hash(ROOT / "test_cold_measure.py")}
            and started["driver_sha256"] == _json_hash(ROOT / "cold_measure.py")
            and started["test_driver_sha256"] == _json_hash(ROOT / "test_cold_measure.py"),
            f"{directory}: cold helper custody differs")
    tmp_root = _path(started["tmpdir"], f"{directory}.tmpdir")
    require(tmp_root == _run_root(spec["attempt"], spec["label"], int(directory.name)),
            f"{directory}: private scratch path differs")
    cleanup = terminal["cleanup"]
    _exact(cleanup, {"path", "removed", "remaining"}, f"{directory}.cleanup")
    require(cleanup["path"] == str(tmp_root) and cleanup["removed"] is True
            and isinstance(cleanup["remaining"], list) and not tmp_root.exists(),
            f"{directory}: private scratch cleanup not proven")
    require(started["argv"] == _command(spec, build, directory / "report.json",
                                      directory / "resource.txt", build["git_revision"]),
            f"{directory}: capture command differs")
    expected_env = {key: ENV[key] for key in ENV_KEYS}
    expected_env["TMPDIR"] = str(tmp_root)
    require(started["environment"] == expected_env and terminal["environment"] == expected_env
            and started["environment_artifact"] ==
            {"path": "environment.json", "sha256": _json_hash(ROOT / "environment.json")},
            f"{directory}: child environment differs")
    require(started["argv"] == terminal["argv"] and started["timeout_seconds"] == DEFAULT_TIMEOUT_SECONDS
            and terminal["timeout_seconds"] == DEFAULT_TIMEOUT_SECONDS
            and started["parent_tree_rss_scope"] ==
            "GNU time maximum resident set size for the parent invocation and its waited child; it is not operation-local RSS",
            f"{directory}: argv/timeout/RSS scope differs")
    _uint(terminal["exit_code"], f"{directory}.exit_code")
    require(terminal["exit_code"] == 0 and terminal["timed_out"] is False
            and terminal["termination"] is None and terminal["source_before"] == started["source"]
            and terminal["source_after"] == started["source"] and terminal["source_unchanged"] is True
            and terminal["missing_artifacts"] == [],
            f"{directory}: terminal process custody failed")
    require(_timestamp(terminal["finished_utc"], f"{directory}.finished_utc") >
            _timestamp(started["started_utc"], f"{directory}.started_utc"),
            f"{directory}: chronology invalid")
    artifacts = terminal["artifacts"]
    _exact(artifacts, ("report.json", "resource.txt", "stdout.txt", "stderr.txt"),
           f"{directory}.artifacts")
    for name, item in artifacts.items():
        _exact(item, ("bytes", "path", "sha256"), f"{directory}.artifacts.{name}")
        artifact_path = _path(item["path"], f"{directory}.artifacts.{name}.path")
        require(artifact_path == directory / name, f"{directory}: artifact escaped sample")
        actual = _file_meta(artifact_path)
        require(actual["bytes"] == item["bytes"] and actual["sha256"] == item["sha256"],
                f"{directory}: artifact changed")
    require(terminal["started_artifact"] == _file_meta(directory / "started.json"),
            f"{directory}: started receipt hash differs")
    resource = terminal["resource"]
    require(isinstance(resource, dict) and resource["scope"] ==
            "GNU time maximum resident set size for the parent invocation and its waited child; it is not operation-local RSS",
            f"{directory}: resource scope missing")
    actual_resource = _resource_values(directory / "resource.txt")
    require(resource == actual_resource, f"{directory}: resource receipt differs")
    _uint(resource.get("maximum_resident_set_size_(kbytes)"), f"{directory}: resource RSS")
    report_path = directory / "report.json"
    result = validate_report(report_path, role=spec["role"], source_revision=build["git_revision"],
                             allow_ineligible=True)
    require(terminal["cold_status"] == result["proof"]["status"]
            and terminal["eligible"] is result["eligible"],
            f"{directory}: terminal cold status differs from report")
    require(not terminal["validation_error"], f"{directory}: validation error recorded")
    require({item.name for item in directory.iterdir()} ==
            {"started.json", "terminal.json", "report.json", "resource.txt", "stdout.txt", "stderr.txt"},
            f"{directory}: sample file inventory differs")
    return {"directory": directory, "started": started, "terminal": terminal,
            "resource": resource, **result}


def _remove_private(path: Path) -> dict[str, Any]:
    try:
        relative = path.relative_to(COLD_MANAGED_ROOT)
    except ValueError:
        fail(f"private root escaped cold manager: {path}")
    require(len(relative.parts) == 3, f"private root escaped cold manager: {path}")
    if not path.exists():
        return {"path": str(path), "removed": False, "remaining": []}
    require(path.is_dir() and not path.is_symlink(), f"private root is unsafe: {path}")
    remaining = [str(item) for item in path.rglob("*")]
    shutil.rmtree(path)
    # Keep the owned TEMP root compatible with the canonical 0494 cleanup
    # receipt: after a sample, no empty cold-manager hierarchy is retained.
    for ancestor in (path.parent, path.parent.parent, COLD_MANAGED_ROOT):
        if ancestor.is_dir() and not ancestor.is_symlink() and not any(ancestor.iterdir()):
            ancestor.rmdir()
    return {"path": str(path), "removed": True, "remaining": remaining}


def _run_sample(spec: dict[str, Any], build: dict[str, Any], protocol_hash: str,
                index: int, timeout_seconds: int) -> dict[str, Any]:
    attempt = _attempt(spec["attempt"])
    cell = _capture_root(attempt, spec["label"])
    directory = cell / "samples" / f"{index:03d}"
    require(not directory.exists(), f"refusing to replace immutable sample: {directory}")
    directory.mkdir(parents=True, exist_ok=False)
    run_root = _run_root(attempt, spec["label"], index)
    require(not run_root.exists(), f"refusing to replace private scratch: {run_root}")
    run_root.mkdir(parents=True, exist_ok=False)
    report, resource = directory / "report.json", directory / "resource.txt"
    stdout, stderr = directory / "stdout.txt", directory / "stderr.txt"
    stdout.touch(); stderr.touch()
    revision = _source_revision()
    require(revision == build["git_revision"], "capture HEAD differs from retained build")
    current = warm._normalized_snapshot()
    require(current == build["source"], "capture source differs from retained build")
    argv = _command(spec, build, report, resource, revision)
    tmp_root = run_root
    environment = {key: ENV[key] for key in ENV_KEYS}
    environment["TMPDIR"] = str(tmp_root)
    started = {
        "schema": COLD_SAMPLE_SCHEMA, "version": VERSION, "status": "running",
        "attempt": attempt,
        "run": {key: spec[key] for key in
                 ("kind", "repeat", "role", "provider", "samples", "warmups", "label")},
        "protocol": {"path": "cold-protocol.json", "sha256": protocol_hash},
        "build": {"path": build["path"], "sha256": build["receipt_sha256"], "source": build["source"]},
        "binary": build["binary"], "source": current, "argv": argv, "cwd": str(REPO),
        "environment": environment, "environment_artifact": {"path": "environment.json",
                                                               "sha256": _json_hash(ROOT / "environment.json")},
        "driver_bindings": {"cold_measure.py": _json_hash(ROOT / "cold_measure.py"),
                             "test_cold_measure.py": _json_hash(ROOT / "test_cold_measure.py")},
        "driver_sha256": _json_hash(ROOT / "cold_measure.py"),
        "test_driver_sha256": _json_hash(ROOT / "test_cold_measure.py"),
        "started_utc": now(), "timeout_seconds": timeout_seconds, "tmpdir": str(tmp_root),
        "parent_tree_rss_scope": (
            "GNU time maximum resident set size for the parent invocation and its waited child; "
            "it is not operation-local RSS"
        ),
    }
    _write(directory / "started.json", started)
    process: subprocess.Popen[bytes] | None = None
    timed_out = False
    termination: str | None = None
    exit_code: int | None = None
    validation_error = ""
    run_environment = dict(ENV)
    run_environment["TMPDIR"] = str(tmp_root)
    try:
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
        validation_error = f"{type(error).__name__}: {error}"
        if process is not None and process.poll() is None:
            termination = _kill_group(process)
            exit_code = process.returncode
    resource_value: dict[str, Any] = {}
    result: dict[str, Any] | None = None
    if exit_code == 0 and not timed_out and termination is None and not validation_error:
        try:
            resource_value = _resource_values(resource)
            result = validate_report(directory / "report.json", role=spec["role"],
                                     source_revision=revision, allow_ineligible=True)
        except (ColdMeasureError, OSError, ValueError) as error:
            validation_error = str(error)
    try:
        after = warm._normalized_snapshot()
        unchanged = after == current
    except (ColdMeasureError, OSError, ValueError) as error:
        after, unchanged = {"error": f"{type(error).__name__}: {error}"}, False
    cleanup = _remove_private(run_root)
    eligible = bool(result and result["eligible"])
    cold_status = result["proof"]["status"] if result else "failed"
    artifacts = {name: _file_meta(directory / name) for name in
                 ("report.json", "resource.txt", "stdout.txt", "stderr.txt")
                 if (directory / name).is_file()}
    missing = [name for name in ("report.json", "resource.txt", "stdout.txt", "stderr.txt")
               if not (directory / name).is_file()]
    terminal = {
        "schema": COLD_TERMINAL_SCHEMA, "version": VERSION, "status": "pass" if
        (exit_code == 0 and not timed_out and termination is None and not validation_error
         and unchanged and not missing and result is not None) else "failed",
        "attempt": attempt, "run": started["run"], "protocol": started["protocol"],
        "build": started["build"], "binary": started["binary"], "source": started["source"],
        "argv": argv, "cwd": str(REPO), "environment": environment,
        "environment_artifact": started["environment_artifact"],
        "driver_bindings": started["driver_bindings"], "driver_sha256": started["driver_sha256"],
        "test_driver_sha256": started["test_driver_sha256"], "started_utc": started["started_utc"],
        "timeout_seconds": timeout_seconds, "tmpdir": str(tmp_root),
        "parent_tree_rss_scope": started["parent_tree_rss_scope"], "exit_code": exit_code,
        "timed_out": timed_out, "termination": termination, "finished_utc": now(),
        "artifacts": artifacts, "missing_artifacts": missing, "source_before": current,
        "source_after": after, "source_unchanged": unchanged, "resource": resource_value,
        "cold_status": cold_status, "eligible": eligible, "validation_error": validation_error,
        "started_artifact": _file_meta(directory / "started.json"), "cleanup": cleanup,
    }
    _write(directory / "terminal.json", terminal)
    if terminal["status"] != "pass":
        fail(f"{spec['label']} sample {index} failed; receipt retained at {directory / 'terminal.json'}")
    return {"directory": directory, "started": started, "terminal": terminal,
            "resource": resource_value, **(result or {"eligible": False, "rows": [],
                                                     "proof": {"status": "failed"}, "report": None})}


def _write_cell_receipts(spec: dict[str, Any], build: dict[str, Any], protocol_hash: str,
                         entries: list[dict[str, Any]], *, stopped: bool) -> Path:
    cell = _capture_root(spec["attempt"], spec["label"])
    reports = []
    rows = []
    pids: set[int] = set()
    for index, entry in enumerate(entries):
        report_path = entry["directory"] / "report.json"
        report_meta = _file_meta(report_path)
        report = entry["report"]
        row = entry["rows"][0] if entry["rows"] else None
        if row is not None:
            require(row["child_process_id"] not in pids,
                    f"{cell}: duplicate measured child PID")
            pids.add(row["child_process_id"])
            rows.append(row)
        reports.append({"index": index, "path": str(report_path.relative_to(cell)),
                        "sha256": report_meta["sha256"], "status": entry["proof"]["status"],
                        "eligible": entry["eligible"],
                        "terminal_sha256": _json_hash(entry["directory"] / "terminal.json")})
    all_eligible = len(entries) == spec["samples"] and all(item["eligible"] for item in entries)
    status = "pass" if all_eligible else ("ineligible" if entries and
                                           any(not item["eligible"] for item in entries) else "failed")
    value = {
        "schema": COLD_CAPTURE_SCHEMA, "version": VERSION, "status": status,
        "attempt": spec["attempt"], "run": {key: spec[key] for key in
                                               ("kind", "repeat", "role", "provider", "samples", "warmups", "label")},
        "protocol": {"path": "cold-protocol.json", "sha256": protocol_hash},
        "build": {"path": build["path"], "sha256": build["receipt_sha256"], "source": build["source"]},
        "binary": build["binary"], "source": build["source"],
        "expected_invocations": spec["samples"], "completed_invocations": len(entries),
        "warmups": 0, "stopped_on_ineligible": stopped,
        "reports": reports, "rows": rows,
        "scope": "fresh parent invocation per sample; each parent starts one fresh measured cold child",
        "claim": "same-cold-cell repeat variance only; no warm-versus-cold timing comparison",
    }
    _write_or_match(cell / "report.json", value)
    return cell / "report.json"


def _with_cpu_lock(action: Callable[[], Any]) -> Any:
    path = Path(CPU_LOCK)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a+") as lock:
        fcntl.flock(lock.fileno(), fcntl.LOCK_EX)
        try:
            return action()
        finally:
            fcntl.flock(lock.fileno(), fcntl.LOCK_UN)


def capture_one(attempt: str, role: str, repeat: int, *, pilot: bool = False,
                build_dir: Path | None = None,
                timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS) -> Path:
    require(role in ROLES and repeat in ((1,) if pilot else REPEATS),
            "unknown cold capture role or repeat")
    builds = warm.load_builds(build_dir)
    protocol, protocol_hash = _load_protocol(builds)
    match = [item for item in formal_inventory(pilot=pilot)
             if item["role"] == role and item["repeat"] == repeat]
    require(len(match) == 1, "cold capture specification is not frozen")
    spec = dict(match[0], attempt=_attempt(attempt))

    def action() -> Path:
        cell = _capture_root(spec["attempt"], spec["label"])
        require(not cell.exists(), f"refusing to replace immutable cold cell: {cell}")
        (cell / "samples").mkdir(parents=True, exist_ok=False)
        entries: list[dict[str, Any]] = []
        stopped = False
        for index in range(spec["samples"]):
            entry = _run_sample(spec, builds[role], protocol_hash, index, timeout_seconds)
            entries.append(entry)
            if pilot and not entry["eligible"]:
                stopped = True
                break
        path = _write_cell_receipts(spec, builds[role], protocol_hash, entries, stopped=stopped)
        if not entries or any(item["terminal"]["status"] != "pass" for item in entries):
            fail(f"{spec['label']}: one or more cold sample invocations failed")
        return path
    return _with_cpu_lock(action)


def capture_all(attempt: str, build_dir: Path | None = None, *, pilot: bool = False,
                timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS) -> None:
    builds = warm.load_builds(build_dir)
    _load_protocol(builds)
    for spec in formal_inventory(pilot=pilot):
        try:
            capture_one(attempt, spec["role"], spec["repeat"], pilot=pilot,
                        build_dir=build_dir, timeout_seconds=timeout_seconds)
        except ColdMeasureError:
            if pilot:
                break
            raise
        if pilot:
            cell = _capture_root(_attempt(attempt), spec["label"]) / "report.json"
            if cell.is_file() and _json(cell)["status"] == "ineligible":
                break


def _collect(attempt: str, builds: dict[str, dict[str, Any]], protocol: dict[str, Any],
             protocol_hash: str, *, pilot: bool = False) -> list[dict[str, Any]]:
    root = COLD_CAPTURE_ROOT / _attempt(attempt)
    require(root.is_dir() and not root.is_symlink(), f"cold capture attempt is missing: {root}")
    entries: list[dict[str, Any]] = []
    all_pids: set[int] = set()
    for item in formal_inventory(pilot=pilot):
        spec = dict(item, attempt=attempt)
        cell = root / spec["label"]
        if not cell.exists():
            if pilot and entries and any(not entry["eligible"] for entry in entries):
                break
            fail(f"cold capture cell is missing: {cell}")
        aggregate = _json(cell / "report.json")
        aggregate_fields = {"schema", "version", "status", "attempt", "run", "protocol", "build",
                            "binary", "source", "expected_invocations", "completed_invocations",
                            "warmups", "stopped_on_ineligible", "reports", "rows", "scope", "claim"}
        _exact(aggregate, aggregate_fields, f"{cell}/report.json")
        require(aggregate["schema"] == COLD_CAPTURE_SCHEMA and aggregate["version"] == VERSION
                and aggregate["attempt"] == attempt
                and aggregate["protocol"] == {"path": "cold-protocol.json", "sha256": protocol_hash},
                f"{cell}: aggregate custody differs")
        require(aggregate["run"] == {key: spec[key] for key in
                                      ("kind", "repeat", "role", "provider", "samples", "warmups", "label")},
                f"{cell}: aggregate run differs")
        require(aggregate["build"] == {"path": builds[spec["role"]]["path"],
                                        "sha256": builds[spec["role"]]["receipt_sha256"],
                                        "source": builds[spec["role"]]["source"]}
                and aggregate["binary"] == builds[spec["role"]]["binary"]
                and aggregate["source"] == builds[spec["role"]]["source"],
                f"{cell}: aggregate build binding differs")
        _uint(aggregate["expected_invocations"], f"{cell}.expected_invocations", positive=True)
        _uint(aggregate["completed_invocations"], f"{cell}.completed_invocations")
        require(aggregate["expected_invocations"] == spec["samples"]
                and 0 < aggregate["completed_invocations"] <= aggregate["expected_invocations"]
                and aggregate["warmups"] == 0 and aggregate["stopped_on_ineligible"] in (True, False),
                f"{cell}: aggregate invocation counts differ")
        require(aggregate["scope"] ==
                "fresh parent invocation per sample; each parent starts one fresh measured cold child"
                and aggregate["claim"] ==
                "same-cold-cell repeat variance only; no warm-versus-cold timing comparison",
                f"{cell}: aggregate scope differs")
        reports: list[dict[str, Any]] = []
        sample_root = cell / "samples"
        require(sample_root.is_dir(), f"{cell}: sample root missing")
        for index in range(int(aggregate["completed_invocations"])):
            directory = sample_root / f"{index:03d}"
            started, terminal = _json(directory / "started.json"), _json(directory / "terminal.json")
            sample = _check_sample_terminal(started, terminal, spec, builds[spec["role"]],
                                            protocol_hash, directory)
            for row in sample["rows"]:
                pid = row["child_process_id"]
                require(pid not in all_pids, f"{root}: measured child PID is not unique: {pid}")
                all_pids.add(pid)
            reports.append(sample)
        require({item.name for item in sample_root.iterdir()} ==
                {f"{index:03d}" for index in range(aggregate["completed_invocations"])},
                f"{cell}: sample inventory differs")
        require({item.name for item in cell.iterdir()} == {"report.json", "samples"},
                f"{cell}: aggregate directory inventory differs")
        require(aggregate["completed_invocations"] == len(reports)
                and isinstance(aggregate["reports"], list) and aggregate["reports"]
                and len(aggregate["reports"]) == len(reports),
                f"{cell}: aggregate report inventory differs")
        for index, (receipt, sample) in enumerate(zip(aggregate["reports"], reports)):
            _exact(receipt, ("index", "path", "sha256", "status", "eligible", "terminal_sha256"),
                   f"{cell}.reports[{index}]")
            report_path = sample["directory"] / "report.json"
            require(receipt["index"] == index
                    and receipt["path"] == str(report_path.relative_to(cell))
                    and receipt["sha256"] == _json_hash(report_path)
                    and receipt["status"] == sample["proof"]["status"]
                    and receipt["eligible"] is sample["eligible"]
                    and receipt["terminal_sha256"] == _json_hash(sample["directory"] / "terminal.json"),
                    f"{cell}.reports[{index}]: receipt differs")
        require(aggregate["rows"] == [entry["rows"][0] for entry in reports if entry["rows"]],
                f"{cell}: aggregate row evidence differs")
        expected_status = "pass" if len(reports) == spec["samples"] and all(item["eligible"] for item in reports) else "ineligible"
        require(aggregate["status"] == expected_status,
                f"{cell}: aggregate status differs")
        require(aggregate["stopped_on_ineligible"] is
                (pilot and any(not item["eligible"] for item in reports)),
                f"{cell}: aggregate stop marker differs")
        entries.append({"spec": spec, "directory": cell, "aggregate": aggregate,
                        "samples": reports, "eligible": aggregate["status"] == "pass",
                        "rows": [entry["rows"][0] for entry in reports if entry["rows"]]})
    labels = {item["label"] for item in formal_inventory(pilot=pilot)}
    actual_labels = {item["spec"]["label"] for item in entries}
    expected_labels = {item["label"] for item in formal_inventory(pilot=pilot)}
    if pilot and any(not item["eligible"] for item in entries):
        require(actual_labels <= expected_labels, "cold pilot capture labels differ")
    else:
        require(actual_labels == expected_labels, "cold capture labels differ")
    # Ineligible pilot capture may intentionally stop before the second role.
    if pilot and any(not item["eligible"] for item in entries):
        return entries
    require({path.name for path in root.iterdir()} == labels, f"{root}: extra/missing cold cells")
    return entries


def _percentiles(values: list[int]) -> dict[str, Any]:
    require(values, "cannot summarize an empty vector")
    ordered = sorted(values)
    middle = len(ordered) // 2
    p50 = ordered[middle] if len(ordered) % 2 else (ordered[middle - 1] + ordered[middle]) // 2
    return {"n": len(values), "min": ordered[0], "max": ordered[-1], "mean": statistics.fmean(values),
            "p50": p50, "p95": ordered[math.ceil(len(ordered) * .95) - 1],
            "p99": ordered[math.ceil(len(ordered) * .99) - 1]}


def _bootstrap_median(values: list[int], *, repetitions: int = 2000) -> dict[str, Any]:
    require(values, "cannot bootstrap an empty vector")
    seed = (0x0494_C01D ^ len(values) ^ sum(values)) & 0xFFFFFFFFFFFFFFFF
    medians: list[int] = []
    for _ in range(repetitions):
        sample: list[int] = []
        for _ in values:
            seed = (seed * 6364136223846793005 + 1442695040888963407) & 0xFFFFFFFFFFFFFFFF
            sample.append(values[(seed >> 32) % len(values)])
        sample.sort()
        index = len(sample) // 2
        medians.append(sample[index] if len(sample) % 2 else (sample[index - 1] + sample[index]) // 2)
    medians.sort()
    return {"method": "deterministic_percentile_bootstrap_median", "seed":
            (0x0494_C01D ^ len(values) ^ sum(values)) & 0xFFFFFFFFFFFFFFFF,
            "resamples": repetitions, "confidence": 0.95,
            "median": _percentiles(values)["p50"], "ci_low": medians[(repetitions * 25) // 1000],
            "ci_high": medians[(repetitions * 975) // 1000 - 1]}


def _relative(first: int | float, second: int | float) -> float | None:
    if first == 0:
        return 0.0 if second == 0 else None
    return (float(second) - float(first)) / abs(float(first)) * 100.0


def _cell_stats(entry: dict[str, Any]) -> dict[str, Any] | None:
    if not entry["eligible"]:
        return None
    samples = entry["samples"]
    require(len(samples) == entry["spec"]["samples"],
            f"{entry['directory']}: ineligible/short cell cannot be analyzed as formal")
    values: dict[str, list[int]] = {
        "parent_tree_rss_kbytes": [
            _uint(sample["resource"].get("maximum_resident_set_size_(kbytes)"),
                  f"{entry['directory']}: parent RSS") for sample in samples
        ]
    }
    for name, getter in (
        ("elapsed_ns", lambda row: row["latency_ns"]),
        ("process_rss_bytes", lambda row: row["process_rss_bytes"]),
        ("process_peak_rss_bytes", lambda row: row["process_peak_rss_bytes"]),
        ("process_read_bytes", lambda row: row["process_read_bytes"]),
        ("source_read_calls", lambda row: row["source_calls"]),
        ("source_read_returned_bytes", lambda row: row["source_returned_bytes"]),
        ("output_bytes", lambda row: row["output_bytes"]),
    ):
        values[name] = [getter(sample["rows"][0]) for sample in samples]
    if entry["spec"]["role"] == "allocator":
        for name, key in (("allocation_calls", "allocation_calls"),
                          ("deallocation_calls", "deallocation_calls"),
                          ("reallocation_calls", "reallocation_calls"),
                          ("allocated_bytes", "allocated_bytes"),
                          ("region_peak_live_bytes", "region_peak_live_bytes")):
            values[name] = [sample["rows"][0]["allocation"][key] for sample in samples]
    stats = {name: _percentiles(value) for name, value in values.items()}
    bootstrap = {name: _bootstrap_median(value) for name, value in values.items()}
    return {"label": entry["spec"]["label"], "role": entry["spec"]["role"],
            "repeat": entry["spec"]["repeat"], "provider": entry["spec"]["provider"],
            "samples": len(samples), "warmups": 0, "raw_vectors": values,
            "percentiles": stats, "bootstrap_median_ci": bootstrap,
            "child_pids": sorted(sample["rows"][0]["child_process_id"] for sample in samples),
            "claim_scope": "same-cold-cell repeat variance only; no warm-versus-cold timing comparison"}


def analyze_data(entries: list[dict[str, Any]], builds: dict[str, dict[str, Any]], *,
                 pilot: bool = False, protocol_hash: str = "",
                 warm_protocol_hash: str | None = None) -> dict[str, Any]:
    rows = [_cell_stats(entry) for entry in entries]
    summaries = [row for row in rows if row is not None]
    repeat_variance: list[dict[str, Any]] = []
    if not pilot:
        for role in ROLES:
            matched = [row for row in summaries if row["role"] == role]
            if len(matched) == 2:
                first, second = sorted(matched, key=lambda row: row["repeat"])
                metrics: dict[str, Any] = {}
                for name in ("elapsed_ns", "process_rss_bytes", "process_peak_rss_bytes",
                             "parent_tree_rss_kbytes",
                             "process_read_bytes", "source_read_calls", "source_read_returned_bytes",
                             "output_bytes", "allocation_calls", "allocated_bytes"):
                    if name not in first["percentiles"] or name not in second["percentiles"]:
                        continue
                    left, right = first["percentiles"][name], second["percentiles"][name]
                    change = {key: _relative(left[key], right[key]) for key in ("p50", "p95", "p99")}
                    metrics[name] = {"repeat1": {key: left[key] for key in ("p50", "p95", "p99")},
                                     "repeat2": {key: right[key] for key in ("p50", "p95", "p99")},
                                     "relative_percent": change,
                                     "flag_over_5_percent": any(item is not None and abs(item) > 5
                                                                  for item in change.values())}
                repeat_variance.append({"role": role, "metrics": metrics,
                                        "scope": "descriptive same-cold-cell repeat variance; no warm-versus-cold timing claim"})
    return {"schema": COLD_ANALYSIS_SCHEMA, "version": VERSION, "change": CHANGE,
            "status": "pass" if len(summaries) == len(entries) and all(entry["eligible"] for entry in entries)
            else "ineligible", "claim_authorized": False,
            "performance_claim": "descriptive verified-cold evidence; repeat variance only; no warm-versus-cold timing claim",
            "protocol_sha256": protocol_hash, "warm_protocol_sha256": warm_protocol_hash,
            "pilot": pilot, "case": "docx_opened_document_one_paragraph_edit_save",
            "roles": list(ROLES), "repeats": [1] if pilot else list(REPEATS),
            "samples": PILOT_SAMPLES if pilot else FORMAL_SAMPLES, "warmups": 0,
            "builds": {role: _build_binding(builds[role]) for role in ROLES},
            "cells": rows, "repeat_variance": repeat_variance,
            "comparisons": [],
            "inventory": {"expected_cells": len(formal_inventory(pilot=pilot)),
                          "completed_cells": len(entries),
                          "expected_invocations": len(formal_inventory(pilot=pilot)) *
                          (PILOT_SAMPLES if pilot else FORMAL_SAMPLES),
                          "eligible_invocations": sum(len(entry["rows"]) for entry in entries)}}


def analyze(attempt: str, build_dir: Path | None = None, *, pilot: bool = False) -> Path:
    builds = warm.load_builds(build_dir)
    protocol, protocol_hash = _load_protocol(builds)
    entries = _collect(attempt, builds, protocol, protocol_hash, pilot=pilot)
    summary = analyze_data(entries, builds, pilot=pilot, protocol_hash=protocol_hash,
                           warm_protocol_hash=protocol["warm_protocol"]["sha256"])
    path = COLD_ANALYSIS_ROOT / f"{attempt}{'-pilot' if pilot else ''}.json"
    _write_or_match(path, summary)
    return path


def verify(attempt: str, build_dir: Path | None = None, *, pilot: bool = False) -> Path:
    builds = warm.load_builds(build_dir)
    protocol, protocol_hash = _load_protocol(builds)
    entries = _collect(attempt, builds, protocol, protocol_hash, pilot=pilot)
    expected = analyze_data(entries, builds, pilot=pilot, protocol_hash=protocol_hash,
                            warm_protocol_hash=protocol["warm_protocol"]["sha256"])
    summary_path = COLD_ANALYSIS_ROOT / f"{attempt}{'-pilot' if pilot else ''}.json"
    require(summary_path.is_file(), f"cold analysis summary is missing: {summary_path}")
    require(_json(summary_path) == expected, "cold analysis summary differs from raw evidence")
    all_eligible = all(entry["eligible"] and len(entry["rows"]) == entry["spec"]["samples"]
                       for entry in entries)
    status = "pass" if all_eligible else "ineligible"
    receipt = {"schema": COLD_VERIFICATION_SCHEMA, "version": VERSION, "status": status,
               "attempt": attempt, "pilot": pilot, "protocol_sha256": protocol_hash,
               "warm_protocol_sha256": protocol["warm_protocol"]["sha256"],
               "summary_sha256": _json_hash(summary_path),
               "raw_receipts": [{"label": entry["spec"]["label"],
                                 "aggregate_sha256": _json_hash(entry["directory"] / "report.json"),
                                 "eligible": entry["eligible"],
                                 "sample_count": len(entry["samples"])} for entry in entries],
               "verified_utc": now()}
    path = COLD_VERIFICATION_ROOT / f"{attempt}{'-pilot' if pilot else ''}.json"
    _write_or_match(path, receipt, volatile=("verified_utc",))
    if not pilot:
        require(status == "pass", "formal cold verification is ineligible and cannot pass")
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
            builds = (warm.load_builds() if (ROOT / "build-normal.json").is_file()
                      and (ROOT / "build-allocator.json").is_file() else None)
            if builds is None:
                print(json.dumps(protocol_value(), indent=2, sort_keys=True))
            else:
                warm_protocol, warm_hash = warm._load_protocol(builds)
                print(json.dumps(protocol_value(builds, warm_protocol, warm_hash),
                                 indent=2, sort_keys=True))
        elif args.command == "freeze":
            print(create_freeze())
        elif args.command == "capture":
            capture_one(args.attempt, args.role, args.repeat, pilot=args.pilot,
                        build_dir=args.build_dir, timeout_seconds=args.timeout_seconds)
        elif args.command == "capture-all":
            capture_all(args.attempt, args.build_dir, pilot=args.pilot,
                        timeout_seconds=args.timeout_seconds)
        elif args.command == "analyze":
            print(analyze(args.attempt, args.build_dir, pilot=args.pilot))
        elif args.command == "verify":
            print(verify(args.attempt, args.build_dir, pilot=args.pilot))
    except (ColdMeasureError, OSError, ValueError, subprocess.SubprocessError) as error:
        print(f"cold_measure.py: FAIL: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
