#!/usr/bin/env python3
"""Plan and retain 0484 replayable DOCX stream measurements.

This driver intentionally keeps the measurement surface smaller than the
correctness matrix.  Each child receives one explicit source/authored/chunk/
text case, so the report's source, authored, window, proof, and oracle records
remain attributable to one process.  Builds and captures are separate lanes;
the caller supplies an attempt token so a retry can never replace an earlier
binary, receipt, or capture directory.

The commands which build binaries or run captures are provided for the root
coordinator.  Importing this module and running ``plan`` or ``--help`` does
not execute Cargo or the benchmark.
"""

from __future__ import annotations

import argparse
import fcntl
import json
import math
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
from contextlib import contextmanager
from typing import Any, Iterable

from common import (
    DRIVER_SCRIPTS,
    ENV,
    ENV_KEYS,
    REPO,
    ROOT,
    TEMP,
    meta,
    now,
    read,
    sha,
    write,
)


SCHEMA = "docx-replayable-tail-append-measurement-v1"
REPORT_SCHEMA = "docx-replayable-tail-append-v1"
BUILD_SCHEMA = "docx-replayable-tail-append-build-v1"
CAPTURE_SCHEMA = "docx-replayable-tail-append-capture-v1"
BINARY_NAME = "docx_replayable_tail_append"
CPU = 2
SINK_WRITE_BYTES = 4 * 1024
FORMAL_SAMPLES = 30
FORMAL_WARMUPS = 3
PILOT_SAMPLES = 3
PILOT_WARMUPS = 1
ROLES = ("normal", "allocator")
REPEATS = (1, 2)
SOURCE_COUNTS = (64, 8_192, 131_072)
AUTHORED_COUNTS = (64, 256, 4_096, 16_384)
CHUNK_MODES = ("one", "64", "window")
TEXT_MODES = ("empty", "short", "near")
CHUNK_BYTES = {"one": 0, "64": 64, "window": 8 * 1024}
TEXT_BYTES = {"empty": 0, "near": 60 * 1024}
REPORT_CHUNK_MODES = {"one": "one", "64": "fixed64", "window": "replay_window"}
REPORT_TEXT_MODES = {"empty": "empty", "short": "short", "near": "near_limit"}
LIFECYCLE = ["source_admission", "prepare", "publish", "drop"]
SOURCE_CONTRACT = "caller_owned_arc_positional_read_at_requested_returned_fixed_histograms"
AUTHORED_PROVIDER = "deterministic_replayable_bounded_cursor"
SINK_CONTRACT = "non_seek_hashing_sha256_short_write_no_archive_retention"
ATTEMPT = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]*$")
SHA256 = re.compile(r"^[0-9a-f]{64}$")
EXPECTED_AUTHORED_OPENS = 5
EXPECTED_XML_DEPTH = 16
# These are the finite policy bounds used by the benchmark harness.  They
# deliberately leave room for the largest selected case while rejecting a
# report that has silently changed a bounded window into an unbounded value.
MAX_SELECTED_REPLAY_WINDOW_BYTES = 4 * 1024 * 1024
MAX_SELECTED_AUTHORED_CHUNK_BYTES = 1024 * 1024
MAX_SELECTED_XML_TOKEN_BYTES = 4 * 1024 * 1024
MAX_SELECTED_PARSER_WORKSPACE_BYTES = 64 * 1024 * 1024
SCANNER_WORKSPACE_TARGET = (
    "64-bit Rust target (usize=8, Range<usize>=16, &[u8] pair=32, "
    "(usize, usize)=16)"
)
# The scanner workspace formula below is only portable across targets with
# this layout.  Formal captures are pinned to the 64-bit target above; using
# the host's native sizes implicitly would make an accepted report ambiguous.
TARGET_USIZE_BYTES = 8
TARGET_RANGE_BYTES = 16
TARGET_SLICE_PAIR_BYTES = 32
TARGET_USIZE_PAIR_BYTES = 16
HISTOGRAM_FIELDS = (
    "bytes_0",
    "bytes_1_to_512",
    "bytes_513_to_4096",
    "bytes_4097_to_16384",
    "bytes_16385_to_65536",
    "bytes_over_65536",
)


class MeasureError(RuntimeError):
    """The requested plan or retained measurement is invalid."""


def fail(message: str) -> None:
    raise MeasureError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _case(source_count: int, authored_count: int, text_mode: str, chunk_mode: str) -> dict[str, Any]:
    label = f"s{source_count}-a{authored_count}-{text_mode}-c{chunk_mode}"
    return {
        "label": label,
        "source_count": source_count,
        "authored_count": authored_count,
        "text_mode": text_mode,
        "chunk_mode": chunk_mode,
        "chunk_bytes": CHUNK_BYTES[chunk_mode],
        "text_bytes": TEXT_BYTES.get(text_mode),
        "cli": [
            "--source-counts", str(source_count),
            "--authored-counts", str(authored_count),
            "--chunks", chunk_mode,
            "--text", text_mode,
            "--sink-write", str(SINK_WRITE_BYTES),
        ],
    }


def _deduplicated_cases() -> tuple[dict[str, Any], ...]:
    candidates: list[dict[str, Any]] = []
    candidates.extend(
        _case(source, 64, "short", "64")
        for source in SOURCE_COUNTS
    )
    candidates.extend(
        _case(64, authored, "short", "64")
        for authored in AUTHORED_COUNTS
    )
    candidates.extend(
        _case(64, 64, "near", chunk)
        for chunk in CHUNK_MODES
    )
    candidates.extend(
        _case(64, 64, text, "64")
        for text in TEXT_MODES
    )
    result: list[dict[str, Any]] = []
    seen: set[tuple[int, int, str, str]] = set()
    for value in candidates:
        key = (
            value["source_count"], value["authored_count"],
            value["text_mode"], value["chunk_mode"],
        )
        if key in seen:
            continue
        seen.add(key)
        result.append(value)
    return tuple(result)


CASES = _deduplicated_cases()
CASE_BY_LABEL = {case["label"]: case for case in CASES}
if len(CASES) != 10 or len(CASE_BY_LABEL) != len(CASES):
    raise RuntimeError("0484 measurement matrix did not deduplicate to ten cases")


def _attempt(value: str | None) -> str:
    if value is None or ATTEMPT.fullmatch(value) is None:
        fail("--attempt must be a unique path-safe token")
    return value


def _attempt_root(attempt: str) -> Path:
    path = ROOT / "attempts" / _attempt(attempt)
    path.mkdir(parents=True, exist_ok=True)
    return path


@contextmanager
def _cpu_lock() -> Iterable[None]:
    TEMP.mkdir(parents=True, exist_ok=True)
    with (TEMP / "cpu.lock").open("a") as lock:
        fcntl.flock(lock, fcntl.LOCK_EX)
        yield


def _json_hash(path: Path) -> str:
    value = sha(path)
    require(SHA256.fullmatch(value) is not None, f"{path}: invalid SHA-256")
    return value


def _protocol_value() -> dict[str, Any]:
    cases = [dict(case) for case in CASES]
    formal_base = [
        (role, case["label"])
        for role in ROLES
        for case in CASES
    ]
    formal_runs: list[dict[str, Any]] = []
    for repeat, sequence in ((1, formal_base), (2, tuple(reversed(formal_base)))):
        for role, case_label in sequence:
            formal_runs.append({
                "label": f"r{repeat}-{role}-{case_label}",
                "repeat": repeat,
                "role": role,
                "case": case_label,
            })
    pilots = [
        {
            "label": f"pilot-{role}-{case['label']}",
            "role": role,
            "case": case["label"],
        }
        for role in ROLES
        for case in CASES
    ]
    return {
        "schema": SCHEMA,
        "version": 1,
        "change": 484,
        "frozen_utc": now(),
        "claim_authorized": False,
        "performance_claim": "none",
        "comparison": "none; standalone bounded-stream enabler measurements",
        "scope": "existing source-backed DOCX tail append with caller-authored replayable plain paragraphs; source, authored, artifact, and replay-window metadata retained per case",
        "no_before_after_speedup_baseline": True,
        "cpu": CPU,
        "sink_write_bytes": SINK_WRITE_BYTES,
        "source_counts": list(SOURCE_COUNTS),
        "authored_counts": list(AUTHORED_COUNTS),
        "chunk_modes": list(CHUNK_MODES),
        "text_modes": list(TEXT_MODES),
        "text_contracts": {
            "empty": {"maximum_text_bytes": 0},
            "short": {"maximum_text_bytes": None},
            "near": {"maximum_text_bytes": 60 * 1024},
        },
        "chunk_contracts": {
            mode: {"configured_chunk_bytes": CHUNK_BYTES[mode]}
            for mode in CHUNK_MODES
        },
        "lifecycle": list(LIFECYCLE),
        "source_contract": SOURCE_CONTRACT,
        "authored_provider": AUTHORED_PROVIDER,
        "sink_contract": SINK_CONTRACT,
        "formal": {"samples": FORMAL_SAMPLES, "warmups": FORMAL_WARMUPS},
        "pilot": {"samples": PILOT_SAMPLES, "warmups": PILOT_WARMUPS},
        "repeats": list(REPEATS),
        "roles": list(ROLES),
        "expected_formal_processes": len(formal_runs),
        "expected_pilot_processes": len(pilots),
        "cases": cases,
        "formal_runs": formal_runs,
        "pilot_runs": pilots,
        "environment": {key: ENV[key] for key in ENV_KEYS},
        "scripts": {
            **{name: sha(ROOT / name) for name in DRIVER_SCRIPTS},
            "measure.py": sha(Path(__file__)),
        },
        "binary": {
            "name": BINARY_NAME,
            "cargo_manifest": "tools/perf-baseline/Cargo.toml",
            "release_profile": True,
            "allocator_feature": "allocator-metrics",
        },
    }


def protocol_path() -> Path:
    return ROOT / "protocol.json"


def load_protocol() -> tuple[dict[str, Any], str]:
    path = protocol_path()
    require(path.is_file(), f"measurement protocol is missing: {path}")
    value = read(path)
    require(isinstance(value, dict), "protocol.json must be an object")
    require(value.get("schema") == SCHEMA and value.get("version") == 1, "protocol schema/version differs")
    require(value.get("cases") == list(CASES), "protocol case matrix differs from the driver")
    formal = value.get("formal")
    pilot = value.get("pilot")
    require(isinstance(formal, dict) and isinstance(pilot, dict), "protocol sample settings are missing")
    require(formal.get("samples") == FORMAL_SAMPLES, "protocol formal sample count differs")
    require(formal.get("warmups") == FORMAL_WARMUPS, "protocol formal warmup count differs")
    require(pilot.get("samples") == PILOT_SAMPLES, "protocol pilot sample count differs")
    require(pilot.get("warmups") == PILOT_WARMUPS, "protocol pilot warmup count differs")
    expected = _protocol_value()
    require(value.get("roles") == list(ROLES), "protocol roles differ")
    require(value.get("repeats") == list(REPEATS), "protocol repeats differ")
    require(value.get("formal_runs") == expected["formal_runs"], "protocol formal run inventory differs")
    require(value.get("pilot_runs") == expected["pilot_runs"], "protocol pilot run inventory differs")
    require(value.get("expected_formal_processes") == len(expected["formal_runs"]), "protocol formal process count differs")
    require(value.get("expected_pilot_processes") == len(expected["pilot_runs"]), "protocol pilot process count differs")
    require(value.get("lifecycle") == LIFECYCLE, "protocol lifecycle differs")
    require(value.get("source_contract") == SOURCE_CONTRACT, "protocol source contract differs")
    require(value.get("authored_provider") == AUTHORED_PROVIDER, "protocol authored provider differs")
    require(value.get("sink_contract") == SINK_CONTRACT, "protocol sink contract differs")
    scripts = value.get("scripts")
    require(isinstance(scripts, dict), "protocol script bindings are missing")
    current_scripts = {
        **{name: sha(ROOT / name) for name in DRIVER_SCRIPTS},
        "measure.py": sha(Path(__file__)),
    }
    require(scripts == current_scripts, "measurement helpers changed after protocol freeze")
    return value, _json_hash(path)


def freeze() -> None:
    path = protocol_path()
    require(not path.exists(), f"refusing to replace existing protocol: {path}")
    value = _protocol_value()
    write(path, value)
    print(f"wrote {path} with {len(value['cases'])} cases, {len(value['formal_runs'])} formal processes, and {len(value['pilot_runs'])} pilots")


def _gate_receipt(label: str, attempt: str) -> Path:
    return ROOT / "validation" / f"{label}-{attempt}.json"


def _run_gate(label: str, attempt: str, command: list[str]) -> Path:
    receipt = _gate_receipt(label, attempt)
    require(not receipt.exists(), f"validation attempt already exists: {receipt}")
    process = subprocess.run(
        [sys.executable, "-B", str(ROOT / "gate.py"), "--attempt", attempt, label, *command],
        cwd=REPO,
        env=ENV,
        check=False,
    )
    require(receipt.is_file(), f"gate did not retain its receipt: {receipt}")
    require(process.returncode == 0, f"{label}-{attempt} failed with exit {process.returncode}")
    return receipt


def _build_command(role: str) -> list[str]:
    command = [
        "cargo", "build", "--release", "--locked",
        "--manifest-path", "tools/perf-baseline/Cargo.toml",
        "--bin", BINARY_NAME,
    ]
    if role == "allocator":
        command.extend(["--features", "allocator-metrics"])
    return command


def _binary_metadata(path: Path, label: str) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"{label}: copied binary is missing")
    require(os.access(path, os.X_OK), f"{label}: copied binary is not executable")
    details = meta(path)
    require(details["bytes"] > 0 and SHA256.fullmatch(details["sha256"]) is not None, f"{label}: binary metadata is invalid")
    return {"path": str(path), **details, "executable": True}


def build(role: str, attempt: str) -> None:
    require(role in ROLES, f"unknown build role {role}")
    protocol, protocol_sha256 = load_protocol()
    attempt_root = _attempt_root(attempt)
    record_path = attempt_root / f"build-{role}.json"
    require(not record_path.exists(), f"refusing to replace build record: {record_path}")
    gate_path = _run_gate(f"build-{role}", attempt, _build_command(role))
    gate = read(gate_path)
    require(isinstance(gate, dict), f"{gate_path}: gate receipt must be an object")
    require(gate.get("exit_code") == 0 and gate.get("source_unchanged") is True, f"{gate_path}: build did not pass custody")
    source_before = gate.get("source_before")
    source_after = gate.get("source_after")
    require(isinstance(source_before, dict) and isinstance(source_after, dict), f"{gate_path}: source identity is missing")
    require(source_before == source_after, f"{gate_path}: source identity changed during build")
    origin = REPO / "tools" / "perf-baseline" / "target" / "release" / BINARY_NAME
    origin_binary = _binary_metadata(origin, f"build-{role}.origin")
    destination = TEMP / "0484" / attempt / role / BINARY_NAME
    require(not destination.exists(), f"refusing to replace copied binary: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(origin, destination)
    copied_binary = _binary_metadata(destination, f"build-{role}.copy")
    require(
        copied_binary["bytes"] == origin_binary["bytes"]
        and copied_binary["sha256"] == origin_binary["sha256"],
        f"build-{role}: copied binary differs from release output",
    )
    record = {
        "schema": BUILD_SCHEMA,
        "version": 1,
        "attempt": attempt,
        "role": role,
        "protocol": {"path": "protocol.json", "sha256": protocol_sha256},
        "command": _build_command(role),
        "gate": {
            "path": gate_path.relative_to(ROOT).as_posix(),
            "sha256": sha(gate_path),
        },
        "source_before": source_before,
        "source_after": source_after,
        "source_unchanged": True,
        "binary": copied_binary,
        "original_binary": origin_binary,
        "copied_utc": now(),
        "environment": {key: ENV[key] for key in ENV_KEYS},
    }
    write(record_path, record)
    print(f"built {role}: {record_path}")


def _load_builds(attempt: str, protocol_sha256: str) -> dict[str, dict[str, Any]]:
    attempt_root = _attempt_root(attempt)
    result: dict[str, dict[str, Any]] = {}
    for role in ROLES:
        path = attempt_root / f"build-{role}.json"
        require(path.is_file(), f"{role} build record is missing: {path}")
        value = read(path)
        require(isinstance(value, dict) and value.get("schema") == BUILD_SCHEMA, f"{path}: build schema differs")
        require(value.get("role") == role and value.get("attempt") == attempt, f"{path}: build identity differs")
        protocol = value.get("protocol")
        require(isinstance(protocol, dict) and protocol.get("sha256") == protocol_sha256, f"{path}: protocol binding differs")
        require(value.get("source_unchanged") is True and value.get("source_before") == value.get("source_after"), f"{path}: source custody failed")
        binary = value.get("binary")
        require(isinstance(binary, dict), f"{path}: copied binary metadata is missing")
        actual = _binary_metadata(Path(binary.get("path", "")), f"{path}.binary")
        require(actual == binary, f"{path}: copied binary metadata changed")
        result[role] = value
    require(result["normal"]["source_after"] == result["allocator"]["source_after"], "normal/allocator source identities differ")
    return result


def _case_from_label(label: str) -> dict[str, Any]:
    try:
        return dict(CASE_BY_LABEL[label])
    except KeyError as error:
        fail(f"unknown case label {label!r}; use plan to list explicit cases")
        raise error


def _run_identity(kind: str, role: str, case: dict[str, Any], *, repeat: int | None = None) -> dict[str, Any]:
    label = f"{kind}-{role}-{case['label']}" if repeat is None else f"r{repeat}-{role}-{case['label']}"
    return {
        "kind": kind,
        "label": label,
        "role": role,
        "case": case["label"],
        "source_count": case["source_count"],
        "authored_count": case["authored_count"],
        "chunk_mode": case["chunk_mode"],
        "text_mode": case["text_mode"],
        "repeat": repeat,
    }


def _run_directory(attempt: str, identity: dict[str, Any]) -> Path:
    directory = ROOT / ("pilots" if identity["kind"] == "pilot" else "captures") / attempt / identity["label"]
    require(not directory.exists(), f"refusing to replace run directory: {directory}")
    directory.mkdir(parents=True)
    return directory


def _argv(
    binary: dict[str, Any], case: dict[str, Any], *, samples: int, warmups: int,
    report: Path, resource: Path,
) -> list[str]:
    return [
        "/usr/bin/time", "-v", "-o", str(resource),
        "/usr/bin/taskset", "-c", str(CPU), binary["path"],
        *case["cli"], "--samples", str(samples), "--warmups", str(warmups),
        "--json", str(report),
    ]


def _check_finite_json(value: Any, path: str = "report") -> None:
    """Reject JSON extensions such as NaN and Infinity before validation.

    ``json.load`` accepts those values by default even though they are not
    valid JSON.  A retained measurement must be portable and must never make
    a verifier choose between a finite scalar and a non-finite sentinel.
    """
    if isinstance(value, float):
        require(math.isfinite(value), f"{path}: non-finite number")
    elif isinstance(value, dict):
        for key, child in value.items():
            _check_finite_json(child, f"{path}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            _check_finite_json(child, f"{path}[{index}]")


def _int_field(value: Any, path: str, *, minimum: int = 0) -> int:
    require(type(value) is int, f"{path}: expected an integer")
    require(value >= minimum, f"{path}: expected a value >= {minimum}")
    return value


def _positive_int(value: Any, path: str) -> int:
    return _int_field(value, path, minimum=1)


def _object(value: Any, path: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{path}: expected an object")
    return value


def _sha256_field(value: Any, path: str) -> str:
    require(isinstance(value, str) and SHA256.fullmatch(value) is not None, f"{path}: invalid SHA-256")
    return value


def _check_histogram(
    value: Any,
    path: str,
    expected_calls: int,
    *,
    observed_bytes: int | None = None,
    largest_write: int | None = None,
    maximum: int | None = None,
) -> None:
    histogram = _object(value, path)
    counts = []
    for field in HISTOGRAM_FIELDS:
        counts.append(_int_field(histogram.get(field), f"{path}.{field}"))
    require(sum(counts) == expected_calls, f"{path}: bins do not sum to {expected_calls}")
    lower_bounds = (0, 1, 513, 4_097, 16_385, 65_537)
    upper_bounds = (0, 512, 4_096, 16_384, 65_536, None)
    lower_total = sum(count * lower for count, lower in zip(counts, lower_bounds))
    upper_total = None if counts[5] else sum(count * upper for count, upper in zip(counts[:-1], upper_bounds[:-1]))
    if observed_bytes is not None:
        _int_field(observed_bytes, f"{path}.observed_bytes")
        require(observed_bytes >= lower_total, f"{path}: aggregate bytes are below bucket lower bounds")
        if upper_total is not None:
            require(observed_bytes <= upper_total, f"{path}: aggregate bytes exceed bucket upper bounds")
    if largest_write is not None:
        largest = _int_field(largest_write, f"{path}.largest_write")
        require(largest > 0, f"{path}: largest write must be positive")
        if observed_bytes is not None:
            require(largest <= observed_bytes, f"{path}: largest write exceeds aggregate bytes")
        if largest <= 512:
            largest_bucket = 1
        elif largest <= 4_096:
            largest_bucket = 2
        elif largest <= 16_384:
            largest_bucket = 3
        elif largest <= 65_536:
            largest_bucket = 4
        else:
            largest_bucket = 5
        highest_bucket = max((index for index, count in enumerate(counts) if count), default=0)
        require(largest_bucket == highest_bucket, f"{path}: largest write does not match highest nonzero bucket")
    if maximum is not None:
        require(counts[0] == 0, f"{path}: zero-sized writes are not emitted by the sink")
        if maximum <= 512:
            require(sum(counts[2:]) == 0, f"{path}: write exceeds configured sink bound")
        elif maximum <= 4096:
            require(sum(counts[3:]) == 0, f"{path}: write exceeds configured sink bound")
        elif maximum <= 16384:
            require(sum(counts[4:]) == 0, f"{path}: write exceeds configured sink bound")
        elif maximum <= 65536:
            require(counts[5] == 0, f"{path}: write exceeds configured sink bound")


def _scanner_workspace_limit(max_token_bytes: int, max_depth: int) -> int:
    """Mirror the harness scanner owner formula for its fixed 64-bit target."""
    require(sys.maxsize >= 2**63 - 1, f"scanner workspace validation requires {SCANNER_WORKSPACE_TARGET}")
    max_token = _positive_int(max_token_bytes, "scanner target max token bytes")
    depth = _positive_int(max_depth, "scanner target max depth")
    token = max_token + 1
    levels = depth + 1
    namespace_declarations = min(256, token)
    attributes = token
    binding_bytes = TARGET_USIZE_BYTES * 4
    hash_entry_bytes = 8 + 1
    token_windows = token * 4
    opened_names = max(levels * token * 2, 8)
    opened_indexes = max(4 * TARGET_USIZE_BYTES, 2 * levels * TARGET_USIZE_BYTES)
    namespace_bytes = max((73 + levels * token) * 2, 128)
    namespace_binding_floor = 8 * binding_bytes
    namespace_bindings = max(
        (2 + levels * namespace_declarations) * 2 * binding_bytes,
        namespace_binding_floor,
    )
    attribute_ranges = max((attributes + 1) * TARGET_RANGE_BYTES * 2, 4 * TARGET_RANGE_BYTES)
    attribute_hash = max((attributes + 1) * hash_entry_bytes * 4, 8 * hash_entry_bytes)
    attribute_seen = max((attributes + 1) * TARGET_SLICE_PAIR_BYTES * 2, 8 * TARGET_SLICE_PAIR_BYTES)
    scope_stack = max(levels * 1 * 2, 8)
    namespace_scope_stack = levels * TARGET_USIZE_PAIR_BYTES * 2
    return sum(
        (
            token_windows,
            opened_names,
            opened_indexes,
            namespace_bytes,
            namespace_bindings,
            attribute_ranges,
            attribute_hash,
            attribute_seen,
            scope_stack,
            namespace_scope_stack,
            3,
        )
    )


def _check_member_identities(source: dict[str, Any], path: str) -> None:
    members = source.get("members")
    require(isinstance(members, list) and members, f"{path}.members: member metadata is missing")
    member_count = _positive_int(source.get("member_count"), f"{path}.member_count")
    require(len(members) == member_count, f"{path}.members: count differs from member_count")
    paths: set[str] = set()
    for index, raw in enumerate(members):
        member = _object(raw, f"{path}.members[{index}]")
        name = member.get("path")
        require(isinstance(name, str) and name and name not in paths, f"{path}.members[{index}].path: missing or duplicate")
        paths.add(name)
        require(isinstance(member.get("compression_method"), str) and member["compression_method"], f"{path}.members[{index}].compression_method: missing")
        require(type(member.get("data_descriptor")) is bool, f"{path}.members[{index}].data_descriptor: expected boolean")
        for field in ("crc32", "decoded_bytes", "compressed_bytes"):
            _int_field(member.get(field), f"{path}.members[{index}].{field}")
        _sha256_field(member.get("decoded_sha256"), f"{path}.members[{index}].decoded_sha256")
        _sha256_field(member.get("compressed_sha256"), f"{path}.members[{index}].compressed_sha256")


def _check_source_identity(source: dict[str, Any], case: dict[str, Any], path: str) -> None:
    _positive_int(source.get("archive_bytes"), f"{path}.archive_bytes")
    _positive_int(source.get("main_xml_bytes"), f"{path}.main_xml_bytes")
    _sha256_field(source.get("archive_sha256"), f"{path}.archive_sha256")
    _sha256_field(source.get("main_xml_sha256"), f"{path}.main_xml_sha256")
    require(source.get("unchanged_oracle") is True, f"{path}.unchanged_oracle: source was not unchanged")
    require(source.get("opaque_member_exact") is True, f"{path}.opaque_member_exact: opaque member changed")
    semantic = _object(source.get("semantic"), f"{path}.semantic")
    require(semantic.get("paragraph_count") == case["source_count"], f"{path}.semantic.paragraph_count: differs from case")
    require(semantic.get("text_bytes") == case["source_count"] * len("source-00000000"), f"{path}.semantic.text_bytes: source text identity differs")
    _positive_int(semantic.get("paragraph_count"), f"{path}.semantic.paragraph_count")
    _int_field(semantic.get("text_bytes"), f"{path}.semantic.text_bytes")
    _sha256_field(semantic.get("order_sha256"), f"{path}.semantic.order_sha256")
    _sha256_field(semantic.get("text_sha256"), f"{path}.semantic.text_sha256")
    _check_member_identities(source, path)


def _check_authored_identity(authored: dict[str, Any], case: dict[str, Any], path: str) -> None:
    authored_count = _positive_int(authored.get("authored_count"), f"{path}.authored_count")
    require(authored_count == case["authored_count"], f"{path}.authored_count: differs from case")
    require(authored.get("chunk_mode") == REPORT_CHUNK_MODES[case["chunk_mode"]], f"{path}.chunk_mode: differs from case")
    require(authored.get("text_mode") == REPORT_TEXT_MODES[case["text_mode"]], f"{path}.text_mode: differs from case")
    max_chunk = _positive_int(authored.get("max_chunk_bytes"), f"{path}.max_chunk_bytes")
    replay_window = _positive_int(authored.get("replay_window_bytes"), f"{path}.replay_window_bytes")
    max_encoded = _positive_int(authored.get("max_encoded_paragraph_bytes"), f"{path}.max_encoded_paragraph_bytes")
    entity_refs = _int_field(authored.get("xml_entity_reference_count"), f"{path}.xml_entity_reference_count")
    text_bytes = _int_field(authored.get("text_bytes"), f"{path}.text_bytes")
    encoded_bytes = _positive_int(authored.get("encoded_xml_bytes"), f"{path}.encoded_xml_bytes")
    event_count = _positive_int(authored.get("event_count"), f"{path}.event_count")
    if case["text_mode"] == "near":
        require(
            text_bytes == authored_count * TEXT_BYTES["near"],
            f"{path}.text_bytes: near-limit payload must be exactly authored_count * 60 KiB",
        )
    require(max_chunk <= MAX_SELECTED_AUTHORED_CHUNK_BYTES, f"{path}.max_chunk_bytes: exceeds selected bound")
    require(replay_window <= MAX_SELECTED_REPLAY_WINDOW_BYTES, f"{path}.replay_window_bytes: exceeds selected bound")
    require(6 * max_chunk <= replay_window, f"{path}: authored chunk does not fit replay window")
    require(encoded_bytes >= max_encoded * authored_count, f"{path}: encoded bytes below per-paragraph maximum")
    require(text_bytes % authored_count == 0, f"{path}.text_bytes: not divisible by authored count")
    require(entity_refs % authored_count == 0, f"{path}.xml_entity_reference_count: not divisible by authored count")
    require(entity_refs <= text_bytes, f"{path}.xml_entity_reference_count: exceeds authored text bytes")
    text_per_paragraph = text_bytes // authored_count
    expected_text_chunks = 0 if text_per_paragraph == 0 else (text_per_paragraph + max_chunk - 1) // max_chunk * authored_count
    require(event_count == 2 * authored_count + expected_text_chunks, f"{path}.event_count: differs from authored framing")
    if case["text_mode"] == "empty":
        require(text_bytes == 0 and entity_refs == 0, f"{path}: empty text identity contains payload")
    else:
        require(text_bytes > 0 and entity_refs > 0, f"{path}: authored text identity is empty")
    _sha256_field(authored.get("expected_event_sha256"), f"{path}.expected_event_sha256")
    _sha256_field(authored.get("expected_encoded_sha256"), f"{path}.expected_encoded_sha256")


def _check_limits(
    limits: dict[str, Any], authored: dict[str, Any], case: dict[str, Any], path: str,
) -> None:
    event_limit = _positive_int(limits.get("parser_event_limit"), f"{path}.parser_event_limit")
    token_limit = _positive_int(limits.get("parser_token_bytes"), f"{path}.parser_token_bytes")
    workspace_limit = _positive_int(limits.get("parser_workspace_bytes"), f"{path}.parser_workspace_bytes")
    depth = _positive_int(limits.get("max_xml_depth"), f"{path}.max_xml_depth")
    chunk_limit = _positive_int(limits.get("max_authored_chunk_bytes"), f"{path}.max_authored_chunk_bytes")
    replay_window = _positive_int(limits.get("replay_window_bytes"), f"{path}.replay_window_bytes")
    entity_refs = _int_field(authored.get("xml_entity_reference_count"), f"authored.xml_entity_reference_count")
    expected_events = (
        case["source_count"] * 8 + 128
        + case["authored_count"] * 8 + entity_refs * 2 + 128
    )
    require(event_limit == expected_events, f"{path}.parser_event_limit: differs from sealed harness bound")
    expected_token = max(authored["max_encoded_paragraph_bytes"] + 1024, 64 * 1024)
    require(token_limit == expected_token, f"{path}.parser_token_bytes: differs from sealed harness bound")
    require(depth == EXPECTED_XML_DEPTH, f"{path}.max_xml_depth: differs from selected parser bound")
    require(
        workspace_limit == _scanner_workspace_limit(token_limit, depth),
        f"{path}.parser_workspace_bytes: differs from {SCANNER_WORKSPACE_TARGET} harness bound",
    )
    require(chunk_limit == authored["max_chunk_bytes"], f"{path}.max_authored_chunk_bytes: differs from authored proof")
    require(replay_window == authored["replay_window_bytes"], f"{path}.replay_window_bytes: differs from authored proof")
    require(replay_window <= MAX_SELECTED_REPLAY_WINDOW_BYTES, f"{path}.replay_window_bytes: exceeds selected bound")
    require(token_limit <= MAX_SELECTED_XML_TOKEN_BYTES, f"{path}.parser_token_bytes: exceeds selected bound")
    require(workspace_limit <= MAX_SELECTED_PARSER_WORKSPACE_BYTES, f"{path}.parser_workspace_bytes: exceeds selected bound")


def _check_oracle_and_proof(
    source: dict[str, Any], authored: dict[str, Any], limits: dict[str, Any],
    oracle: dict[str, Any], proof: dict[str, Any], case: dict[str, Any], path: str,
) -> None:
    candidate_semantic = _object(oracle.get("candidate_semantic"), f"{path}.oracle.candidate_semantic")
    _positive_int(oracle.get("candidate_archive_bytes"), f"{path}.oracle.candidate_archive_bytes")
    candidate_main_bytes = _positive_int(oracle.get("candidate_main_xml_bytes"), f"{path}.oracle.candidate_main_xml_bytes")
    _sha256_field(oracle.get("candidate_archive_sha256"), f"{path}.oracle.candidate_archive_sha256")
    _sha256_field(oracle.get("candidate_main_xml_sha256"), f"{path}.oracle.candidate_main_xml_sha256")
    candidate_member_count = _positive_int(oracle.get("candidate_member_count"), f"{path}.oracle.candidate_member_count")
    require(candidate_member_count == source["member_count"], f"{path}.oracle.candidate_member_count: differs from source")
    require(candidate_semantic.get("paragraph_count") == case["source_count"] + case["authored_count"], f"{path}.oracle.candidate_semantic.paragraph_count: differs")
    require(candidate_semantic.get("text_bytes") == source["semantic"]["text_bytes"] + authored["text_bytes"], f"{path}.oracle.candidate_semantic.text_bytes: differs")
    _int_field(candidate_semantic.get("paragraph_count"), f"{path}.oracle.candidate_semantic.paragraph_count")
    _int_field(candidate_semantic.get("text_bytes"), f"{path}.oracle.candidate_semantic.text_bytes")
    _sha256_field(candidate_semantic.get("order_sha256"), f"{path}.oracle.candidate_semantic.order_sha256")
    _sha256_field(candidate_semantic.get("text_sha256"), f"{path}.oracle.candidate_semantic.text_sha256")
    for field in (
        "candidate_xml_exact",
        "candidate_semantic_exact",
        "untouched_member_metadata_exact",
        "untouched_raw_members_preserved",
        "physical_order_exact",
        "opaque_member_exact",
        "source_unchanged",
        "inverse_exact",
    ):
        require(oracle.get(field) is True, f"{path}.oracle.{field}: oracle failed")

    proof_source_len = _positive_int(proof.get("source_len"), f"{path}.proof.source_len")
    proof_candidate_len = _positive_int(proof.get("candidate_len"), f"{path}.proof.candidate_len")
    _sha256_field(proof.get("source_sha256"), f"{path}.proof.source_sha256")
    _sha256_field(proof.get("candidate_sha256"), f"{path}.proof.candidate_sha256")
    require(proof_source_len == source["main_xml_bytes"], f"{path}.proof.source_len: differs from source XML")
    require(proof.get("source_sha256") == source["main_xml_sha256"], f"{path}.proof.source_sha256: differs from source XML")
    require(proof_candidate_len == candidate_main_bytes, f"{path}.proof.candidate_len: differs from candidate XML")
    require(proof.get("candidate_sha256") == oracle["candidate_main_xml_sha256"], f"{path}.proof.candidate_sha256: differs from candidate XML")
    source_paragraph_count = _positive_int(proof.get("source_paragraph_count"), f"{path}.proof.source_paragraph_count")
    candidate_paragraph_count = _positive_int(proof.get("candidate_paragraph_count"), f"{path}.proof.candidate_paragraph_count")
    require(source_paragraph_count == case["source_count"], f"{path}.proof.source_paragraph_count: differs")
    require(candidate_paragraph_count == case["source_count"] + case["authored_count"], f"{path}.proof.candidate_paragraph_count: differs")
    _positive_int(proof.get("source_event_count"), f"{path}.proof.source_event_count")
    _positive_int(proof.get("candidate_event_count"), f"{path}.proof.candidate_event_count")
    require(proof["candidate_event_count"] > proof["source_event_count"], f"{path}.proof.candidate_event_count: candidate did not add events")
    require(proof["candidate_event_count"] <= limits["parser_event_limit"], f"{path}.proof.candidate_event_count: exceeds parser bound")
    insertion = _int_field(proof.get("insertion_offset"), f"{path}.proof.insertion_offset")
    generated = _int_field(proof.get("generated_offset"), f"{path}.proof.generated_offset")
    require(insertion == generated and insertion <= proof_source_len, f"{path}.proof: generated offset is not a source insertion point")
    require(proof.get("generated_once") is True, f"{path}.proof.generated_once: generation proof failed")
    proof_authored = _object(proof.get("authored"), f"{path}.proof.authored")
    for field in (
        "authored_count", "chunk_mode", "text_mode", "max_chunk_bytes", "replay_window_bytes",
        "max_encoded_paragraph_bytes", "xml_entity_reference_count", "text_bytes", "encoded_xml_bytes",
        "event_count", "expected_event_sha256", "expected_encoded_sha256",
    ):
        require(proof_authored.get(field) == authored.get(field), f"{path}.proof.authored.{field}: differs from sealed authored proof")


def _check_allocator_sample(sample: dict[str, Any], role: str, path: str) -> None:
    allocation = sample.get("allocation")
    if role == "normal":
        require(allocation is None, f"{path}.allocation: normal run must omit allocator metrics")
        return
    metrics = _object(allocation, f"{path}.allocation")
    require(metrics.get("status") == "measured", f"{path}.allocation.status: allocator sample is not measured")
    require(metrics.get("scope") == "operation_global_system_allocator", f"{path}.allocation.scope: scope differs")
    counter_fields = (
        "allocation_calls", "deallocation_calls", "reallocation_calls", "failed_allocation_calls",
        "allocated_bytes", "deallocated_bytes", "live_bytes_before", "live_bytes_after",
        "peak_live_bytes_before", "peak_live_bytes_after", "region_peak_live_bytes",
    )
    for field in counter_fields:
        _int_field(metrics.get(field), f"{path}.allocation.{field}")
    require(metrics["failed_allocation_calls"] == 0, f"{path}.allocation: successful formal sample recorded a failed allocation")
    require(metrics["allocation_calls"] >= metrics["reallocation_calls"], f"{path}.allocation: reallocations exceed allocation calls")
    require(metrics["live_bytes_after"] == metrics["live_bytes_before"] + metrics["allocated_bytes"] - metrics["deallocated_bytes"], f"{path}.allocation: live-byte counter relation failed")
    require(metrics["live_bytes_after"] == metrics["live_bytes_before"], f"{path}.allocation: lifecycle net live bytes are not zero")
    require(metrics["allocated_bytes"] == metrics["deallocated_bytes"], f"{path}.allocation: zero-net lifecycle bytes are not balanced")
    require(metrics["peak_live_bytes_before"] >= metrics["live_bytes_before"], f"{path}.allocation: pre-operation peak below live bytes")
    require(metrics["peak_live_bytes_after"] >= metrics["live_bytes_after"], f"{path}.allocation: post-operation peak below live bytes")
    require(metrics["peak_live_bytes_after"] >= metrics["peak_live_bytes_before"], f"{path}.allocation: peak counter moved backwards")
    require(metrics["region_peak_live_bytes"] >= max(metrics["live_bytes_before"], metrics["live_bytes_after"]), f"{path}.allocation: region peak below live endpoints")
    require(metrics["region_peak_live_bytes"] <= metrics["peak_live_bytes_after"], f"{path}.allocation: region peak exceeds process peak")


def _check_process_sample(process: Any, path: str) -> None:
    if process is None:
        return
    metrics = _object(process, path)
    fields = (
        "rchar", "wchar", "read_bytes", "write_bytes", "cancelled_write_bytes", "syscr", "syscw",
        "minor_faults", "major_faults", "user_cpu_ticks", "system_cpu_ticks", "clock_ticks_per_second",
        "voluntary_context_switches", "nonvoluntary_context_switches", "rss_bytes", "peak_rss_bytes",
    )
    for field in fields:
        _int_field(metrics.get(field), f"{path}.{field}")
    require(metrics["clock_ticks_per_second"] > 0, f"{path}.clock_ticks_per_second: must be positive")
    require(metrics["peak_rss_bytes"] >= metrics["rss_bytes"], f"{path}: peak RSS below RSS delta")


def _check_sample(
    sample: Any, index: int, role: str, case: dict[str, Any], oracle: dict[str, Any],
    authored: dict[str, Any], *, sink_write_bytes: int, path: str,
) -> None:
    item = _object(sample, path)
    sample_index = _int_field(item.get("sample"), f"{path}.sample")
    require(sample_index == index, f"{path}.sample: expected index {index}")
    sample_source_count = _positive_int(item.get("source_count"), f"{path}.source_count")
    sample_authored_count = _positive_int(item.get("authored_count"), f"{path}.authored_count")
    require(sample_source_count == case["source_count"], f"{path}.source_count: differs from case")
    require(sample_authored_count == case["authored_count"], f"{path}.authored_count: differs from case")
    require(item.get("chunk_mode") == REPORT_CHUNK_MODES[case["chunk_mode"]], f"{path}.chunk_mode: differs from case")
    require(item.get("text_mode") == REPORT_TEXT_MODES[case["text_mode"]], f"{path}.text_mode: differs from case")
    _positive_int(item.get("elapsed_ns"), f"{path}.elapsed_ns")
    reads = _object(item.get("source_reads"), f"{path}.source_reads")
    calls = _positive_int(reads.get("calls"), f"{path}.source_reads.calls")
    requested = _int_field(reads.get("requested_bytes"), f"{path}.source_reads.requested_bytes")
    returned = _int_field(reads.get("returned_bytes"), f"{path}.source_reads.returned_bytes")
    require(requested > 0 and returned > 0 and returned <= requested, f"{path}.source_reads: invalid requested/returned totals")
    _check_histogram(
        reads.get("request_histogram"),
        f"{path}.source_reads.request_histogram",
        calls,
        observed_bytes=requested,
    )
    _check_histogram(
        reads.get("returned_histogram"),
        f"{path}.source_reads.returned_histogram",
        calls,
        observed_bytes=returned,
    )
    authored_observation = _object(item.get("authored"), f"{path}.authored")
    opens = _int_field(authored_observation.get("opens"), f"{path}.authored.opens")
    require(opens == EXPECTED_AUTHORED_OPENS, f"{path}.authored.opens: expected {EXPECTED_AUTHORED_OPENS}")
    expected_events = authored["event_count"] * opens
    expected_text_chunks = (authored["event_count"] - 2 * authored["authored_count"]) * opens
    expected_text_bytes = authored["text_bytes"] * opens
    require(authored_observation.get("events") == expected_events, f"{path}.authored.events: differs from sealed multiple")
    require(authored_observation.get("text_chunks") == expected_text_chunks, f"{path}.authored.text_chunks: differs from sealed multiple")
    require(authored_observation.get("text_bytes") == expected_text_bytes, f"{path}.authored.text_bytes: differs from sealed multiple")
    _int_field(authored_observation.get("events"), f"{path}.authored.events")
    _int_field(authored_observation.get("text_chunks"), f"{path}.authored.text_chunks")
    _int_field(authored_observation.get("text_bytes"), f"{path}.authored.text_bytes")
    sink = _object(item.get("sink"), f"{path}.sink")
    accepted = _positive_int(sink.get("accepted_bytes"), f"{path}.sink.accepted_bytes")
    write_calls = _positive_int(sink.get("write_calls"), f"{path}.sink.write_calls")
    largest = _positive_int(sink.get("largest_write"), f"{path}.sink.largest_write")
    require(accepted == oracle["candidate_archive_bytes"], f"{path}.sink.accepted_bytes: final output differs from candidate archive")
    require(sink.get("sha256") == oracle["candidate_archive_sha256"], f"{path}.sink.sha256: final output identity differs")
    require(largest <= sink_write_bytes, f"{path}.sink.largest_write: exceeds sink bound")
    _sha256_field(sink.get("sha256"), f"{path}.sink.sha256")
    _check_histogram(
        sink.get("histogram"),
        f"{path}.sink.histogram",
        write_calls,
        observed_bytes=accepted,
        largest_write=largest,
        maximum=sink_write_bytes,
    )
    _check_allocator_sample(item, role, path)
    _check_process_sample(item.get("process"), f"{path}.process")


def _check_report_shell(
    report: Path,
    role: str,
    case: dict[str, Any],
    *,
    samples: int,
    warmups: int,
    binary: dict[str, Any],
    argv: list[str],
) -> dict[str, Any]:
    value = read(report)
    _check_finite_json(value)
    require(isinstance(value, dict) and value.get("schema") == REPORT_SCHEMA and value.get("version") == 1, f"{report}: report schema differs")
    expected_binary = {
        "binary": "litchi-perf-baseline" if role == "normal" else "litchi-perf-baseline-alloc",
        "allocator": "Rust system allocator" if role == "normal" else "CountingSystemAllocator(std::alloc::System)",
        "instrumentation": "none" if role == "normal" else "system_allocator_operation_scoped",
        "counter_revision": None if role == "normal" else "serialized_region_peak_v3",
    }
    require(value.get("binary") == expected_binary, f"{report}: report binary identity differs")
    config = value.get("config")
    require(isinstance(config, dict), f"{report}: report config is missing")
    require(config.get("source_counts") == [case["source_count"]], f"{report}: source count differs")
    require(config.get("authored_counts") == [case["authored_count"]], f"{report}: authored count differs")
    require(config.get("chunk_modes") == [REPORT_CHUNK_MODES[case["chunk_mode"]]], f"{report}: chunk mode differs")
    require(config.get("text_modes") == [REPORT_TEXT_MODES[case["text_mode"]]], f"{report}: text mode differs")
    require(_positive_int(config.get("samples"), f"{report}.config.samples") == samples, f"{report}: sample contract differs")
    require(_positive_int(config.get("warmups"), f"{report}.config.warmups") == warmups, f"{report}: warmup contract differs")
    require(_positive_int(config.get("sink_write_bytes"), f"{report}.config.sink_write_bytes") == SINK_WRITE_BYTES, f"{report}: sink bound differs")
    require(_positive_int(config.get("expected_authored_opens"), f"{report}.config.expected_authored_opens") == EXPECTED_AUTHORED_OPENS, f"{report}: authored-open contract differs")
    require(config.get("lifecycle") == LIFECYCLE, f"{report}: lifecycle differs")
    require(config.get("source") == SOURCE_CONTRACT, f"{report}: source contract differs")
    require(config.get("authored_provider") == AUTHORED_PROVIDER, f"{report}: authored provider differs")
    require(config.get("sink") == SINK_CONTRACT, f"{report}: sink contract differs")
    require(config.get("fixture_dir") is None, f"{report}: fixture export must remain outside formal capture")
    cases = value.get("cases")
    require(isinstance(cases, list) and len(cases) == 1, f"{report}: expected one case per process")
    observed = cases[0]
    require(
        isinstance(observed, dict)
        and observed.get("source_count") == case["source_count"]
        and observed.get("authored_count") == case["authored_count"]
        and observed.get("chunk_mode") == REPORT_CHUNK_MODES[case["chunk_mode"]]
        and observed.get("text_mode") == REPORT_TEXT_MODES[case["text_mode"]],
        f"{report}: case identity differs",
    )
    source = observed.get("source")
    authored = observed.get("authored")
    limits = observed.get("limits")
    oracle = observed.get("oracle")
    proof = observed.get("proof")
    require(isinstance(source, dict), f"{report}: source metadata is missing")
    require(isinstance(authored, dict), f"{report}: authored metadata is missing")
    require(isinstance(limits, dict), f"{report}: stream-window metadata is missing")
    require(isinstance(oracle, dict), f"{report}: artifact oracle metadata is missing")
    require(isinstance(proof, dict), f"{report}: lifecycle proof metadata is missing")
    _check_source_identity(source, case, f"{report}.cases[0].source")
    _check_authored_identity(authored, case, f"{report}.cases[0].authored")
    _check_limits(limits, authored, case, f"{report}.cases[0].limits")
    _check_oracle_and_proof(
        source,
        authored,
        limits,
        oracle,
        proof,
        case,
        f"{report}.cases[0]",
    )
    observed_samples = observed.get("samples")
    require(isinstance(observed_samples, list) and len(observed_samples) == samples, f"{report}: sample cardinality differs")
    for index, sample in enumerate(observed_samples):
        _check_sample(
            sample,
            index,
            role,
            case,
            oracle,
            authored,
            sink_write_bytes=SINK_WRITE_BYTES,
            path=f"{report}.cases[0].samples[{index}]",
        )
    require(SHA256.fullmatch(binary.get("sha256", "")) is not None, f"{report}: binary metadata is malformed")
    require(len(argv) > 7 and argv[7] == binary.get("path"), f"{report}: argv binary binding differs")
    return value


def capture_one(
    attempt: str,
    role: str,
    case_label: str,
    *,
    pilot: bool,
    repeat: int | None,
) -> Path:
    require(role in ROLES, f"unknown role {role}")
    if pilot:
        require(repeat is None, "pilot runs do not take a formal repeat")
    else:
        require(repeat in REPEATS, "formal repeat must be 1 or 2")
    case = _case_from_label(case_label)
    protocol, protocol_sha256 = load_protocol()
    builds = _load_builds(attempt, protocol_sha256)
    identity = _run_identity("pilot" if pilot else "formal", role, case, repeat=repeat)
    directory = _run_directory(attempt, identity)
    samples = PILOT_SAMPLES if pilot else FORMAL_SAMPLES
    warmups = PILOT_WARMUPS if pilot else FORMAL_WARMUPS
    binary = builds[role]["binary"]
    report = directory / "report.json"
    resource = directory / "resource.txt"
    stdout = directory / "stdout.txt"
    stderr = directory / "stderr.txt"
    argv = _argv(binary, case, samples=samples, warmups=warmups, report=report, resource=resource)
    started = {
        "schema": CAPTURE_SCHEMA,
        "version": 1,
        "status": "running",
        "attempt": attempt,
        "run": identity,
        "protocol": {"path": "protocol.json", "sha256": protocol_sha256},
        "build": {
            "path": (Path("attempts") / attempt / f"build-{role}.json").as_posix(),
            "sha256": _json_hash(ROOT / "attempts" / attempt / f"build-{role}.json"),
        },
        "binary": binary,
        "argv": argv,
        "cwd": str(REPO),
        "environment": {key: ENV[key] for key in ENV_KEYS},
        "started_utc": now(),
    }
    started_path = directory / "started.json"
    write(started_path, started)
    launch_error: str | None = None
    try:
        with stdout.open("xb") as out, stderr.open("xb") as err:
            process = subprocess.run(argv, cwd=REPO, env=ENV, stdout=out, stderr=err, check=False)
        exit_code: int | None = process.returncode
    except (OSError, subprocess.SubprocessError) as error:
        process = None
        exit_code = None
        launch_error = f"{type(error).__name__}: {error}"
    artifacts: dict[str, dict[str, Any]] = {}
    for path in (stdout, stderr, resource, report):
        if path.is_file():
            artifacts[path.name] = meta(path)
    missing = [path.name for path in (stdout, stderr, resource, report) if not path.is_file()]
    validation_error: str | None = None
    if exit_code == 0:
        try:
            _check_report_shell(
                report,
                role,
                case,
                samples=samples,
                warmups=warmups,
                binary=binary,
                argv=argv,
            )
        except (MeasureError, OSError, ValueError) as error:
            validation_error = str(error)
    passed = exit_code == 0 and validation_error is None and not missing
    finished = dict(
        started,
        status="pass" if passed else "failed",
        exit_code=exit_code,
        finished_utc=now(),
        artifacts=artifacts,
        missing_artifacts=missing,
    )
    if launch_error is not None:
        finished["launch_error"] = launch_error
    if validation_error is not None:
        finished["validation_error"] = validation_error
    write(directory / "receipt.json", finished)
    if not passed:
        detail = launch_error or validation_error or f"exit {exit_code}"
        fail(f"{identity['label']} failed ({detail}); receipt retained")
    print(f"captured {identity['label']} ({samples} samples, {warmups} warmups)")
    return directory / "receipt.json"


def _formal_runs(protocol: dict[str, Any]) -> Iterable[dict[str, Any]]:
    runs = protocol.get("formal_runs")
    require(isinstance(runs, list) and len(runs) == len(CASES) * len(ROLES) * len(REPEATS), "protocol formal run inventory differs")
    return runs


def _pilot_runs(protocol: dict[str, Any]) -> Iterable[dict[str, Any]]:
    runs = protocol.get("pilot_runs")
    require(isinstance(runs, list) and len(runs) == len(CASES) * len(ROLES), "protocol pilot run inventory differs")
    return runs


def run_all(attempt: str, *, pilot: bool) -> None:
    protocol, _ = load_protocol()
    runs = _pilot_runs(protocol) if pilot else _formal_runs(protocol)
    for run in runs:
        capture_one(
            attempt,
            run["role"],
            run["case"],
            pilot=pilot,
            repeat=None if pilot else run["repeat"],
        )


def plan() -> None:
    value = _protocol_value()
    print(json.dumps(value, indent=2, sort_keys=True))


def parser() -> argparse.ArgumentParser:
    command = argparse.ArgumentParser(description=__doc__)
    subcommands = command.add_subparsers(dest="command", required=True)
    subcommands.add_parser("freeze", help="write the immutable measurement protocol")
    subcommands.add_parser("plan", help="print the deduplicated protocol without writing it")
    for name, help_text in (
        ("build-normal", "build and retain the normal binary"),
        ("build-allocator", "build and retain the allocator binary"),
        ("capture-all", "run all formal one-case processes"),
        ("pilot-all", "run all excluded pilot one-case processes"),
    ):
        sub = subcommands.add_parser(name, help=help_text)
        sub.add_argument("--attempt", required=True)
    for name, help_text in (
        ("capture", "run one formal one-case process"),
        ("pilot", "run one excluded pilot one-case process"),
    ):
        sub = subcommands.add_parser(name, help=help_text)
        sub.add_argument("--attempt", required=True)
        sub.add_argument("--role", choices=ROLES, required=True)
        sub.add_argument("--case", choices=tuple(CASE_BY_LABEL), required=True)
        if name == "capture":
            sub.add_argument("--repeat", type=int, choices=REPEATS, required=True)
    return command


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    try:
        if args.command == "freeze":
            freeze()
        elif args.command == "plan":
            plan()
        elif args.command == "build-normal":
            build("normal", _attempt(args.attempt))
        elif args.command == "build-allocator":
            build("allocator", _attempt(args.attempt))
        elif args.command == "capture-all":
            with _cpu_lock():
                run_all(_attempt(args.attempt), pilot=False)
        elif args.command == "pilot-all":
            with _cpu_lock():
                run_all(_attempt(args.attempt), pilot=True)
        elif args.command == "capture":
            with _cpu_lock():
                capture_one(_attempt(args.attempt), args.role, args.case, pilot=False, repeat=args.repeat)
        elif args.command == "pilot":
            with _cpu_lock():
                capture_one(_attempt(args.attempt), args.role, args.case, pilot=True, repeat=None)
        else:
            fail(f"unknown command {args.command!r}")
    except (MeasureError, OSError, ValueError, subprocess.SubprocessError) as error:
        print(f"measure.py: FAIL: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
