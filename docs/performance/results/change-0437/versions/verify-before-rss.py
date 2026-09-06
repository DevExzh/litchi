#!/usr/bin/env python3
"""Portable, read-only verifier for the 0437 ODP formal matrix.

The copied ODP oracle owns report and catalog semantics.  This outer verifier
owns the frozen matrix, build/source/binary bindings, process receipts, phase
ordering, profile custody, and cross-role identity.  It never starts a
workload, builds a binary, invokes Git, or requires a live checkout after the
binary/source manifests have been retained.
"""

from __future__ import annotations

import argparse
import datetime as dt
import gzip
import hashlib
import importlib.util
import json
from pathlib import Path
import re
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent
CHANGE = 437
ROLES = ("before-buffered", "after-buffered", "after-streaming")
MODES = ("normal", "allocator")
SHAPES = {"tiny": 64, "medium": 4_096, "large": 8_192}
PROFILE_KINDS = ("stat", "record")
PROFILE_EVENTS = [
    "cycles:u", "instructions:u", "branches:u", "branch-misses:u",
    "L1-dcache-load-misses:u",
]
PHASES = {
    "A1": ("before-buffered", 0, 6, "R1"),
    "B1": ("after-buffered", 0, 6, "R1"),
    "C1": ("after-streaming", 0, 6, "R1"),
    "C2": ("after-streaming", 6, 12, "R2"),
    "B2": ("after-buffered", 6, 12, "R2"),
    "A2": ("before-buffered", 6, 12, "R2"),
}
PHASE_ORDER = tuple(PHASES)
HEX40 = re.compile(r"^[0-9a-fA-F]{40}$")
HEX64 = re.compile(r"^[0-9a-fA-F]{64}$")
MAX_JSON_BYTES = 512 * 1024 * 1024
MAX_ARTIFACT_BYTES = 512 * 1024 * 1024
RSS_RE = re.compile(r"^Maximum resident set size \(kbytes\):\s*(\d+)\s*$")


class VerificationError(ValueError):
    pass


def fail(label: str, message: str) -> None:
    raise VerificationError(f"{label}: {message}")


def duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise VerificationError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def reject_constant(value: str) -> Any:
    raise VerificationError(f"non-finite JSON value {value!r}")


def load(path: Path, label: str | None = None) -> Any:
    label = label or str(path)
    if not path.is_file():
        fail(label, "file is missing")
    if path.stat().st_size > MAX_JSON_BYTES:
        fail(label, f"JSON exceeds {MAX_JSON_BYTES} bytes")
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=duplicate_pairs,
            parse_constant=reject_constant,
        )
    except (OSError, UnicodeError, json.JSONDecodeError, VerificationError) as error:
        fail(label, f"invalid JSON: {error}")
    raise AssertionError("unreachable")


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(label, "expected an object")
    return value


def array(value: Any, label: str) -> list[Any]:
    if not isinstance(value, list):
        fail(label, "expected an array")
    return value


def u64(value: Any, label: str, *, positive: bool = False) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < (1 if positive else 0) or value > (1 << 64) - 1:
        fail(label, "expected a u64")
    return value


def digest(value: Any, label: str) -> str:
    if not isinstance(value, str) or HEX64.fullmatch(value) is None:
        fail(label, "expected a SHA-256 digest")
    return value.lower()


def timestamp(value: Any, label: str) -> dt.datetime:
    if not isinstance(value, str) or not value:
        fail(label, "expected an ISO-8601 timestamp")
    try:
        parsed = dt.datetime.fromisoformat(value)
    except ValueError as error:
        fail(label, f"invalid timestamp: {error}")
    if parsed.tzinfo is None or parsed.utcoffset() is None:
        fail(label, "timestamp must include a timezone")
    return parsed


def sha(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            value.update(block)
    return value.hexdigest()


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def bundle_path(value: Any, label: str) -> Path:
    if not isinstance(value, str) or not value or Path(value).is_absolute():
        fail(label, "must be a non-empty bundle-relative path")
    path = (ROOT / value).resolve()
    if not path.is_relative_to(ROOT.resolve()):
        fail(label, "path escapes the evidence bundle")
    return path


def protocol_path() -> Path:
    path = ROOT / "protocol.json"
    if not path.is_file():
        fail("protocol.json", "frozen protocol is missing")
    return path


def required_paths(value: Any, key: str, count: int, label: str) -> tuple[str, ...]:
    contract = obj(value, label)
    names = contract.get(key)
    if not isinstance(names, list) or len(names) != count or any(not isinstance(name, str) or not name for name in names):
        fail(f"{label}.{key}", f"expected exactly {count} non-empty paths")
    if len(set(names)) != len(names):
        fail(f"{label}.{key}", "contains duplicate paths")
    for name in names:
        path = Path(name)
        if path.is_absolute() or ".." in path.parts:
            fail(f"{label}.{key}", f"unsafe path {name!r}")
    return tuple(names)


def validate_protocol(value: dict[str, Any]) -> None:
    if value.get("change") != CHANGE:
        fail("protocol.change", f"must be {CHANGE}")
    if value.get("cpu") != 2 or value.get("workers") != 1:
        fail("protocol", "must bind CPU 2 and one worker")
    if value.get("samples") != 30 or value.get("warmups") != 3 or value.get("repeats") != 2:
        fail("protocol", "must bind samples=30, warmups=3, repeats=2")
    if value.get("modes") != list(MODES) or value.get("shapes") != SHAPES:
        fail("protocol", "mode or shape matrix differs from 64/4096/8192 contract")
    roles = obj(value.get("roles"), "protocol.roles")
    if set(roles) != set(ROLES):
        fail("protocol.roles", "must contain the three ODP roles")
    expected_roles = {
        "before-buffered": ("odp_buffered_create", "buffered"),
        "after-buffered": ("odp_buffered_create", "buffered"),
        "after-streaming": ("odp_streaming_create", "streaming"),
    }
    for role, (selector, source_role) in expected_roles.items():
        spec = obj(roles[role], f"protocol.roles.{role}")
        if spec.get("selector") != selector or spec.get("source_field") != "odp_slides" or spec.get("source_role") != source_role:
            fail(f"protocol.roles.{role}", "selector/source binding differs from ODP contract")
        if spec.get("build_directory") not in {"before", "after"}:
            fail(f"protocol.roles.{role}.build_directory", "must be before or after")
    order = array(value.get("order"), "protocol.order")
    if len(order) != 12:
        fail("protocol.order", "must contain 12 lanes")
    expected_order = [
        ("normal", "tiny", "R1"), ("normal", "medium", "R1"), ("normal", "large", "R1"),
        ("allocator", "tiny", "R1"), ("allocator", "medium", "R1"), ("allocator", "large", "R1"),
        ("allocator", "large", "R2"), ("allocator", "medium", "R2"), ("allocator", "tiny", "R2"),
        ("normal", "large", "R2"), ("normal", "medium", "R2"), ("normal", "tiny", "R2"),
    ]
    actual = []
    for index, lane in enumerate(order):
        lane = obj(lane, f"protocol.order[{index}]")
        actual.append((lane.get("mode"), lane.get("shape"), lane.get("repeat")))
    if actual != expected_order:
        fail("protocol.order", f"expected {expected_order!r}, got {actual!r}")
    matrix = obj(value.get("matrix"), "protocol.matrix")
    if tuple(matrix.get("roles", ())) != ROLES or matrix.get("formal_reports") != 36 or matrix.get("retained_samples") != 1080:
        fail("protocol.matrix", "must bind ordered roles, 36 reports, and 1080 samples")
    report_contract = obj(value.get("report_contract"), "protocol.report_contract")
    if report_contract.get("required_reports") != 36 or report_contract.get("retained_samples") != 1080:
        fail("protocol.report_contract", "must bind 36 reports and 1080 samples")
    oracle = obj(value.get("oracle"), "protocol.oracle")
    oracle_protocol = bundle_path(oracle.get("path"), "protocol.oracle.path")
    oracle_verifier = bundle_path(oracle.get("verifier_path"), "protocol.oracle.verifier_path")
    if oracle.get("sha256") != sha(oracle_protocol) or oracle.get("verifier_sha256") != sha(oracle_verifier):
        fail("protocol.oracle", "retained oracle hashes are stale")
    required_paths(value.get("pilots"), "required_receipts", 18, "protocol.pilots")
    required_paths(value.get("preparatory_profiles"), "required_receipts", 2, "protocol.preparatory_profiles")
    required_paths(value.get("formal_profiles"), "required_receipts", 6, "protocol.formal_profiles")
    replay = value.get("replay_drivers")
    if not isinstance(replay, list) or not replay or any(not isinstance(name, str) or Path(name).is_absolute() or ".." in Path(name).parts for name in replay):
        fail("protocol.replay_drivers", "must be a safe nonempty path list")
    cleanup = value.get("cleanup_paths")
    expected_cleanup = value.get("cleanup_expected_paths")
    if not isinstance(cleanup, list) or not cleanup or any(not isinstance(name, str) or not name or not name.startswith("/tmp/litchi-goal-0437-") for name in cleanup):
        fail("protocol.cleanup_paths", "must be a nonempty /tmp/litchi-goal-0437-* path list")
    if cleanup != expected_cleanup:
        fail("protocol.cleanup_paths", "cleanup_paths and cleanup_expected_paths must be identical")


def check_source_manifest(value: Any, label: str) -> dict[str, Any]:
    manifest = obj(value, label)
    path = bundle_path(manifest.get("path"), f"{label}.path")
    expected = digest(manifest.get("sha256"), f"{label}.sha256")
    files = manifest.get("files")
    if isinstance(files, bool) or not isinstance(files, int) or files < 1:
        fail(f"{label}.files", "expected a positive count")
    if sha(path) != expected:
        fail(label, "retained manifest hash differs")
    content = obj(load(path, str(path)), str(path))
    if len(content) != files:
        fail(label, "retained manifest file count differs")
    return {"path": str(path.relative_to(ROOT)), "sha256": expected, "files": files}


def check_build(role: str, protocol_sha: str, require_binaries: bool) -> tuple[Path, dict[str, Any]]:
    protocol = load(protocol_path(), "protocol")
    role_spec = obj(obj(protocol.get("roles"), "protocol.roles")[role], f"protocol.roles.{role}")
    directory = bundle_path(role_spec.get("build_directory"), f"protocol.roles.{role}.build_directory")
    path = directory / "build.json"
    build = obj(load(path, str(path)), str(path))
    if build.get("schema") != "litchi-0437-build-descriptor-v1" or build.get("change") != CHANGE or build.get("role") not in {role, role_spec.get("build_directory")}:
        fail(str(path), "build descriptor identity differs")
    revision = build.get("revision")
    if not isinstance(revision, str) or HEX40.fullmatch(revision) is None:
        fail(str(path), "malformed build revision")
    if build.get("protocol_sha256") != protocol_sha:
        fail(str(path), "protocol binding is stale")
    source = check_source_manifest(build.get("source_manifest"), f"{path}.source_manifest")
    if source != build.get("source_manifest"):
        fail(str(path), "source manifest descriptor is not canonical")
    build_receipt = bundle_path(build.get("build_receipt"), f"{path}.build_receipt")
    if build.get("build_receipt_sha256") != sha(build_receipt):
        fail(str(path), "build receipt hash is stale")
    build_row = obj(load(build_receipt, str(build_receipt)), str(build_receipt))
    if build_row.get("status") != "pass" or build_row.get("exit_code") != 0 or build_row.get("source_unchanged") is not True:
        fail(str(build_receipt), "build receipt is not passing")
    if build_row.get("revision") != revision or build_row.get("source_before") != build_row.get("source_after") or build_row.get("source_before") != source:
        fail(str(build_receipt), "build receipt does not bind descriptor")
    copies_path = bundle_path(build.get("binary_copies_receipt"), f"{path}.binary_copies_receipt")
    if build.get("binary_copies_receipt_sha256") != sha(copies_path):
        fail(str(path), "binary-copy receipt hash is stale")
    binaries = obj(build.get("binaries"), f"{path}.binaries")
    copies = obj(load(copies_path, str(copies_path)), str(copies_path))
    if copies != binaries or set(binaries) != set(MODES):
        fail(str(path), "binary descriptor differs from retained copies")
    for mode in MODES:
        identity = obj(binaries[mode], f"{path}.binaries.{mode}")
        binary = Path(identity.get("path", ""))
        if not binary.is_absolute() or not isinstance(identity.get("bytes"), int) or identity["bytes"] <= 0 or HEX64.fullmatch(str(identity.get("sha256", ""))) is None:
            fail(f"{path}.binaries.{mode}", "malformed binary identity")
        if require_binaries:
            if not binary.is_file() or binary.stat().st_size != identity["bytes"] or sha(binary) != identity["sha256"].lower():
                fail(f"{path}.binaries.{mode}", "binary is missing or stale")
    oracle = obj(build.get("oracle"), f"{path}.oracle")
    protocol_value = obj(load(protocol_path(), "protocol"), "protocol")
    frozen_oracle = obj(protocol_value.get("oracle"), "protocol.oracle")
    if oracle.get("verifier_sha256", oracle.get("sha256")) != frozen_oracle.get("verifier_sha256"):
        fail(str(path), "oracle verifier binding is stale")
    drivers = obj(build.get("driver_hashes"), f"{path}.driver_hashes")
    for name in ("capture.py", "profile.py"):
        if drivers.get(name) != sha(ROOT / name):
            fail(str(path), f"{name} binding is stale")
    return directory, build


def artifact_data(receipt_path: Path, value: Any, label: str) -> tuple[Path, bytes]:
    record = obj(value, label)
    raw_path = bundle_path(record.get("path"), f"{label}.path")
    selected = raw_path
    if not selected.is_file() and Path(str(raw_path) + ".gz").is_file():
        selected = Path(str(raw_path) + ".gz")
    if not selected.is_file():
        fail(label, "artifact is missing in raw and gzip form")
    try:
        content = gzip.decompress(selected.read_bytes()) if selected.name.endswith(".gz") else selected.read_bytes()
    except (OSError, EOFError, gzip.BadGzipFile) as error:
        fail(label, f"invalid gzip artifact: {error}")
    if len(content) > MAX_ARTIFACT_BYTES:
        fail(label, "artifact exceeds bounded verifier size")
    if record.get("bytes") != len(content) or record.get("sha256") != sha_bytes(content):
        fail(label, "artifact hash or size differs from receipt")
    return selected, content


def read_resource(receipt_path: Path, value: Any, label: str) -> dict[str, Any]:
    path, content = artifact_data(receipt_path, value, label)
    text = content.decode("utf-8", errors="replace")
    rss = None
    for line in text.splitlines():
        match = RSS_RE.match(line)
        if match:
            rss = int(match.group(1)) * 1024
            break
    return {"path": str(path.relative_to(ROOT)), "bytes": len(content), "sha256": sha_bytes(content), "peak_rss_bytes": rss, "scope": "gnu_time_v_verbose_whole_fresh_process"}


def oracle_role(protocol: dict[str, Any], role: str) -> str:
    spec = obj(obj(protocol["roles"], "protocol.roles")[role], f"protocol.roles.{role}")
    if isinstance(spec.get("oracle_role"), str):
        return spec["oracle_role"]
    roles = protocol.get("oracle_roles")
    if isinstance(roles, dict) and isinstance(roles.get(role), str):
        return roles[role]
    oracle = protocol.get("oracle")
    if isinstance(oracle, dict) and isinstance(oracle.get("roles"), dict) and isinstance(oracle["roles"].get(role), str):
        return oracle["roles"][role]
    return role


def normalized_identity(verified: dict[str, Any], report_path: Path) -> dict[str, Any]:
    identity = dict(obj(verified.get("identity"), f"{report_path}.identity"))
    identity["output_sha256"] = digest(verified["result"].get("output_sha256"), f"{report_path}.result.output_sha256")
    identity["canonical_text_bytes"] = u64(verified["corpus"].get("input_bytes"), f"{report_path}.corpus.input_bytes")
    source_summary = obj(obj(verified["result"].get("source"), f"{report_path}.result.source").get("odp_slides"), f"{report_path}.result.source.odp_slides")
    for gate in ("page_structure_verified", "page_geometry_verified"):
        if not isinstance(source_summary.get(gate), bool):
            fail(f"{report_path}.result.source.odp_slides.{gate}", "expected a boolean gate")
        identity[gate] = source_summary[gate]
    return identity


def load_oracle(protocol: dict[str, Any]) -> Any:
    oracle = obj(protocol.get("oracle"), "protocol.oracle")
    path = bundle_path(oracle.get("verifier_path"), "protocol.oracle.verifier_path")
    if oracle.get("verifier_sha256") != sha(path):
        fail("protocol.oracle.verifier_sha256", "verifier hash is stale")
    spec = importlib.util.spec_from_file_location("change0437_odp_oracle", path)
    if spec is None or spec.loader is None:
        fail(str(path), "cannot load copied ODP oracle")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    if not callable(getattr(module, "validate_report", None)):
        fail(str(path), "oracle has no validate_report API")
    return module


def expected_lanes(protocol: dict[str, Any], phase: str) -> list[dict[str, Any]]:
    role, first, last, repeat = PHASES[phase]
    order = array(protocol.get("order"), "protocol.order")
    rows = [obj(value, f"protocol.order[{index}]") for index, value in enumerate(order[first:last], first)]
    expected = {(mode, shape, repeat) for mode in MODES for shape in SHAPES}
    actual = {(row.get("mode"), row.get("shape"), row.get("repeat")) for row in rows}
    if actual != expected or len(actual) != len(rows):
        fail(f"protocol.phase.{phase}", "lane set is not the declared six-lane matrix")
    return rows


def check_argv_path(value: Any, artifact: tuple[Path, bytes], label: str) -> None:
    if not isinstance(value, str) or not Path(value).is_absolute():
        fail(label, "argv path must be absolute")
    actual = Path(value)
    logical = artifact[0]
    if logical.name.endswith(".gz"):
        logical = Path(str(logical)[:-3])
    expected = logical.relative_to(ROOT).parts
    if len(actual.parts) < len(expected) or actual.parts[-len(expected):] != expected:
        fail(label, "argv path does not bind the retained artifact")


def check_binary_argv(value: Any, binary: dict[str, Any], label: str) -> None:
    if not isinstance(value, str) or not Path(value).is_absolute():
        fail(label, "binary argv path must be absolute")
    if Path(value) != Path(binary["path"]):
        fail(label, "binary argv path differs from the build descriptor")


def check_capture_argv(receipt: dict[str, Any], protocol: dict[str, Any], lane: dict[str, Any], binary: dict[str, Any], artifacts: dict[str, tuple[Path, bytes]], label: str) -> None:
    argv = array(receipt.get("argv"), f"{label}.argv")
    if len(argv) != 22 or any(not isinstance(value, str) or not value for value in argv):
        fail(f"{label}.argv", "capture argv does not match the frozen command shape")
    fixed = {
        0: "taskset", 1: "-c", 2: str(protocol["cpu"]), 3: "/usr/bin/time", 4: "-v", 5: "-o",
        8: "--case", 9: protocol["roles"][receipt["role"]]["selector"], 10: "--semantic-shape", 11: lane["shape"],
        12: "--workers", 13: str(protocol["workers"]), 14: "--samples", 15: str(protocol["samples"]),
        16: "--warmup", 17: str(protocol["warmups"]), 18: "--json", 20: "--corpus-manifest",
    }
    if argv[2] != str(protocol["cpu"]):
        fail(f"{label}.argv", "CPU affinity differs from the frozen protocol")
    for index, expected in fixed.items():
        if argv[index] != expected:
            fail(f"{label}.argv[{index}]", f"must be {expected!r}")
    check_argv_path(argv[6], artifacts["resource_log"], f"{label}.argv[6]")
    check_binary_argv(argv[7], binary, f"{label}.argv[7]")
    check_argv_path(argv[19], artifacts["report"], f"{label}.argv[19]")
    check_argv_path(argv[21], artifacts["catalog"], f"{label}.argv[21]")


def check_profile_argv(receipt: dict[str, Any], protocol: dict[str, Any], binary: dict[str, Any], artifacts: dict[str, tuple[Path, bytes]], kind: str, label: str) -> None:
    if kind == "record":
        if receipt.get("record_event") != "cycles:u" or receipt.get("record_frequency_hz") != 999 or receipt.get("call_graph") != "fp,127" or receipt.get("stat_events") != []:
            fail(label, "perf-record sampling/call-graph binding is stale")
    elif receipt.get("stat_events") != PROFILE_EVENTS or receipt.get("record_event") is not None or receipt.get("record_frequency_hz") is not None or receipt.get("call_graph") is not None:
        fail(label, "perf-stat event binding is stale")
    argv = array(receipt.get("argv"), f"{label}.argv")
    expected_length = 31 if kind == "stat" else 34
    if len(argv) != expected_length or any(not isinstance(value, str) or not value for value in argv):
        fail(f"{label}.argv", "profile argv contains malformed arguments")
    if argv[:6] != ["taskset", "-c", str(protocol["cpu"]), "/usr/bin/time", "-v", "-o"]:
        fail(f"{label}.argv", "profile argv prefix differs from protocol")
    check_argv_path(argv[6], artifacts["resource"], f"{label}.argv[6]")
    if kind == "stat":
        expected = ["perf", "stat", "--no-big-num", "-x,", "-e", ",".join(PROFILE_EVENTS), "-o"]
        if argv[7:14] != expected:
            fail(f"{label}.argv", "perf-stat argv differs from the frozen event command")
        check_argv_path(argv[14], artifacts["perf_stat"], f"{label}.argv[14]")
        if argv[15] != "--":
            fail(f"{label}.argv", "perf-stat must terminate its profiler options with --")
        tail = 16
    else:
        if argv[7:11] != ["perf", "record", "--no-buildid-cache", "-o"]:
            fail(f"{label}.argv", "perf-record argv prefix differs from the frozen command")
        check_argv_path(argv[11], artifacts["perf_data"], f"{label}.argv[11]")
        if argv[12:19] != ["-F", "999", "-e", "cycles:u", "--call-graph", "fp,127", "--"]:
            fail(f"{label}.argv", "perf-record sampling/call-graph binding is stale")
        tail = 19
    base = ["--case", receipt["selector"], "--semantic-shape", "large", "--workers", str(protocol["workers"]), "--samples", str(protocol["samples"]), "--warmup", str(protocol["warmups"]), "--json"]
    if argv[tail + 1:tail + 1 + len(base)] != base:
        fail(f"{label}.argv", "profile workload arguments differ from the protocol")
    check_binary_argv(argv[tail], binary, f"{label}.argv[{tail}]")
    report_index = tail + 1 + len(base)
    check_argv_path(argv[report_index], artifacts["report"], f"{label}.argv[{report_index}]")
    if len(argv) != report_index + 3 or argv[report_index + 1] != "--corpus-manifest":
        fail(f"{label}.argv", "profile corpus-manifest arguments are malformed")
    check_argv_path(argv[report_index + 2], artifacts["catalog"], f"{label}.argv[{report_index + 2}]")


def check_state(path: Path, phase: str, role: str, build: dict[str, Any], ambient: dict[str, Any], protocol_sha: str) -> list[str]:
    state = obj(load(path, str(path)), str(path))
    if state.get("schema") != "litchi-0437-capture-state-v1" or state.get("change") != CHANGE or state.get("status") != "pass":
        fail(str(path), "capture state is not a passing 0437 state")
    if state.get("phase") != phase or state.get("attempt") != "formal" or state.get("role") != role or state.get("expected_lanes") != 6 or state.get("completed_lanes") != 6:
        fail(str(path), "capture state phase/role/completion differs")
    if state.get("source_manifest") != build.get("source_manifest") or state.get("ambient_source_manifest") != ambient.get("source_manifest"):
        fail(str(path), "capture state source binding differs")
    if state.get("source_before") != state.get("source_after") or state.get("source_before") != ambient.get("source_manifest"):
        fail(str(path), "capture state source custody differs")
    if state.get("protocol_sha256") != protocol_sha or state.get("driver_sha256") != sha(ROOT / "capture.py"):
        fail(str(path), "capture state driver/protocol binding is stale")
    timestamp(state.get("finished_utc"), f"{path}.finished_utc")
    index = array(state.get("index"), f"{path}.index")
    if len(index) != 6 or any(not isinstance(value, str) or not value.endswith("-receipt.json") for value in index):
        fail(str(path), "capture state receipt index is malformed")
    return index


def check_capture_receipt(path: Path, phase: str, role: str, lane: dict[str, Any], build: dict[str, Any], ambient: dict[str, Any], protocol: dict[str, Any], protocol_sha: str, oracle: Any) -> dict[str, Any]:
    receipt = obj(load(path, str(path)), str(path))
    if receipt.get("schema") != "litchi-0437-capture-receipt-v1" or receipt.get("change") != CHANGE:
        fail(str(path), "capture receipt schema differs")
    if receipt.get("status") != "pass" or receipt.get("exit_code") != 0 or receipt.get("oracle_exit_code") != 0:
        fail(str(path), "capture lane is not passing")
    if receipt.get("phase") != phase or receipt.get("attempt") != "formal" or receipt.get("role") != role or receipt.get("lane") != lane:
        fail(str(path), "capture phase/role/lane binding differs")
    spec = obj(protocol["roles"][role], f"protocol.roles.{role}")
    if receipt.get("selector") != spec.get("selector") or receipt.get("source_field") != "odp_slides":
        fail(str(path), "capture selector/source binding differs")
    if receipt.get("revision") != build.get("revision") or receipt.get("source_manifest") != build.get("source_manifest") or receipt.get("ambient_source_manifest") != ambient.get("source_manifest"):
        fail(str(path), "capture executable/source binding differs")
    if receipt.get("binary") != build.get("binaries", {}).get(lane["mode"]):
        fail(str(path), "capture receipt binary identity differs from descriptor")
    if receipt.get("protocol_sha256") != protocol_sha or receipt.get("driver_sha256") != sha(ROOT / "capture.py"):
        fail(str(path), "capture protocol/driver binding is stale")
    if receipt.get("source_before") != receipt.get("source_after") or receipt.get("source_unchanged") is not True or receipt.get("outside_bundle_status_unchanged") is not True:
        fail(str(path), "capture source custody changed")
    started = timestamp(receipt.get("started_utc"), f"{path}.started_utc")
    finished = timestamp(receipt.get("finished_utc"), f"{path}.finished_utc")
    if finished < started:
        fail(str(path), "capture finished before it started")
    artifacts = obj(receipt.get("artifacts"), f"{path}.artifacts")
    expected = {"report", "catalog", "workload_log", "resource_log", "oracle_log"}
    if set(artifacts) != expected:
        fail(str(path), "capture artifact inventory differs")
    artifact_paths: dict[str, tuple[Path, bytes]] = {
        name: artifact_data(path, artifacts[name], f"{path}.artifacts.{name}") for name in expected
    }
    check_capture_argv(receipt, protocol, lane, build["binaries"][lane["mode"]], artifact_paths, str(path))
    oracle_log = artifact_paths["oracle_log"][1]
    if b"stdout=VALID" not in oracle_log:
        fail(f"{path}.artifacts.oracle_log", "oracle log does not retain VALID")
    if receipt.get("oracle_stdout_sha256") != sha_bytes(oracle_log):
        fail(str(path), "oracle log hash differs")
    report_path = artifact_paths["report"][0]
    if report_path.name.endswith(".gz"):
        fail(str(path), "formal report must remain directly readable; only logs may be sealed")
    mode = lane["mode"]
    shape = lane["shape"]
    verified = oracle.validate_report(report_path, mode, shape, oracle_role(protocol, role))
    if mode == "allocator":
        allocation = obj(verified["metrics"].get("allocation"), f"{report_path}.operation_metrics.allocation")
        before = obj(allocation.get("live_bytes_before"), f"{report_path}.allocation.live_bytes_before").get("values")
        after = obj(allocation.get("live_bytes_after"), f"{report_path}.allocation.live_bytes_after").get("values")
        if not isinstance(before, list) or not isinstance(after, list) or len(before) != len(after) or any(not isinstance(value, int) for value in before + after):
            fail(str(report_path), "allocator live-byte vectors are malformed")
        if any(exit_value - entry != 0 for entry, exit_value in zip(before, after, strict=True)):
            fail(str(report_path), "fresh complete ODP operation retained nonzero live-byte delta")
    report = obj(verified["report"], f"{report_path}.report")
    binary = obj(build["binaries"][mode], f"{path}.binary")
    reported_binary = obj(report.get("binary_identity"), f"{report_path}.binary_identity")
    if reported_binary.get("path") != binary.get("path") or reported_binary.get("binary_sha256") != binary.get("sha256") or reported_binary.get("binary_bytes") != binary.get("bytes"):
        fail(str(report_path), "report binary identity differs from descriptor")
    environment = obj(report.get("environment"), f"{report_path}.environment")
    if not isinstance(environment.get("git_revision"), str) or HEX40.fullmatch(environment["git_revision"]) is None or not isinstance(environment.get("git_worktree_dirty"), bool):
        fail(str(report_path), "runtime ambient identity is malformed")
    identity = normalized_identity(verified, report_path)
    return {
        "phase": phase, "role": role, "lane": lane, "name": receipt.get("name"),
        "receipt": receipt, "report": report, "verified": verified,
        "report_path": str(report_path.relative_to(ROOT)),
        "catalog_path": str(artifact_paths["catalog"][0].relative_to(ROOT)),
        "resource": read_resource(path, artifacts["resource_log"], f"{path}.artifacts.resource_log"),
        "started": started, "finished": finished, "identity": identity,
        "environment": {"git_revision": environment["git_revision"], "git_worktree_dirty": environment["git_worktree_dirty"]},
    }


def verify_profiles(protocol: dict[str, Any], builds: dict[str, tuple[Path, dict[str, Any]]], ambient: set[tuple[str, bool]], protocol_sha: str, oracle: Any) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for role in ROLES:
        for kind in PROFILE_KINDS:
            path = ROOT / "profiles" / role / kind / "receipt.json"
            receipt = obj(load(path, str(path)), str(path))
            if receipt.get("schema") != "litchi-0437-profile-receipt-v1" or receipt.get("change") != CHANGE or receipt.get("status") != "pass":
                fail(str(path), "formal profile is not passing")
            if receipt.get("role") != role or receipt.get("kind") != kind or receipt.get("attempt") != "formal" or receipt.get("preparatory") is not False or receipt.get("shape") != "large":
                fail(str(path), "formal profile identity differs")
            if receipt.get("protocol_sha256") != protocol_sha or receipt.get("driver_sha256") != sha(ROOT / "profile.py") or receipt.get("capture_helper_sha256") != sha(ROOT / "capture.py"):
                fail(str(path), "profile driver/protocol binding is stale")
            oracle_spec = obj(protocol.get("oracle"), "protocol.oracle")
            if receipt.get("oracle_verifier_sha256") != oracle_spec.get("verifier_sha256"):
                fail(str(path), "profile oracle binding is stale")
            build_receipt = bundle_path(receipt.get("build_receipt"), f"{path}.build_receipt")
            copies_receipt = bundle_path(receipt.get("binary_copies_receipt"), f"{path}.binary_copies_receipt")
            if receipt.get("build_receipt_sha256") != sha(build_receipt) or receipt.get("binary_copies_receipt_sha256") != sha(copies_receipt):
                fail(str(path), "profile build/copy receipt binding is stale")
            if receipt.get("source_manifest") != builds[role][1].get("source_manifest") or receipt.get("ambient_source_manifest") != builds["after-streaming"][1].get("source_manifest"):
                fail(str(path), "profile source binding differs")
            if receipt.get("source_before") != receipt.get("source_after") or receipt.get("source_unchanged") is not True:
                fail(str(path), "profile source custody changed")
            started = timestamp(receipt.get("started_utc"), f"{path}.started_utc")
            finished = timestamp(receipt.get("finished_utc"), f"{path}.finished_utc")
            if finished < started or receipt.get("exit_code") != 0 or receipt.get("oracle_exit_code") != 0:
                fail(str(path), "profile terminal bindings are invalid")
            artifacts = obj(receipt.get("artifacts"), f"{path}.artifacts")
            required = {"report", "catalog", "resource", "workload_log", "oracle_log", "perf_stat" if kind == "stat" else "perf_data"}
            if kind == "record":
                required.update({"perf_script", "perf_report"})
            if set(artifacts) != required:
                fail(str(path), "profile artifact inventory differs")
            data = {key: artifact_data(path, artifacts[key], f"{path}.artifacts.{key}") for key in required}
            check_profile_argv(receipt, protocol, builds[role][1]["binaries"]["normal"], data, kind, str(path))
            if not data["oracle_log"][1].startswith(b"argv=") or b"stdout=VALID" not in data["oracle_log"][1]:
                fail(f"{path}.artifacts.oracle_log", "profile oracle log does not retain VALID")
            profiler_artifacts = ("perf_stat",) if kind == "stat" else ("perf_data", "perf_script", "perf_report")
            if any(not data[name][1] for name in profiler_artifacts):
                fail(str(path), "retained profiler artifact is empty")
            report_path = data["report"][0]
            verified = oracle.validate_report(report_path, "normal", "large", oracle_role(protocol, role))
            report = obj(verified["report"], f"{report_path}.report")
            binary = builds[role][1]["binaries"]["normal"]
            if receipt.get("binary") != binary:
                fail(str(path), "profile receipt binary identity differs from descriptor")
            reported_binary = obj(report.get("binary_identity"), f"{report_path}.binary_identity")
            if reported_binary.get("path") != binary.get("path") or reported_binary.get("binary_sha256") != binary.get("sha256") or reported_binary.get("binary_bytes") != binary.get("bytes"):
                fail(str(report_path), "profile report binary identity differs")
            environment = obj(report.get("environment"), f"{report_path}.environment")
            if (environment.get("git_revision"), environment.get("git_worktree_dirty")) not in ambient:
                fail(str(report_path), "profile ambient identity differs")
            rows.append({"role": role, "kind": kind, "receipt": receipt, "report": verified, "identity": normalized_identity(verified, report_path), "resource": read_resource(path, artifacts["resource"], f"{path}.artifacts.resource"), "started": started, "finished": finished})
    return rows


def compare_identities(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    common = ("slide_count", "title_count", "body_count", "title_text_bytes", "body_text_bytes", "canonical_text_bytes", "title_variant_counts", "body_variant_counts", "semantic_sha256", "styles_xml_sha256", "meta_xml_sha256", "member_names", "mimetype", "manifest_entry_count", "page_structure_verified", "page_geometry_verified")
    output_identity = ("archive_bytes", "archive_sha256", "target_payload_bytes", "target_payload_sha256", "output_sha256", "content_xml_sha256", *common)
    buffered_control_identity = ("archive_bytes", "archive_sha256", "target_payload_bytes", "target_payload_sha256", "output_sha256", "content_xml_sha256")
    identity_pairs: list[dict[str, Any]] = []
    by_key: dict[tuple[str, str, str], dict[str, dict[str, Any]]] = {}
    for row in rows:
        lane = row["lane"]
        by_key.setdefault((lane["mode"], lane["shape"], lane["repeat"]), {})[row["role"]] = row
    for key, roles in sorted(by_key.items()):
        if set(roles) != set(ROLES):
            fail(f"cross_role.{key}", "one report per declared role is required")
        reference = roles["after-buffered"]["identity"]
        for role in ROLES:
            actual = roles[role]["identity"]
            mismatches = {field: {"reference": reference.get(field), "actual": actual.get(field)} for field in common if actual.get(field) != reference.get(field)}
            if mismatches:
                fail(f"cross_role.{key}.{role}", f"semantic/topology identity differs: {mismatches}")
        buffered = roles["before-buffered"]["identity"]
        buffered_after = roles["after-buffered"]["identity"]
        buffered_mismatches = {field: {"before": buffered.get(field), "after": buffered_after.get(field)} for field in buffered_control_identity if buffered.get(field) != buffered_after.get(field)}
        if buffered_mismatches:
            fail(f"buffered_control.{key}", f"same-builder archive/content/output identity differs: {buffered_mismatches}")
        identity_pairs.append({"mode": key[0], "shape": key[1], "repeat": key[2], "common_identity": {field: reference.get(field) for field in common}, "buffered_control_exact_identity": {field: buffered_after.get(field) for field in buffered_control_identity}, "cross_api_scope": "semantic/style/meta/frame/page/topology only; archive and lexical content remain role-local", "archive_bytes_role_local": True, "content_xml_role_local": True})
    for role in ROLES:
        for mode in MODES:
            for shape in SHAPES:
                selected = [row for row in rows if row["role"] == role and row["lane"]["mode"] == mode and row["lane"]["shape"] == shape]
                if len(selected) != 2:
                    fail(f"repeat.{role}.{mode}.{shape}", "must contain R1 and R2")
                first, second = sorted(selected, key=lambda row: row["lane"]["repeat"])
                full = ("archive_bytes", "archive_sha256", "target_payload_bytes", "target_payload_sha256", "output_sha256", *common)
                for field in full:
                    if first["identity"].get(field) != second["identity"].get(field):
                        fail(f"repeat.{role}.{mode}.{shape}", f"identity {field} changed")
        for shape in SHAPES:
            reference = next(row for row in rows if row["role"] == role and row["lane"]["mode"] == "normal" and row["lane"]["shape"] == shape and row["lane"]["repeat"] == "R1")["identity"]
            for mode in MODES:
                for repeat in ("R1", "R2"):
                    candidate = next(row for row in rows if row["role"] == role and row["lane"]["mode"] == mode and row["lane"]["shape"] == shape and row["lane"]["repeat"] == repeat)["identity"]
                    mismatches = {field: {"normal_R1": reference.get(field), "actual": candidate.get(field)} for field in output_identity if reference.get(field) != candidate.get(field)}
                    if mismatches:
                        fail(f"same_role_mode.{role}.{shape}.{mode}.{repeat}", f"normal/allocator output identity differs: {mismatches}")
    return identity_pairs


def verify_lifecycle_receipt_contract(protocol: dict[str, Any]) -> dict[str, Any]:
    return {
        "pilots": list(required_paths(protocol["pilots"], "required_receipts", 18, "protocol.pilots")),
        "preparatory_profiles": list(required_paths(protocol["preparatory_profiles"], "required_receipts", 2, "protocol.preparatory_profiles")),
        "formal_profiles": list(required_paths(protocol["formal_profiles"], "required_receipts", 6, "protocol.formal_profiles")),
        "replay_drivers": list(protocol["replay_drivers"]),
        "cleanup_paths": list(protocol["cleanup_paths"]),
    }


def verify_compression(required: bool) -> int:
    path = ROOT / "compression.json"
    if not path.is_file():
        if required:
            fail("compression.json", "sealed compression inventory is required")
        return 0
    records = obj(load(path, str(path)), str(path))
    seen: set[str] = set()
    for stored_name, value in records.items():
        if not isinstance(stored_name, str) or not stored_name.endswith(".gz") or stored_name in seen:
            fail(f"{path}.records", "stored compression paths are malformed or duplicated")
        seen.add(stored_name)
        row = obj(value, f"{path}.{stored_name}")
        if set(row) != {"original_path", "original_sha256", "original_bytes", "stored_sha256", "stored_bytes"}:
            fail(f"{path}.{stored_name}", "compression record fields differ")
        original_name = row.get("original_path")
        if not isinstance(original_name, str) or stored_name[:-3] != original_name:
            fail(f"{path}.{stored_name}.original_path", "does not match stored path")
        stored = bundle_path(stored_name, f"{path}.{stored_name}")
        if not stored.is_file():
            fail(f"{path}.{stored_name}", "stored compressed artifact is missing")
        stored_raw = stored.read_bytes()
        if row.get("stored_bytes") != len(stored_raw) or row.get("stored_sha256") != sha_bytes(stored_raw):
            fail(f"{path}.{stored_name}", "stored artifact binding differs")
        try:
            original_raw = gzip.decompress(stored_raw)
        except (OSError, EOFError, gzip.BadGzipFile) as error:
            fail(f"{path}.{stored_name}", f"invalid gzip stream: {error}")
        original = bundle_path(original_name, f"{path}.{stored_name}.original_path")
        if original.is_file() and original.read_bytes() != original_raw:
            fail(f"{path}.{stored_name}", "raw and compressed artifact bytes differ")
        if row.get("original_bytes") != len(original_raw) or row.get("original_sha256") != sha_bytes(original_raw):
            fail(f"{path}.{stored_name}", "original artifact binding differs")
    actual = {str(candidate.relative_to(ROOT)) for candidate in ROOT.rglob("*.gz") if candidate.is_file()}
    if actual != seen:
        fail(str(path), "compressed artifact inventory differs from compression.json")
    return len(seen)


def verify_inventory(required: bool) -> int:
    path = ROOT / "SHA256SUMS"
    if not path.is_file():
        if required:
            fail("SHA256SUMS", "sealed inventory is required")
        return 0
    rows: dict[str, str] = {}
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as error:
        fail("SHA256SUMS", f"cannot read inventory: {error}")
    for line in lines:
        if "  " not in line:
            fail("SHA256SUMS", "malformed inventory line")
        expected, name = line.split("  ", 1)
        if not HEX64.fullmatch(expected) or not name or name == "SHA256SUMS" or name in rows or Path(name).is_absolute() or ".." in Path(name).parts:
            fail("SHA256SUMS", "duplicate or unsafe inventory entry")
        target = bundle_path(name, f"SHA256SUMS.{name}")
        if not target.is_file() or sha(target) != expected.lower():
            fail(f"SHA256SUMS.{name}", "inventory digest differs")
        rows[name] = expected.lower()
    expected_names = {str(candidate.relative_to(ROOT)) for candidate in ROOT.rglob("*") if candidate.is_file() and candidate.name != "SHA256SUMS"}
    if set(rows) != expected_names:
        fail("SHA256SUMS", "inventory does not cover the bundle exactly")
    return len(rows)


def verify_matrix(*, require_binaries: bool = False) -> dict[str, Any]:
    path = protocol_path()
    protocol = obj(load(path, str(path)), str(path))
    validate_protocol(protocol)
    protocol_sha = sha(path)
    oracle = load_oracle(protocol)
    builds = {role: check_build(role, protocol_sha, require_binaries) for role in ROLES}
    if builds["before-buffered"][1].get("revision") == builds["after-buffered"][1].get("revision"):
        fail("builds", "before and after revisions must remain distinct")
    rows: list[dict[str, Any]] = []
    phase_bounds: dict[str, tuple[dt.datetime, dt.datetime]] = {}
    for phase in PHASE_ORDER:
        role, _, _, _ = PHASES[phase]
        lanes = expected_lanes(protocol, phase)
        state_path = ROOT / "runs" / phase / "formal" / "capture-state.json"
        index = check_state(state_path, phase, role, builds[role][1], builds["after-streaming"][1], protocol_sha)
        phase_dir = state_path.parent
        phase_rows: list[dict[str, Any]] = []
        for lane, indexed in zip(lanes, index, strict=True):
            receipt_path = bundle_path(indexed, f"{state_path}.index")
            if receipt_path.parent != phase_dir:
                fail(str(receipt_path), "capture index points outside its phase directory")
            phase_rows.append(check_capture_receipt(receipt_path, phase, role, lane, builds[role][1], builds["after-streaming"][1], protocol, protocol_sha, oracle))
        for previous, current in zip(phase_rows, phase_rows[1:]):
            if previous["finished"] > current["started"]:
                fail(f"phase.{phase}", "formal lanes overlap or are out of capture order")
        phase_bounds[phase] = (min(row["started"] for row in phase_rows), max(row["finished"] for row in phase_rows))
        rows.extend(phase_rows)
    for previous, current in zip(PHASE_ORDER, PHASE_ORDER[1:]):
        if phase_bounds[previous][1] > phase_bounds[current][0]:
            fail("phase.order", f"{previous} overlaps {current}")
    if len(rows) != 36:
        fail("matrix", "formal matrix must contain 36 reports")
    ambient = {(row["environment"]["git_revision"], row["environment"]["git_worktree_dirty"]) for row in rows}
    if len(ambient) != 1:
        fail("ambient_identity", "runtime revision/dirty state differs across reports")
    profiles = verify_profiles(protocol, builds, ambient, protocol_sha, oracle)
    formal_finished = max(bound[1] for bound in phase_bounds.values())
    previous_profile_finished: dt.datetime | None = None
    profile_identity_fields = ("archive_bytes", "archive_sha256", "target_payload_bytes", "target_payload_sha256", "output_sha256", "content_xml_sha256", "canonical_text_bytes", "semantic_sha256", "styles_xml_sha256", "meta_xml_sha256", "member_names", "mimetype", "manifest_entry_count", "page_structure_verified", "page_geometry_verified")
    for profile in sorted(profiles, key=lambda row: row["started"]):
        if profile["started"] < formal_finished:
            fail(f"profile.{profile['role']}.{profile['kind']}", "profile started before the formal matrix finished")
        if previous_profile_finished is not None and profile["started"] < previous_profile_finished:
            fail(f"profile.{profile['role']}.{profile['kind']}", "profile intervals overlap")
        previous_profile_finished = profile["finished"]
        reference = next(row for row in rows if row["role"] == profile["role"] and row["lane"]["mode"] == "normal" and row["lane"]["shape"] == "large" and row["lane"]["repeat"] == "R1")
        mismatches = {field: {"formal_R1": reference["identity"].get(field), "profile": profile["identity"].get(field)} for field in profile_identity_fields if reference["identity"].get(field) != profile["identity"].get(field)}
        if mismatches:
            fail(f"profile.{profile['role']}.{profile['kind']}", f"large normal profile identity differs from formal R1: {mismatches}")
    return {
        "schema": "litchi-0437-verified-matrix-v1",
        "change": CHANGE,
        "protocol_sha256": protocol_sha,
        "oracle": {"protocol_sha256": digest(protocol["oracle"]["sha256"], "protocol.oracle.sha256"), "verifier_sha256": digest(protocol["oracle"]["verifier_sha256"], "protocol.oracle.verifier_sha256")},
        "builds": {role: build for role, (_, build) in builds.items()},
        "ambient": {"git_revision": next(iter(ambient))[0], "git_worktree_dirty": next(iter(ambient))[1]},
        "matrix": {"formal_reports": len(rows), "retained_samples": len(rows) * protocol["samples"], "roles": list(ROLES), "phases": list(PHASE_ORDER)},
        "rows": rows,
        "identity_pairs": compare_identities(rows),
        "profiles": profiles,
        "lifecycle": verify_lifecycle_receipt_contract(protocol),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--require-binaries", action="store_true")
    parser.add_argument("--portable-check", action="store_true")
    parser.add_argument("--require-inventory", action="store_true")
    parser.add_argument("--stage", choices=("precleanup", "aftercleanup", "final"), default="final")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    try:
        result = verify_matrix(require_binaries=args.require_binaries and not args.portable_check)
        result["stage"] = args.stage
        result["inventory_files"] = verify_inventory(args.require_inventory)
        result["compressed_artifacts"] = verify_compression(args.stage != "precleanup")
    except (OSError, KeyError, TypeError, ValueError, AssertionError) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 1
    if args.json:
        print(json.dumps(result, indent=2, sort_keys=True, allow_nan=False, default=str))
    else:
        print("VALID")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
