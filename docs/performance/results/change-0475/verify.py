#!/usr/bin/env python3
"""Portable, independent verifier for the 0475 profiling evidence bundle.

The capture driver records what was executed, while this module checks the
records without importing the driver, an earlier bundle, or any path in
``/tmp``.  A copied bundle can therefore be checked on a machine that does
not have Rust, perf, Heaptrack, or the source checkout.  ``--live`` is an
explicit opt-in for checking the authenticated source tree and binary.
"""

from __future__ import annotations

import argparse
import copy
import datetime as _datetime
import gzip
import hashlib
import json
import math
import os
from pathlib import Path, PurePosixPath
import re
import subprocess
import sys
from typing import Any, Iterable, Mapping, NoReturn


ROOT = Path(__file__).resolve().parent
SHA_RE = re.compile(r"^[0-9a-f]{64}$")

PROTOCOL_SCHEMA = "litchi-0475-protocol-v1"
BINDING_SCHEMA = "litchi-0475-reused-build-v1"
CAPTURE_SCHEMA = "litchi-0475-capture-v1"
COMPRESSION_SCHEMA = "litchi-0475-compression-v1"
VERIFICATION_SCHEMA = "litchi-0475-verification-v1"
SUPPLEMENT_SCHEMA = "litchi-0475-supplement-protocol-v1"
CHUNKING_SCHEMA = "litchi-0475-chunking-v1"

EXPECTED_ORDER = (
    ("normal-R1", "normal", 30, 3),
    ("counters", "counters", 30, 3),
    ("cpu-P1", "cpu", 30, 3),
    ("heap-H1", "heap", 1, 0),
    ("heap-H2", "heap", 1, 0),
    ("cpu-P2", "cpu", 30, 3),
    ("normal-R2", "normal", 30, 3),
)
EXPECTED_EVENTS = (
    "cycles:u",
    "instructions:u",
    "branches:u",
    "branch-misses:u",
    "cache-misses:u",
    "L1-dcache-load-misses:u",
    "LLC-load-misses:u",
    "page-faults",
    "context-switches",
)
EXPECTED_REUSE_FILES = (
    "SHA256SUMS",
    "build-command.json",
    "build-command.started.json",
    "build-command.stderr",
    "build-command.stdout",
    "build.json",
    "build.py",
    "captures/R1-allocator-large/report.json",
    "captures/R1-normal-large/report.json",
    "captures/R2-normal-large/report.json",
    "protocol.json",
    "source-binding.json",
    "sources/source.json",
)
EXPECTED_COMPRESSION_PATHS = frozenset(
    {
        *(f"cpu-P{i}/{name}" for i in (1, 2) for name in ("perf.data", "perf-script.stdout")),
        *(
            f"heap-H{i}/{name}"
            for i in (1, 2)
            for name in ("decoded.stdout", "print.stdout", "runner-allocations.stdout", "allocation-stacks.txt")
        ),
    }
)
EXPECTED_SUPPLEMENT_PATHS = frozenset(
    {
        *(f"heap-H{i}/{name}" for i in (1, 2) for name in ("runner-mangled.stdout", "runner-mangled-stacks.txt")),
    }
)
EXPECTED_TOOL = {
    "name": "litchi-perf-baseline",
    "version": "0.1.0",
    "binary": "litchi-perf-baseline",
    "profile": "release",
    "target_os": "linux",
    "target_arch": "x86_64",
    "instrumentation": "none",
}
EXPECTED_ENVIRONMENT_KEYS = {
    "rustc_version",
    "git_revision",
    "git_worktree_dirty",
    "logical_cpus_available",
    "allocator",
    "rustflags",
    "cargo_build_target",
    "perf_event_paranoid",
    "os",
    "kernel",
    "cpu_model",
    "total_memory_bytes",
    "page_size_bytes",
    "filesystem_type",
    "source_destination_same_device",
    "cpu_affinity",
    "storage_identifier",
}


class VerificationError(ValueError):
    """A bundle does not meet the frozen evidence contract."""


def fail(message: str) -> NoReturn:
    raise VerificationError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _reject_constant(value: str) -> NoReturn:
    raise ValueError(f"non-finite JSON constant {value}")


def _pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def read_json(path: Path) -> Any:
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=_pairs,
            parse_constant=_reject_constant,
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as error:
        fail(f"cannot read JSON {path}: {error}")


def object_value(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} must be an object")
    return value


def list_value(value: Any, label: str) -> list[Any]:
    require(isinstance(value, list), f"{label} must be an array")
    return value


def exact_keys(value: Mapping[str, Any], expected: Iterable[str], label: str) -> None:
    actual = set(value)
    wanted = set(expected)
    require(actual == wanted, f"{label} keys differ: expected {sorted(wanted)}, got {sorted(actual)}")


def canonical(value: Any) -> bytes:
    try:
        return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), allow_nan=False).encode("utf-8")
    except (TypeError, ValueError) as error:
        fail(f"value is not canonical JSON: {error}")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def file_meta(path: Path) -> dict[str, Any]:
    require(not path.is_symlink(), f"symlink is not an evidence file: {path}")
    require(path.is_file(), f"missing regular evidence file: {path}")
    try:
        size = path.stat().st_size
    except OSError as error:
        fail(f"cannot stat {path}: {error}")
    return {"sha256": sha256(path), "bytes": size}


def safe_relative(raw: Any, label: str) -> str:
    require(isinstance(raw, str) and raw, f"{label} must be a non-empty relative path")
    require("\\" not in raw, f"{label} contains a backslash")
    path = PurePosixPath(raw)
    require(not path.is_absolute(), f"{label} must be relative")
    require(raw not in (".", ".."), f"{label} is not a file path")
    require(all(part not in ("", ".", "..") for part in path.parts), f"{label} escapes the bundle")
    return path.as_posix()


def bundle_path(root: Path, raw: Any, label: str) -> Path:
    relative = safe_relative(raw, label)
    path = root / relative
    # Check every path component, including parents, without following an
    # untrusted symlink into a different tree.
    current = root
    for component in PurePosixPath(relative).parts:
        current = current / component
        require(not current.is_symlink(), f"symlink in {label}: {relative}")
    return path


def check_meta(path: Path, expected: Mapping[str, Any], label: str) -> None:
    exact_keys(expected, ("sha256", "bytes"), label)
    digest = expected["sha256"]
    size = expected["bytes"]
    require(isinstance(digest, str) and SHA_RE.fullmatch(digest), f"{label}.sha256 is invalid")
    require(isinstance(size, int) and not isinstance(size, bool) and size >= 0, f"{label}.bytes is invalid")
    actual = file_meta(path)
    require(actual == dict(sha256=digest, bytes=size), f"{label} hash/size mismatch")


def parse_utc(value: Any, label: str) -> _datetime.datetime:
    require(isinstance(value, str), f"{label} must be an ISO timestamp")
    try:
        parsed = _datetime.datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as error:
        fail(f"{label} is not an ISO timestamp: {error}")
    require(parsed.tzinfo is not None, f"{label} has no timezone")
    return parsed.astimezone(_datetime.timezone.utc)


def suffix_path(raw: Any, expected_relative: str, label: str) -> None:
    """Check a recorded absolute path by its portable bundle suffix.

    Capture records contain the original checkout's absolute path.  Requiring
    that path verbatim would make a copied bundle unverifiable; its lane and
    filename suffix are the stable identity tied to the receipt.
    """

    require(isinstance(raw, str) and raw, f"{label} must be a path")
    path = PurePosixPath(raw)
    require(path.is_absolute(), f"{label} must be absolute")
    wanted = PurePosixPath(expected_relative).parts
    require(tuple(path.parts[-len(wanted) :]) == wanted, f"{label} does not end in {expected_relative!r}")


def verify_seal(root: Path, *, required: bool) -> dict[str, Any]:
    seal = root / "SHA256SUMS"
    if not seal.exists():
        require(not required, "SHA256SUMS is absent; rerun with --unsealed to inspect an unsealed bundle")
        return {"status": "skipped", "files": 0}
    require(not seal.is_symlink() and seal.is_file(), "SHA256SUMS must be a regular file")
    entries: dict[str, str] = {}
    try:
        lines = seal.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as error:
        fail(f"cannot read SHA256SUMS: {error}")
    for number, line in enumerate(lines, 1):
        fields = line.split(maxsplit=1)
        require(len(fields) == 2, f"SHA256SUMS line {number} is malformed")
        digest, raw_path = fields
        require(SHA_RE.fullmatch(digest) is not None, f"SHA256SUMS line {number} has invalid digest")
        relative = safe_relative(raw_path, f"SHA256SUMS line {number}")
        require(relative != "SHA256SUMS", "SHA256SUMS must not self-seal")
        require(relative not in entries, f"duplicate SHA256SUMS path {relative}")
        entries[relative] = digest
    for relative, digest in entries.items():
        path = bundle_path(root, relative, "SHA256SUMS entry")
        require(sha256(path) == digest, f"sealed digest mismatch: {relative}")
    actual: set[str] = set()
    for path in root.rglob("*"):
        if path.is_symlink():
            fail(f"symlink in sealed bundle: {path.relative_to(root)}")
        if path.is_file() and path != seal:
            actual.add(path.relative_to(root).as_posix())
    require(actual == set(entries), f"SHA256SUMS coverage differs: missing={sorted(actual - set(entries))}, extra={sorted(set(entries) - actual)}")
    return {"status": "pass", "files": len(entries)}


def verify_protocol(root: Path) -> tuple[dict[str, Any], str]:
    path = root / "protocol.json"
    protocol = object_value(read_json(path), "protocol")
    exact_keys(protocol, {
        "acceptance", "capture_driver_sha256", "claim", "corpus_sha256", "counter_events", "cpu",
        "declared_utc", "members", "normal_repeat_review_threshold_percent", "order", "schema", "scope",
        "selector", "shape", "slides", "workers",
    }, "protocol")
    require(protocol["schema"] == PROTOCOL_SCHEMA, "unexpected protocol schema")
    require(protocol["selector"] == "pptx_streaming_create", "protocol selector changed")
    require(protocol["shape"] == "large" and protocol["slides"] == 8192 and protocol["members"] == 16421, "protocol corpus shape changed")
    require(protocol["cpu"] == 2 and protocol["workers"] == 1, "protocol CPU/worker binding changed")
    require(protocol["counter_events"] == ",".join(EXPECTED_EVENTS), "protocol counter event order changed")
    require(protocol["normal_repeat_review_threshold_percent"] == 5, "normal repeat threshold changed")
    parse_utc(protocol["declared_utc"], "protocol.declared_utc")
    rows = list_value(protocol["order"], "protocol.order")
    actual: list[tuple[str, str, int, int]] = []
    for index, row_value in enumerate(rows):
        row = object_value(row_value, f"protocol.order[{index}]")
        exact_keys(row, ("kind", "lane", "samples", "warmups"), f"protocol.order[{index}]")
        actual.append((row["lane"], row["kind"], row["samples"], row["warmups"]))
    require(tuple(actual) == EXPECTED_ORDER, "protocol lane order or sample counts changed")
    return protocol, sha256(path)


def verify_binding(root: Path, protocol: Mapping[str, Any], protocol_hash: str) -> tuple[dict[str, Any], str, dict[str, str]]:
    path = root / "binding.json"
    binding = object_value(read_json(path), "binding")
    exact_keys(binding, {
        "binary", "cached_binary_path", "capture_driver_sha256", "fixtures", "fresh_build", "prepared_finished_utc",
        "prepared_started_utc", "protocol_sha256", "reused_build_sha256", "revision", "schema", "source_files",
        "source_manifest_sha256", "tree",
    }, "binding")
    require(binding["schema"] == BINDING_SCHEMA, "unexpected binding schema")
    require(binding["fresh_build"] is False, "0475 unexpectedly claims a fresh build")
    require(isinstance(binding["revision"], str) and re.fullmatch(r"[0-9a-f]{40}", binding["revision"]), "invalid source revision")
    require(binding["protocol_sha256"] == protocol_hash, "binding does not bind protocol.json")
    capture_driver = root / "capture.py"
    driver_hash = sha256(capture_driver)
    require(binding["capture_driver_sha256"] == driver_hash == protocol["capture_driver_sha256"], "capture driver hash is stale")
    parse_utc(binding["prepared_started_utc"], "binding.prepared_started_utc")
    parse_utc(binding["prepared_finished_utc"], "binding.prepared_finished_utc")
    require(parse_utc(binding["prepared_finished_utc"], "binding.prepared_finished_utc") >= parse_utc(binding["prepared_started_utc"], "binding.prepared_started_utc"), "binding preparation times reverse")
    tree = binding["tree"]
    cached = binding["cached_binary_path"]
    require(isinstance(tree, str) and Path(tree).is_absolute(), "binding.tree must be an absolute recorded path")
    require(isinstance(cached, str) and Path(cached).is_absolute(), "binding.cached_binary_path must be absolute")
    binary = object_value(binding["binary"], "binding.binary")
    exact_keys(binary, ("bytes", "path", "sha256"), "binding.binary")
    require(isinstance(binary["path"], str) and Path(binary["path"]).is_absolute(), "binding.binary.path must be absolute")
    require(isinstance(binary["bytes"], int) and binary["bytes"] > 0, "binding binary size is invalid")
    require(isinstance(binary["sha256"], str) and SHA_RE.fullmatch(binary["sha256"]), "binding binary hash is invalid")
    require(binding["source_files"] == 7034, "source manifest file count is stale")
    require(binding["reused_build_sha256"] == sha256(root / "reuse/build.json"), "reused build hash is stale")
    fixtures = object_value(binding["fixtures"], "binding.fixtures")
    require(set(fixtures) == {
        "test-data/poi/test-data/spreadsheet/54016.xls",
        "test-data/rtf/watermark.rtf",
    }, "fixture identity set changed")
    for fixture, digest in fixtures.items():
        safe_relative(fixture, "binding fixture")
        require(isinstance(digest, str) and SHA_RE.fullmatch(digest), f"fixture hash is invalid: {fixture}")
    source_path = bundle_path(root, "reuse/sources/source.json", "source manifest")
    require(sha256(source_path) == binding["source_manifest_sha256"], "source manifest hash is stale")
    source = object_value(read_json(source_path), "source manifest")
    require(len(source) == binding["source_files"], "source manifest file count does not match binding")
    for relative, digest in source.items():
        safe_relative(relative, "source manifest path")
        require(isinstance(digest, str) and SHA_RE.fullmatch(digest), f"source hash is invalid: {relative}")
    return binding, sha256(path), {str(k): str(v) for k, v in source.items()}


def parse_sha256sums(path: Path) -> dict[str, str]:
    entries: dict[str, str] = {}
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as error:
        fail(f"cannot read {path}: {error}")
    for number, line in enumerate(lines, 1):
        fields = line.split(maxsplit=1)
        require(len(fields) == 2 and SHA_RE.fullmatch(fields[0]) is not None, f"malformed seal {path}:{number}")
        relative = safe_relative(fields[1], f"{path}:{number}")
        require(relative not in entries, f"duplicate seal path {relative}")
        entries[relative] = fields[0]
    return entries


def verify_reuse(root: Path, binding: Mapping[str, Any], protocol: Mapping[str, Any]) -> dict[str, Any]:
    reuse = root / "reuse"
    validation = object_value(read_json(root / "reuse-validation.json"), "reuse validation")
    exact_keys(validation, ("argv", "exit_code", "selected_files", "stderr", "stdout", "verified_utc"), "reuse validation")
    parse_utc(validation["verified_utc"], "reuse validation verified_utc")
    require(validation["argv"] == ["python3", "-B", "docs/performance/results/change-0474/verify.py", "--sealed"], "reuse validation command is stale")
    require(validation["exit_code"] == 0 and validation["stderr"] == "", "reuse validation did not pass cleanly")
    selected = object_value(validation["selected_files"], "reuse selected_files")
    require(tuple(selected) == EXPECTED_REUSE_FILES, "reuse selected artifact set is stale")
    for relative, digest in selected.items():
        require(isinstance(digest, str) and SHA_RE.fullmatch(digest), f"invalid reuse selected hash: {relative}")
        path = bundle_path(reuse, relative, "reuse selected path")
        require(sha256(path) == digest, f"reuse selected hash mismatch: {relative}")
    old_seal = parse_sha256sums(reuse / "SHA256SUMS")
    for relative in EXPECTED_REUSE_FILES:
        if relative != "SHA256SUMS":
            require(old_seal.get(relative) == selected[relative], f"reuse/SHA256SUMS does not bind {relative}")
    require(isinstance(validation["stdout"], str), "reuse verification stdout is not text")
    try:
        prior_result = object_value(json.loads(validation["stdout"]), "reuse verification stdout")
    except (TypeError, json.JSONDecodeError) as error:
        fail(f"reuse verification stdout is not JSON: {error}")
    require(prior_result.get("status") == "pass", "prior sealed verification result was not pass")
    require(prior_result.get("schema") == "litchi-0474-verification-v1", "prior verification schema is stale")
    require(prior_result.get("lanes") == 12 and prior_result.get("samples") == 360, "prior verification totals are stale")
    sealed = prior_result.get("sealed")
    require(isinstance(sealed, dict) and sealed.get("status") == "pass", "prior verification did not seal the selected build")

    old_protocol = object_value(read_json(reuse / "protocol.json"), "reused protocol")
    require(old_protocol.get("schema") == "litchi-0474-protocol-v1", "unexpected reused protocol")
    require(old_protocol.get("selector") == protocol["selector"], "reused protocol selector disagrees")
    old_shapes = object_value(old_protocol.get("shapes"), "reused protocol shapes")
    require(old_shapes.get("large") == protocol["slides"], "reused protocol large shape disagrees")
    require(old_protocol.get("samples") == 30 and old_protocol.get("warmups") == 3 and old_protocol.get("workers") == protocol["workers"], "reused protocol sample/worker configuration disagrees")
    old_build = object_value(read_json(reuse / "build.json"), "reused build")
    require(old_build.get("schema") == "litchi-0474-build-v1", "unexpected reused build schema")
    require(old_build.get("revision") == binding["revision"], "reused build revision disagrees")
    source_manifest = object_value(old_build.get("source_manifest"), "reused build source_manifest")
    require(source_manifest.get("files") == binding["source_files"], "reused build source count disagrees")
    require(source_manifest.get("sha256") == binding["source_manifest_sha256"], "reused build source hash disagrees")
    normal_binary = object_value(old_build.get("binaries", {}).get("normal"), "reused normal binary")
    require(normal_binary.get("sha256") == binding["binary"]["sha256"] and normal_binary.get("bytes") == binding["binary"]["bytes"], "reused binary identity disagrees")
    old_binding = object_value(read_json(reuse / "source-binding.json"), "reused source binding")
    require(old_binding.get("revision") == binding["revision"], "reused source binding revision disagrees")
    old_source = object_value(old_binding.get("source_manifest"), "reused source binding source_manifest")
    require(old_source.get("files") == binding["source_files"] and old_source.get("sha256") == binding["source_manifest_sha256"], "reused source binding manifest disagrees")
    require(old_binding.get("fixtures") == binding["fixtures"], "reused fixture identity disagrees")
    return {"files": len(selected), "prior_schema": prior_result["schema"]}


def _compression_records(root: Path, path: Path, expected: frozenset[str], label: str) -> dict[str, dict[str, Any]]:
    data = object_value(read_json(path), label)
    exact_keys(data, ("artifacts", "schema"), label)
    require(data["schema"] == COMPRESSION_SCHEMA, f"unexpected {label} schema")
    records = list_value(data["artifacts"], f"{label}.artifacts")
    result: dict[str, dict[str, Any]] = {}
    compressed_paths: set[str] = set()
    for index, item_value in enumerate(records):
        item = object_value(item_value, f"{label}.artifacts[{index}]")
        exact_keys(item, ("bytes", "compressed_bytes", "compressed_path", "compressed_sha256", "path", "sha256"), f"{label}.artifacts[{index}]")
        relative = safe_relative(item["path"], f"{label} path {index}")
        compressed = safe_relative(item["compressed_path"], f"{label} compressed_path {index}")
        require(relative in expected, f"unexpected compressed original {relative}")
        require(relative not in result and compressed not in compressed_paths, "duplicate compression artifact")
        require(compressed.endswith(".gz"), f"unsupported compression format: {compressed}")
        require(isinstance(item["bytes"], int) and item["bytes"] >= 0, f"invalid original byte count {relative}")
        require(isinstance(item["compressed_bytes"], int) and item["compressed_bytes"] >= 0, f"invalid compressed byte count {compressed}")
        require(isinstance(item["sha256"], str) and SHA_RE.fullmatch(item["sha256"]), f"invalid original hash {relative}")
        require(isinstance(item["compressed_sha256"], str) and SHA_RE.fullmatch(item["compressed_sha256"]), f"invalid compressed hash {compressed}")
        original = bundle_path(root, relative, f"{label} original")
        compressed_path = bundle_path(root, compressed, f"{label} compressed artifact")
        require(not original.exists(), f"compressed original was not deleted: {relative}")
        compressed_meta = file_meta(compressed_path)
        require(compressed_meta == {"sha256": item["compressed_sha256"], "bytes": item["compressed_bytes"]}, f"compressed artifact hash/size mismatch: {compressed}")
        try:
            with gzip.open(compressed_path, "rb") as stream:
                digest = hashlib.sha256()
                size = 0
                for block in iter(lambda: stream.read(1024 * 1024), b""):
                    digest.update(block)
                    size += len(block)
        except (OSError, EOFError) as error:
            fail(f"cannot decompress {compressed}: {error}")
        require(size == item["bytes"] and digest.hexdigest() == item["sha256"], f"decompressed hash/size mismatch: {compressed}")
        result[relative] = item
        compressed_paths.add(compressed)
    require(set(result) == expected, f"{label} coverage differs: missing={sorted(expected - set(result))}, extra={sorted(set(result) - expected)}")
    return result


def compression_record_map(root: Path) -> dict[str, dict[str, Any]]:
    result = _compression_records(root, root / "compression.json", EXPECTED_COMPRESSION_PATHS, "compression")
    supplement = root / "supplement-compression.json"
    require(supplement.is_file(), "supplement-compression.json is missing")
    supplemental = _compression_records(root, supplement, EXPECTED_SUPPLEMENT_PATHS, "supplement compression")
    for relative, item in supplemental.items():
        require(relative not in result, f"duplicate compression original across manifests: {relative}")
        result[relative] = item
    return result


def verify_chunking(root: Path) -> dict[str, dict[str, Any]]:
    """Verify ordered split parts for the retained raw Heaptrack traces.

    The logical trace is intentionally absent from the final package.  Hashing
    each part and feeding the parts into one digest preserves the original
    receipt binding without allocating the complete trace in memory.
    """

    path = root / "chunking.json"
    data = object_value(read_json(path), "chunking")
    exact_keys(data, ("artifacts", "chunk_bytes", "driver_sha256", "recorded_utc", "schema"), "chunking")
    require(data["schema"] == CHUNKING_SCHEMA, "unexpected chunking schema")
    require(data["chunk_bytes"] == 67108864, "chunk size changed")
    require(isinstance(data["driver_sha256"], str) and SHA_RE.fullmatch(data["driver_sha256"]), "chunking driver hash is invalid")
    parse_utc(data["recorded_utc"], "chunking.recorded_utc")
    records = list_value(data["artifacts"], "chunking.artifacts")
    result: dict[str, dict[str, Any]] = {}
    listed_parts: set[str] = set()
    for index, item_value in enumerate(records):
        item = object_value(item_value, f"chunking.artifacts[{index}]")
        exact_keys(item, ("bytes", "parts", "path", "sha256"), f"chunking.artifacts[{index}]")
        original = safe_relative(item["path"], f"chunking original path {index}")
        require(original in {"heap-H1/heaptrack.zst", "heap-H2/heaptrack.zst"}, f"unexpected chunked original {original}")
        require(original not in result, f"duplicate chunked original {original}")
        require(isinstance(item["bytes"], int) and item["bytes"] > 0, f"invalid chunked byte count {original}")
        require(isinstance(item["sha256"], str) and SHA_RE.fullmatch(item["sha256"]), f"invalid chunked hash {original}")
        require(not bundle_path(root, original, "chunked original").exists(), f"chunked original still exists: {original}")
        parts = list_value(item["parts"], f"chunking parts {original}")
        require(parts, f"chunking parts are empty: {original}")
        digest = hashlib.sha256()
        total = 0
        expected_names: list[str] = []
        for part_index, part_value in enumerate(parts):
            part = object_value(part_value, f"chunking {original} part {part_index}")
            exact_keys(part, ("bytes", "path", "sha256"), f"chunking {original} part {part_index}")
            relative = safe_relative(part["path"], f"chunking part {original} {part_index}")
            require(relative not in listed_parts, f"duplicate chunk part {relative}")
            expected_name = f"{original}.part-{part_index:03d}"
            require(relative == expected_name, f"chunk part order/name changed: expected {expected_name}, got {relative}")
            expected_names.append(relative)
            listed_parts.add(relative)
            require(isinstance(part["bytes"], int) and part["bytes"] >= 0, f"invalid chunk byte count {relative}")
            require(isinstance(part["sha256"], str) and SHA_RE.fullmatch(part["sha256"]), f"invalid chunk hash {relative}")
            part_path = bundle_path(root, relative, "chunk part")
            actual = file_meta(part_path)
            require(actual == {"sha256": part["sha256"], "bytes": part["bytes"]}, f"chunk hash/size mismatch: {relative}")
            try:
                with part_path.open("rb") as stream:
                    for block in iter(lambda: stream.read(1024 * 1024), b""):
                        digest.update(block)
                        total += len(block)
            except OSError as error:
                fail(f"cannot read chunk {relative}: {error}")
        require(total == item["bytes"] and digest.hexdigest() == item["sha256"], f"chunked reconstruction mismatch: {original}")
        result[original] = item
    require(set(result) == {"heap-H1/heaptrack.zst", "heap-H2/heaptrack.zst"}, "chunking coverage does not include both raw Heaptrack traces")
    actual_parts = {
        path.relative_to(root).as_posix()
        for lane in ("heap-H1", "heap-H2")
        for path in (root / lane).glob("heaptrack.zst.part-*")
        if path.is_file()
    }
    require(actual_parts == listed_parts, f"chunk part file coverage differs: missing={sorted(listed_parts - actual_parts)}, extra={sorted(actual_parts - listed_parts)}")
    return result


def expected_capture_argv(root: Path, binding: Mapping[str, Any], item: tuple[str, str, int, int]) -> list[str]:
    lane, kind, samples, warmups = item
    output = f"{lane}/"
    argv = ["taskset", "-c", "2", "/usr/bin/time", "-v", "-o", f"{output}resource.log"]
    if kind == "counters":
        argv += ["perf", "stat", "-x", ";", "-o", f"{output}counters.csv", "-e", ",".join(EXPECTED_EVENTS), "--"]
    elif kind == "cpu":
        argv += ["perf", "record", "-F", "499", "-e", "cycles:u", "--call-graph", "fp", "-o", f"{output}perf.data", "--"]
    elif kind == "heap":
        argv += ["heaptrack", "--record-only", "-o", f"{output}heaptrack"]
    argv += [
        str(binding["binary"]["path"]), "--case", "pptx_streaming_create", "--semantic-shape", "large", "--workers", "1",
        "--warmup", str(warmups), "--samples", str(samples), "--json", f"{output}report.json", "--corpus-manifest", f"{output}corpus-catalog.json",
    ]
    return argv


def normalize_capture_argv(argv: Any, item: tuple[str, str, int, int], binding: Mapping[str, Any]) -> list[str]:
    require(isinstance(argv, list) and all(isinstance(x, str) for x in argv), "capture argv must be a string array")
    lane, kind, samples, warmups = item
    expected = expected_capture_argv(Path("."), binding, item)
    require(len(argv) == len(expected), f"{lane} capture argv length changed")
    result: list[str] = []
    dynamic_suffixes = {
        f"{lane}/resource.log",
        f"{lane}/counters.csv",
        f"{lane}/perf.data",
        f"{lane}/report.json",
        f"{lane}/corpus-catalog.json",
        f"{lane}/heaptrack",
    }
    for actual, wanted in zip(argv, expected):
        if wanted in dynamic_suffixes:
            suffix_path(actual, wanted, f"{lane} capture argv path")
            result.append(wanted)
        elif wanted == str(binding["binary"]["path"]):
            require(actual == wanted, f"{lane} binary argv changed")
            result.append(actual)
        else:
            require(actual == wanted, f"{lane} capture argv changed: {actual!r} != {wanted!r}")
            result.append(actual)
    return result


def verify_resource(path: Path, lane: str) -> None:
    text = path.read_text(encoding="utf-8", errors="strict")
    require("Command being timed:" in text, f"{lane} resource log lacks command")
    require("Exit status: 0" in text, f"{lane} resource log does not record status zero")
    match = re.search(r"Maximum resident set size \(kbytes\): (\d+)", text)
    require(match is not None and int(match.group(1)) >= 0, f"{lane} resource log lacks RSS")


def verify_command_receipt(root: Path, path: Path, binding: Mapping[str, Any], protocol_hash: str, *, lane: str, prefix: str, compression: Mapping[str, Mapping[str, Any]]) -> None:
    data = object_value(read_json(path), f"{path.name}")
    required = ("argv", "artifacts", "cwd", "driver_sha256", "environment", "exit_code", "finished_utc", "started_utc")
    for key in required:
        require(key in data, f"{path.name} lacks {key}")
    require(data["cwd"] == binding["tree"], f"{path.name} cwd is stale")
    require(data["driver_sha256"] == binding["capture_driver_sha256"], f"{path.name} driver hash is stale")
    require(data["environment"] == {
        "RUSTUP_TOOLCHAIN": "1.98.1",
        "CARGO_PROFILE_RELEASE_DEBUG": "1",
        "RUSTFLAGS": "-C force-frame-pointers=yes -C force-unwind-tables=yes",
        "DEBUGINFOD_URLS": "",
        "LC_ALL": "C",
    }, f"{path.name} environment is stale")
    require(data["exit_code"] == 0, f"{path.name} exit status is nonzero")
    started = parse_utc(data["started_utc"], f"{path.name}.started_utc")
    finished = parse_utc(data["finished_utc"], f"{path.name}.finished_utc")
    require(finished >= started, f"{path.name} timestamps reverse")
    artifacts = object_value(data["artifacts"], f"{path.name}.artifacts")
    exact_keys(artifacts, (f"{prefix}.stdout", f"{prefix}.stderr"), f"{path.name}.artifacts")
    for name, meta in artifacts.items():
        current = path.parent / name
        if current.exists():
            check_meta(current, meta, f"{path.name}.{name}")
        else:
            original = f"{lane}/{name}"
            record = compression.get(original)
            require(record is not None, f"{path.name}.{name} is deleted without compression evidence")
            require(record["sha256"] == meta["sha256"] and record["bytes"] == meta["bytes"], f"{path.name}.{name} compression does not match receipt")
    started_path = path.with_name(f"{prefix}.started.json")
    started_data = object_value(read_json(started_path), f"{started_path.name}")
    for key in ("argv", "cwd", "environment", "started_utc", "driver_sha256"):
        require(started_data.get(key) == data.get(key), f"{started_path.name}.{key} disagrees with receipt")
    # The receipt itself may be checked through its capture receipt; command
    # output files are the only artifacts emitted by command().
    del root, protocol_hash, lane, compression


def verify_capture_receipt(root: Path, item: tuple[str, str, int, int], binding: Mapping[str, Any], binding_hash: str, protocol_hash: str, compression: Mapping[str, Mapping[str, Any]], chunks: Mapping[str, Mapping[str, Any]]) -> tuple[_datetime.datetime, _datetime.datetime]:
    lane, kind, samples, warmups = item
    directory = root / lane
    capture = object_value(read_json(directory / "capture.json"), f"{lane}/capture.json")
    started_data = object_value(read_json(directory / "capture.started.json"), f"{lane}/capture.started.json")
    require(capture["lane"] == lane and started_data["lane"] == lane, f"{lane} capture lane identity changed")
    require(capture["binding_sha256"] == binding_hash and capture["protocol_sha256"] == protocol_hash, f"{lane} capture binding hashes are stale")
    require(capture["binary_sha256"] == binding["binary"]["sha256"], f"{lane} capture binary hash is stale")
    require(capture["clean_before"] is True and capture["exit_code"] == 0, f"{lane} capture status is not clean")
    normalize_capture_argv(capture["argv"], item, binding)
    require(capture["cwd"] == binding["tree"], f"{lane} capture cwd is stale")
    expected_environment = {
        "RUSTUP_TOOLCHAIN": "1.98.1", "CARGO_PROFILE_RELEASE_DEBUG": "1",
        "RUSTFLAGS": "-C force-frame-pointers=yes -C force-unwind-tables=yes", "DEBUGINFOD_URLS": "", "LC_ALL": "C",
    }
    require(capture["environment"] == expected_environment, f"{lane} capture environment is stale")
    require(capture["driver_sha256"] == binding["capture_driver_sha256"], f"{lane} capture driver hash is stale")
    started = parse_utc(capture["started_utc"], f"{lane} capture.started_utc")
    finished = parse_utc(capture["finished_utc"], f"{lane} capture.finished_utc")
    require(finished >= started, f"{lane} capture timestamps reverse")
    for key in ("argv", "cwd", "environment", "started_utc", "driver_sha256", "lane", "binding_sha256", "protocol_sha256", "binary_sha256", "clean_before"):
        require(started_data.get(key) == capture.get(key), f"{lane} capture.started.json {key} disagrees")
    command_artifacts = object_value(capture["artifacts"], f"{lane}/capture.json.artifacts")
    exact_keys(command_artifacts, ("capture.stdout", "capture.stderr"), f"{lane}/capture.json.artifacts")
    for name, meta in command_artifacts.items():
        check_meta(directory / name, meta, f"{lane}/capture.json.{name}")

    receipt = object_value(read_json(directory / "receipt.json"), f"{lane}/receipt.json")
    exact_keys(receipt, ("artifacts", "binary_unchanged", "binding_sha256", "clean_after", "exit_code", "kind", "lane", "protocol_sha256", "schema", "source_unchanged"), f"{lane}/receipt.json")
    require(receipt["schema"] == CAPTURE_SCHEMA and receipt["lane"] == lane and receipt["kind"] == kind, f"{lane} capture receipt identity changed")
    require(receipt["exit_code"] == 0 and receipt["clean_after"] is True and receipt["binary_unchanged"] is True and receipt["source_unchanged"] is True, f"{lane} capture receipt status is not clean")
    require(receipt["binding_sha256"] == binding_hash and receipt["protocol_sha256"] == protocol_hash, f"{lane} receipt binding hashes are stale")
    artifacts = object_value(receipt["artifacts"], f"{lane}/receipt.json.artifacts")
    mandatory = {"capture.json", "capture.started.json", "capture.stderr", "capture.stdout", "corpus-catalog.json", "report.json", "resource.log"}
    if kind == "counters":
        mandatory.add("counters.csv")
    if kind == "cpu":
        mandatory.add("perf.data")
    if kind == "heap":
        mandatory.add("heaptrack.zst")
    require(mandatory <= set(artifacts), f"{lane} receipt omits mandatory artifacts")
    for name, meta in artifacts.items():
        safe_relative(name, f"{lane} receipt artifact")
        # capture.py records basenames only; do not allow a receipt to bind a
        # file outside its lane.
        require("/" not in name, f"{lane} receipt artifact is not a basename: {name}")
        current = directory / name
        if current.exists():
            check_meta(current, meta, f"{lane}/{name}")
        else:
            original = f"{lane}/{name}"
            record = compression.get(original)
            if record is not None:
                require(record["sha256"] == meta["sha256"] and record["bytes"] == meta["bytes"], f"{lane}/{name} compression does not match receipt")
            else:
                chunk = chunks.get(original)
                require(chunk is not None, f"{lane}/{name} is deleted without compression or chunk evidence")
                require(chunk["sha256"] == meta["sha256"] and chunk["bytes"] == meta["bytes"], f"{lane}/{name} chunks do not match receipt")
    verify_resource(directory / "resource.log", lane)
    return started, finished


def verify_export_receipts(root: Path, binding: Mapping[str, Any], protocol_hash: str, compression: Mapping[str, Mapping[str, Any]]) -> None:
    expected: dict[str, tuple[str, ...]] = {
        "cpu-P1": ("perf-script", "top-symbols"),
        "cpu-P2": ("perf-script", "top-symbols"),
        "heap-H1": ("print", "runner-allocations", "decoded"),
        "heap-H2": ("print", "runner-allocations", "decoded"),
    }
    for lane, prefixes in expected.items():
        directory = root / lane
        for prefix in prefixes:
            path = directory / f"{prefix}.json"
            verify_command_receipt(root, path, binding, protocol_hash, lane=lane, prefix=prefix, compression=compression)
        # The command receipt for an export binds stdout/stderr.  The
        # output-file allocation-stacks.txt is bound by compression.json.
        if lane.startswith("cpu"):
            require((directory / "top-symbols.stdout").is_file() and (directory / "top-symbols.stderr").is_file(), f"{lane} top-symbol export output missing")
            require((directory / "perf-script.stderr").is_file(), f"{lane} perf-script stderr missing")
        else:
            require((directory / "print.stderr").is_file() and (directory / "runner-allocations.stderr").is_file() and (directory / "decoded.stderr").is_file(), f"{lane} export stderr missing")
    # Check all command receipts have a matching started receipt and no export
    # was silently omitted.  Unknown auxiliary files are retained for review.


def verify_supplement(root: Path, binding: Mapping[str, Any], compression: Mapping[str, Mapping[str, Any]]) -> dict[str, Any]:
    protocol_path = root / "supplement-protocol.json"
    driver_path = root / "supplement.py"
    declaration = object_value(read_json(protocol_path), "supplement protocol")
    exact_keys(declaration, ("declared_utc", "driver_sha256", "filter_token", "reason", "schema", "scope"), "supplement protocol")
    require(declaration["schema"] == SUPPLEMENT_SCHEMA, "unexpected supplement protocol schema")
    require(declaration["filter_token"] == "21pptx_streaming_create3run", "supplement filter token changed")
    require(declaration["driver_sha256"] == sha256(driver_path), "supplement driver hash is stale")
    declared = parse_utc(declaration["declared_utc"], "supplement protocol declared_utc")
    require("same raw traces" in declaration["reason"] and "no recapture" in declaration["scope"], "supplement custody statement changed")
    supplement_hash = sha256(protocol_path)
    driver_hash = sha256(driver_path)
    expected_argv: list[str] = []
    for lane in ("heap-H1", "heap-H2"):
        directory = root / lane
        receipt_path = directory / "runner-mangled.json"
        verify_command_receipt(root, receipt_path, binding, sha256(root / "protocol.json"), lane=lane, prefix="runner-mangled", compression=compression)
        receipt = object_value(read_json(receipt_path), f"{lane}/runner-mangled.json")
        require(receipt.get("supplemental_protocol_sha256") == supplement_hash and receipt.get("supplemental_driver_sha256") == driver_hash, f"{lane} runner-mangled receipt supplement binding is stale")
        started_receipt = object_value(read_json(root / lane / "runner-mangled.started.json"), f"{lane}/runner-mangled.started.json")
        require(started_receipt.get("supplemental_protocol_sha256") == supplement_hash and started_receipt.get("supplemental_driver_sha256") == driver_hash, f"{lane} runner-mangled started receipt supplement binding is stale")
        argv = receipt.get("argv")
        require(isinstance(argv, list) and all(isinstance(value, str) for value in argv), f"{lane} runner-mangled argv is malformed")
        expected = [
            "heaptrack_print", "-f", f"{lane}/heaptrack.zst", "-t", "0", "-m", "0", "-p", "0", "-T", "0", "-l", "0",
            "--filter-bt-function", declaration["filter_token"], "-n", "100", "-F", f"{lane}/runner-mangled-stacks.txt",
            "--flamegraph-cost-type", "allocations",
        ]
        require(len(argv) == len(expected), f"{lane} runner-mangled argv length changed")
        for actual, wanted in zip(argv, expected):
            if wanted.startswith(f"{lane}/"):
                suffix_path(actual, wanted, f"{lane} runner-mangled argv path")
            else:
                require(actual == wanted, f"{lane} runner-mangled argv changed")
        started = parse_utc(receipt["started_utc"], f"{lane}/runner-mangled.json.started_utc")
        require(started >= declared, f"{lane} supplement command predates supplement declaration")
        for original in (f"{lane}/runner-mangled.stdout", f"{lane}/runner-mangled-stacks.txt"):
            require(original in compression, f"{lane} supplement output lacks compression binding")
        expected_argv.extend(argv)
    return {"schema": declaration["schema"], "filter_token": declaration["filter_token"], "artifacts": len(EXPECTED_SUPPLEMENT_PATHS)}


def parse_counter_text(text: str) -> dict[str, dict[str, Any]]:
    rows: dict[str, dict[str, Any]] = {}
    lines = text.splitlines()
    raw_rows: list[str] = []
    for line_number, line in enumerate(lines, 1):
        if not line or line.startswith("#"):
            continue
        fields = line.split(";")
        require(len(fields) == 7 and fields[1] == fields[5] == fields[6] == "", f"malformed perf counter row {line_number}")
        value, _, event, runtime, percent, _, _ = fields
        require(event not in rows, f"duplicate perf counter event {event}")
        require(event in EXPECTED_EVENTS, f"unknown perf counter event {event}; unsupported rows must remain explicit")
        require(runtime.isdecimal(), f"invalid runtime for counter {event}")
        try:
            running_percent = float(percent)
        except ValueError:
            fail(f"invalid running percentage for counter {event}")
        require(math.isfinite(running_percent) and 0 <= running_percent <= 100, f"invalid running percentage for counter {event}")
        if value in ("<not supported>", "<not counted>"):
            status = value[1:-1].replace(" ", "_")
            count = None
        else:
            require(value.isdecimal(), f"invalid count for counter {event}")
            status = "reported"
            count = int(value)
        rows[event] = {"status": status, "count": count, "event_runtime_ns": int(runtime), "running_percent": running_percent}
        raw_rows.append(line)
    require(set(rows) == set(EXPECTED_EVENTS), f"counter event coverage differs: missing={sorted(set(EXPECTED_EVENTS)-set(rows))}, extra={sorted(set(rows)-set(EXPECTED_EVENTS))}")
    # Return every retained row, including explicit unsupported/not-counted
    # statuses.  No unsupported event is coerced to zero.
    require(raw_rows, "counters.csv has no event rows")
    return rows


def verify_counter_file(root: Path, protocol: Mapping[str, Any]) -> dict[str, dict[str, Any]]:
    path = root / "counters/counters.csv"
    try:
        text = path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as error:
        fail(f"cannot read counters.csv: {error}")
    return parse_counter_text(text)


def check_metric_vectors(value: Any, sample_count: int, label: str) -> None:
    if isinstance(value, dict):
        if "values" in value:
            values = value["values"]
            require(isinstance(values, list) and len(values) == sample_count, f"{label}.values length is not {sample_count}")
            for index, item in enumerate(values):
                require(isinstance(item, (int, float)) and not isinstance(item, bool) and math.isfinite(float(item)) and item >= 0, f"{label}.values[{index}] is invalid")
        for key, child in value.items():
            check_metric_vectors(child, sample_count, f"{label}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            check_metric_vectors(child, sample_count, f"{label}[{index}]")


def collapse_vectors(value: Any, sample_count: int, label: str) -> Any:
    if isinstance(value, dict):
        if "values" in value:
            values = value["values"]
            require(isinstance(values, list) and len(values) == sample_count, f"{label}.values length is not {sample_count}")
            require(values and all(item == values[0] for item in values), f"{label}.values are not deterministic")
            result = {key: collapse_vectors(child, sample_count, f"{label}.{key}") for key, child in value.items() if key != "values"}
            result["value"] = values[0]
            return result
        return {key: collapse_vectors(child, sample_count, f"{label}.{key}") for key, child in value.items()}
    if isinstance(value, list):
        return [collapse_vectors(child, sample_count, f"{label}[{index}]") for index, child in enumerate(value)]
    return value


def verify_elapsed(elapsed: Mapping[str, Any], sample_count: int, label: str) -> None:
    exact_keys(elapsed, ("confidence_interval_95", "max", "mean", "min", "p50", "p95", "p99", "sample_order", "samples", "standard_deviation", "unit"), label)
    require(elapsed["unit"] == "ns", f"{label}.unit changed")
    samples = elapsed["samples"]
    order = elapsed["sample_order"]
    require(isinstance(samples, list) and len(samples) == sample_count, f"{label}.samples count changed")
    require(all(isinstance(value, int) and not isinstance(value, bool) and value > 0 for value in samples), f"{label}.samples contains invalid values")
    require(samples == sorted(samples), f"{label}.samples is not sorted")
    require(isinstance(order, list) and sorted(order) == list(range(sample_count)), f"{label}.sample_order is not a permutation")
    require(elapsed["min"] == min(samples) and elapsed["max"] == max(samples), f"{label} min/max arithmetic mismatch")
    median = (samples[(sample_count - 1) // 2] + samples[sample_count // 2]) / 2
    p95 = samples[max(0, math.ceil(sample_count * 0.95) - 1)]
    p99 = samples[max(0, math.ceil(sample_count * 0.99) - 1)]
    require(elapsed["p50"] == median and elapsed["p95"] == p95 and elapsed["p99"] == p99, f"{label} percentile arithmetic mismatch")
    mean = sum(samples) / sample_count
    deviation = math.sqrt(sum((value - mean) ** 2 for value in samples) / (sample_count - 1)) if sample_count > 1 else 0.0
    require(math.isclose(float(elapsed["mean"]), mean, rel_tol=1e-12, abs_tol=1e-6), f"{label}.mean arithmetic mismatch")
    require(math.isclose(float(elapsed["standard_deviation"]), deviation, rel_tol=1e-12, abs_tol=1e-6), f"{label}.standard_deviation arithmetic mismatch")
    ci = object_value(elapsed["confidence_interval_95"], f"{label}.confidence_interval_95")
    exact_keys(ci, ("lower", "method", "upper"), f"{label}.confidence_interval_95")
    require(ci["method"] == "two-sided Student's t interval for the mean", f"{label} CI method changed")
    require(math.isfinite(float(ci["lower"])) and math.isfinite(float(ci["upper"])), f"{label} CI is non-finite")
    if sample_count == 1:
        require(ci["lower"] == ci["upper"] == elapsed["mean"], f"{label} one-sample CI arithmetic mismatch")
    else:
        require(ci["lower"] <= elapsed["mean"] <= ci["upper"], f"{label} CI does not contain the mean")
        # The harness uses the pinned 95% t critical value for df=29.
        if sample_count == 30:
            standard_error = deviation / math.sqrt(sample_count)
            require(math.isclose((float(ci["upper"]) - mean) / standard_error, 2.045, rel_tol=2e-12, abs_tol=1e-9), f"{label} CI t arithmetic mismatch")


def expected_producer_projection(row: Mapping[str, Any], sample_count: int) -> dict[str, Any]:
    operation = object_value(row["operation_metrics"], "operation_metrics")
    projection = {
        "case": row["case"],
        "corpus": row["corpus"],
        "source": row["source"],
        "sink": row["sink"],
        "output_sha256": row["output_sha256"],
        "operation": {
            key: operation[key]
            for key in ("alignment", "latency_claim", "source", "publication", "materialization", "cfb_phases")
        },
        "operation_sink": collapse_vectors(operation["sink"], sample_count, "operation_metrics.sink"),
    }
    return projection


def verify_catalog(root: Path, lane: str, report: Mapping[str, Any], reference_catalog: Mapping[str, Any], receipt_artifacts: Mapping[str, Any]) -> None:
    path = root / lane / "corpus-catalog.json"
    catalog = object_value(read_json(path), f"{lane}/corpus-catalog.json")
    require(catalog == reference_catalog, f"{lane} corpus catalog identity differs")
    require(report["corpus_catalog"] == {
        "manifest_version": catalog["manifest_version"],
        "catalog_id": catalog["catalog_id"],
        "catalog_sha256": catalog["catalog_sha256"],
        "content_set_sha256": catalog["content_set_sha256"],
    }, f"{lane} report catalog summary differs")
    check_meta(path, receipt_artifacts["corpus-catalog.json"], f"{lane}/corpus-catalog.json")


def verify_report(root: Path, lane: str, kind: str, sample_count: int, warmups: int, report: Mapping[str, Any], reference: Mapping[str, Any], reference_row: Mapping[str, Any], reference_catalog: Mapping[str, Any], binding: Mapping[str, Any], receipt: Mapping[str, Any]) -> None:
    exact_keys(report, ("binary_identity", "configuration", "corpus_catalog", "environment", "parallel_metrics", "results", "schema_version", "tool"), f"{lane}/report.json")
    require(report["schema_version"] == reference["schema_version"] == 1, f"{lane} report schema changed")
    require(report["tool"] == EXPECTED_TOOL == reference["tool"], f"{lane} producer tool identity changed")
    binary = object_value(report["binary_identity"], f"{lane}.binary_identity")
    exact_keys(binary, ("binary_bytes", "binary_sha256", "executable", "mode_bits", "path", "profile"), f"{lane}.binary_identity")
    require(binary["path"] == binding["binary"]["path"], f"{lane} binary path does not bind current profile")
    require(binary["binary_sha256"] == binding["binary"]["sha256"] and binary["binary_bytes"] == binding["binary"]["bytes"], f"{lane} binary hash/size identity changed")
    require(binary["executable"] is True and binary["profile"] == "release" and isinstance(binary["mode_bits"], int), f"{lane} binary metadata invalid")
    reference_binary = object_value(reference["binary_identity"], "reused R1 binary_identity")
    for key in ("binary_bytes", "binary_sha256", "executable", "mode_bits", "profile"):
        require(binary[key] == reference_binary[key], f"{lane} binary identity field {key} differs from reused R1")
    environment = object_value(report["environment"], f"{lane}.environment")
    require(set(environment) == EXPECTED_ENVIRONMENT_KEYS, f"{lane} environment keys changed")
    require(environment == reference["environment"], f"{lane} environment producer identity differs from reused R1")
    require(environment["git_revision"] == binding["revision"] and environment["git_worktree_dirty"] is False and environment["allocator"] == "Rust system allocator", f"{lane} source/allocator identity changed")
    require(environment["rustflags"] == "-C force-frame-pointers=yes -C force-unwind-tables=yes" and environment["cpu_affinity"] == "2", f"{lane} instrumentation environment changed")
    config = object_value(report["configuration"], f"{lane}.configuration")
    reference_config = copy.deepcopy(reference["configuration"])
    reference_config["samples_per_case"] = sample_count
    reference_config["warmup_iterations_per_case"] = warmups
    require(config == reference_config, f"{lane} report configuration differs from frozen producer configuration")
    require(report["parallel_metrics"] == reference["parallel_metrics"], f"{lane} parallel producer metrics identity differs")
    rows = list_value(report["results"], f"{lane}.results")
    require(len(rows) == 1, f"{lane} must contain exactly one result")
    row = object_value(rows[0], f"{lane}.results[0]")
    exact_keys(row, ("case", "corpus", "elapsed_ns", "operation_metrics", "output_sha256", "sink", "source"), f"{lane}.results[0]")
    require(row["case"] == "pptx_streaming_create" and row["corpus"] == reference_row["corpus"] and row["source"] == reference_row["source"], f"{lane} corpus/source identity differs from reused R1")
    require(row["sink"] == reference_row["sink"] and row["output_sha256"] == reference_row["output_sha256"] == protocol_corpus_hash(reference_row), f"{lane} sink/output identity differs from reused R1")
    elapsed = object_value(row["elapsed_ns"], f"{lane}.elapsed_ns")
    verify_elapsed(elapsed, sample_count, f"{lane}.elapsed_ns")
    operation = object_value(row["operation_metrics"], f"{lane}.operation_metrics")
    exact_keys(operation, ("alignment", "cfb_phases", "latency_claim", "materialization", "process", "publication", "sample_count", "sample_indices", "sink", "source"), f"{lane}.operation_metrics")
    require(operation["sample_count"] == sample_count and operation["sample_indices"] == elapsed["sample_order"], f"{lane} operation sample identity changed")
    check_metric_vectors(operation["process"], sample_count, f"{lane}.operation_metrics.process")
    current_projection = expected_producer_projection(row, sample_count)
    reference_elapsed = object_value(reference_row["elapsed_ns"], "reference elapsed")
    reference_projection = expected_producer_projection(reference_row, len(reference_elapsed["samples"]))
    require(current_projection == reference_projection, f"{lane} deterministic producer counters/source/sink projection differs from reused R1")
    verify_catalog(root, lane, report, reference_catalog, receipt["artifacts"])


def protocol_corpus_hash(reference_row: Mapping[str, Any]) -> str:
    corpus = object_value(reference_row["corpus"], "reference corpus")
    digest = corpus.get("archive_sha256")
    require(isinstance(digest, str) and SHA_RE.fullmatch(digest), "reference corpus archive hash is invalid")
    return digest


def verify_reports(root: Path, protocol: Mapping[str, Any], binding: Mapping[str, Any], binding_hash: str, protocol_hash: str, compression: Mapping[str, Mapping[str, Any]], source_manifest: Mapping[str, str]) -> dict[str, Any]:
    del binding_hash, compression, source_manifest
    reference_path = root / "reuse/captures/R1-normal-large/report.json"
    reference = object_value(read_json(reference_path), "reused R1 report")
    reference_row = object_value(list_value(reference.get("results"), "reused R1 results")[0], "reused R1 row")
    require(reference["tool"] == EXPECTED_TOOL, "reused R1 producer tool is stale")
    require(reference_row["output_sha256"] == protocol["corpus_sha256"] == protocol_corpus_hash(reference_row), "reused R1 corpus hash does not bind protocol")
    require(reference_row["corpus"]["entry_count"] == protocol["slides"] and reference_row["corpus"]["archive_member_count"] == protocol["members"], "reused R1 corpus shape is stale")
    require(reference_row["sink"]["accepted_bytes"] == 7940406 and reference_row["sink"]["write_calls"] == 114958, "reused R1 sink counters are stale")
    reference_catalog = object_value(read_json(root / "normal-R1/corpus-catalog.json"), "reference corpus catalog")
    reports: dict[str, Any] = {}
    for lane, kind, samples, warmups in EXPECTED_ORDER:
        report = object_value(read_json(root / lane / "report.json"), f"{lane}/report.json")
        receipt = object_value(read_json(root / lane / "receipt.json"), f"{lane}/receipt.json")
        verify_report(root, lane, kind, samples, warmups, report, reference, reference_row, reference_catalog, binding, receipt)
        reports[lane] = {"kind": kind, "samples": samples, "warmups": warmups}
    return reports


def local_binding(root: Path, path: Path) -> dict[str, Any]:
    relative = path.relative_to(root).as_posix()
    meta = file_meta(path)
    return {"path": relative, **meta}


def verify_summary(root: Path, protocol: Mapping[str, Any], reports: Mapping[str, Any], counters: Mapping[str, dict[str, Any]]) -> dict[str, Any]:
    path = root / "summary.json"
    require(path.is_file(), "summary.json is missing")
    summary = object_value(read_json(path), "summary")
    expected_lanes: dict[str, Any] = {}
    for lane, kind, samples, warmups in EXPECTED_ORDER:
        report_path = root / lane / "report.json"
        report = object_value(read_json(report_path), f"{lane}/report.json")
        row = object_value(list_value(report["results"], f"{lane}.results")[0], f"{lane}.row")
        resource = (root / lane / "resource.log").read_text(encoding="utf-8")
        rss = re.search(r"Maximum resident set size \(kbytes\): (\d+)", resource)
        require(rss is not None, f"{lane} resource RSS missing for summary")
        expected_lanes[lane] = {
            "kind": kind,
            "samples": samples,
            "warmups": warmups,
            "report": local_binding(root, report_path),
            "elapsed_ns": row["elapsed_ns"],
            "whole_process_max_rss_kib": int(rss.group(1)),
            "output_sha256": row["output_sha256"],
            "sink": row["sink"],
        }
    r1 = expected_lanes["normal-R1"]["elapsed_ns"]
    r2 = expected_lanes["normal-R2"]["elapsed_ns"]
    drift: dict[str, Any] = {}
    for key in ("mean", "p50", "p95", "p99"):
        a = r1[key]
        b = r2[key]
        percent = 100 * (b / a - 1)
        drift[key] = {"r1_ns": a, "r2_ns": b, "signed_percent": percent, "exceeds_review_threshold": abs(percent) > protocol["normal_repeat_review_threshold_percent"]}
    expected = {
        "schema": "litchi-0475-normal-counter-context-v1",
        "lanes": expected_lanes,
        "normal_repeat_drift": drift,
        "normal_repeat_review_required": any(item["exceeds_review_threshold"] for item in drift.values()),
        "whole_process_counters": counters,
        "counter_source": local_binding(root, root / "counters/counters.csv"),
        "limitations": [
            "Normal runs bracket profilers on unchanged source and binary; this is descriptive repeatability, not a before/after or registered latency comparison.",
            "Profiler-instrumented elapsed vectors are diagnostic and remain separate from normal timing.",
            "perf stat values cover whole process including preflight and are scaled by perf for multiplexing; they are not operation-local counts.",
            "A reported zero hardware event is retained as reported and does not establish absence of cache misses; unsupported events remain unavailable.",
            "Whole-process RSS includes setup and profilers and cannot isolate writer heap or attribute its peak.",
        ],
    }
    require(summary == expected, "summary.json does not replay exactly from the raw reports")
    return {"status": "pass", "lanes": len(expected_lanes)}


def verify_file_binding(root: Path, value: Any, label: str, expected_path: str | None = None) -> None:
    binding = object_value(value, label)
    require({"bytes", "path", "sha256"} <= set(binding), f"{label} lacks path/bytes/sha256")
    relative = safe_relative(binding["path"], f"{label}.path")
    if expected_path is not None:
        require(relative == expected_path, f"{label}.path is not {expected_path}")
    path = bundle_path(root, relative, f"{label}.path")
    require(isinstance(binding["sha256"], str) and SHA_RE.fullmatch(binding["sha256"]), f"{label}.sha256 is invalid")
    require(isinstance(binding["bytes"], int) and binding["bytes"] >= 0, f"{label}.bytes is invalid")
    compression = binding.get("compression")
    if compression is None:
        check_meta(path, {"sha256": binding["sha256"], "bytes": binding["bytes"]}, label)
        if "compressed_sha256" in binding or "compressed_bytes" in binding:
            require(binding.get("compressed_sha256") == binding["sha256"] and binding.get("compressed_bytes") == binding["bytes"], f"{label} uncompressed binding has inconsistent compressed fields")
    elif compression == "gzip":
        require(binding.get("outside_bundle") is False, f"{label} gzip input is outside the bundle")
        require(isinstance(binding.get("compressed_sha256"), str) and SHA_RE.fullmatch(binding["compressed_sha256"]), f"{label}.compressed_sha256 is invalid")
        require(isinstance(binding.get("compressed_bytes"), int) and binding["compressed_bytes"] >= 0, f"{label}.compressed_bytes is invalid")
        current = file_meta(path)
        require(current == {"sha256": binding["compressed_sha256"], "bytes": binding["compressed_bytes"]}, f"{label} compressed file binding mismatch")
        digest = hashlib.sha256()
        size = 0
        try:
            with gzip.open(path, "rb") as stream:
                for block in iter(lambda: stream.read(1024 * 1024), b""):
                    digest.update(block)
                    size += len(block)
        except (OSError, EOFError) as error:
            fail(f"cannot decompress {label}: {error}")
        require(size == binding["bytes"] and digest.hexdigest() == binding["sha256"], f"{label} decompressed binding mismatch")
    else:
        fail(f"{label} declares unsupported compression {compression!r}")


def verify_cpu_analysis(root: Path, reports: Mapping[str, Any]) -> dict[str, Any]:
    outputs = {}
    for lane in ("cpu-P1", "cpu-P2"):
        path = root / f"{lane}-attribution.json"
        require(path.is_file(), f"{path.name} is missing")
        data = object_value(read_json(path), path.name)
        exact_keys(data, ("contexts", "inputs", "limitations", "materialized_preflight", "owner_families", "profile", "purpose", "sample_parser", "schema", "timing_semantics", "whole_process", "writer_context"), path.name)
        require(data["schema"] == "litchi-0475-pptx-streaming-cpu-attribution-v1", f"{path.name} schema changed")
        profile = object_value(data["profile"], f"{path.name}.profile")
        require(profile.get("event") == "cycles:u" and profile.get("sampled_period_is_elapsed_time") is False, f"{path.name} profile semantics changed")
        timing = object_value(data["timing_semantics"], f"{path.name}.timing_semantics")
        for key in ("phase_latency", "operation_counter", "sample_period_to_elapsed_conversion"):
            require(timing.get(key) is False, f"{path.name} converts sampled periods into timing")
        inputs = object_value(data["inputs"], f"{path.name}.inputs")
        require(inputs.get("bundle_root") == ".", f"{path.name} bundle root is not portable")
        verify_file_binding(root, inputs.get("perf_script"), f"{path.name}.inputs.perf_script", f"{lane}/perf-script.stdout.gz")
        verify_file_binding(root, inputs.get("parser"), f"{path.name}.inputs.parser", "cpu_parser.py")
        verify_file_binding(root, inputs.get("analyzer"), f"{path.name}.inputs.analyzer", "cpu_analyze.py")
        anchors = inputs.get("source_anchors")
        if anchors is not None:
            verify_file_binding(root, anchors, f"{path.name}.inputs.source_anchors")
        normal_reports = list_value(inputs.get("normal_reports"), f"{path.name}.inputs.normal_reports")
        require(normal_reports, f"{path.name} has no normal-report context")
        for index, report_binding in enumerate(normal_reports):
            value = object_value(report_binding, f"{path.name}.inputs.normal_reports[{index}]")
            file_info = value.get("file")
            verify_file_binding(root, file_info, f"{path.name}.inputs.normal_reports[{index}].file")
            report_path = root / file_info["path"]
            require(report_path.name == "report.json", f"{path.name} normal context is not a report")
        coverage = object_value(data["sample_parser"], f"{path.name}.sample_parser")
        for key in ("unknown_frame_weighted_event_period", "unparsed_frame_weighted_event_period", "empty_stack_weighted_event_period", "truncated_weighted_event_period", "lost_sample_count", "invalid_cycle_period", "parser_diagnostics"):
            require(key in coverage, f"{path.name} drops parser coverage field {key}")
        require(coverage.get("lost_period_is_unavailable") is True, f"{path.name} hides lost-period uncertainty")
        owner = object_value(data["owner_families"], f"{path.name}.owner_families")
        require(owner.get("overlap_allowed") is True and owner.get("period_sum_is_not_a_partition") is True, f"{path.name} owner overlap semantics changed")
        whole = object_value(data["whole_process"], f"{path.name}.whole_process")
        require(isinstance(whole.get("weighted_event_period"), int) and whole["weighted_event_period"] > 0, f"{path.name} has no positive whole-process period")
        limitations = list_value(data["limitations"], f"{path.name}.limitations")
        joined = " ".join(item for item in limitations if isinstance(item, str)).lower()
        for token in ("unknown", "truncated", "lost", "not", "elapsed"):
            require(token in joined, f"{path.name} limitations omit {token} coverage")
        outputs[lane] = {"schema": data["schema"], "perf_script": inputs["perf_script"]["path"]}
    return outputs


def verify_heap_analysis(root: Path, compression: Mapping[str, Mapping[str, Any]]) -> dict[str, Any]:
    path = root / "heap-attribution.json"
    require(path.is_file(), "heap-attribution.json is missing")
    data = object_value(read_json(path), "heap-attribution.json")
    exact_keys(data, ("schema", "tool", "traces"), "heap-attribution.json")
    require(data["schema"] == "litchi-0475-heap-attribution-v1", "heap attribution schema changed")
    tool = object_value(data["tool"], "heap-attribution.tool")
    require(tool.get("format") == "interpreted_heaptrack_v3", "heap attribution format changed")
    traces = list_value(data["traces"], "heap-attribution.traces")
    require({object_value(item, "heap trace").get("lane") for item in traces} == {"heap-H1", "heap-H2"}, "heap attribution lane coverage changed")
    for item_value in traces:
        item = object_value(item_value, "heap trace")
        lane = item.get("lane")
        trace = object_value(item.get("trace"), f"{lane}.trace")
        expected_path = f"{lane}/decoded.stdout.gz"
        trace_path = trace.get("path")
        require(isinstance(trace_path, str), f"{lane} heap trace path is malformed")
        if trace_path != expected_path:
            suffix_path(trace_path, expected_path, f"{lane} heap trace path")
        record = compression[f"{lane}/decoded.stdout"]
        require(trace.get("compressed_sha256") == record["compressed_sha256"] and trace.get("compressed_bytes") == record["compressed_bytes"], f"{lane} heap trace compressed binding changed")
        require(trace.get("format") == "interpreted_heaptrack_v3" and trace.get("file_version") == 3, f"{lane} heap trace format identity changed")
        records = object_value(item.get("records"), f"{lane}.records")
        for key in ("lines", "comments", "strings", "instruction_pointers", "traces", "allocation_descriptors", "allocation_events", "deallocation_events", "timestamps", "rss_events"):
            require(isinstance(records.get(key), int) and records[key] >= 0, f"{lane}.records.{key} is invalid")
        scope = object_value(item.get("scope"), f"{lane}.scope")
        for key in ("primary_phase_allocation_events", "unknown_scope_allocation_events", "unresolved_trace_ids", "classification", "whole_process"):
            require(key in scope, f"{lane}.scope drops {key}")
        require(isinstance(scope["unresolved_trace_ids"], list), f"{lane}.scope unresolved ids are not explicit")
        require(scope['whole_process']['status'] == 'unavailable' and scope['whole_process']['requested_bytes'] is None,
                f'{lane} scoped heap scan claims whole-process totals')
        require(records['complete_event_scan'] is False and records['event_scope'] == 'exact_phase_matching_descriptors_only',
                f'{lane} scoped heap record counts lost their boundary')
        phases = object_value(item.get('phase_attribution'), f'{lane}.phase_attribution')
        require(set(phases) == {'writer-under-run'}, f'{lane} unexpected heap phase')
        phase = phases['writer-under-run']
        categories = object_value(item.get('category_attribution'), f'{lane}.category_attribution')
        stacks = list_value(item.get('stacks'), f'{lane}.stacks')
        for metric in ('allocation_calls', 'deallocation_calls', 'requested_bytes', 'deallocated_bytes', 'live_bytes_at_end', 'outstanding_allocations'):
            require(sum(row[metric] for row in stacks) == phase[metric], f'{lane} heap stack {metric} total differs')
            require(sum(row[metric] for row in categories.values()) == phase[metric], f'{lane} heap category {metric} total differs')
            for category, aggregate in categories.items():
                require(sum(row[metric] for row in stacks if row['category'] == category) == aggregate[metric],
                        f'{lane} heap category {category} {metric} differs from its stacks')
        require(phase['allocation_calls'] == records['allocation_events'] == scope['primary_phase_allocation_events'], f'{lane} heap allocation event count differs')
        require(phase['deallocation_calls'] == records['deallocation_events'], f'{lane} heap deallocation count differs')
        require(phase['requested_bytes'] - phase['deallocated_bytes'] == phase['live_bytes_at_end'], f'{lane} heap byte balance differs')
        require(phase['peak_live_bytes'] == item['timeline']['peak_live_bytes'], f'{lane} heap projected peak differs')
        limitations = list_value(item.get("limitations"), f"{lane}.limitations")
        require(limitations, f"{lane}.limitations is empty")
    return {"traces": len(traces), "schema": data["schema"]}


def live_check(binding: Mapping[str, Any], source_manifest: Mapping[str, str]) -> dict[str, Any]:
    tree = Path(binding["tree"])
    binary = Path(binding["binary"]["path"])
    require(tree.is_dir(), f"live source tree is absent: {tree}")
    require(binary.is_file() and not binary.is_symlink(), f"live binary is absent: {binary}")
    try:
        revision = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=tree, text=True).strip()
        dirty = subprocess.check_output(["git", "status", "--porcelain"], cwd=tree, text=True).strip()
    except (OSError, subprocess.CalledProcessError) as error:
        fail(f"cannot inspect live source tree: {error}")
    require(revision == binding["revision"] and not dirty, "live source revision/worktree is not clean")
    require(binary.stat().st_size == binding["binary"]["bytes"] and sha256(binary) == binding["binary"]["sha256"], "live binary identity changed")
    for relative, digest in source_manifest.items():
        path = tree / relative
        require(path.is_file() and not path.is_symlink() and sha256(path) == digest, f"live source hash mismatch: {relative}")
    for relative, digest in binding["fixtures"].items():
        path = tree / relative
        require(path.is_file() and not path.is_symlink() and sha256(path) == digest, f"live fixture hash mismatch: {relative}")
    return {"status": "pass", "source_files": len(source_manifest), "binary": binding["binary"]["sha256"]}


def verify_bundle(root: Path = ROOT, *, sealed: bool = True, live: bool = False) -> dict[str, Any]:
    root = root.resolve()
    seal = verify_seal(root, required=sealed)
    protocol, protocol_hash = verify_protocol(root)
    binding, binding_hash, source_manifest = verify_binding(root, protocol, protocol_hash)
    reuse = verify_reuse(root, binding, protocol)
    chunks = verify_chunking(root)
    compression = compression_record_map(root)
    times: list[tuple[str, _datetime.datetime, _datetime.datetime]] = []
    for item in EXPECTED_ORDER:
        start, finish = verify_capture_receipt(root, item, binding, binding_hash, protocol_hash, compression, chunks)
        times.append((item[0], start, finish))
    declared = parse_utc(protocol["declared_utc"], "protocol.declared_utc")
    prepared = parse_utc(binding["prepared_finished_utc"], "binding.prepared_finished_utc")
    previous_finish: _datetime.datetime | None = None
    for lane, started, finished in times:
        require(started >= declared and started >= prepared, f"{lane} starts before frozen preparation")
        if previous_finish is not None:
            require(started >= previous_finish, f"capture order overlaps or is reversed at {lane}")
        previous_finish = finished
    verify_export_receipts(root, binding, protocol_hash, compression)
    supplement = verify_supplement(root, binding, compression)
    counters = verify_counter_file(root, protocol)
    reports = verify_reports(root, protocol, binding, binding_hash, protocol_hash, compression, source_manifest)
    summary = verify_summary(root, protocol, reports, counters)
    cpu = verify_cpu_analysis(root, reports)
    heap = verify_heap_analysis(root, compression)
    live_result = live_check(binding, source_manifest) if live else {"status": "not_checked"}
    return {
        "schema": VERIFICATION_SCHEMA,
        "status": "pass",
        "sealed": seal,
        "reuse": reuse,
        "lanes": len(EXPECTED_ORDER),
        "samples": sum(item[2] for item in EXPECTED_ORDER),
        "compressed_originals": len(compression),
        "chunked_originals": len(chunks),
        "supplement": supplement,
        "reports": reports,
        "summary": summary,
        "cpu": cpu,
        "heap": heap,
        "live": live_result,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--sealed", action="store_true", help="require and verify SHA256SUMS")
    group.add_argument("--unsealed", action="store_true", help="allow a bundle without SHA256SUMS")
    parser.add_argument("--live", action="store_true", help="also inspect binding.tree and binding.binary.path")
    parser.add_argument("--root", type=Path, default=ROOT, help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    try:
        result = verify_bundle(args.root, sealed=not args.unsealed, live=args.live)
    except VerificationError as error:
        parser.exit(1, f"verify.py: error: {error}\n")
    print(json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
