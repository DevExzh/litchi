#!/usr/bin/env python3
"""Pure portable custody checks for the 0476 evidence bundle.

This module authenticates the frozen protocol and the provenance surrounding
each retained measurement.  It deliberately does not import a capture/build
driver, execute a command, inspect a temporary checkout, or require Rust,
``perf``, or allocator tooling.  Producer arithmetic and paired comparisons
belong to the report/analyzer boundary; this file checks that those reports
come from the declared command, binary, source, and lane.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import math
from pathlib import Path, PurePosixPath
import re
import sys
from typing import Any, Iterable, Mapping, NoReturn, Sequence


ROOT = Path(__file__).resolve().parent
SCHEMA = "litchi-0476-custody-v1"
PROTOCOL_SCHEMA = "litchi-0476-protocol-v1"
SOURCE_SCHEMA = "litchi-0476-source-v1"
CAPTURE_SCHEMA = "litchi-0476-capture-v1"
SHA = re.compile(r"^[0-9a-fA-F]{64}$")
REVISION = re.compile(r"^[0-9a-fA-F]{40}$")

SHAPES = {"tiny": 8, "medium": 256, "large": 8192}
MODES = ("normal", "allocator")
ARMS = ("control", "candidate")
COUNTER_EVENTS = "cycles:u,instructions:u,branches:u,branch-misses:u,cache-misses:u,L1-dcache-load-misses:u,LLC-load-misses:u,page-faults,context-switches"
GUARD_SELECTORS = (
    "docx_streaming_create",
    "xlsx_streaming_create",
    "odt_streaming_create",
    "ods_streaming_create",
    "odp_streaming_create",
)
DRIVERS = {
    "build.py": "3991732207d16f385692f32d02fd4134b5141b7ed3673360243b798c62b3d167",
    "capture.py": "7b9984eeedafa94d1b04777c79136db7a19eab5288364a0d10d91863a93f05d6",
    "common.py": "865ac2ff6d8ae8c91334555a37aff93fd30c552da8a7c21dd82d17a42ecd2db6",
    "prepare.py": "b23a074b152dccbbd0eb79fae3cc584ad0a8a7133204034b6322fa2795a26efd",
}

FIXTURES = {
    "test-data/poi/test-data/spreadsheet/54016.xls": "2e050f1fbb31868b097aa6d4d0fe0a16af8e39c252af82d01cd8f93c4f9a911a",
    "test-data/rtf/watermark.rtf": "48d62dcd959e737b06ebb8255780bcaaf1e88056ff9c3d5a21d3ff5cd3ddf9cb",
}

CONTROL_REVISION = "60a1a4300a2844f372c0c916b5be69c50ce38668"
CANDIDATE_REVISION = "bbbdedaa4cd2229403cc6c6f39d66548d3887816"
SOURCE_EXPECTATIONS = {
    "control": {
        "revision": CONTROL_REVISION,
        "build_path": "/tmp/litchi-goal-0474/tree",
        "manifest_path": "sources/control.json",
        "manifest_files": 7034,
        "manifest_sha256": "17db263e9709a9965f71f12efe2bd411a568e15689713312e98f51bcb0bcc423",
        "selected_files": 9777,
    },
    "candidate": {
        "revision": CANDIDATE_REVISION,
        "build_path": "/tmp/litchi-goal-0476/candidate",
        "manifest_path": "sources/candidate.json",
        "manifest_files": 7035,
        "manifest_sha256": "9913faf6155e4296eae6f1573795fad815611d397a96fdf11e3fc910164540ce",
        "selected_files": 9778,
    },
}

BUILD_ENV = {
    "CARGO_BUILD_JOBS": "4",
    "CARGO_INCREMENTAL": "0",
    "CARGO_PROFILE_RELEASE_DEBUG": "1",
    "DEBUGINFOD_URLS": "",
    "LC_ALL": "C",
    "RUSTFLAGS": "-C force-frame-pointers=yes -C force-unwind-tables=yes",
    "RUSTUP_TOOLCHAIN": "1.98.1",
}
CAPTURE_ENV = dict(BUILD_ENV)
BUILD_ARGV = [
    "cargo",
    "build",
    "--release",
    "--locked",
    "--manifest-path",
    "tools/perf-baseline/Cargo.toml",
    "--target-dir",
    "/home/zhuhe/code/litchi/tools/perf-baseline/target",
    "--features",
    "allocator-metrics",
    "--bin",
    "litchi-perf-baseline",
    "--bin",
    "litchi-perf-baseline-alloc",
]

CORPORA = {
    "tiny": {
        "archive_bytes": 36259,
        "archive_member_count": 53,
        "archive_sha256": "951505889af106f032241c30f07b5d237e54822dade768c911aca2d0f68c22c5",
        "compression": "deflate",
        "entry_bytes": 46,
        "entry_count": 8,
        "generator": "litchi-pptx-streaming-plaintext-slides-v1",
        "name": "pptx-streaming-slides-tiny",
        "package_format": "PPTX/OOXML/ZIP",
        "payload_kind": "deterministic-plain-unicode-xml-significant-text-box-per-slide",
        "shape": "tiny",
        "target_entry": "ppt/presentation.xml",
        "target_payload_bytes": 848,
        "target_payload_sha256": "1c935108c7ac1b7d74416b9cf7b5edbbe858774f9605eec898b2dc762f0c7712",
        "uncompressed_payload_bytes": 368,
        "xlsx": None,
    },
    "medium": {
        "archive_bytes": 274398,
        "archive_member_count": 549,
        "archive_sha256": "1f33f8b2c36a2a51abc62d827e4c915dd3b4e52859b300323f250ab561942cf2",
        "compression": "deflate",
        "entry_bytes": 46,
        "entry_count": 256,
        "generator": "litchi-pptx-streaming-plaintext-slides-v1",
        "name": "pptx-streaming-slides-medium",
        "package_format": "PPTX/OOXML/ZIP",
        "payload_kind": "deterministic-plain-unicode-xml-significant-text-box-per-slide",
        "shape": "medium",
        "target_entry": "ppt/presentation.xml",
        "target_payload_bytes": 8944,
        "target_payload_sha256": "b4bc0b783f4309e21a34237444b12da891dbb31f1b5fda2ce3afcaa132578987",
        "uncompressed_payload_bytes": 11776,
        "xlsx": None,
    },
    "large": {
        "archive_bytes": 7940406,
        "archive_member_count": 16421,
        "archive_sha256": "c7b08da644e651046d368b1baaff9a12c6d7218c4f96914dacb1033722e4b527",
        "compression": "deflate",
        "entry_bytes": 46,
        "entry_count": 8192,
        "generator": "litchi-pptx-streaming-plaintext-slides-v1",
        "name": "pptx-streaming-slides-large",
        "package_format": "PPTX/OOXML/ZIP",
        "payload_kind": "deterministic-plain-unicode-xml-significant-text-box-per-slide",
        "shape": "large",
        "target_entry": "ppt/presentation.xml",
        "target_payload_bytes": 285476,
        "target_payload_sha256": "22f4de48dc64e3136b3c48fd9dcea578cab85c22b76f31a61afba8956a9def77",
        "uncompressed_payload_bytes": 376832,
        "xlsx": None,
    },
}


class CustodyError(ValueError):
    """A retained provenance artifact is missing or no longer authenticated."""


def fail(message: str) -> NoReturn:
    raise CustodyError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _integer(value: Any, label: str, expected: int | None = None) -> int:
    require(isinstance(value, int) and not isinstance(value, bool), f"{label}: expected an integer")
    if expected is not None:
        require(value == expected, f"{label}: expected {expected}")
    return value


def _success(value: Any, label: str) -> None:
    _integer(value, f"{label}.exit_code", 0)


def _pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def read_json(path: Path, label: str | None = None) -> Any:
    label = label or str(path)
    require(path.is_file() and not path.is_symlink(), f"{label}: regular file required")
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=_pairs,
            parse_constant=lambda value: (_ for _ in ()).throw(ValueError(f"non-finite JSON {value}")),
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as error:
        fail(f"{label}: cannot read JSON: {error}")


def canonical(value: Any) -> bytes:
    try:
        return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), allow_nan=False).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        fail(f"cannot canonicalize JSON: {error}")


def sha256_file(path: Path) -> tuple[str, int]:
    require(path.is_file() and not path.is_symlink(), f"{path}: regular file required")
    digest = hashlib.sha256()
    size = 0
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
                size += len(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest(), size


def digest(value: Any, label: str) -> str:
    require(isinstance(value, str) and SHA.fullmatch(value) is not None, f"{label}: expected SHA-256")
    return value.lower()


def _safe_relative(value: Any, label: str) -> str:
    require(isinstance(value, str) and value, f"{label}: expected non-empty relative POSIX path")
    require("\\" not in value, f"{label}: backslashes are not allowed")
    path = PurePosixPath(value)
    require(not path.is_absolute(), f"{label}: absolute path is not allowed")
    require(value not in {".", ".."} and all(part not in {"", ".", ".."} for part in path.parts), f"{label}: path escapes its root")
    return path.as_posix()


def _bundle_file(root: Path, value: Any, label: str) -> Path:
    relative = _safe_relative(value, label)
    current = root
    for part in PurePosixPath(relative).parts:
        current = current / part
        require(not current.is_symlink(), f"{label}: symlink component is not allowed")
    try:
        resolved = current.resolve(strict=True)
        require(resolved == root.resolve() or root.resolve() in resolved.parents, f"{label}: path escapes bundle")
    except OSError as error:
        fail(f"{label}: cannot resolve path: {error}")
    require(current.is_file(), f"{label}: regular file is missing")
    return current


def _metadata(path: Path, label: str) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"{label}: regular file required")
    sha, size = sha256_file(path)
    return {"bytes": size, "sha256": sha}


def _check_metadata(path: Path, expected: Any, label: str) -> None:
    require(isinstance(expected, dict) and set(expected) == {"bytes", "sha256"}, f"{label}: artifact metadata keys differ")
    expected_size = expected["bytes"]
    require(isinstance(expected_size, int) and not isinstance(expected_size, bool) and expected_size >= 0, f"{label}.bytes: invalid size")
    expected_hash = digest(expected["sha256"], f"{label}.sha256")
    actual = _metadata(path, label)
    require(actual == {"bytes": expected_size, "sha256": expected_hash}, f"{label}: hash or size differs")


def _keys(value: Any, required: Iterable[str], label: str, *, exact: bool = True) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: expected object")
    required_set = set(required)
    require(required_set <= set(value), f"{label}: missing required fields {sorted(required_set - set(value))}")
    if exact:
        require(set(value) == required_set, f"{label}: unexpected fields {sorted(set(value) - required_set)}")
    return value


def _timestamp(value: Any, label: str) -> _datetime.datetime:
    require(isinstance(value, str) and value, f"{label}: timestamp is missing")
    try:
        parsed = _datetime.datetime.fromisoformat(value)
    except ValueError as error:
        fail(f"{label}: invalid timestamp: {error}")
    require(parsed.tzinfo is not None, f"{label}: timestamp must include timezone")
    return parsed


def _interval(start: Any, finish: Any, label: str) -> tuple[_datetime.datetime, _datetime.datetime]:
    left = _timestamp(start, f"{label}.started_utc")
    right = _timestamp(finish, f"{label}.finished_utc")
    require(right >= left, f"{label}: finished before started")
    return left, right


def _expected_lane(suite: str, arm: str, lane: str, mode: str, repeat: str, shape: str) -> dict[str, str]:
    return {"arm": arm, "lane": lane, "mode": mode, "repeat": repeat, "shape": shape}


def _make_lanes() -> tuple[tuple[str, dict[str, str]], ...]:
    result: list[tuple[str, dict[str, str]]] = []
    cases = [(mode, shape) for mode in MODES for shape in ("tiny", "medium", "large")]
    for mode, shape in cases:
        result.append(("main", _expected_lane("main", "control", f"R1-control-{mode}-{shape}", mode, "R1", shape)))
        result.append(("main", _expected_lane("main", "candidate", f"R1-candidate-{mode}-{shape}", mode, "R1", shape)))
    for mode, shape in reversed(cases):
        result.append(("main", _expected_lane("main", "candidate", f"R2-candidate-{mode}-{shape}", mode, "R2", shape)))
        result.append(("main", _expected_lane("main", "control", f"R2-control-{mode}-{shape}", mode, "R2", shape)))
    for arm, lane, mode, shape in (
        ("candidate", "P-candidate-normal-tiny", "normal", "tiny"),
        ("candidate", "P-candidate-normal-large", "normal", "large"),
        ("candidate", "P-candidate-allocator-tiny", "allocator", "tiny"),
        ("candidate", "P-candidate-allocator-large", "allocator", "large"),
    ):
        result.append(("pilot", _expected_lane("pilot", arm, lane, mode, "pilot", shape)))
    for arm, lane, repeat in (
        ("control", "C1-control", "C1"),
        ("candidate", "C2-candidate", "C2"),
        ("candidate", "C3-candidate", "C3"),
        ("control", "C4-control", "C4"),
    ):
        result.append(("counter", _expected_lane("counter", arm, lane, "normal", repeat, "large")))
    for arm, lane, repeat in (
        ("control", "G1-control", "G1"),
        ("candidate", "G2-candidate", "G2"),
        ("candidate", "G3-candidate", "G3"),
        ("control", "G4-control", "G4"),
    ):
        result.append(("guard", _expected_lane("guard", arm, lane, "normal", repeat, "tiny,large")))
    return tuple(result)


EXPECTED_LANES = _make_lanes()
EXPECTED_BY_LANE = {row["lane"]: (suite, row) for suite, row in EXPECTED_LANES for row in (row,)}


def _verify_protocol(root: Path) -> tuple[dict[str, Any], str]:
    path = root / "protocol.json"
    require(path.is_file() and not path.is_symlink(), "protocol.json: missing")
    protocol = read_json(path, "protocol.json")
    _keys(protocol, ("schema", "selector", "shapes", "samples", "warmups", "workers", "cpu", "pilot_samples", "pilot_warmups", "counter_events", "guard_selectors", "drivers", "corpora", "order", "pilot_order", "counter_order", "guard_order", "large_requested_bytes_reduction_required_percent", "normal_regression_review_threshold_percent", "normal_repeat_review_threshold_percent", "peak_rss_regression_review_threshold_percent", "declared_utc", "acceptance", "claim", "control", "scope"), "protocol", exact=False)
    require(protocol.get("schema") == PROTOCOL_SCHEMA, "protocol.schema differs")
    require(protocol.get("selector") == "pptx_streaming_create", "protocol.selector differs")
    require(protocol.get("shapes") == SHAPES, "protocol.shapes differs")
    for key, expected in (("samples", 30), ("warmups", 3), ("workers", 1), ("cpu", 2), ("pilot_samples", 1), ("pilot_warmups", 0), ("large_requested_bytes_reduction_required_percent", 95), ("normal_regression_review_threshold_percent", 5), ("normal_repeat_review_threshold_percent", 5), ("peak_rss_regression_review_threshold_percent", 5)):
        _integer(protocol.get(key), f"protocol.{key}", expected)
    require(protocol.get("counter_events") == COUNTER_EVENTS, "protocol.counter_events differs")
    require(protocol.get("guard_selectors") == list(GUARD_SELECTORS), "protocol.guard_selectors differs")
    require(protocol.get("drivers") == DRIVERS, "protocol.drivers set or hash differs")

    arrays = {
        "order": [row for suite, row in EXPECTED_LANES if suite == "main"],
        "pilot_order": [row for suite, row in EXPECTED_LANES if suite == "pilot"],
        "counter_order": [row for suite, row in EXPECTED_LANES if suite == "counter"],
        "guard_order": [row for suite, row in EXPECTED_LANES if suite == "guard"],
    }
    for key, expected in arrays.items():
        require(protocol.get(key) == expected, f"protocol.{key} does not match the frozen lane order")
    require(isinstance(protocol.get("corpora"), dict) and protocol["corpora"] == CORPORA, "protocol.corpora differs from the frozen PPTX corpus")

    for name, expected_hash in DRIVERS.items():
        driver = root / name
        require(driver.is_file() and not driver.is_symlink(), f"protocol driver missing: {name}")
        actual, _ = sha256_file(driver)
        require(actual == expected_hash, f"protocol driver hash differs: {name}")
    protocol_hash, _ = sha256_file(path)
    return protocol, protocol_hash


def _verify_manifest(root: Path, arm: str, source: Mapping[str, Any]) -> dict[str, str]:
    expected = SOURCE_EXPECTATIONS[arm]
    _keys(source.get("source_manifest"), ("files", "path", "sha256"), f"{arm}.source_manifest")
    manifest_ref = source["source_manifest"]
    require(manifest_ref == {"files": expected["manifest_files"], "path": expected["manifest_path"], "sha256": expected["manifest_sha256"]}, f"{arm}.source_manifest differs")
    manifest_path = _bundle_file(root, manifest_ref["path"], f"{arm}.source_manifest.path")
    actual_hash, _ = sha256_file(manifest_path)
    require(actual_hash == expected["manifest_sha256"], f"{arm} source manifest hash differs")
    manifest = read_json(manifest_path, f"{arm} source manifest")
    require(isinstance(manifest, dict) and len(manifest) == expected["manifest_files"], f"{arm} source manifest file count differs")
    normalized: dict[str, str] = {}
    for name, value in manifest.items():
        relative = _safe_relative(name, f"{arm} source manifest path")
        require(relative not in normalized, f"{arm} source manifest contains duplicate path")
        normalized[relative] = digest(value, f"{arm} source manifest {relative}")
    return normalized


def _verify_source(root: Path, arm: str) -> tuple[dict[str, Any], dict[str, str]]:
    expected = SOURCE_EXPECTATIONS[arm]
    path = root / f"{arm}-source.json"
    source = read_json(path, f"{arm}-source.json")
    _keys(source, ("arm", "build_path", "fixtures", "revision", "schema", "source_manifest"), f"{arm}-source.json")
    require(source["arm"] == arm and source["schema"] == SOURCE_SCHEMA, f"{arm} source identity differs")
    require(source["revision"] == expected["revision"] and REVISION.fullmatch(source["revision"]) is not None, f"{arm} source revision differs")
    require(source["build_path"] == expected["build_path"], f"{arm} source build path differs")
    require(source["fixtures"] == FIXTURES, f"{arm} source fixtures differ")
    manifest = _verify_manifest(root, arm, source)
    return source, manifest


def _verify_prepare(root: Path, arm: str, source: Mapping[str, Any], source_hash: str, protocol: Mapping[str, Any]) -> tuple[dict[str, Any], tuple[_datetime.datetime, _datetime.datetime]]:
    expected = SOURCE_EXPECTATIONS[arm]
    path = root / f"{arm}-prepare.json"
    record = read_json(path, f"{arm}-prepare.json")
    _keys(record, ("clean_after", "common_sha256", "driver_sha256", "finished_utc", "selected_files", "source_binding_sha256", "started_utc"), f"{arm}-prepare.json")
    require(record["clean_after"] is True, f"{arm} prepare did not finish clean")
    require(record["common_sha256"] == protocol["drivers"]["common.py"], f"{arm} prepare common driver differs")
    require(record["driver_sha256"] == protocol["drivers"]["prepare.py"], f"{arm} prepare driver differs")
    require(record["selected_files"] == expected["selected_files"], f"{arm} selected file count differs")
    require(record["source_binding_sha256"] == source_hash, f"{arm} prepare source binding hash differs")
    times = _interval(record["started_utc"], record["finished_utc"], f"{arm}-prepare.json")
    return record, times


def _verify_binary(item: Any, arm: str, mode: str) -> dict[str, Any]:
    require(isinstance(item, dict) and set(item) == {"bytes", "path", "sha256"}, f"{arm}.binaries.{mode}: metadata keys differ")
    expected_name = f"/tmp/litchi-goal-0476/{arm}-{mode}"
    require(item["path"] == expected_name, f"{arm}.binaries.{mode}.path differs")
    require(isinstance(item["bytes"], int) and not isinstance(item["bytes"], bool) and item["bytes"] > 0, f"{arm}.binaries.{mode}.bytes is invalid")
    digest(item["sha256"], f"{arm}.binaries.{mode}.sha256")
    return item


def _verify_build(root: Path, arm: str, source: Mapping[str, Any], source_hash: str, protocol_hash: str, protocol: Mapping[str, Any]) -> tuple[dict[str, Any], str | None, tuple[_datetime.datetime, _datetime.datetime] | None]:
    expected = SOURCE_EXPECTATIONS[arm]
    path = root / f"{arm}-build.json"
    build = read_json(path, f"{arm}-build.json")
    base = ("arm", "binaries", "build_path", "fixtures", "fresh_build", "revision", "schema", "source_manifest")
    if arm == "control":
        _keys(build, (*base, "prepared_utc", "reused_build_sha256"), f"{arm}-build.json")
    else:
        _keys(build, (*base, "protocol_sha256", "receipt"), f"{arm}-build.json")
    require(build["arm"] == arm and build["schema"] == SOURCE_SCHEMA, f"{arm} build identity differs")
    require(build["build_path"] == expected["build_path"] and build["revision"] == expected["revision"], f"{arm} build source identity differs")
    require(build["fixtures"] == FIXTURES and build["source_manifest"] == source["source_manifest"], f"{arm} build source binding differs")
    require(build["fresh_build"] is (arm == "candidate"), f"{arm} fresh_build differs")
    require(isinstance(build["binaries"], dict) and set(build["binaries"]) == set(MODES), f"{arm} binary mode set differs")
    binaries = {mode: _verify_binary(build["binaries"][mode], arm, mode) for mode in MODES}
    build["binaries"] = binaries
    if arm == "control":
        _timestamp(build["prepared_utc"], "control-build.json.prepared_utc")
        digest_value, _ = sha256_file(root / "reuse" / "build.json")
        require(build["reused_build_sha256"] == digest_value, "control reused build hash differs")
        return build, None, None
    require(build["protocol_sha256"] == protocol_hash, "candidate build protocol hash differs")
    receipt_ref = build["receipt"]
    require(isinstance(receipt_ref, dict) and set(receipt_ref) == {"path", "sha256"}, "candidate build receipt reference is malformed")
    receipt_path = _bundle_file(root, receipt_ref["path"], "candidate build receipt path")
    receipt_hash, _ = sha256_file(receipt_path)
    require(receipt_hash == digest(receipt_ref["sha256"], "candidate build receipt hash"), "candidate build receipt hash differs")
    return build, receipt_ref["path"], None


def _verify_command_artifacts(root: Path, directory: Path, artifacts: Any, names: set[str], label: str) -> None:
    require(isinstance(artifacts, dict) and set(artifacts) == names, f"{label}: artifact set differs")
    for name, metadata in artifacts.items():
        relative = _safe_relative(name, f"{label}.artifacts.{name}")
        require(relative == name, f"{label}: artifact name is not canonical")
        _check_metadata(directory / relative, metadata, f"{label}/{name}")


def _verify_build_command(root: Path, build: Mapping[str, Any], source_hash: str, protocol_hash: str, protocol: Mapping[str, Any]) -> tuple[_datetime.datetime, _datetime.datetime]:
    started_path = root / "build-command.started.json"
    finished_path = root / "build-command.json"
    started = read_json(started_path, "build-command.started.json")
    finished = read_json(finished_path, "build-command.json")
    started_keys = {"argv", "common_sha256", "cwd", "driver_sha256", "environment", "protocol_sha256", "source_binding_sha256", "started_utc"}
    finished_keys = started_keys | {"artifacts", "exit_code", "finished_utc"}
    _keys(started, started_keys, "build-command.started.json")
    _keys(finished, finished_keys, "build-command.json")
    require(started["argv"] == BUILD_ARGV and finished["argv"] == BUILD_ARGV, "candidate build argv differs")
    for record, label in ((started, "build-command.started.json"), (finished, "build-command.json")):
        require(record["cwd"] == build["build_path"], f"{label}.cwd differs")
        require(record["environment"] == BUILD_ENV, f"{label}.environment differs")
        require(record["driver_sha256"] == protocol["drivers"]["build.py"], f"{label}.driver_sha256 differs")
        require(record["common_sha256"] == protocol["drivers"]["common.py"], f"{label}.common_sha256 differs")
        require(record["protocol_sha256"] == protocol_hash, f"{label}.protocol_sha256 differs")
        require(record["source_binding_sha256"] == source_hash, f"{label}.source_binding_sha256 differs")
    _success(finished["exit_code"], "build-command")
    start, finish = _interval(started["started_utc"], finished["finished_utc"], "build-command")
    for key in started_keys:
        require(started[key] == finished[key], f"build command started/final {key} differs")
    _verify_command_artifacts(root, root, finished["artifacts"], {"build-command.stdout", "build-command.stderr"}, "build-command")
    return start, finish


def _verify_reuse(root: Path, control_build: Mapping[str, Any]) -> None:
    reuse = root / "reuse"
    require(reuse.is_dir() and not reuse.is_symlink(), "reuse directory is missing")
    reuse_build_path = reuse / "build.json"
    reuse_build = read_json(reuse_build_path, "reuse/build.json")
    _keys(reuse_build, ("binaries", "build_path", "fixtures", "protocol_sha256", "receipt", "revision", "schema", "source_manifest"), "reuse/build.json")
    require(reuse_build["schema"] == "litchi-0474-build-v1" and reuse_build["revision"] == CONTROL_REVISION, "reuse build identity differs")
    require(reuse_build["build_path"] == "/tmp/litchi-goal-0474/tree" and reuse_build["fixtures"] == FIXTURES, "reuse build path or fixtures differ")
    require(reuse_build["source_manifest"] == {"files": 7034, "path": "sources/source.json", "sha256": "17db263e9709a9965f71f12efe2bd411a568e15689713312e98f51bcb0bcc423"}, "reuse source manifest reference differs")
    require(reuse_build["binaries"] == {"normal": {"bytes": 466033000, "path": "/tmp/litchi-goal-0474/normal", "sha256": "d205421b297b54a2f167ff0abde4b3814ef8c07c5883550b168d78d1af4f0935"}, "allocator": {"bytes": 465741960, "path": "/tmp/litchi-goal-0474/allocator", "sha256": "e7f50f245bf0a525cc24967f8dcf1604d53d5e94fd1fd36341d73506211b404d"}}, "reuse binaries differ")
    reuse_source_ref = root / "reuse" / "source-binding.json"
    reuse_source = read_json(reuse_source_ref, "reuse/source-binding.json")
    _keys(reuse_source, ("fixtures", "revision", "source_manifest"), "reuse/source-binding.json")
    require(reuse_source["revision"] == reuse_build["revision"] and reuse_source["fixtures"] == reuse_build["fixtures"] and reuse_source["source_manifest"] == reuse_build["source_manifest"], "reuse source binding differs")
    manifest_path = _bundle_file(reuse, reuse_build["source_manifest"]["path"], "reuse source manifest path")
    require(sha256_file(manifest_path)[0] == reuse_build["source_manifest"]["sha256"], "reuse source manifest hash differs")
    manifest = read_json(manifest_path, "reuse source manifest")
    require(isinstance(manifest, dict) and len(manifest) == 7034, "reuse source manifest file count differs")
    for name, value in manifest.items():
        _safe_relative(name, f"reuse source manifest {name}")
        digest(value, f"reuse source manifest {name}")
    require(reuse_build["protocol_sha256"] == "e4ea1bd41bde36c6eed1f8405dcd8ac7988ff9a373d8f8fb33a8c900d8961745", "reuse protocol hash differs")
    reuse_protocol = read_json(reuse / "protocol.json", "reuse/protocol.json")
    require(sha256_file(reuse / "protocol.json")[0] == reuse_build["protocol_sha256"], "reuse protocol file hash differs")
    require(isinstance(reuse_protocol, dict), "reuse protocol is malformed")

    receipt_ref = reuse_build["receipt"]
    require(isinstance(receipt_ref, dict) and set(receipt_ref) == {"path", "sha256"}, "reuse build receipt reference differs")
    receipt_path = _bundle_file(reuse, receipt_ref["path"], "reuse build receipt path")
    require(sha256_file(receipt_path)[0] == receipt_ref["sha256"], "reuse build receipt hash differs")
    started = read_json(reuse / "build-command.started.json", "reuse build command started")
    finished = read_json(receipt_path, "reuse build command")
    started_keys = {"argv", "cwd", "driver_sha256", "environment", "protocol_sha256", "source_binding_sha256", "started_utc"}
    finished_keys = started_keys | {"artifacts", "exit_code", "finished_utc"}
    _keys(started, started_keys, "reuse build-command.started.json")
    _keys(finished, finished_keys, "reuse build-command.json")
    require(finished["argv"] == BUILD_ARGV and started["argv"] == BUILD_ARGV, "reuse build argv differs")
    require(finished["cwd"] == reuse_build["build_path"] and started["cwd"] == reuse_build["build_path"], "reuse build cwd differs")
    require(finished["environment"] == BUILD_ENV and started["environment"] == BUILD_ENV, "reuse build environment differs")
    require(finished["driver_sha256"] == "0af06d660f176641a628dfff314554398494bcc70dbfc375e626aa1e170baf79", "reuse build driver differs")
    require(finished["protocol_sha256"] == reuse_build["protocol_sha256"] and finished["source_binding_sha256"] == sha256_file(reuse_source_ref)[0], "reuse build command bindings differ")
    _success(finished["exit_code"], "reuse build command")
    _interval(started["started_utc"], finished["finished_utc"], "reuse build command")
    for key in started_keys:
        require(started[key] == finished[key], f"reuse build started/final {key} differs")
    _verify_command_artifacts(reuse, reuse, finished["artifacts"], {"build-command.stdout", "build-command.stderr"}, "reuse build command")

    seal_path = reuse / "SHA256SUMS"
    require(seal_path.is_file() and not seal_path.is_symlink(), "reuse/SHA256SUMS is missing")
    seal_entries = _read_seal(reuse, seal_path)
    selected = (
        "build-command.json", "build-command.started.json", "build-command.stderr", "build-command.stdout", "build.json", "build.py",
        "captures/R1-allocator-large/report.json", "captures/R1-allocator-medium/report.json", "captures/R1-allocator-tiny/report.json",
        "captures/R1-normal-large/report.json", "captures/R1-normal-medium/report.json", "captures/R1-normal-tiny/report.json",
        "protocol.json", "source-binding.json", "sources/source.json",
    )
    for name in selected:
        path = _bundle_file(reuse, name, f"reuse selected artifact {name}")
        actual, _ = sha256_file(path)
        require(seal_entries.get(name) == actual, f"reuse/SHA256SUMS does not authenticate {name}")
    require(seal_entries.get("SHA256SUMS") is None, "reuse/SHA256SUMS must not self-reference")
    for mode in MODES:
        require(control_build["binaries"][mode]["sha256"] == reuse_build["binaries"][mode]["sha256"] and control_build["binaries"][mode]["bytes"] == reuse_build["binaries"][mode]["bytes"], f"control {mode} binary differs from authenticated reuse")


def _read_seal(root: Path, path: Path) -> dict[str, str]:
    entries: dict[str, str] = {}
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as error:
        fail(f"{path}: cannot read seal: {error}")
    for index, line in enumerate(lines, 1):
        if not line.strip():
            continue
        parts = line.split(maxsplit=1)
        require(len(parts) == 2, f"{path.name}:{index}: malformed seal line")
        value = digest(parts[0], f"{path.name}:{index}.sha256")
        name = parts[1][1:] if parts[1].startswith("*") else parts[1]
        name = _safe_relative(name, f"{path.name}:{index}.path")
        require(name not in entries and name != path.name, f"{path.name}:{index}: duplicate or self entry")
        entries[name] = value
    return entries


def _lane_directory(root: Path, suite: str, lane: str) -> Path:
    base = "pilots" if suite == "pilot" else "captures"
    path = root / base / lane
    require(path.is_dir() and not path.is_symlink(), f"{lane}: capture directory is missing")
    for item in path.iterdir():
        require(not item.is_symlink(), f"{lane}: symlink artifact is not allowed")
    return path


def _path_suffix(value: Any, suite: str, lane: str, filename: str, label: str) -> None:
    require(isinstance(value, str) and Path(value).is_absolute(), f"{label}: expected absolute retained command path")
    path = Path(value)
    expected_parts = ("pilots" if suite == "pilot" else "captures", lane, filename)
    require(tuple(path.parts[-len(expected_parts):]) == expected_parts, f"{label}: path does not bind {filename} to {lane}")


def _capture_argv(receipt: Mapping[str, Any], suite: str, row: Mapping[str, str], binary: Mapping[str, Any]) -> None:
    lane = row["lane"]
    samples = 1 if suite == "pilot" else 30
    warmups = 0 if suite == "pilot" else 3
    case_filter = ",".join(GUARD_SELECTORS) if suite == "guard" else "pptx_streaming_create"
    expected = ["taskset", "-c", "2", "/usr/bin/time", "-v", "-o", None]
    if suite == "counter":
        expected += ["perf", "stat", "-x", ";", "-o", None, "-e", COUNTER_EVENTS, "--"]
    expected += [binary["path"], "--workers", "1", "--warmup", str(warmups), "--samples", str(samples), "--case", case_filter, "--semantic-shape", row["shape"], "--json", None, "--corpus-manifest", None]
    argv = receipt.get("argv")
    require(isinstance(argv, list) and len(argv) == len(expected), f"{lane}: argv length differs")
    for index, (actual, wanted) in enumerate(zip(argv, expected)):
        if wanted is None:
            continue
        require(actual == wanted, f"{lane}: argv[{index}] differs")
    _path_suffix(argv[6], suite, lane, "resource.log", f"{lane}.argv.resource")
    offset = 0
    if suite == "counter":
        offset = 9
        _path_suffix(argv[12], suite, lane, "counters.csv", f"{lane}.argv.counters")
    json_index = 19 + offset
    catalog_index = json_index + 2
    _path_suffix(argv[json_index], suite, lane, "report.json", f"{lane}.argv.report")
    _path_suffix(argv[catalog_index], suite, lane, "corpus-catalog.json", f"{lane}.argv.catalog")


def _verify_catalog(path: Path, report: Mapping[str, Any], label: str) -> dict[str, Any]:
    catalog = read_json(path, f"{label}/corpus-catalog.json")
    _keys(catalog, ("manifest_version", "manifest_kind", "catalog_id", "canonicalization", "catalog_sha256", "content_set_sha256", "build", "corpora", "case_bindings"), label)
    _integer(catalog["manifest_version"], f"{label}.manifest_version", 2)
    require(catalog["manifest_kind"] == "corpus-catalog" and catalog["catalog_id"] == "litchi-perf-corpus-v2", f"{label}: catalog identity differs")
    require(catalog["canonicalization"] == {"algorithm": "sorted-json-utf8-compact-v1", "hash": "sha256"}, f"{label}: canonicalization differs")
    catalog_hash = digest(catalog["catalog_sha256"], f"{label}.catalog_sha256")
    content_hash = digest(catalog["content_set_sha256"], f"{label}.content_set_sha256")
    without = dict(catalog)
    del without["catalog_sha256"]
    require(hashlib.sha256(canonical(without)).hexdigest() == catalog_hash, f"{label}: catalog_sha256 canonical hash differs")

    corpora = catalog["corpora"]
    require(isinstance(corpora, list) and corpora, f"{label}: corpora must be non-empty")
    content_rows: list[dict[str, Any]] = []
    expected_archives: set[str] = set()
    report_results = report.get("results")
    require(isinstance(report_results, list) and report_results, f"{label}: report results are missing")
    for report_index, report_row in enumerate(report_results):
        require(isinstance(report_row, dict) and isinstance(report_row.get("corpus"), dict), f"{label}: report result {report_index} lacks corpus identity")
        expected_archives.add(digest(report_row["corpus"].get("archive_sha256"), f"{label}.report.results[{report_index}].corpus.archive_sha256"))
    seen_archives: set[str] = set()
    for index, value in enumerate(corpora):
        require(isinstance(value, dict), f"{label}.corpora[{index}]: expected object")
        require(isinstance(value.get("id"), str) and isinstance(value.get("bytes"), dict) and isinstance(value.get("members"), dict), f"{label}.corpora[{index}: content identity missing")
        bytes_value = value["bytes"]
        archive_hash = digest(bytes_value.get("archive_sha256"), f"{label}.corpora[{index}].bytes.archive_sha256")
        seen_archives.add(archive_hash)
        items = value["members"].get("items")
        require(isinstance(items, list), f"{label}.corpora[{index}].members.items: expected array")
        member_rows: list[dict[str, Any]] = []
        for member_index, item in enumerate(items):
            require(isinstance(item, dict), f"{label}.corpora[{index}].members.items[{member_index}]: expected object")
            require(set(item) == {"ordinal", "name", "sha256"}, f"{label}.corpora[{index}].members.items[{member_index}]: fields differ")
            require(isinstance(item["ordinal"], int) and not isinstance(item["ordinal"], bool) and item["ordinal"] >= 0, f"{label}.corpora[{index}].members.items[{member_index}].ordinal invalid")
            require(isinstance(item["name"], str) and item["name"], f"{label}.corpora[{index}].members.items[{member_index}].name invalid")
            member_rows.append({"ordinal": item["ordinal"], "name": item["name"], "sha256": digest(item["sha256"], f"{label}.corpora[{index}].members.items[{member_index}].sha256")})
        content_rows.append({"id": value["id"], "archive_sha256": archive_hash, "members": member_rows})
    require(expected_archives <= seen_archives, f"{label}: catalog omits one or more frozen corpus archives")
    bindings = catalog["case_bindings"]
    require(isinstance(bindings, list) and bindings, f"{label}: case_bindings must be non-empty")
    binding_rows: list[dict[str, Any]] = []
    for index, item in enumerate(bindings):
        require(isinstance(item, dict) and set(item) == {"case", "corpus_id", "legacy_name", "legacy_archive_sha256", "role"}, f"{label}.case_bindings[{index}]: fields differ")
        binding_rows.append({"case": item["case"], "corpus_id": item["corpus_id"], "role": item["role"]})
        digest(item["legacy_archive_sha256"], f"{label}.case_bindings[{index}].legacy_archive_sha256")
    require(hashlib.sha256(canonical({"corpora": content_rows, "case_bindings": binding_rows})).hexdigest() == content_hash, f"{label}: content_set_sha256 canonical hash differs")
    reference = report.get("corpus_catalog")
    require(isinstance(reference, dict) and reference == {"manifest_version": 2, "catalog_id": "litchi-perf-corpus-v2", "catalog_sha256": catalog_hash, "content_set_sha256": content_hash}, f"{label}: report catalog reference differs")
    return catalog


def _verify_report(root: Path, directory: Path, suite: str, row: Mapping[str, str], samples: int, warmups: int, build: Mapping[str, Any]) -> None:
    report_path = directory / "report.json"
    report = read_json(report_path, f"{row['lane']}/report.json")
    _keys(report, ("schema_version", "tool", "binary_identity", "environment", "configuration", "parallel_metrics", "results", "corpus_catalog"), f"{row['lane']}.report")
    _integer(report["schema_version"], f"{row['lane']}.report.schema_version", 1)
    binary = report["binary_identity"]
    _keys(binary, ("path", "binary_sha256", "binary_bytes", "mode_bits", "executable", "profile"), f"{row['lane']}.binary_identity")
    expected_binary = build["binaries"][row["mode"]]
    require(binary["path"] == expected_binary["path"] and binary["binary_sha256"] == expected_binary["sha256"] and binary["binary_bytes"] == expected_binary["bytes"], f"{row['lane']}: report binary identity differs")
    require(binary["executable"] is True and binary["profile"] == "release", f"{row['lane']}: report binary execution identity differs")
    environment = report["environment"]
    require(isinstance(environment, dict) and environment.get("git_revision") == build["revision"] and environment.get("git_worktree_dirty") is False, f"{row['lane']}: report source identity differs")
    configuration = report["configuration"]
    require(isinstance(configuration, dict), f"{row['lane']}: report configuration differs")
    _integer(configuration.get("samples_per_case"), f"{row['lane']}.report.configuration.samples_per_case", samples)
    _integer(configuration.get("warmup_iterations_per_case"), f"{row['lane']}.report.configuration.warmup_iterations_per_case", warmups)
    workers = configuration.get("execution_workers")
    require(isinstance(workers, list) and len(workers) == 1, f"{row['lane']}.report.configuration.execution_workers: expected one worker")
    _integer(workers[0], f"{row['lane']}.report.configuration.execution_workers[0]", 1)
    results = report["results"]
    expected_count = len(GUARD_SELECTORS) * 2 if suite == "guard" else 1
    require(isinstance(results, list) and len(results) == expected_count, f"{row['lane']}: report result count differs")
    if suite == "guard":
        expected_pairs = {(case, shape) for case in GUARD_SELECTORS for shape in ("tiny", "large")}
        actual_pairs = {(item.get("case"), item.get("corpus", {}).get("shape")) for item in results if isinstance(item, dict) and isinstance(item.get("corpus"), dict)}
        require(actual_pairs == expected_pairs, f"{row['lane']}: guard selector/shape result set differs")
    else:
        require(results[0].get("case") == "pptx_streaming_create", f"{row['lane']}: report case differs")
        corpus = results[0].get("corpus")
        require(isinstance(corpus, dict) and corpus == CORPORA[row["shape"]], f"{row['lane']}: report corpus differs from protocol")
        require(results[0].get("output_sha256") == CORPORA[row["shape"]]["archive_sha256"], f"{row['lane']}: output identity differs")
    _verify_catalog(directory / "corpus-catalog.json", report, row["lane"])


CAPTURE_SHARED_FIELDS = ("argv", "arm", "binary_sha256", "build_sha256", "case_filter", "clean_before", "common_sha256", "cwd", "driver_sha256", "environment", "lane", "mode", "protocol_sha256", "repeat", "revision", "samples", "schema", "shape", "started_utc", "suite", "warmups")


def _verify_capture(root: Path, protocol_hash: str, protocol: Mapping[str, Any], suite: str, row: Mapping[str, str], builds: Mapping[str, Mapping[str, Any]]) -> tuple[_datetime.datetime, _datetime.datetime]:
    directory = _lane_directory(root, suite, row["lane"])
    started_path = directory / "started.json"
    receipt_path = directory / "receipt.json"
    started = read_json(started_path, f"{row['lane']}/started.json")
    receipt = read_json(receipt_path, f"{row['lane']}/receipt.json")
    started_keys = set(CAPTURE_SHARED_FIELDS)
    receipt_keys = started_keys | {"artifacts", "binary_unchanged", "clean_after", "exit_code", "finished_utc", "source_unchanged"}
    _keys(started, started_keys, f"{row['lane']}/started.json")
    _keys(receipt, receipt_keys, f"{row['lane']}/receipt.json")
    for key in CAPTURE_SHARED_FIELDS:
        require(started[key] == receipt[key], f"{row['lane']}: started/final {key} differs")
    require(started["schema"] == CAPTURE_SCHEMA and started["suite"] == suite, f"{row['lane']}: capture schema/suite differs")
    require({key: started[key] for key in ("arm", "lane", "mode", "repeat", "shape")} == row, f"{row['lane']}: receipt lane identity differs")
    samples = 1 if suite == "pilot" else 30
    warmups = 0 if suite == "pilot" else 3
    expected_filter = ",".join(GUARD_SELECTORS) if suite == "guard" else "pptx_streaming_create"
    _integer(started["samples"], f"{row['lane']}.samples", samples)
    _integer(started["warmups"], f"{row['lane']}.warmups", warmups)
    require(started["case_filter"] == expected_filter, f"{row['lane']}: capture sample/filter configuration differs")
    build = builds[row["arm"]]
    require(started["revision"] == build["revision"] and started["binary_sha256"] == build["binaries"][row["mode"]]["sha256"], f"{row['lane']}: revision or binary binding differs")
    build_path = root / f"{row['arm']}-build.json"
    require(started["build_sha256"] == sha256_file(build_path)[0], f"{row['lane']}: build binding differs")
    require(started["protocol_sha256"] == protocol_hash and started["driver_sha256"] == DRIVERS["capture.py"] and started["common_sha256"] == DRIVERS["common.py"], f"{row['lane']}: driver/protocol binding differs")
    require(started["environment"] == CAPTURE_ENV and started["cwd"] == build["build_path"], f"{row['lane']}: capture environment or cwd differs")
    require(started["clean_before"] is True and receipt["clean_after"] is True and receipt["binary_unchanged"] is True and receipt["source_unchanged"] is True, f"{row['lane']}: capture status is not clean")
    _success(receipt["exit_code"], f"{row['lane']} capture")
    start, finish = _interval(started["started_utc"], receipt["finished_utc"], f"{row['lane']} capture")
    _capture_argv(receipt, suite, row, build["binaries"][row["mode"]])
    names = {"started.json", "stdout.log", "stderr.log", "report.json", "corpus-catalog.json", "resource.log"}
    if suite == "counter":
        names.add("counters.csv")
    _verify_command_artifacts(root, directory, receipt["artifacts"], names, f"{row['lane']} receipt")
    _verify_report(root, directory, suite, row, samples, warmups, build)
    return start, finish


def _verify_gate_receipts(root: Path, protocol: Mapping[str, Any]) -> int:
    directory = root / "validation"
    require(directory.is_dir() and not directory.is_symlink(), "validation directory is missing")
    finals = sorted(path for path in directory.glob("*.json") if not path.name.endswith(".started.json"))
    require(finals, "validation has no finished receipts")
    started_names = {path.name[:-len(".json")] + ".started.json" for path in finals}
    actual_started = {path.name for path in directory.glob("*.started.json")}
    require(actual_started == started_names, "validation started/final receipt set differs")
    for path in finals:
        label = path.stem
        record = read_json(path, f"validation/{path.name}")
        started_path = directory / f"{label}.started.json"
        started = read_json(started_path, f"validation/{started_path.name}")
        started_keys = {"argv", "common_sha256", "cwd", "driver_sha256", "environment", "source_before", "started_utc"}
        finished_keys = started_keys | {"artifacts", "exit_code", "finished_utc", "source_after", "source_unchanged"}
        _keys(started, started_keys, f"validation/{started_path.name}")
        _keys(record, finished_keys, f"validation/{path.name}")
        for key in started_keys:
            require(started[key] == record[key], f"validation/{label}: started/final {key} differs")
        for command, command_label in ((started, f"validation/{started_path.name}"), (record, f"validation/{path.name}")):
            require(isinstance(command["argv"], list) and command["argv"] and all(isinstance(value, str) for value in command["argv"]), f"{command_label}.argv: expected a non-empty string array")
            require(isinstance(command["cwd"], str) and Path(command["cwd"]).is_absolute(), f"{command_label}.cwd: expected an absolute path")
        # Validation history intentionally retains rejected attempts.  Custody
        # authenticates their recorded status and output; the final acceptance
        # gate decides which retained labels must have exit_code == 0.
        require(isinstance(record["exit_code"], int) and not isinstance(record["exit_code"], bool) and record["source_unchanged"] is True, f"validation/{label}: command/source status is malformed")
        _interval(started["started_utc"], record["finished_utc"], f"validation/{label}")
        require(record["common_sha256"] == DRIVERS["common.py"], f"validation/{label}: common driver differs")
        gate_hash, _ = sha256_file(root / "gate.py")
        require(record["driver_sha256"] == gate_hash, f"validation/{label}: gate driver differs")
        require(record["environment"] == BUILD_ENV, f"validation/{label}: environment differs")
        _verify_command_artifacts(root, directory, record["artifacts"], {f"{label}.stdout", f"{label}.stderr"}, f"validation/{label}")
        before = record["source_before"]
        after = record["source_after"]
        require(before == after, f"validation/{label}: source snapshot changed")
        _verify_snapshot(root, before, f"validation/{label}.source_before")
    return len(finals)


def _verify_snapshot(root: Path, value: Any, label: str) -> None:
    require(isinstance(value, dict) and set(value) == {"files", "path", "sha256"}, f"{label}: snapshot fields differ")
    path = _bundle_file(root, value["path"], f"{label}.path")
    expected_hash = digest(value["sha256"], f"{label}.sha256")
    require(sha256_file(path)[0] == expected_hash, f"{label}: snapshot hash differs")
    manifest = read_json(path, label)
    expected_files = _integer(value["files"], f"{label}.files")
    require(isinstance(manifest, dict) and len(manifest) == expected_files, f"{label}: snapshot file count differs")
    for name, file_hash in manifest.items():
        _safe_relative(name, f"{label}.{name}")
        digest(file_hash, f"{label}.{name}")


def verify(root: Path = ROOT) -> dict[str, Any]:
    """Verify every frozen provenance boundary and return a compact receipt."""

    root = Path(root).resolve()
    protocol, protocol_hash = _verify_protocol(root)
    sources: dict[str, dict[str, Any]] = {}
    manifests: dict[str, dict[str, str]] = {}
    source_hashes: dict[str, str] = {}
    prepare_times: dict[str, tuple[_datetime.datetime, _datetime.datetime]] = {}
    builds: dict[str, dict[str, Any]] = {}
    for arm in ARMS:
        source_path = root / f"{arm}-source.json"
        source, manifest = _verify_source(root, arm)
        sources[arm] = source
        manifests[arm] = manifest
        source_hashes[arm] = sha256_file(source_path)[0]
        _prepare, prepare_times[arm] = _verify_prepare(root, arm, source, source_hashes[arm], protocol)
    builds["control"], _, _ = _verify_build(root, "control", sources["control"], source_hashes["control"], protocol_hash, protocol)
    builds["candidate"], _, _ = _verify_build(root, "candidate", sources["candidate"], source_hashes["candidate"], protocol_hash, protocol)
    control_prepared = _timestamp(builds["control"]["prepared_utc"], "control-build.json.prepared_utc")
    require(control_prepared >= prepare_times["control"][1], "control build was prepared before control preparation finished")
    _verify_reuse(root, builds["control"])
    build_start, build_finish = _verify_build_command(root, builds["candidate"], source_hashes["candidate"], protocol_hash, protocol)
    require(prepare_times["candidate"][1] <= build_start, "candidate build started before candidate preparation finished")

    intervals: list[tuple[_datetime.datetime, _datetime.datetime, str]] = []
    capture_count = 0
    for suite, row in EXPECTED_LANES:
        start, finish = _verify_capture(root, protocol_hash, protocol, suite, row, builds)
        intervals.append((start, finish, row["lane"]))
        capture_count += 1
    declared = _timestamp(protocol["declared_utc"], "protocol.declared_utc")
    require(build_finish >= declared, "candidate build finished before protocol declaration")
    require(all(start >= declared for start, _finish, _lane in intervals), "capture started before protocol declaration")
    require(all(start >= build_finish for start, _finish, _lane in intervals), "capture started before candidate build finished")
    intervals.sort()
    for (_start, previous_finish, previous_lane), (next_start, _finish, next_lane) in zip(intervals, intervals[1:]):
        require(next_start >= previous_finish, f"capture chronology overlaps: {previous_lane} and {next_lane}")

    gate_count = _verify_gate_receipts(root, protocol)
    return {
        "schema": SCHEMA,
        "status": "pass",
        "protocol_sha256": protocol_hash,
        "lanes": capture_count,
        "main_lanes": 24,
        "pilot_lanes": 4,
        "counter_lanes": 4,
        "guard_lanes": 4,
        "validation_receipts": gate_count,
        "source_manifests": {arm: len(manifests[arm]) for arm in ARMS},
    }


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    try:
        print(json.dumps(verify(args.root), indent=2, sort_keys=True))
    except (CustodyError, OSError, TypeError, ValueError, KeyError) as error:
        print(json.dumps({"schema": SCHEMA, "status": "fail", "error": str(error)}, sort_keys=True))
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
