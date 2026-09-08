#!/usr/bin/env python3
"""Portable verifier for the 0477 ZIP directory-spool evidence bundle.

The verifier reads only files below the evidence directory.  It authenticates
the frozen capture plan, build-gate and source-manifest references, binary
metadata, capture receipts, producer reports, and the derived summary.  It
does not run a build or capture command and it never follows an absolute path
from a receipt into the worktree or a temporary directory.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import math
import re
import statistics
import sys
from pathlib import Path, PurePosixPath
from typing import Any, Iterable, NoReturn


ROOT = Path(__file__).resolve().parent
PROTOCOL_SCHEMA = "zip-directory-spool-capture-v1"
REPORT_SCHEMA = "zip-directory-spool-v1"
SUMMARY_SCHEMA = "zip-directory-spool-summary-v1"
SHA256 = re.compile(r"^[0-9a-fA-F]{64}$")
COUNT_VALUES = (8, 256, 8192)
METHOD_VALUES = ("store", "deflate")
POLICY_VALUES = ("control", "spool")
INSTRUMENTATION_VALUES = ("normal", "allocator")
SAMPLES = 30
WARMUPS = 3
REPEATS = 1
MAX_SPOOL_BYTES = 64 * 1024 * 1024
SPOOL_BUFFER_BYTES = 16 * 1024
PAYLOAD_BYTES = 256
CENTRAL_FIXED_BYTES = 46
SLIDE_NAME = "ppt/slides/slide-00000.xml"
REQUIRED_FINAL_VALIDATION_LABELS = (
    "libraries-clippy-corrected",
    "opc-tests",
    "shared-streaming-unit",
    "docx-streaming",
    "pptx-streaming",
    "odf-streaming",
    "odp-streaming",
    "ods-streaming",
    "odt-streaming",
    "harness-tests",
    "harness-clippy-corrected",
    "harness-focused-final",
    "libraries-rustdoc",
    "libraries-format",
    "harness-format",
    "zip-tests-final",
    "boundaries",
    "registry-strict",
    "build-normal",
    "build-allocator",
    "pilot-normal",
    "pilot-allocator",
    "evidence-tests",
)
REQUIRED_FINAL_NO_SOURCE_EXCLUSIONS = frozenset({
    "harness-clippy-corrected",
    "harness-focused-final",
    "build-normal",
    "build-allocator",
    "pilot-normal",
    "pilot-allocator",
    "evidence-tests",
})


class VerificationError(ValueError):
    """A retained evidence artifact failed an independent check."""


def fail(message: str) -> NoReturn:
    raise VerificationError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _pairs(items: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in items:
        if key in result:
            raise ValueError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def read_json(path: Path, label: str) -> Any:
    require(path.is_file() and not path.is_symlink(), f"{label}: regular file required")
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=_pairs,
            parse_constant=lambda value: (_ for _ in ()).throw(
                ValueError(f"non-finite JSON value {value}")
            ),
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as error:
        fail(f"{label}: invalid JSON: {error}")


def sha256_file(path: Path, label: str) -> tuple[str, int]:
    require(path.is_file() and not path.is_symlink(), f"{label}: regular file required")
    digest = hashlib.sha256()
    size = 0
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
                size += len(block)
    except OSError as error:
        fail(f"{label}: cannot read: {error}")
    return digest.hexdigest(), size


def digest(value: Any, label: str) -> str:
    require(isinstance(value, str) and SHA256.fullmatch(value) is not None,
            f"{label}: expected a SHA-256 digest")
    return value.lower()


def safe_relative(value: Any, label: str) -> str:
    require(isinstance(value, str) and value,
            f"{label}: expected a non-empty relative POSIX path")
    require("\\" not in value, f"{label}: backslashes are not allowed")
    path = PurePosixPath(value)
    require(not path.is_absolute(), f"{label}: absolute path is not allowed")
    require(value not in {".", ".."} and all(
        part not in {"", ".", ".."} for part in path.parts
    ), f"{label}: path escapes its bundle root")
    return path.as_posix()


def bundle_file(root: Path, value: Any, label: str) -> Path:
    relative = safe_relative(value, label)
    root = root.resolve()
    current = root
    for part in PurePosixPath(relative).parts:
        current = current / part
        require(not current.is_symlink(), f"{label}: symlink component is not allowed")
    try:
        resolved = current.resolve(strict=True)
    except OSError as error:
        fail(f"{label}: cannot resolve path: {error}")
    require(resolved == root or root in resolved.parents,
            f"{label}: path escapes the evidence bundle")
    require(current.is_file(), f"{label}: regular file is missing")
    return current


def metadata(path: Path, label: str) -> dict[str, Any]:
    value, size = sha256_file(path, label)
    return {"bytes": size, "sha256": value}


def check_metadata(path: Path, expected: Any, label: str) -> None:
    require(isinstance(expected, dict), f"{label}: metadata must be an object")
    require(set(expected) == {"bytes", "sha256"},
            f"{label}: metadata fields differ")
    size = expected["bytes"]
    require(isinstance(size, int) and not isinstance(size, bool) and size >= 0,
            f"{label}.bytes: invalid size")
    expected_hash = digest(expected["sha256"], f"{label}.sha256")
    require(metadata(path, label) == {"bytes": size, "sha256": expected_hash},
            f"{label}: metadata does not match the retained file")


def integer(value: Any, label: str, expected: int | None = None) -> int:
    require(isinstance(value, int) and not isinstance(value, bool),
            f"{label}: expected an integer")
    require(value >= 0, f"{label}: expected a non-negative integer")
    if expected is not None:
        require(value == expected, f"{label}: expected {expected}, got {value}")
    return value


def timestamp(value: Any, label: str) -> _datetime.datetime:
    require(isinstance(value, str) and value, f"{label}: missing timestamp")
    try:
        parsed = _datetime.datetime.fromisoformat(value)
    except ValueError as error:
        fail(f"{label}: invalid timestamp: {error}")
    require(parsed.tzinfo is not None, f"{label}: timezone is required")
    return parsed


def interval(start: Any, finish: Any, label: str) -> None:
    require(timestamp(finish, f"{label}.finished_utc") >=
            timestamp(start, f"{label}.started_utc"),
            f"{label}: finished before started")


def absolute_path(value: Any, label: str) -> str:
    require(isinstance(value, str) and value, f"{label}: expected a path string")
    require(Path(value).is_absolute(), f"{label}: command path must be absolute")
    return value


def expected_captures() -> list[dict[str, Any]]:
    forward = [
        (instrumentation, count, method, policy)
        for instrumentation in INSTRUMENTATION_VALUES
        for count in COUNT_VALUES
        for method in METHOD_VALUES
        for policy in POLICY_VALUES
    ]
    result: list[dict[str, Any]] = []
    for repeat, sequence in ((1, forward), (2, list(reversed(forward)))):
        for instrumentation, count, method, policy in sequence:
            label = f"r{repeat}-{instrumentation}-{count}-{method}-{policy}"
            result.append({
                "label": label,
                "instrumentation": instrumentation,
                "count": count,
                "method": method,
                "policy": policy,
                "repeat": repeat,
            })
    return result


def check_command_argv(argv: Any, spec: dict[str, Any], label: str) -> None:
    require(isinstance(argv, list), f"{label}.argv: expected an array")
    require(all(isinstance(value, str) for value in argv),
            f"{label}.argv: every argument must be a string")
    require(len(argv) == 28, f"{label}.argv: expected 28 arguments")
    fixed = {
        0: "/usr/bin/time", 1: "-v", 2: "-o", 4: "/usr/bin/taskset",
        5: "-c", 6: "2", 8: "--mode", 10: "--counts", 12: "--methods",
        14: "--samples", 15: str(SAMPLES), 16: "--warmups", 17: str(WARMUPS),
        18: "--repeats", 19: str(REPEATS), 20: "--max-spool-bytes",
        21: str(MAX_SPOOL_BYTES), 22: "--spool-buffer-bytes",
        23: str(SPOOL_BUFFER_BYTES), 24: "--json", 26: "--spool-dir",
    }
    for index, expected in fixed.items():
        require(argv[index] == expected,
                f"{label}.argv[{index}]: expected {expected!r}, got {argv[index]!r}")
    require(argv[9] == spec["policy"], f"{label}: mode flag is not bound to policy")
    require(argv[11] == str(spec["count"]), f"{label}: count flag is not bound")
    require(argv[13] == spec["method"], f"{label}: method flag is not bound")
    absolute_path(argv[3], f"{label}.argv.resource")
    absolute_path(argv[7], f"{label}.argv.binary")
    absolute_path(argv[25], f"{label}.argv.report")
    absolute_path(argv[27], f"{label}.argv.spool_dir")
    require(argv[3].endswith(f"/captures/{spec['label']}.resource"),
            f"{label}: resource path does not bind the capture label")
    require(argv[25].endswith(f"/captures/{spec['label']}.report.json"),
            f"{label}: report path does not bind the capture label")
    require(argv[27].endswith(f"/spools/{spec['label']}"),
            f"{label}: spool path does not bind the capture label")


def check_protocol(root: Path) -> tuple[dict[str, Any], str]:
    path = root / "protocol.json"
    protocol = read_json(path, "protocol.json")
    require(isinstance(protocol, dict), "protocol.json: expected an object")
    require(protocol.get("schema") == PROTOCOL_SCHEMA,
            "protocol.schema differs")
    integer(protocol.get("samples"), "protocol.samples", SAMPLES)
    integer(protocol.get("warmups"), "protocol.warmups", WARMUPS)
    integer(protocol.get("cpu"), "protocol.cpu", 2)
    require(protocol.get("normal_and_allocator_timings_separate") is True,
            "protocol must separate normal and allocator timings")
    require(protocol.get("process_rss_includes_setup_oracles_and_teardown") is True,
            "protocol must identify whole-process RSS scope")
    integer(protocol.get("regression_review_percent"),
            "protocol.regression_review_percent", 5)
    require(isinstance(protocol.get("environment"), dict),
            "protocol.environment is missing")
    for key in ("RUSTUP_TOOLCHAIN", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
                "CARGO_PROFILE_RELEASE_DEBUG", "RUSTFLAGS", "DEBUGINFOD_URLS",
                "LC_ALL"):
        require(key in protocol["environment"],
                f"protocol.environment.{key} is missing")
    scope = protocol.get("scope")
    require(isinstance(scope, str), "protocol.scope is missing")
    for phrase in ("complete", "file open/close", "oracle", "cleanup"):
        require(phrase.lower() in scope.lower(),
                f"protocol.scope does not declare {phrase!r}")
    scripts = protocol.get("scripts")
    require(isinstance(scripts, dict) and scripts,
            "protocol.scripts is missing")
    for required in ("common.py", "capture.py"):
        require(required in scripts,
                f"protocol.scripts.{required} is missing")
    for name, expected_hash in scripts.items():
        relative = safe_relative(name, "protocol.scripts path")
        script_path = bundle_file(root, relative, f"protocol.scripts.{name}")
        require(sha256_file(script_path, f"protocol script {name}")[0] ==
                digest(expected_hash, f"protocol.scripts.{name}"),
                f"protocol script hash differs: {name}")

    expected = expected_captures()
    captures = protocol.get("captures")
    require(isinstance(captures, list) and len(captures) == len(expected),
            f"protocol.captures: expected exactly {len(expected)} rows")
    for index, (actual, wanted) in enumerate(zip(captures, expected)):
        require(isinstance(actual, dict), f"protocol.captures[{index}]: expected object")
        for key, value in wanted.items():
            require(actual.get(key) == value,
                    f"protocol.captures[{index}].{key}: expected {value!r}")
        check_command_argv(actual.get("argv"), wanted,
                           f"protocol.captures[{index}]")
    protocol_hash = sha256_file(path, "protocol.json")[0]
    return protocol, protocol_hash


def _walk(value: Any) -> Iterable[tuple[str, Any]]:
    if isinstance(value, dict):
        for key, child in value.items():
            yield key, child
            yield from _walk(child)
    elif isinstance(value, list):
        for child in value:
            yield from _walk(child)


def check_source_manifest(root: Path, reference: Any, label: str) -> str:
    return read_source_manifest(root, reference, label)[0]


def read_source_manifest(root: Path, reference: Any,
                         label: str) -> tuple[str, dict[str, str]]:
    require(isinstance(reference, dict), f"{label}: manifest reference must be an object")
    require("path" in reference and "sha256" in reference,
            f"{label}: manifest path/hash is missing")
    path = bundle_file(root, reference["path"], f"{label}.path")
    expected_hash = digest(reference["sha256"], f"{label}.sha256")
    actual_hash = sha256_file(path, label)[0]
    require(actual_hash == expected_hash, f"{label}: manifest hash differs")
    manifest = read_json(path, label)
    require(isinstance(manifest, dict), f"{label}: manifest must be an object")
    for name, value in manifest.items():
        safe_relative(name, f"{label} entry")
        digest(value, f"{label} entry {name}")
    if "files" in reference:
        integer(reference["files"], f"{label}.files", len(manifest))
    return expected_hash, manifest


def binary_binding_present(value: Any, spec: dict[str, Any]) -> bool:
    if isinstance(value, dict):
        actual_hash = value.get("sha256")
        if (isinstance(actual_hash, str)
                and actual_hash.lower() == spec["sha256"].lower()
                and value.get("bytes") == spec["bytes"]):
            return True
        return any(binary_binding_present(child, spec) for child in value.values())
    if isinstance(value, list):
        return any(binary_binding_present(child, spec) for child in value)
    return False


def binary_path_binding_present(value: Any, spec: dict[str, Any]) -> bool:
    """Require the final build wrapper to name the copied binary it sealed."""
    if isinstance(value, dict):
        actual_hash = value.get("sha256")
        if (value.get("path") == spec["path"]
                and isinstance(actual_hash, str)
                and actual_hash.lower() == spec["sha256"].lower()
                and value.get("bytes") == spec["bytes"]):
            return True
        return any(binary_path_binding_present(child, spec) for child in value.values())
    if isinstance(value, list):
        return any(binary_path_binding_present(child, spec) for child in value)
    return False


def check_wrapper_binary_metadata(gate: dict[str, Any], spec: dict[str, Any],
                                 instrumentation: str) -> None:
    """Check the wrapper's recorded origin and copied binary identities."""
    copied = gate.get("binary")
    require(isinstance(copied, dict),
            f"{instrumentation}: build wrapper binary metadata is missing")
    for key in ("path", "bytes", "sha256"):
        require(key in copied,
                f"{instrumentation}: copied binary metadata omits {key}")
    require(copied["path"] == spec["path"],
            f"{instrumentation}: copied binary path differs from binaries.json")
    absolute_path(copied["path"], f"{instrumentation}.binary.path")
    integer(copied["bytes"], f"{instrumentation}.binary.bytes", spec["bytes"])
    require(digest(copied["sha256"], f"{instrumentation}.binary.sha256") ==
            spec["sha256"],
            f"{instrumentation}: copied binary digest differs")

    origin = gate.get("original_binary")
    require(isinstance(origin, dict),
            f"{instrumentation}: original binary metadata is missing")
    for key in ("path", "bytes", "sha256"):
        require(key in origin,
                f"{instrumentation}: original binary metadata omits {key}")
    absolute_path(origin["path"], f"{instrumentation}.original_binary.path")
    require(origin["path"] != copied["path"],
            f"{instrumentation}: origin and copied binary paths are identical")
    integer(origin["bytes"], f"{instrumentation}.original_binary.bytes",
            spec["bytes"])
    require(digest(origin["sha256"],
                   f"{instrumentation}.original_binary.sha256") == spec["sha256"],
            f"{instrumentation}: origin binary digest differs")


def check_validation_gate_reference(root: Path, wrapper: dict[str, Any],
                                    instrumentation: str) -> dict[str, Any]:
    """Check the final wrapper's content-addressed underlying gate receipt."""
    reference = wrapper.get("validation_gate")
    require(isinstance(reference, dict),
            f"{instrumentation}: final build lacks validation_gate reference")
    require(set(reference) == {"path", "sha256"},
            f"{instrumentation}: validation_gate reference fields differ")
    gate_path = bundle_file(root, reference["path"],
                            f"{instrumentation}.validation_gate.path")
    expected_hash = digest(reference["sha256"],
                           f"{instrumentation}.validation_gate.sha256")
    require(sha256_file(gate_path, f"{instrumentation} validation gate")[0] ==
            expected_hash,
            f"{instrumentation}: validation_gate hash differs")
    receipt = read_json(gate_path, f"{instrumentation} validation gate")
    require(isinstance(receipt, dict),
            f"{instrumentation}: validation_gate receipt is malformed")
    require(receipt.get("exit_code") == 0,
            f"{instrumentation}: designated final validation gate did not succeed")
    require(receipt.get("source_unchanged") is True,
            f"{instrumentation}: designated final validation gate changed source")
    mirrored = (
        "argv", "cwd", "environment", "driver_sha256", "common_sha256",
        "source_before", "source_after", "started_utc", "finished_utc",
        "exit_code", "source_unchanged",
    )
    for key in mirrored:
        require(key in wrapper and key in receipt,
                f"{instrumentation}: validation_gate mirror omits {key}")
        require(wrapper[key] == receipt[key],
                f"{instrumentation}: validation_gate mirror differs for {key}")
    return receipt


def check_gate(root: Path, spec: dict[str, Any], instrumentation: str,
               protocol: dict[str, Any]) -> dict[str, Any]:
    gate_path = bundle_file(root, spec["build_gate"],
                            f"binaries.{instrumentation}.build_gate")
    expected_gate_hash = digest(spec["build_gate_sha256"],
                                f"binaries.{instrumentation}.build_gate_sha256")
    require(sha256_file(gate_path, f"build gate {instrumentation}")[0] ==
            expected_gate_hash,
            f"{instrumentation}: build gate hash differs")
    gate = read_json(gate_path, f"build gate {instrumentation}")
    require(isinstance(gate, dict), f"{instrumentation}: build gate is malformed")
    integer(gate.get("exit_code"), f"build gate {instrumentation}.exit_code", 0)
    require(gate.get("source_unchanged") is True,
            f"{instrumentation}: build gate source changed")
    if "clean_after" in gate:
        require(gate["clean_after"] is True,
                f"{instrumentation}: build gate did not finish clean")
    require("started_utc" in gate and "finished_utc" in gate,
            f"{instrumentation}: final build timestamps are missing")
    interval(gate["started_utc"], gate["finished_utc"],
             f"build gate {instrumentation}")
    argv = gate.get("argv")
    require(isinstance(argv, list) and all(isinstance(item, str) for item in argv),
            f"{instrumentation}: build gate command is missing")
    expected_argv = [
        "cargo", "build", "--release", "--locked", "--manifest-path",
        "tools/perf-baseline/Cargo.toml", "--bin", "zip_directory_spool",
    ]
    if instrumentation == "allocator":
        expected_argv += ["--features", "allocator-metrics"]
    require(argv == expected_argv,
            f"{instrumentation}: build gate argv differs from the frozen command")
    if "cwd" in gate:
        absolute_path(gate["cwd"], f"build gate {instrumentation}.cwd")
    if "environment" in gate:
        require(gate["environment"] == protocol["environment"],
                f"{instrumentation}: build gate environment differs from protocol")
    driver_hash = digest(gate.get("driver_sha256"),
                         f"{instrumentation}.build_gate.driver_sha256")
    common_hash = digest(gate.get("common_sha256"),
                        f"{instrumentation}.build_gate.common_sha256")
    require(driver_hash == sha256_file(root / "gate.py", "gate.py")[0],
            f"{instrumentation}: build gate driver hash differs")
    require(common_hash == sha256_file(root / "common.py", "common.py")[0],
            f"{instrumentation}: build gate common helper hash differs")
    check_wrapper_binary_metadata(gate, spec, instrumentation)
    require(binary_binding_present(gate, spec),
            f"{instrumentation}: build gate does not bind the retained binary")
    require(binary_path_binding_present(gate, spec),
            f"{instrumentation}: build gate does not bind the copied binary path")

    check_validation_gate_reference(root, gate, instrumentation)

    manifests: list[str] = []
    seen_manifest_refs: set[tuple[str, str]] = set()
    for key, value in _walk(gate):
        if key not in {"source_manifest", "source_before", "source_after"} \
                or not isinstance(value, dict) \
                or "path" not in value or "sha256" not in value:
            continue
        reference_key = (str(value["path"]), str(value["sha256"]))
        if reference_key in seen_manifest_refs:
            continue
        seen_manifest_refs.add(reference_key)
        manifests.append(check_source_manifest(
            root, value, f"build gate {instrumentation}.{key}"))
    require(manifests, f"{instrumentation}: build gate has no source-manifest binding")
    before_hash = check_source_manifest(
        root, gate["source_before"], f"build gate {instrumentation}.source_before")
    after_hash = check_source_manifest(
        root, gate["source_after"], f"build gate {instrumentation}.source_after")
    require(before_hash == after_hash,
            f"{instrumentation}: final build source snapshots differ")
    require(digest(spec["source_manifest_sha256"],
                   f"binaries.{instrumentation}.source_manifest_sha256") == before_hash,
            f"{instrumentation}: binary source-manifest binding differs")
    return gate


def check_binaries(root: Path, protocol: dict[str, Any]) -> dict[str, dict[str, Any]]:
    path = root / "binaries.json"
    value = read_json(path, "binaries.json")
    require(isinstance(value, dict), "binaries.json: expected object")
    require(set(value) == set(INSTRUMENTATION_VALUES),
            "binaries.json: expected normal and allocator identities")
    result: dict[str, dict[str, Any]] = {}
    for instrumentation in INSTRUMENTATION_VALUES:
        spec = value[instrumentation]
        require(isinstance(spec, dict), f"binaries.{instrumentation}: expected object")
        for key in ("path", "bytes", "sha256", "build_gate",
                    "build_gate_sha256", "source_manifest_sha256"):
            require(key in spec, f"binaries.{instrumentation}.{key} is missing")
        absolute_path(spec["path"], f"binaries.{instrumentation}.path")
        integer(spec["bytes"], f"binaries.{instrumentation}.bytes")
        require(spec["bytes"] > 0, f"binaries.{instrumentation}.bytes must be positive")
        spec = dict(spec)
        spec["sha256"] = digest(spec["sha256"], f"binaries.{instrumentation}.sha256")
        spec["build_gate_sha256"] = digest(
            spec["build_gate_sha256"],
            f"binaries.{instrumentation}.build_gate_sha256",
        )
        spec["source_manifest_sha256"] = digest(
            spec["source_manifest_sha256"],
            f"binaries.{instrumentation}.source_manifest_sha256",
        )
        safe_relative(spec["build_gate"], f"binaries.{instrumentation}.build_gate")
        if not Path(spec["path"]).is_absolute():
            binary_path = bundle_file(root, spec["path"],
                                      f"binaries.{instrumentation}.path")
            check_metadata(binary_path,
                           {"bytes": spec["bytes"], "sha256": spec["sha256"]},
                           f"binaries.{instrumentation}.path")
        check_gate(root, spec, instrumentation, protocol)
        result[instrumentation] = spec
    return result


def check_artifact_map(root: Path, directory: Path, artifacts: Any,
                       expected_names: set[str], label: str) -> None:
    require(isinstance(artifacts, dict) and set(artifacts) == expected_names,
            f"{label}: artifact set differs")
    for name, expected in artifacts.items():
        require(name == Path(name).name and "\\" not in name,
                f"{label}.{name}: artifact name is not local")
        check_metadata(directory / name, expected, f"{label}/{name}")


def parse_resource(path: Path, label: str) -> int:
    require(path.is_file() and not path.is_symlink(), f"{label}: resource file missing")
    try:
        text = path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as error:
        fail(f"{label}: cannot read resource file: {error}")
    match = re.search(r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$",
                      text, re.MULTILINE)
    require(match is not None, f"{label}: maximum RSS is missing")
    return int(match.group(1))


def check_process_delta(value: Any, label: str) -> None:
    if value is None:
        return
    require(isinstance(value, dict), f"{label}: expected object or null")
    for key in ("rchar", "wchar", "syscr", "syscw", "read_bytes",
                "write_bytes", "minor_faults", "major_faults"):
        require(key in value, f"{label}.{key}: required process counter is missing")
    for key, child in value.items():
        integer(child, f"{label}.{key}")


def check_allocation(value: Any, instrumentation: str, label: str) -> None:
    if instrumentation == "normal":
        require(value is None, f"{label}: normal operation must omit allocator metrics")
        return
    require(isinstance(value, dict), f"{label}: allocator sample is missing")
    require(value.get("status") == "measured",
            f"{label}.status: allocator sample is not measured")
    require(value.get("scope") == "operation_global_system_allocator",
            f"{label}.scope differs")
    for key, child in value.items():
        if key in {"status", "scope"}:
            continue
        integer(child, f"{label}.{key}")
    require(value.get("failed_allocation_calls") == 0,
            f"{label}: failed allocation count is non-zero")
    require(value.get("live_bytes_before") == value.get("live_bytes_after"),
            f"{label}: operation did not return to the same live-byte level")
    require(value.get("region_peak_live_bytes", -1) >=
            value.get("live_bytes_before", 0),
            f"{label}: region peak is below its starting live bytes")


def check_report(path: Path, spec: dict[str, Any], instrumentation: str,
                 source_identities: dict[str, tuple[Any, ...]]) -> dict[str, Any]:
    label = spec["label"]
    report = read_json(path, f"{label}.report.json")
    require(isinstance(report, dict), f"{label}: report must be an object")
    require(report.get("schema") == REPORT_SCHEMA, f"{label}: report schema differs")
    expected_instrumentation = (
        "none" if instrumentation == "normal" else "system_allocator_operation_scoped"
    )
    require(report.get("instrumentation") == expected_instrumentation,
            f"{label}: instrumentation identity differs")
    expected_allocator = (
        "Rust system allocator"
        if instrumentation == "normal"
        else "CountingSystemAllocator(std::alloc::System)"
    )
    require(report.get("allocator") == expected_allocator,
            f"{label}: allocator identity differs")
    require(report.get("samples") == SAMPLES and report.get("warmups") == WARMUPS
            and report.get("repeats") == REPEATS,
            f"{label}: report sample protocol differs")
    require(report.get("counts") == [spec["count"]], f"{label}: counts differ")
    require(report.get("methods") == [spec["method"]], f"{label}: methods differ")
    require(report.get("modes") == [spec["policy"]], f"{label}: modes differ")
    require(report.get("spool_max_bytes") == MAX_SPOOL_BYTES,
            f"{label}: spool maximum differs")
    require(report.get("spool_buffer_bytes") == SPOOL_BUFFER_BYTES,
            f"{label}: spool buffer differs")
    for key in ("timing_scope", "oracle_scope", "cleanup_scope",
                "control_storage_policy", "storage_capability"):
        require(isinstance(report.get(key), str), f"{label}: {key} is missing")
    require(all(phrase in report["timing_scope"] for phrase in
                ("File creation/open", "complete writer finish", "flush", "close")),
            f"{label}: timing scope omits part of the complete operation")
    require("process and allocator endpoint snapshots" in report["timing_scope"] and
            "outside" in report["timing_scope"],
            f"{label}: endpoint and whole-process peak boundaries are unclear")
    require("every generated member" in report["oracle_scope"] and
            "decompressed" in report["oracle_scope"],
            f"{label}: oracle scope does not prove full-member validation")
    require("unlink" in report["cleanup_scope"] and
            "excluded" in report["cleanup_scope"],
            f"{label}: cleanup scope is not explicit")
    require("in-memory" in report["control_storage_policy"] and
            "Vec<FileHeader>" in report["control_storage_policy"],
            f"{label}: control memory policy is not explicit")
    require("explicit caller-owned" in report["storage_capability"] and
            "std::fs::File" in report["storage_capability"] and
            "no sync_all" in report["storage_capability"],
            f"{label}: storage capability is not explicit")

    cases = report.get("cases")
    require(isinstance(cases, list) and len(cases) == 2,
            f"{label}: expected control and spool oracle rows")
    case_identity: tuple[int, str, str, int, str] | None = None
    case_projection: tuple[int, str] | None = None
    for index, case in enumerate(cases):
        require(isinstance(case, dict), f"{label}.cases[{index}]: expected object")
        require(case.get("mode") in POLICY_VALUES and
                case.get("method") == spec["method"] and
                case.get("member_count") == spec["count"],
                f"{label}.cases[{index}]: identity differs")
        integer(case.get("output_bytes"), f"{label}.cases[{index}].output_bytes")
        output_hash = digest(case.get("output_sha256"),
                             f"{label}.cases[{index}].output_sha256")
        require(case.get("every_member_reopened") is True,
                f"{label}.cases[{index}]: member reopen oracle is absent")
        require(case.get("byte_exact_control_match") is True,
                f"{label}.cases[{index}]: byte oracle is absent")
        current = (case["output_bytes"], output_hash)
        require(case_projection is None or case_projection == current,
                f"{label}: control and spool oracle bytes differ")
        case_projection = current
        case_identity = (case["member_count"], case["method"], case["mode"],
                         case["output_bytes"], output_hash)
    require({case["mode"] for case in cases} == set(POLICY_VALUES),
            f"{label}: oracle policy set differs")

    operations = report.get("operations")
    require(isinstance(operations, list) and len(operations) == SAMPLES,
            f"{label}: expected {SAMPLES} measured operations")
    require({row.get("sample") for row in operations} == set(range(SAMPLES)),
            f"{label}: measured sample indices are incomplete")
    for index, row in enumerate(operations):
        require(isinstance(row, dict), f"{label}.operations[{index}]: expected object")
        require(row.get("mode") == spec["policy"] and row.get("method") == spec["method"]
                and row.get("member_count") == spec["count"] and row.get("repeat") == 0,
                f"{label}.operations[{index}]: identity differs")
        integer(row.get("sample"), f"{label}.operations[{index}].sample")
        integer(row.get("elapsed_ns"), f"{label}.operations[{index}].elapsed_ns")
        integer(row.get("output_bytes"), f"{label}.operations[{index}].output_bytes")
        integer(row.get("output_write_calls"),
                f"{label}.operations[{index}].output_write_calls")
        output_hash = digest(row.get("output_sha256"),
                             f"{label}.operations[{index}].output_sha256")
        require(row.get("output_matches_oracle") is True,
                f"{label}.operations[{index}]: output oracle failed")
        require((row["output_bytes"], output_hash) == case_projection,
                f"{label}.operations[{index}]: output differs from oracle")
        require(row.get("source_bytes") == spec["count"] * PAYLOAD_BYTES,
                f"{label}.operations[{index}]: source byte count differs")
        source_hash = digest(row.get("source_sha256"),
                             f"{label}.operations[{index}].source_sha256")
        identity = (source_hash, row["source_bytes"], row["output_bytes"], output_hash)
        identity_key = f"{spec['count']}-{spec['method']}"
        require(identity_key not in source_identities or
                source_identities[identity_key] == identity,
                f"{label}: source/output identity drifted across captures")
        source_identities[identity_key] = identity
        if spec["policy"] == "control":
            require(row.get("spool_bytes") is None,
                    f"{label}.operations[{index}]: control has spool bytes")
        else:
            expected_spool = spec["count"] * (CENTRAL_FIXED_BYTES + len(SLIDE_NAME.encode()))
            require(row.get("spool_bytes") == expected_spool,
                    f"{label}.operations[{index}]: spool extent differs")
        check_allocation(row.get("allocation"), instrumentation,
                         f"{label}.operations[{index}].allocation")
        process_label = f"{label}.operations[{index}].process"
        require(row.get("process") is not None,
                f"{process_label}: process I/O observer is missing")
        check_process_delta(row["process"], process_label)
    return report


def check_captures(root: Path, protocol: dict[str, Any], protocol_hash: str,
                   binaries: dict[str, dict[str, Any]]) -> tuple[dict[str, dict[str, Any]], dict[str, int]]:
    captures_dir = root / "captures"
    require(captures_dir.is_dir() and not captures_dir.is_symlink(),
            "captures directory is missing")
    reports: dict[str, dict[str, Any]] = {}
    rss: dict[str, int] = {}
    identities: dict[str, tuple[Any, ...]] = {}
    for spec in protocol["captures"]:
        label = spec["label"]
        started_path = captures_dir / f"{label}.started.json"
        receipt_path = captures_dir / f"{label}.json"
        started = read_json(started_path, f"captures/{label}.started.json")
        receipt = read_json(receipt_path, f"captures/{label}.json")
        require(started.get("capture") == spec,
                f"{label}: started capture plan differs")
        require(started.get("protocol_sha256") == protocol_hash,
                f"{label}: protocol binding differs")
        require(started.get("binary") == binaries[spec["instrumentation"]],
                f"{label}: binary binding differs")
        require(spec["argv"][7] == binaries[spec["instrumentation"]]["path"],
                f"{label}: command binary path differs from binary identity")
        require(isinstance(started.get("cwd"), str) and
                Path(started["cwd"]).is_absolute(), f"{label}: cwd is not absolute")
        require(started.get("environment") == protocol["environment"],
                f"{label}: capture environment differs from protocol")
        require(started.get("started_utc"), f"{label}: start timestamp is missing")
        check_command_argv(spec["argv"], spec, f"captures/{label}")

        for key in ("capture", "cwd", "environment", "protocol_sha256", "binary"):
            require(receipt.get(key) == started.get(key),
                    f"{label}: started/final {key} differs")
        require(receipt.get("exit_code") == 0, f"{label}: capture exit code is not zero")
        interval(started.get("started_utc"), receipt.get("finished_utc"),
                 f"capture {label}")
        expected_artifacts = {
            f"{label}.stdout", f"{label}.stderr", f"{label}.resource",
            f"{label}.report.json",
        }
        check_artifact_map(root, captures_dir, receipt.get("artifacts"),
                           expected_artifacts, f"capture {label}.artifacts")
        rss[label] = parse_resource(captures_dir / f"{label}.resource",
                                    f"capture {label}")
        reports[label] = check_report(
            captures_dir / f"{label}.report.json", spec,
            spec["instrumentation"], identities,
        )
    return reports, rss


def check_validation_receipts(root: Path, protocol: dict[str, Any]) -> dict[str, int]:
    """Authenticate every retained validation attempt, including failures.

    Validation receipts are evidence of what was attempted.  A non-zero exit
    status is therefore retained and checked rather than treated as a verifier
    failure or silently replaced by a later successful attempt.
    """
    directory = root / "validation"
    require(directory.is_dir() and not directory.is_symlink(),
            "validation directory is missing")
    finished_paths = sorted(
        path for path in directory.glob("*.json")
        if not path.name.endswith(".started.json")
    )
    require(finished_paths, "validation has no finished receipts")
    started_names = {f"{path.stem}.started.json" for path in finished_paths}
    actual_started = {path.name for path in directory.glob("*.started.json")}
    require(actual_started == started_names,
            "validation started/final receipt sets differ")
    counts = {"total": 0, "successful": 0, "failed": 0}
    for finished_path in finished_paths:
        label = finished_path.stem
        started_path = directory / f"{label}.started.json"
        started = read_json(started_path, f"validation/{started_path.name}")
        finished = read_json(finished_path, f"validation/{finished_path.name}")
        required_started = {
            "argv", "common_sha256", "cwd", "driver_sha256", "environment",
            "source_before", "started_utc",
        }
        require(required_started <= set(started),
                f"validation/{label}.started.json: required fields are missing")
        required_finished = required_started | {
            "artifacts", "exit_code", "finished_utc", "source_after",
            "source_unchanged",
        }
        require(required_finished <= set(finished),
                f"validation/{label}.json: required fields are missing")
        for key in required_started:
            require(finished.get(key) == started.get(key),
                    f"validation/{label}: started/final {key} differs")
        require(isinstance(started["argv"], list) and started["argv"] and
                all(isinstance(item, str) for item in started["argv"]),
                f"validation/{label}: argv is malformed")
        absolute_path(started["cwd"], f"validation/{label}.cwd")
        driver_hash = digest(started["driver_sha256"],
                             f"validation/{label}.driver_sha256")
        common_hash = digest(started["common_sha256"],
                             f"validation/{label}.common_sha256")
        require(driver_hash == sha256_file(root / "gate.py", "gate.py")[0],
                f"validation/{label}: gate driver hash differs")
        require(common_hash == sha256_file(root / "common.py", "common.py")[0],
                f"validation/{label}: common helper hash differs")
        require(started["environment"] == protocol["environment"],
                f"validation/{label}: environment differs from protocol")
        before_hash = check_source_manifest(
            root, started["source_before"], f"validation/{label}.source_before")
        after_hash = check_source_manifest(
            root, finished["source_after"], f"validation/{label}.source_after")
        require(finished["source_unchanged"] ==
                (before_hash == after_hash),
                f"validation/{label}: source_unchanged is inconsistent")
        require(isinstance(finished["exit_code"], int) and
                not isinstance(finished["exit_code"], bool),
                f"validation/{label}.exit_code: expected integer status")
        interval(started["started_utc"], finished["finished_utc"],
                 f"validation/{label}")
        expected_artifacts = {f"{label}.stdout", f"{label}.stderr"}
        check_artifact_map(root, directory, finished["artifacts"],
                           expected_artifacts, f"validation/{label}.artifacts")
        counts["total"] += 1
        if finished["exit_code"] == 0:
            counts["successful"] += 1
        else:
            counts["failed"] += 1
    require(counts["failed"] > 0,
            "validation bundle does not retain any non-zero development attempt")
    return counts


def check_rust_validation(root: Path, binaries: dict[str, dict[str, Any]]) -> dict[str, int]:
    """Bind required final Rust gates to the complete validation ledger."""
    path = root / "rust-validation.json"
    ledger = read_json(path, "rust-validation.json")
    require(isinstance(ledger, dict), "rust-validation.json: expected an object")
    if "schema" in ledger:
        require(ledger["schema"] == "zip-directory-spool-rust-validation-v1",
                "rust-validation.json: schema differs")
    required = ledger.get("required")
    require(isinstance(required, list) and required and
            all(isinstance(label, str) and label for label in required),
            "rust-validation.required: expected non-empty labels")
    require(len(set(required)) == len(required),
            "rust-validation.required: duplicate labels")
    for label in required:
        require(re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]*", label) is not None,
                f"rust-validation.required: unsafe label {label!r}")
    final_source = digest(ledger.get("final_source_sha256"),
                          "rust-validation.final_source_sha256")
    missing_required = set(REQUIRED_FINAL_VALIDATION_LABELS) - set(required)
    require(not missing_required,
            "rust-validation.required: mandatory final gate is missing: "
            + ", ".join(sorted(missing_required)))
    require(final_source == binaries["normal"]["source_manifest_sha256"] ==
            binaries["allocator"]["source_manifest_sha256"],
            "rust-validation.final_source_sha256 does not match build gates")
    final_gate_path = bundle_file(
        root, binaries["normal"]["build_gate"],
        "rust-validation.final_source.build_gate")
    final_gate = read_json(final_gate_path, "rust-validation.final_source.build_gate")
    require(isinstance(final_gate, dict),
            "rust-validation.final_source.build_gate: expected object")
    final_manifest_hash, final_manifest = read_source_manifest(
        root, final_gate.get("source_after"),
        "rust-validation.final_source_manifest")
    require(final_manifest_hash == final_source,
            "rust-validation.final_source_manifest: hash differs")

    attempts = ledger.get("attempts")
    require(isinstance(attempts, dict) and attempts,
            "rust-validation.attempts: expected non-empty object")
    finished_paths = sorted(
        path for path in (root / "validation").glob("*.json")
        if not path.name.endswith(".started.json")
    )
    actual_labels = {path.stem for path in finished_paths}
    require(set(attempts) == actual_labels,
            "rust-validation.attempts: ledger does not cover the exact receipt inventory")
    required_set = set(required)
    require(required_set <= actual_labels,
            "rust-validation.required: a required gate is absent from attempts")

    expected_argv_map = ledger.get("expected_argv", {})
    require(isinstance(expected_argv_map, dict),
            "rust-validation.expected_argv: expected object")
    require(set(expected_argv_map) <= set(attempts),
            "rust-validation.expected_argv: contains an unknown receipt label")
    results = {"required_success": len(required_set), "receipts": len(attempts),
               "retained_nonzero_attempts": 0}
    for label, expected in attempts.items():
        require(re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]*", label) is not None,
                f"rust-validation.attempts: unsafe label {label!r}")
        require(isinstance(expected, dict),
                f"rust-validation.attempts.{label}: expected object")
        for key in ("path", "sha256", "exit_code", "source_unchanged"):
            require(key in expected,
                    f"rust-validation.attempts.{label}.{key}: missing")
        relative = safe_relative(expected["path"],
                                 f"rust-validation.attempts.{label}.path")
        require(relative == f"validation/{label}.json",
                f"rust-validation.attempts.{label}.path: does not bind its label")
        receipt_path = bundle_file(
            root, relative, f"rust-validation.attempts.{label}.path")
        require(sha256_file(receipt_path,
                            f"rust-validation.attempts.{label}")[0] ==
                digest(expected["sha256"],
                       f"rust-validation.attempts.{label}.sha256"),
                f"rust-validation.attempts.{label}: receipt hash differs")
        receipt = read_json(receipt_path,
                            f"rust-validation.attempts.{label}.receipt")
        require(isinstance(receipt, dict),
                f"rust-validation.attempts.{label}: receipt is malformed")
        require(receipt.get("exit_code") == expected["exit_code"],
                f"rust-validation.attempts.{label}: exit code differs")
        require(receipt.get("source_unchanged") == expected["source_unchanged"],
                f"rust-validation.attempts.{label}: source status differs")
        require(isinstance(expected["exit_code"], int) and
                not isinstance(expected["exit_code"], bool),
                f"rust-validation.attempts.{label}.exit_code: expected integer")
        require(isinstance(expected["source_unchanged"], bool),
                f"rust-validation.attempts.{label}.source_unchanged: expected boolean")

        entry_argv = expected.get("argv", expected_argv_map.get(label))
        if label in expected_argv_map and "argv" in expected:
            require(expected["argv"] == expected_argv_map[label],
                    f"rust-validation.attempts.{label}: argv sources differ")
        if label in required_set:
            require(isinstance(entry_argv, list) and entry_argv and
                    all(isinstance(item, str) for item in entry_argv),
                    f"rust-validation.attempts.{label}: exact expected argv is missing")
            require(receipt.get("argv") == entry_argv,
                    f"rust-validation.attempts.{label}: argv differs")
            require(expected["exit_code"] == 0,
                    f"rust-validation.attempts.{label}: required gate did not succeed")
            require(expected["source_unchanged"] is True,
                    f"rust-validation.attempts.{label}: required gate changed source")
            exclusions = expected.get("source_exclusions", [])
            require(isinstance(exclusions, list) and
                    len(set(exclusions)) == len(exclusions) and
                    all(isinstance(item, str) for item in exclusions),
                    f"rust-validation.attempts.{label}.source_exclusions: malformed")
            require(set(exclusions) <= {
                "tools/perf-baseline/src/zip_directory_spool.rs"
            },
                    f"rust-validation.attempts.{label}.source_exclusions: path is not approved")
            if exclusions:
                require(label not in REQUIRED_FINAL_NO_SOURCE_EXCLUSIONS,
                        f"rust-validation.attempts.{label}: source exclusions are not allowed for this final gate")
            before_hash, before_manifest = read_source_manifest(
                root, receipt.get("source_before"),
                f"rust-validation.attempts.{label}.source_before")
            after_hash, after_manifest = read_source_manifest(
                root, receipt.get("source_after"),
                f"rust-validation.attempts.{label}.source_after")
            require(before_hash == after_hash,
                    f"rust-validation.attempts.{label}: source snapshots differ")
            differences = {
                name for name in set(before_manifest) | set(final_manifest)
                if before_manifest.get(name) != final_manifest.get(name)
            }
            require(differences == set(exclusions),
                    f"rust-validation.attempts.{label}: source differences are not the declared projection")
            for snapshot_name, snapshot in (("before", before_manifest),
                                            ("after", after_manifest)):
                projected = {
                    name: value for name, value in snapshot.items()
                    if name not in exclusions
                }
                final_projected = {
                    name: value for name, value in final_manifest.items()
                    if name not in exclusions
                }
                require(projected == final_projected,
                        f"rust-validation.attempts.{label}: {snapshot_name} source is not final under projection")
        else:
            require("source_exclusions" not in expected,
                    f"rust-validation.attempts.{label}: source exclusions require a required gate")
        if label not in required_set and expected["exit_code"] != 0:
            results["retained_nonzero_attempts"] += 1
            if entry_argv is not None:
                require(isinstance(entry_argv, list) and entry_argv and
                        all(isinstance(item, str) for item in entry_argv),
                        f"rust-validation.attempts.{label}.argv: malformed")
                require(receipt.get("argv") == entry_argv,
                        f"rust-validation.attempts.{label}: argv differs")

    require(results["retained_nonzero_attempts"] > 0,
            "rust-validation.attempts: no retained non-zero development attempt")
    return results


def stats(values: list[int]) -> dict[str, Any]:
    require(values and all(type(value) is int and value >= 0 for value in values),
            "summary statistic input is invalid")
    ordered = sorted(values)
    mean = statistics.mean(values)
    half = 2.045 * statistics.stdev(values) / math.sqrt(len(values)) \
        if len(values) > 1 else 0
    return {
        "n": len(values),
        "mean": mean,
        "minimum": min(values),
        "maximum": max(values),
        "p50": ordered[math.ceil(len(values) * 0.5) - 1],
        "p95": ordered[math.ceil(len(values) * 0.95) - 1],
        "p99": ordered[math.ceil(len(values) * 0.99) - 1],
        "mean_t95_interval": [mean - half, mean + half],
    }


def percent(before: float, after: float) -> float | None:
    return 100 * (after - before) / before if before else None


def compare(expected: Any, actual: Any, label: str) -> None:
    if isinstance(expected, bool) or isinstance(actual, bool):
        require(expected == actual, f"{label}: value differs")
    elif isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        require(math.isclose(float(expected), float(actual), rel_tol=1e-12,
                             abs_tol=1e-9), f"{label}: numeric value differs")
    elif isinstance(expected, dict) and isinstance(actual, dict):
        require(set(expected) == set(actual),
                f"{label}: object fields differ")
        for key in expected:
            compare(expected[key], actual[key], f"{label}.{key}")
    elif isinstance(expected, list) and isinstance(actual, list):
        require(len(expected) == len(actual), f"{label}: list length differs")
        for index, (left, right) in enumerate(zip(expected, actual)):
            compare(left, right, f"{label}[{index}]")
    else:
        require(expected == actual, f"{label}: value differs")


def derive_summary(protocol: dict[str, Any], reports: dict[str, dict[str, Any]],
                   rss: dict[str, int]) -> dict[str, Any]:
    rows: dict[str, Any] = {}
    for spec in protocol["captures"]:
        label = spec["label"]
        report = reports[label]
        operations = report["operations"]
        projection = (report["cases"][0]["output_bytes"],
                      report["cases"][0]["output_sha256"])
        row: dict[str, Any] = {
            "capture": spec,
            "normal_timing": spec["instrumentation"] == "normal",
            "elapsed_ns": stats([item["elapsed_ns"] for item in operations]),
            "output_bytes": projection[0],
            "output_sha256": projection[1],
            "output_write_calls": stats([item["output_write_calls"] for item in operations]),
            "process_max_rss_kib": rss[label],
        }
        processes = [item["process"] for item in operations]
        if all(process is not None for process in processes):
            row["process_observer_deltas"] = {
                metric: stats([process[metric] for process in processes])
                for metric in ("rchar", "wchar", "syscr", "syscw",
                               "read_bytes", "write_bytes", "minor_faults",
                               "major_faults")
            }
        if spec["instrumentation"] == "allocator":
            samples = [item["allocation"] for item in operations]
            row["allocation_calls"] = stats([item["allocation_calls"] for item in samples])
            row["requested_bytes"] = stats([item["allocated_bytes"] for item in samples])
            row["incremental_peak_live_bytes"] = stats([
                item["region_peak_live_bytes"] - item["live_bytes_before"]
                for item in samples
            ])
        rows[label] = row

    pairs: list[dict[str, Any]] = []
    drifts: list[dict[str, Any]] = []
    for instrumentation in INSTRUMENTATION_VALUES:
        for count in COUNT_VALUES:
            for method in METHOD_VALUES:
                for repeat in (1, 2):
                    before = rows[f"r{repeat}-{instrumentation}-{count}-{method}-control"]
                    after = rows[f"r{repeat}-{instrumentation}-{count}-{method}-spool"]
                    changes: dict[str, Any] = {
                        "process_max_rss_kib": percent(
                            before["process_max_rss_kib"], after["process_max_rss_kib"]),
                    }
                    if instrumentation == "normal":
                        changes.update({
                            f"elapsed_{metric}": percent(
                                before["elapsed_ns"][metric], after["elapsed_ns"][metric])
                            for metric in ("mean", "p50", "p95", "p99")
                        })
                    else:
                        changes.update({
                            metric: percent(before[metric]["mean"], after[metric]["mean"])
                            for metric in ("allocation_calls", "requested_bytes",
                                           "incremental_peak_live_bytes")
                        })
                    pairs.append({
                        "instrumentation": instrumentation,
                        "count": count,
                        "method": method,
                        "repeat": repeat,
                        "percent_changes": changes,
                        "positive_review_flags": [
                            key for key, value in changes.items()
                            if value is not None and value > 5
                        ],
                    })
                for policy in POLICY_VALUES:
                    before = rows[f"r1-{instrumentation}-{count}-{method}-{policy}"]
                    after = rows[f"r2-{instrumentation}-{count}-{method}-{policy}"]
                    changes = {
                        "process_max_rss_kib": percent(
                            before["process_max_rss_kib"], after["process_max_rss_kib"]),
                    }
                    if instrumentation == "normal":
                        changes.update({
                            f"elapsed_{metric}": percent(
                                before["elapsed_ns"][metric], after["elapsed_ns"][metric])
                            for metric in ("mean", "p50", "p95", "p99")
                        })
                    drifts.append({
                        "instrumentation": instrumentation,
                        "count": count,
                        "method": method,
                        "policy": policy,
                        "percent_changes": changes,
                        "absolute_review_flags": [
                            key for key, value in changes.items()
                            if value is not None and abs(value) > 5
                        ],
                    })
    return {
        "schema": SUMMARY_SCHEMA,
        "captures": len(rows),
        "samples": len(rows) * SAMPLES,
        "rows": rows,
        "identities": {
            key: [value for value in identity]
            for key, identity in _identities_from_reports(protocol, reports).items()
        },
        "pairs": pairs,
        "repeat_drifts": drifts,
        "uncertainty": "nearest-rank percentiles; mean interval uses t(29)=2.045; independent process repeats retained, no registered latency claim",
    }


def _identities_from_reports(protocol: dict[str, Any],
                             reports: dict[str, dict[str, Any]]) -> dict[str, tuple[Any, ...]]:
    identities: dict[str, tuple[Any, ...]] = {}
    for spec in protocol["captures"]:
        report = reports[spec["label"]]
        row = report["operations"][0]
        key = f"{spec['count']}-{spec['method']}"
        identity = (row["source_sha256"], row["source_bytes"],
                    report["cases"][0]["output_bytes"],
                    report["cases"][0]["output_sha256"])
        require(key not in identities or identities[key] == identity,
                f"{key}: report identity differs across captures")
        identities[key] = identity
    return identities


def check_summary(root: Path, protocol: dict[str, Any],
                  reports: dict[str, dict[str, Any]], rss: dict[str, int]) -> None:
    path = root / "summary.json"
    summary = read_json(path, "summary.json")
    expected = derive_summary(protocol, reports, rss)
    require(isinstance(summary, dict), "summary.json: expected object")
    compare(expected, summary, "summary")


def check_seal(root: Path) -> None:
    path = root / "SHA256SUMS"
    if not path.exists():
        return
    require(path.is_file() and not path.is_symlink(),
            "SHA256SUMS: regular file required")
    entries: dict[str, str] = {}
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as error:
        fail(f"SHA256SUMS: cannot read: {error}")
    for number, line in enumerate(lines, 1):
        if not line.strip():
            continue
        fields = line.split(maxsplit=1)
        require(len(fields) == 2, f"SHA256SUMS:{number}: malformed line")
        value = digest(fields[0], f"SHA256SUMS:{number}.sha256")
        name = fields[1][1:] if fields[1].startswith("*") else fields[1]
        name = safe_relative(name, f"SHA256SUMS:{number}.path")
        require(name not in entries and name != "SHA256SUMS",
                f"SHA256SUMS:{number}: duplicate or self-reference")
        entries[name] = value
    require(entries, "SHA256SUMS: no entries")
    for name, expected in entries.items():
        actual = sha256_file(bundle_file(root, name, f"SHA256SUMS.{name}"),
                             f"SHA256SUMS.{name}")[0]
        require(actual == expected, f"SHA256SUMS: digest differs for {name}")
    actual_files: set[str] = set()
    for candidate in root.rglob("*"):
        relative = candidate.relative_to(root).as_posix()
        require("__pycache__" not in candidate.parts,
                f"SHA256SUMS: __pycache__ is not allowed: {relative}")
        if candidate.is_symlink():
            fail(f"SHA256SUMS: symlink is not allowed: {relative}")
        require(candidate.is_dir() or candidate.is_file(),
                f"SHA256SUMS: non-regular bundle entry is not allowed: {relative}")
        if candidate.is_file() and relative != "SHA256SUMS":
            actual_files.add(relative)
    require(actual_files == set(entries),
            "SHA256SUMS: regular-file inventory differs from the retained bundle")
    for required in ("protocol.json", "binaries.json", "summary.json"):
        require(required in entries, f"SHA256SUMS: required entry missing: {required}")


def verify_bundle(root: Path = ROOT) -> dict[str, Any]:
    root = Path(root).resolve()
    require(root.is_dir() and not root.is_symlink(), "evidence root is not a directory")
    protocol, protocol_hash = check_protocol(root)
    binaries = check_binaries(root, protocol)
    reports, rss = check_captures(root, protocol, protocol_hash, binaries)
    validation = check_validation_receipts(root, protocol)
    rust_validation = check_rust_validation(root, binaries)
    check_summary(root, protocol, reports, rss)
    check_seal(root)
    return {
        "schema": "zip-directory-spool-verification-v1",
        "protocol_sha256": protocol_hash,
        "captures": len(reports),
        "samples": len(reports) * SAMPLES,
        "validation": validation,
        "rust_validation": rust_validation,
        "sealed": (root / "SHA256SUMS").is_file(),
        "status": "pass",
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT,
                        help="evidence bundle directory (default: script directory)")
    args = parser.parse_args(argv)
    try:
        result = verify_bundle(args.root)
    except VerificationError as error:
        print(f"verify.py: FAIL: {error}", file=sys.stderr)
        return 1
    print(json.dumps(result, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
