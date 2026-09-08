#!/usr/bin/env python3
"""Portable, fail-closed verification for the 0478 evidence bundle.

The default invocation authenticates the frozen protocol, copied binaries,
build and validation custody, every capture receipt, every producer report,
the archived preliminary run, the derived summary, and the final file seal.
``--data-only`` is intended for
the interval after captures and before the Rust validation ledger exists: it
checks only the retained protocol, capture custody, reports, and summary.
No mode follows an absolute path from a receipt into the source checkout or a
temporary directory.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import re
import sys
from pathlib import Path, PurePosixPath
from typing import Any, Iterable, Mapping, NoReturn

import analyze


ROOT = Path(__file__).resolve().parent
SHA256 = re.compile(r"^[0-9a-fA-F]{64}$")
REVISION = re.compile(r"^[0-9a-fA-F]{7,64}$")

# The final validation set is part of the evidence contract.  The ledger may
# retain additional development attempts, but it cannot redefine which gates
# constitute the final result.
REQUIRED_FINAL_LABELS = (
    "libraries-tests-opc-guard",
    "libraries-clippy-opc-guard",
    "libraries-rustdoc-opc-guard",
    "libraries-format-opc-guard",
    "shared-streaming-unit",
    "docx-streaming",
    "odf-streaming",
    "odt-streaming",
    "ods-streaming",
    "odp-streaming",
    "harness-tests",
    "harness-clippy-opc-guard",
    "harness-format-opc-guard",
    "boundaries",
    "registry-strict",
    "build-normal-opc-guard",
    "build-allocator-opc-guard",
    "pilot-normal-opc-guard",
    "pilot-allocator-opc-guard",
    "analyze",
    "evidence-tests",
)

# The embedded presentation and notes templates are part of the measured
# input.  Keep the source inventory fixed here so a producer cannot silently
# omit an include_str! asset while still presenting a self-consistent
# manifest.  The destination names are also fixed: the verifier is portable
# and must not consult the source checkout to infer them.
EMBEDDED_INPUT_SCHEMA = "pptx-embedded-inputs-v1"
EMBEDDED_INPUT_DESTINATIONS = {
    "crates/litchi-pptx/src/notes/resources/generated/notesMaster.xml":
        "embedded-resources/notes/notesMaster.xml",
    "crates/litchi-pptx/src/resources/generated/docProps/app.xml":
        "embedded-resources/generated/docProps/app.xml",
    "crates/litchi-pptx/src/resources/generated/docProps/core.xml":
        "embedded-resources/generated/docProps/core.xml",
    "crates/litchi-pptx/src/resources/generated/presProps.xml":
        "embedded-resources/generated/presProps.xml",
    "crates/litchi-pptx/src/resources/generated/presentation.xml":
        "embedded-resources/generated/presentation.xml",
    **{
        f"crates/litchi-pptx/src/resources/generated/slideLayouts/slideLayout{number}.xml":
        f"embedded-resources/generated/slideLayouts/slideLayout{number}.xml"
        for number in range(1, 12)
    },
    "crates/litchi-pptx/src/resources/generated/slideMasters/slideMaster1.xml":
        "embedded-resources/generated/slideMasters/slideMaster1.xml",
    "crates/litchi-pptx/src/resources/generated/theme/theme1.xml":
        "embedded-resources/generated/theme/theme1.xml",
    "crates/litchi-pptx/src/resources/generated/viewProps.xml":
        "embedded-resources/generated/viewProps.xml",
}
if len(EMBEDDED_INPUT_DESTINATIONS) != 19:
    raise RuntimeError("internal embedded input inventory must contain 19 resources")

HISTORICAL_SOURCE_EXCLUSIONS = (
    "tools/perf-baseline/src/lib.rs",
    "tools/perf-baseline/src/pptx_metadata_spool.rs",
    "tools/perf-baseline/src/zip_directory_spool.rs",
)

ENVIRONMENT_ARTIFACT_SHA256 = (
    "27469836a0b72136b4ad7ce0f85121a2dee3036b066b8c1ac28a3c7103ee4289"
)
ENVIRONMENT_TMPFS_PATH = "/tmp/litchi-goal-0478"
PRELIMINARY_SCHEMA = "pptx-metadata-spool-preliminary-v1"
PRELIMINARY_SCHEMA_V2 = "pptx-metadata-spool-preliminary-v2"
PRELIMINARY_MANIFEST_PATH = "preliminary/manifest.json"
RUST_VALIDATION_SCHEMA = "pptx-metadata-spool-rust-validation-v1"


class VerificationError(ValueError):
    """A retained evidence artifact failed an independent check."""


def fail(message: str) -> NoReturn:
    raise VerificationError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def read_json(path: Path, label: str = "JSON") -> Any:
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


def sha256_file(path: Path, label: str = "file") -> tuple[str, int]:
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


def integer(value: Any, label: str, expected: int | None = None) -> int:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label}: expected a non-negative integer")
    if expected is not None:
        require(value == expected, f"{label}: expected {expected}, got {value}")
    return value


def timestamp(value: Any, label: str) -> _datetime.datetime:
    require(isinstance(value, str) and value, f"{label}: timestamp is missing")
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


def safe_relative(value: Any, label: str) -> str:
    require(isinstance(value, str) and value,
            f"{label}: expected a non-empty relative POSIX path")
    require("\\" not in value, f"{label}: backslashes are not allowed")
    path = PurePosixPath(value)
    require(not path.is_absolute() and value not in {".", ".."},
            f"{label}: absolute or parent path is not allowed")
    require(all(part not in {"", ".", ".."} for part in path.parts),
            f"{label}: path escapes its bundle root")
    return path.as_posix()


def bundle_file(root: Path, value: Any, label: str) -> Path:
    relative = safe_relative(value, label)
    current = root
    for part in PurePosixPath(relative).parts:
        current = current / part
        require(not current.is_symlink(), f"{label}: symlink component is not allowed")
    try:
        resolved = current.resolve(strict=True)
    except OSError as error:
        fail(f"{label}: cannot resolve path: {error}")
    root_resolved = root.resolve()
    require(resolved == root_resolved or root_resolved in resolved.parents,
            f"{label}: path escapes the evidence bundle")
    require(current.is_file(), f"{label}: regular file is missing")
    return current


def bundle_directory(root: Path, value: Any, label: str) -> Path:
    relative = safe_relative(value, label)
    current = root
    for part in PurePosixPath(relative).parts:
        current = current / part
        require(not current.is_symlink(), f"{label}: symlink component is not allowed")
    try:
        resolved = current.resolve(strict=True)
    except OSError as error:
        fail(f"{label}: cannot resolve path: {error}")
    root_resolved = root.resolve()
    require(resolved == root_resolved or root_resolved in resolved.parents,
            f"{label}: path escapes the evidence bundle")
    require(current.is_dir(), f"{label}: directory is missing")
    return current


def absolute_path(value: Any, label: str) -> str:
    require(isinstance(value, str) and value and Path(value).is_absolute(),
            f"{label}: expected an absolute path")
    return value


def metadata(path: Path, label: str) -> dict[str, Any]:
    value, size = sha256_file(path, label)
    return {"bytes": size, "sha256": value}


def check_metadata(path: Path, expected: Any, label: str) -> None:
    require(isinstance(expected, dict) and set(expected) == {"bytes", "sha256"},
            f"{label}: metadata fields differ")
    size = integer(expected.get("bytes"), f"{label}.bytes")
    expected_hash = digest(expected.get("sha256"), f"{label}.sha256")
    require(metadata(path, label) == {"bytes": size, "sha256": expected_hash},
            f"{label}: metadata does not match the retained file")


def _environment(protocol: Mapping[str, Any]) -> dict[str, Any]:
    environment = protocol.get("environment")
    require(isinstance(environment, dict), "protocol.environment is missing")
    for key in (
        "RUSTUP_TOOLCHAIN", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
        "CARGO_PROFILE_RELEASE_DEBUG", "RUSTFLAGS", "DEBUGINFOD_URLS", "LC_ALL",
    ):
        require(key in environment, f"protocol.environment.{key} is missing")
    return dict(environment)


def check_protocol(root: Path) -> tuple[dict[str, Any], str]:
    path = root / "protocol.json"
    protocol = read_json(path, "protocol.json")
    require(isinstance(protocol, dict), "protocol.json: expected an object")
    try:
        rows = analyze.protocol_rows(protocol)
    except analyze.AnalysisError as error:
        fail(str(error))
    _environment(protocol)
    for key in ("scope", "comparison", "memory_gate"):
        require(isinstance(protocol.get(key), str) and protocol[key],
                f"protocol.{key} is missing")
    scripts = protocol.get("scripts")
    require(isinstance(scripts, dict) and scripts,
            "protocol.scripts is missing")
    for name, expected_hash in scripts.items():
        relative = safe_relative(name, f"protocol.scripts[{name!r}]")
        script_path = bundle_file(root, relative, f"protocol.scripts[{name!r}]")
        require(sha256_file(script_path, f"protocol script {name}")[0] ==
                digest(expected_hash, f"protocol.scripts[{name!r}]"),
                f"protocol script hash differs: {name}")
    for spec in rows:
        check_command_argv(spec["argv"], spec, f"protocol.captures[{spec['label']}]", root)
    return protocol, sha256_file(path, "protocol.json")[0]


def check_environment(root: Path, protocol: Mapping[str, Any]) -> dict[str, Any]:
    """Authenticate the recorded tmpfs context used by explicit spools."""
    path = bundle_file(root, "environment.json", "environment.json")
    actual_hash = sha256_file(path, "environment.json")[0]
    require(actual_hash == ENVIRONMENT_ARTIFACT_SHA256,
            "environment.json: retained artifact digest differs")
    value = read_json(path, "environment.json")
    require(isinstance(value, dict), "environment.json: expected an object")
    filesystem = value.get("filesystem")
    require(isinstance(filesystem, dict) and set(filesystem) ==
            {"argv", "exit_code", "stderr", "stdout"},
            "environment.filesystem: retained df record is malformed")
    require(filesystem.get("argv") == ["df", "-T", ENVIRONMENT_TMPFS_PATH],
            "environment.filesystem.argv: scratch path differs")
    require(filesystem.get("exit_code") == 0 and filesystem.get("stderr") == "",
            "environment.filesystem: df did not complete cleanly")
    stdout = filesystem.get("stdout")
    require(isinstance(stdout, str), "environment.filesystem.stdout: expected text")
    lines = stdout.splitlines()
    require(len(lines) == 2, "environment.filesystem.stdout: expected one df row")
    fields = lines[1].split()
    require(len(fields) >= 7 and fields[0] == "tmpfs" and fields[1] == "tmpfs" and
            fields[-1] == "/tmp",
            "environment.filesystem.stdout: /tmp is not authenticated as tmpfs")
    for row in analyze.protocol_rows(protocol):
        spool_path = row["argv"][25]
        require(spool_path.startswith(f"{ENVIRONMENT_TMPFS_PATH}/spools/"),
                f"{row['label']}: spool path is outside retained tmpfs context")
    return {"path": "environment.json", "sha256": actual_hash}


def _check_preliminary_run(
    preliminary_base: Path,
    run: Any,
    index: int,
) -> tuple[str, Path, dict[str, Any], str]:
    """Authenticate one archived run from a preliminary manifest."""
    label = f"preliminary.runs[{index}]"
    require(isinstance(run, dict) and set(run) ==
            {"attempt", "root", "protocol", "files"},
            f"{label}: fields differ")
    require(isinstance(run.get("attempt"), str) and run["attempt"],
            f"{label}.attempt: identity is missing")
    attempt = run["attempt"]
    snapshot_name = safe_relative(run.get("root"), f"{label}.root")
    preliminary_root = bundle_directory(preliminary_base, snapshot_name,
                                         f"{label}.root")
    protocol_reference = run.get("protocol")
    require(isinstance(protocol_reference, dict) and set(protocol_reference) ==
            {"path", "bytes", "sha256"},
            f"{label}.protocol: reference is malformed")
    protocol_relative = safe_relative(protocol_reference["path"],
                                       f"{label}.protocol.path")
    require(protocol_relative == f"{snapshot_name}/protocol.json",
            f"{label}.protocol.path: does not bind snapshot root")
    archived_protocol_path = bundle_file(
        preliminary_base, protocol_relative,
        f"{label}.protocol.path")
    check_metadata(archived_protocol_path,
                   {"bytes": protocol_reference["bytes"],
                    "sha256": protocol_reference["sha256"]},
                   f"{label}.protocol")
    files = run.get("files")
    require(isinstance(files, dict) and files,
            f"{label}.files: non-empty inventory is required")
    for name, expected in files.items():
        relative_file = safe_relative(name, f"{label}.files[{name!r}]")
        require(relative_file.startswith(f"{snapshot_name}/"),
                f"{label}.files[{name!r}]: outside snapshot root")
        path = bundle_file(preliminary_base, relative_file,
                           f"{label}.files[{name!r}]")
        check_metadata(path, expected, f"{label}.files[{name!r}]")
    actual: set[str] = set()
    for candidate in preliminary_root.rglob("*"):
        relative_file = candidate.relative_to(preliminary_base).as_posix()
        require(not candidate.is_symlink(),
                f"preliminary snapshot symlink is present: {relative_file}")
        if candidate.is_file():
            actual.add(relative_file)
    require(actual == set(files),
            f"{label}: snapshot file inventory differs")
    archived_protocol = read_json(
        archived_protocol_path, f"preliminary/{snapshot_name}/protocol.json")
    require(isinstance(archived_protocol, dict),
            f"{label}.protocol: expected an object")
    protocol_hash = sha256_file(archived_protocol_path,
                                f"{label}.protocol")[0]
    return attempt, preliminary_root, archived_protocol, protocol_hash


def check_preliminary_manifest(
    root: Path, protocol: Mapping[str, Any]
) -> list[tuple[str, Path, dict[str, Any], str]]:
    """Authenticate one or more archived runs and return their protocol roots."""
    reference = protocol.get("preliminary")
    require(isinstance(reference, dict) and set(reference) ==
            {"path", "bytes", "sha256"},
            "protocol.preliminary: manifest reference is missing or malformed")
    relative = safe_relative(reference["path"], "protocol.preliminary.path")
    require(relative == PRELIMINARY_MANIFEST_PATH,
            "protocol.preliminary.path: canonical manifest path differs")
    manifest_path = bundle_file(root, relative, "protocol.preliminary.path")
    check_metadata(manifest_path,
                   {"bytes": reference["bytes"], "sha256": reference["sha256"]},
                   "protocol.preliminary")
    manifest = read_json(manifest_path, "preliminary/manifest.json")
    preliminary_base = root / "preliminary"
    require(preliminary_base.is_dir() and not preliminary_base.is_symlink(),
            "preliminary: directory is missing")
    require(isinstance(manifest, dict),
            "preliminary/manifest.json: expected an object")
    schema = manifest.get("schema")
    if schema == PRELIMINARY_SCHEMA:
        require(set(manifest) ==
                {"schema", "attempt", "root", "protocol", "files"},
                "preliminary/manifest.json: fields differ")
        runs: list[Any] = [{
            key: value for key, value in manifest.items() if key != "schema"
        }]
    elif schema == PRELIMINARY_SCHEMA_V2:
        require(set(manifest) == {"schema", "runs"},
                "preliminary/manifest.json: v2 fields differ")
        runs_value = manifest.get("runs")
        require(isinstance(runs_value, list) and runs_value,
                "preliminary/manifest.json.runs: non-empty list is required")
        runs = runs_value
    else:
        fail("preliminary/manifest.json: schema differs")
    roots: list[str] = []
    for index, run in enumerate(runs):
        require(isinstance(run, dict),
                f"preliminary.runs[{index}]: expected an object")
        roots.append(safe_relative(run.get("root"),
                                   f"preliminary.runs[{index}].root"))
    require(len(set(roots)) == len(roots),
            "preliminary/manifest.json: duplicate snapshot roots")
    require({child.name for child in preliminary_base.iterdir()} <=
            {"manifest.json", *roots},
            "preliminary: unbound sibling is present")
    attempts: list[str] = []
    result: list[tuple[str, Path, dict[str, Any], str]] = []
    for index, run in enumerate(runs):
        attempt, preliminary_root, archived_protocol, protocol_hash = (
            _check_preliminary_run(preliminary_base, run, index))
        attempts.append(attempt)
        result.append((attempt, preliminary_root, archived_protocol, protocol_hash))
    require(len(set(attempts)) == len(attempts),
            "preliminary/manifest.json: duplicate attempt identities")
    return result


def check_preliminary(
    root: Path,
    protocol: Mapping[str, Any],
    *,
    _depth: int = 0,
) -> dict[str, Any]:
    require(_depth <= 1,
            "preliminary custody nesting exceeds the two-run limit")
    runs = check_preliminary_manifest(root, protocol)
    results: list[dict[str, Any]] = []
    for attempt, preliminary_root, archived_protocol, protocol_hash in runs:
        if len(runs) > 1 and "preliminary" in archived_protocol:
            fail("preliminary custody exceeds the two-run limit")
        checked_protocol, checked_hash = check_protocol(preliminary_root)
        require(checked_hash == protocol_hash,
                f"preliminary {attempt}: protocol hash changed during validation")
        archived_protocol = checked_protocol
        check_environment(preliminary_root, archived_protocol)
        validation = check_validation_receipts(
            preliminary_root, archived_protocol, allow_incomplete=True)
        reports = check_capture_receipts(
            preliminary_root, archived_protocol, protocol_hash, None)
        result: dict[str, Any] = {
            "attempt": attempt,
            "protocol_sha256": protocol_hash,
            "captures": len(reports),
            "samples": len(reports) * analyze.SAMPLES,
            "validation": validation,
            "status": "authenticated",
        }
        if "preliminary" in archived_protocol:
            require(_depth == 0,
                    f"preliminary {attempt}: nested custody exceeds two runs")
            result["prior"] = check_preliminary(
                preliminary_root, archived_protocol, _depth=_depth + 1)
        results.append(result)
    if len(results) == 1:
        return results[0]
    return {
        "runs": results,
        "captures": sum(item["captures"] for item in results),
        "samples": sum(item["samples"] for item in results),
        "status": "authenticated",
    }


def check_command_argv(argv: Any, spec: Mapping[str, Any], label: str, root: Path) -> None:
    require(isinstance(argv, list) and all(isinstance(value, str) for value in argv),
            f"{label}.argv: expected a string array")
    require(len(argv) == 26, f"{label}.argv: expected 26 arguments")
    fixed = {
        0: "/usr/bin/time", 1: "-v", 2: "-o", 4: "/usr/bin/taskset",
        5: "-c", 6: "2", 8: "--mode", 10: "--counts", 12: "--samples",
        13: "30", 14: "--warmups", 15: "3", 16: "--repeats", 17: "1",
        18: "--max-spool-bytes", 19: str(analyze.MAX_SPOOL_BYTES),
        20: "--spool-buffer-bytes", 21: str(analyze.SPOOL_BUFFER_BYTES),
        22: "--json", 24: "--spool-dir",
    }
    for index, expected in fixed.items():
        require(argv[index] == expected,
                f"{label}.argv[{index}]: expected {expected!r}, got {argv[index]!r}")
    require(argv[9] == spec["policy"], f"{label}: mode is not bound to policy")
    require(argv[11] == str(spec["count"]), f"{label}: count is not bound")
    for index, suffix in (
        (3, f"/captures/{spec['label']}.resource"),
        (23, f"/captures/{spec['label']}.report.json"),
        (25, f"/spools/{spec['label']}"),
    ):
        absolute_path(argv[index], f"{label}.argv[{index}]")
        require(argv[index].endswith(suffix),
                f"{label}.argv[{index}] does not bind the capture label")
    absolute_path(argv[7], f"{label}.argv.binary")


def required_validation_argv(protocol: Mapping[str, Any]) -> dict[str, list[str]]:
    """Return the immutable command contract for every final gate.

    Pilot paths are recovered from the frozen capture paths so a portable
    verifier does not invent a new checkout or temporary root.  The command
    words and all semantic options remain fixed here rather than being
    accepted from ``rust-validation.json``.
    """
    rows = {row["label"]: row for row in analyze.protocol_rows(protocol)}
    normal_binary = rows["r1-normal-8-control"]["argv"][7]
    allocator_binary = rows["r1-allocator-8-control"]["argv"][7]
    evidence_root = str(Path(rows["r1-normal-8-control"]["argv"][23]).parent.parent)
    spool_root = str(Path(rows["r1-normal-8-control"]["argv"][25]).parent)
    libraries = {
        "libraries-tests-opc-guard": [
            "cargo", "test", "--release", "--locked", "-p", "soapberry-zip",
            "-p", "litchi-opc", "-p", "litchi-pptx",
        ],
        "libraries-clippy-opc-guard": [
            "cargo", "clippy", "--release", "--locked", "-p", "soapberry-zip",
            "-p", "litchi-opc", "-p", "litchi-pptx", "--all-targets", "--no-deps",
            "--", "-D", "warnings",
        ],
        "libraries-rustdoc-opc-guard": [
            "env", "RUSTDOCFLAGS=-D warnings", "cargo", "doc", "--locked",
            "-p", "soapberry-zip", "-p", "litchi-opc", "-p", "litchi-pptx",
            "--no-deps",
        ],
        "libraries-format-opc-guard": [
            "cargo", "fmt", "-p", "soapberry-zip", "-p", "litchi-opc",
            "-p", "litchi-pptx", "--", "--check",
        ],
        "shared-streaming-unit": [
            "cargo", "test", "--release", "--locked", "-p", "litchi-docx",
            "-p", "litchi-xlsx", "-p", "litchi-odf-common", "-p", "litchi-odt",
            "-p", "litchi-ods", "-p", "litchi-odp", "--lib", "streaming",
        ],
        "docx-streaming": [
            "cargo", "test", "--release", "--locked", "-p", "litchi-docx",
            "--test", "streaming",
        ],
        "odf-streaming": [
            "cargo", "test", "--release", "--locked", "-p", "litchi-odf-common",
            "--test", "streaming_package_writer",
        ],
        "odt-streaming": [
            "cargo", "test", "--release", "--locked", "-p", "litchi-odt",
            "--test", "streaming_plain_paragraphs",
        ],
        "ods-streaming": [
            "cargo", "test", "--release", "--locked", "-p", "litchi-ods",
            "--test", "streaming_creation", "--test", "streaming_text_spans",
        ],
        "odp-streaming": [
            "cargo", "test", "--release", "--locked", "-p", "litchi-odp",
            "--test", "streaming_provider",
        ],
        "harness-tests": [
            "cargo", "test", "--release", "--locked", "--manifest-path",
            "tools/perf-baseline/Cargo.toml", "--features", "allocator-metrics",
            "--lib", "--bin", "litchi-perf-baseline-alloc", "--bin", "pptx_metadata_spool",
        ],
        "harness-clippy-opc-guard": [
            "cargo", "clippy", "--release", "--locked", "--manifest-path",
            "tools/perf-baseline/Cargo.toml", "--features", "allocator-metrics", "--lib",
            "--bin", "litchi-perf-baseline-alloc", "--bin", "pptx_metadata_spool",
            "--no-deps", "--", "-D", "warnings",
        ],
        "harness-format-opc-guard": [
            "cargo", "fmt", "--manifest-path", "tools/perf-baseline/Cargo.toml",
            "--", "--check",
        ],
        "boundaries": ["python3", "-B", "tools/check_crate_boundaries.py"],
        "registry-strict": [
            "python3", "-B", "tools/check_perf_claims.py", "--registry",
            "docs/performance/claim-registry-v1.json", "--repo-root", ".",
            "--evidence-root", ".", "--mode", "strict",
        ],
    }
    result = dict(libraries)
    result.update({
        "build-normal-opc-guard": [
            "cargo", "build", "--release", "--locked", "--manifest-path",
            "tools/perf-baseline/Cargo.toml", "--bin", "pptx_metadata_spool",
        ],
        "build-allocator-opc-guard": [
            "cargo", "build", "--release", "--locked", "--manifest-path",
            "tools/perf-baseline/Cargo.toml", "--bin", "pptx_metadata_spool",
            "--features", "allocator-metrics",
        ],
        "pilot-normal-opc-guard": [
            normal_binary, "--mode", "both", "--counts", "8", "--samples", "1",
            "--warmups", "1", "--repeats", "1", "--max-spool-bytes",
            str(analyze.MAX_SPOOL_BYTES), "--spool-buffer-bytes", str(analyze.SPOOL_BUFFER_BYTES),
            "--json", f"{evidence_root}/pilot-normal-opc-guard.report.json", "--spool-dir",
            f"{spool_root}/pilot-normal-opc-guard",
        ],
        "pilot-allocator-opc-guard": [
            allocator_binary, "--mode", "both", "--counts", "8", "--samples", "1",
            "--warmups", "1", "--repeats", "1", "--max-spool-bytes",
            str(analyze.MAX_SPOOL_BYTES), "--spool-buffer-bytes", str(analyze.SPOOL_BUFFER_BYTES),
            "--json", f"{evidence_root}/pilot-allocator-opc-guard.report.json", "--spool-dir",
            f"{spool_root}/pilot-allocator-opc-guard",
        ],
        "analyze": ["python3", "-B", "docs/performance/results/change-0478/analyze.py"],
        "evidence-tests": [
            "python3", "-B", "-m", "unittest", "discover", "-s",
            "docs/performance/results/change-0478", "-p", "test_evidence.py",
        ],
    })
    require(tuple(result) == REQUIRED_FINAL_LABELS,
            "internal final validation command table does not match required label set")
    return result


def _manifest(root: Path, reference: Any, label: str) -> tuple[str, dict[str, str]]:
    require(isinstance(reference, dict), f"{label}: manifest reference must be an object")
    require(set(reference) == {"path", "sha256", "files"},
            f"{label}: manifest reference fields differ")
    path = bundle_file(root, reference["path"], f"{label}.path")
    expected_hash = digest(reference["sha256"], f"{label}.sha256")
    actual_hash = sha256_file(path, label)[0]
    require(actual_hash == expected_hash, f"{label}: manifest hash differs")
    value = read_json(path, label)
    require(isinstance(value, dict), f"{label}: expected an object")
    normalized: dict[str, str] = {}
    for name, value_hash in value.items():
        safe_relative(name, f"{label}.{name}")
        normalized[name] = digest(value_hash, f"{label}.{name}")
    integer(reference["files"], f"{label}.files", len(normalized))
    return expected_hash, normalized


def check_embedded_inputs(root: Path, reference: Any, label: str) -> dict[str, Any]:
    """Authenticate the retained XML input manifest and every copied asset.

    The build process records the manifest in both build wrappers and both
    binary records.  This check deliberately accepts no source-checkout path:
    the sealed bundle must carry the bytes and the exact 19-entry source to
    copy mapping needed to reproduce the input identity.
    """
    require(isinstance(reference, dict) and set(reference) ==
            {"path", "bytes", "sha256"},
            f"{label}: embedded input reference fields differ")
    relative = safe_relative(reference["path"], f"{label}.path")
    require(relative == "embedded-inputs.json",
            f"{label}.path: current embedded input manifest is required")
    manifest_path = bundle_file(root, relative, f"{label}.path")
    expected_size = integer(reference["bytes"], f"{label}.bytes")
    expected_hash = digest(reference["sha256"], f"{label}.sha256")
    check_metadata(manifest_path,
                   {"bytes": expected_size, "sha256": expected_hash}, label)
    manifest = read_json(manifest_path, label)
    require(isinstance(manifest, dict) and set(manifest) == {"schema", "files"},
            f"{label}: manifest fields differ")
    require(manifest.get("schema") == EMBEDDED_INPUT_SCHEMA,
            f"{label}: manifest schema differs")
    files = manifest.get("files")
    require(isinstance(files, dict) and
            set(files) == set(EMBEDDED_INPUT_DESTINATIONS),
            f"{label}: exact 19-source inventory is required")
    require(len(set(EMBEDDED_INPUT_DESTINATIONS.values())) == 19,
            f"{label}: embedded destination inventory is not unique")
    for source in sorted(EMBEDDED_INPUT_DESTINATIONS):
        entry = files[source]
        entry_label = f"{label}.files[{source!r}]"
        require(isinstance(entry, dict) and set(entry) ==
                {"path", "bytes", "sha256"},
                f"{entry_label}: metadata fields differ")
        destination = safe_relative(entry["path"], f"{entry_label}.path")
        require(destination == EMBEDDED_INPUT_DESTINATIONS[source],
                f"{entry_label}.path: destination differs")
        copied = bundle_file(root, destination, f"{entry_label}.path")
        check_metadata(copied,
                       {"bytes": entry["bytes"], "sha256": entry["sha256"]},
                       entry_label)
    return {"path": relative, "bytes": expected_size, "sha256": expected_hash}


def _walk(value: Any) -> Iterable[tuple[str, Any]]:
    if isinstance(value, dict):
        for key, child in value.items():
            yield key, child
            yield from _walk(child)
    elif isinstance(value, list):
        for child in value:
            yield from _walk(child)


def _check_wrapper(root: Path, wrapper: Mapping[str, Any], spec: Mapping[str, Any], instrumentation: str, protocol: Mapping[str, Any]) -> None:
    require(wrapper.get("exit_code") == 0 and wrapper.get("source_unchanged") is True,
            f"{instrumentation}: build wrapper did not succeed unchanged")
    require(wrapper.get("embedded_inputs_unchanged") is True,
            f"{instrumentation}: embedded inputs were not retained unchanged")
    wrapper_inputs = check_embedded_inputs(
        root, wrapper.get("embedded_inputs"),
        f"{instrumentation}.embedded_inputs")
    require(wrapper_inputs == spec["embedded_inputs"],
            f"{instrumentation}: wrapper embedded input identity differs")
    interval(wrapper.get("started_utc"), wrapper.get("finished_utc"),
             f"{instrumentation}: build wrapper")
    require(wrapper.get("binary") == {
        "path": spec["path"], "bytes": spec["bytes"], "sha256": spec["sha256"]
    }, f"{instrumentation}: copied binary metadata differs")
    origin = wrapper.get("original_binary")
    require(isinstance(origin, dict), f"{instrumentation}: original binary metadata is missing")
    absolute_path(origin.get("path"), f"{instrumentation}.original_binary.path")
    require(origin.get("path") != spec["path"], f"{instrumentation}: origin and copy paths coincide")
    require(origin.get("bytes") == spec["bytes"] and digest(origin.get("sha256"), f"{instrumentation}.original_binary.sha256") == spec["sha256"],
            f"{instrumentation}: origin binary identity differs")
    reference = wrapper.get("validation_gate")
    require(isinstance(reference, dict) and set(reference) == {"path", "sha256"},
            f"{instrumentation}: validation gate reference is missing")
    gate_path = bundle_file(root, reference["path"], f"{instrumentation}.validation_gate.path")
    require(sha256_file(gate_path, f"{instrumentation}.validation_gate")[0] ==
            digest(reference["sha256"], f"{instrumentation}.validation_gate.sha256"),
            f"{instrumentation}: validation gate digest differs")
    gate = read_json(gate_path, f"{instrumentation}.validation_gate")
    require(isinstance(gate, dict) and gate.get("exit_code") == 0 and gate.get("source_unchanged") is True,
            f"{instrumentation}: underlying validation gate did not succeed")
    for key in ("argv", "cwd", "environment", "driver_sha256", "common_sha256", "source_before", "source_after", "started_utc", "finished_utc", "exit_code", "source_unchanged"):
        require(wrapper.get(key) == gate.get(key),
                f"{instrumentation}: wrapper/gate field {key} differs")
    require(gate.get("environment") == protocol["environment"],
            f"{instrumentation}: build environment differs from protocol")
    before_hash, _ = _manifest(root, wrapper.get("source_before"), f"{instrumentation}.source_before")
    after_hash, _ = _manifest(root, wrapper.get("source_after"), f"{instrumentation}.source_after")
    require(before_hash == after_hash == spec["source_manifest_sha256"],
            f"{instrumentation}: source manifest identity differs")


def check_binaries(root: Path, protocol: Mapping[str, Any]) -> dict[str, dict[str, Any]]:
    value = read_json(root / "binaries.json", "binaries.json")
    require(isinstance(value, dict) and set(value) == set(analyze.INSTRUMENTATIONS),
            "binaries.json: normal and allocator identities are required")
    result: dict[str, dict[str, Any]] = {}
    for instrumentation in analyze.INSTRUMENTATIONS:
        raw = value[instrumentation]
        require(isinstance(raw, dict), f"binaries.{instrumentation}: expected an object")
        for key in ("path", "bytes", "sha256", "build_gate", "build_gate_sha256",
                    "source_manifest_sha256", "embedded_inputs"):
            require(key in raw, f"binaries.{instrumentation}.{key} is missing")
        spec = dict(raw)
        absolute_path(spec["path"], f"binaries.{instrumentation}.path")
        integer(spec["bytes"], f"binaries.{instrumentation}.bytes")
        require(spec["bytes"] > 0, f"binaries.{instrumentation}.bytes must be positive")
        spec["sha256"] = digest(spec["sha256"], f"binaries.{instrumentation}.sha256")
        spec["build_gate_sha256"] = digest(spec["build_gate_sha256"], f"binaries.{instrumentation}.build_gate_sha256")
        spec["source_manifest_sha256"] = digest(spec["source_manifest_sha256"], f"binaries.{instrumentation}.source_manifest_sha256")
        if "embedded_inputs_unchanged" in spec:
            require(spec["embedded_inputs_unchanged"] is True,
                    f"binaries.{instrumentation}: embedded inputs were not retained unchanged")
        spec["embedded_inputs"] = check_embedded_inputs(
            root, spec["embedded_inputs"],
            f"binaries.{instrumentation}.embedded_inputs")
        gate_path = bundle_file(root, spec["build_gate"], f"binaries.{instrumentation}.build_gate")
        require(sha256_file(gate_path, f"binaries.{instrumentation}.build_gate")[0] == spec["build_gate_sha256"],
                f"binaries.{instrumentation}: build wrapper digest differs")
        wrapper = read_json(gate_path, f"binaries.{instrumentation}.build_gate")
        require(isinstance(wrapper, dict), f"binaries.{instrumentation}: build wrapper is malformed")
        _check_wrapper(root, wrapper, spec, instrumentation, protocol)
        result[instrumentation] = spec
    require(result["normal"]["source_manifest_sha256"] == result["allocator"]["source_manifest_sha256"],
            "normal and allocator source manifests differ")
    require(result["normal"]["embedded_inputs"] == result["allocator"]["embedded_inputs"],
            "normal and allocator embedded input identities differ")
    return result


def check_artifacts(root: Path, directory: Path, artifacts: Any, names: set[str], label: str) -> None:
    require(isinstance(artifacts, dict) and set(artifacts) == names,
            f"{label}: artifact inventory differs")
    for name, expected in artifacts.items():
        require(name == Path(name).name and "\\" not in name,
                f"{label}.{name}: artifact name is not local")
        check_metadata(directory / name, expected, f"{label}/{name}")


def check_capture_receipts(root: Path, protocol: Mapping[str, Any], protocol_hash: str,
                           binaries: Mapping[str, Mapping[str, Any]] | None) -> dict[str, dict[str, Any]]:
    directory = root / "captures"
    require(directory.is_dir() and not directory.is_symlink(), "captures directory is missing")
    reports: dict[str, dict[str, Any]] = {}
    for spec in analyze.protocol_rows(protocol):
        label = spec["label"]
        started = read_json(directory / f"{label}.started.json", f"captures/{label}.started.json")
        finished = read_json(directory / f"{label}.json", f"captures/{label}.json")
        require(started.get("capture") == spec, f"{label}: capture plan differs")
        require(started.get("protocol_sha256") == protocol_hash, f"{label}: protocol binding differs")
        require(started.get("environment") == protocol["environment"], f"{label}: environment differs")
        absolute_path(started.get("cwd"), f"{label}.cwd")
        if binaries is not None:
            require(started.get("binary") == binaries[spec["instrumentation"]],
                    f"{label}: binary binding differs")
            require(spec["argv"][7] == binaries[spec["instrumentation"]]["path"],
                    f"{label}: command binary path differs")
        else:
            binary = started.get("binary")
            require(isinstance(binary, dict), f"{label}: binary binding is missing")
            absolute_path(binary.get("path"), f"{label}.binary.path")
            digest(binary.get("sha256"), f"{label}.binary.sha256")
        check_command_argv(spec["argv"], spec, f"captures/{label}", root)
        for key in ("capture", "cwd", "environment", "protocol_sha256", "binary"):
            require(finished.get(key) == started.get(key), f"{label}: started/final {key} differs")
        require(finished.get("exit_code") == 0, f"{label}: capture exit code is not zero")
        interval(started.get("started_utc"), finished.get("finished_utc"), f"capture {label}")
        names = {f"{label}.stdout", f"{label}.stderr", f"{label}.resource", f"{label}.report.json"}
        check_artifacts(root, directory, finished.get("artifacts"), names, f"capture {label}")
        report_path = directory / f"{label}.report.json"
        report = read_json(report_path, f"{label}.report.json")
        require(isinstance(report, dict), f"{label}: report is malformed")
        try:
            reports[label] = analyze.validate_report(report, spec, str(report_path))
        except analyze.AnalysisError as error:
            fail(str(error))
    return reports


def check_validation_receipts(root: Path, protocol: Mapping[str, Any],
                              *, allow_incomplete: bool = False) -> dict[str, int]:
    directory = root / "validation"
    require(directory.is_dir() and not directory.is_symlink(), "validation directory is missing")
    finished = sorted(path for path in directory.glob("*.json") if not path.name.endswith(".started.json"))
    require(finished, "validation has no finished receipts")
    expected_started = {f"{path.stem}.started.json" for path in finished}
    actual_started = {path.name for path in directory.glob("*.started.json")}
    if allow_incomplete:
        require(expected_started <= actual_started,
                "validation finished receipt has no started receipt")
    else:
        require(expected_started == actual_started,
                "validation started/final receipt sets differ")
    counts = {"total": 0, "successful": 0, "failed": 0}
    for path in finished:
        label = path.stem
        started = read_json(directory / f"{label}.started.json", f"validation/{label}.started.json")
        receipt = read_json(path, f"validation/{label}.json")
        for key in ("argv", "common_sha256", "cwd", "driver_sha256", "environment", "source_before", "started_utc"):
            require(key in started, f"validation/{label}.started.json: {key} is missing")
            require(receipt.get(key) == started.get(key), f"validation/{label}: {key} differs")
        require(isinstance(started["argv"], list) and started["argv"] and all(isinstance(item, str) for item in started["argv"]),
                f"validation/{label}: argv is malformed")
        absolute_path(started["cwd"], f"validation/{label}.cwd")
        require(digest(started["driver_sha256"], f"validation/{label}.driver_sha256") == sha256_file(root / "gate.py", "gate.py")[0],
                f"validation/{label}: gate driver differs")
        require(digest(started["common_sha256"], f"validation/{label}.common_sha256") == sha256_file(root / "common.py", "common.py")[0],
                f"validation/{label}: common helper differs")
        require(started["environment"] == protocol["environment"], f"validation/{label}: environment differs")
        before_hash, _ = _manifest(root, started["source_before"], f"validation/{label}.source_before")
        after_hash, _ = _manifest(root, receipt.get("source_after"), f"validation/{label}.source_after")
        require(receipt.get("source_unchanged") == (before_hash == after_hash),
                f"validation/{label}: source_unchanged is inconsistent")
        require(isinstance(receipt.get("exit_code"), int) and not isinstance(receipt["exit_code"], bool),
                f"validation/{label}: exit code is malformed")
        interval(started["started_utc"], receipt.get("finished_utc"), f"validation/{label}")
        check_artifacts(root, directory, receipt.get("artifacts"), {f"{label}.stdout", f"{label}.stderr"}, f"validation {label}")
        counts["total"] += 1
        counts["successful" if receipt["exit_code"] == 0 else "failed"] += 1
    return counts


def check_rust_validation(root: Path, binaries: Mapping[str, Mapping[str, Any]],
                          protocol: Mapping[str, Any]) -> dict[str, int]:
    ledger = read_json(root / "rust-validation.json", "rust-validation.json")
    require(isinstance(ledger, dict), "rust-validation.json: expected an object")
    require(ledger.get("schema") == RUST_VALIDATION_SCHEMA,
            "rust-validation.json: schema differs")
    required = ledger.get("required")
    require(isinstance(required, list) and required and all(isinstance(label, str) for label in required),
            "rust-validation.required: expected unique non-empty labels")
    require(len(set(required)) == len(required) and all(re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]*", label) for label in required),
            "rust-validation.required: unsafe label")
    require(set(required) == set(REQUIRED_FINAL_LABELS),
            "rust-validation.required: final gate set differs")
    expected_argv = required_validation_argv(protocol)
    final_source = digest(ledger.get("final_source_sha256"), "rust-validation.final_source_sha256")
    require(final_source == binaries["normal"]["source_manifest_sha256"] == binaries["allocator"]["source_manifest_sha256"],
            "rust-validation final source differs from build source")
    attempts = ledger.get("attempts")
    require(isinstance(attempts, dict) and attempts, "rust-validation.attempts: expected an object")
    actual_paths = {path.stem for path in (root / "validation").glob("*.json") if not path.name.endswith(".started.json")}
    require(set(attempts) == actual_paths, "rust-validation.attempts does not cover exact receipt inventory")
    final_label = REQUIRED_FINAL_LABELS[0]
    final_receipt_path = bundle_file(
        root, f"validation/{final_label}.json",
        f"rust-validation.final_source.{final_label}")
    final_receipt = read_json(final_receipt_path,
                              f"rust-validation.final_source.{final_label}")
    final_manifest_hash, _ = _manifest(
        root, final_receipt.get("source_after"),
        "rust-validation.final_source_manifest")
    require(final_manifest_hash == final_source,
            "rust-validation.final_source_manifest: hash differs")
    retained_failures = 0
    for label, expected in attempts.items():
        require(isinstance(label, str) and
                re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]*", label),
                f"rust-validation.attempts: unsafe label {label!r}")
        require(isinstance(expected, dict), f"rust-validation.attempts.{label}: expected an object")
        for key in ("path", "sha256", "argv", "exit_code", "source_unchanged"):
            require(key in expected,
                    f"rust-validation.attempts.{label}.{key}: missing")
        relative = safe_relative(expected.get("path"), f"rust-validation.attempts.{label}.path")
        require(relative == f"validation/{label}.json", f"rust-validation.attempts.{label}: path does not bind label")
        path = bundle_file(root, relative, f"rust-validation.attempts.{label}.path")
        require(sha256_file(path, f"rust-validation.attempts.{label}")[0] == digest(expected.get("sha256"), f"rust-validation.attempts.{label}.sha256"),
                f"rust-validation.attempts.{label}: receipt hash differs")
        receipt = read_json(path, f"rust-validation.attempts.{label}.receipt")
        require(receipt.get("exit_code") == expected.get("exit_code") and receipt.get("source_unchanged") == expected.get("source_unchanged"),
                f"rust-validation.attempts.{label}: receipt status differs")
        require(isinstance(expected.get("exit_code"), int) and not isinstance(expected["exit_code"], bool),
                f"rust-validation.attempts.{label}.exit_code: expected integer")
        require(isinstance(expected.get("source_unchanged"), bool) and isinstance(receipt.get("source_unchanged"), bool),
                f"rust-validation.attempts.{label}.source_unchanged: expected boolean")
        exclusions = expected.get("source_exclusions", [])
        require(isinstance(exclusions, list) and all(isinstance(item, str) for item in exclusions) and len(set(exclusions)) == len(exclusions),
                f"rust-validation.attempts.{label}.source_exclusions: malformed")
        for item in exclusions:
            safe_relative(item, f"rust-validation.attempts.{label}.source_exclusions")
        before_hash, before = _manifest(root, receipt.get("source_before"), f"rust-validation.attempts.{label}.source_before")
        after_hash, after = _manifest(root, receipt.get("source_after"), f"rust-validation.attempts.{label}.source_after")
        actual_unchanged = before_hash == after_hash
        require(receipt.get("source_unchanged") is actual_unchanged,
                f"rust-validation.attempts.{label}: source_unchanged does not match manifests")
        entry_argv = expected.get("argv")
        require(isinstance(entry_argv, list) and entry_argv and
                all(isinstance(item, str) for item in entry_argv),
                f"rust-validation.attempts.{label}.argv: malformed")
        require(receipt.get("argv") == entry_argv,
                f"rust-validation.attempts.{label}: argv differs")
        if label in required:
            require(expected.get("exit_code") == 0 and expected.get("source_unchanged") is True,
                    f"rust-validation.required.{label}: required gate did not pass unchanged")
            require(receipt.get("argv") == expected_argv[label],
                    f"rust-validation.required.{label}: argv differs from fixed command")
            require(entry_argv == expected_argv[label],
                    f"rust-validation.required.{label}: ledger argv differs from fixed command")
            require(not exclusions,
                    f"rust-validation.required.{label}: source exclusions are not allowed")
            require(before_hash == after_hash == final_source,
                    f"rust-validation.required.{label}: source is not the final build source")
        else:
            require("source_exclusions" not in expected,
                    f"rust-validation.attempts.{label}: development exclusions are not accepted")
            if expected.get("exit_code") != 0:
                retained_failures += 1
    require(retained_failures > 0, "rust-validation.attempts: no retained non-zero development attempt")
    require(digest(ledger.get("final_source_sha256"), "rust-validation.final_source_sha256") == final_source,
            "rust-validation final source digest is unstable")
    return {"required_success": len(required), "receipts": len(attempts), "retained_nonzero_attempts": retained_failures}


def _compare(expected: Any, actual: Any, label: str) -> None:
    if isinstance(expected, bool) or isinstance(actual, bool):
        require(expected == actual, f"{label}: value differs")
    elif isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        require(expected == actual or abs(float(expected) - float(actual)) <= 1e-9 * max(1.0, abs(float(expected)), abs(float(actual))), f"{label}: numeric value differs")
    elif isinstance(expected, dict) and isinstance(actual, dict):
        require(set(expected) == set(actual), f"{label}: object fields differ")
        for key in expected:
            _compare(expected[key], actual[key], f"{label}.{key}")
    elif isinstance(expected, list) and isinstance(actual, list):
        require(len(expected) == len(actual), f"{label}: list length differs")
        for index, (left, right) in enumerate(zip(expected, actual)):
            _compare(left, right, f"{label}[{index}]")
    else:
        require(expected == actual, f"{label}: value differs")


def check_summary(root: Path, reports: Mapping[str, Mapping[str, Any]]) -> None:
    path = root / "summary.json"
    summary = read_json(path, "summary.json")
    require(isinstance(summary, dict), "summary.json: expected an object")
    # analyze.py intentionally owns the arithmetic; temporarily point it at a
    # copied bundle for portable verification without importing source paths.
    original = analyze.ROOT
    try:
        analyze.ROOT = root
        expected = analyze.derive()
    except analyze.AnalysisError as error:
        fail(str(error))
    finally:
        analyze.ROOT = original
    _compare(expected, summary, "summary")


def check_seal(root: Path, required: bool = True) -> bool:
    path = root / "SHA256SUMS"
    if not path.exists():
        require(not required, "SHA256SUMS is missing")
        return False
    require(path.is_file() and not path.is_symlink(), "SHA256SUMS: regular file required")
    entries: dict[str, str] = {}
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except OSError as error:
        fail(f"SHA256SUMS: cannot read: {error}")
    for number, line in enumerate(lines, 1):
        if not line.strip():
            continue
        fields = line.split(maxsplit=1)
        require(len(fields) == 2, f"SHA256SUMS:{number}: malformed line")
        name = fields[1][1:] if fields[1].startswith("*") else fields[1]
        name = safe_relative(name, f"SHA256SUMS:{number}.path")
        require(name not in entries and name != "SHA256SUMS", f"SHA256SUMS:{number}: duplicate/self reference")
        entries[name] = digest(fields[0], f"SHA256SUMS:{number}.sha256")
    require(entries, "SHA256SUMS: no entries")
    for name, expected in entries.items():
        require(sha256_file(bundle_file(root, name, f"SHA256SUMS.{name}"), f"SHA256SUMS.{name}")[0] == expected,
                f"SHA256SUMS: digest differs for {name}")
    actual: set[str] = set()
    for candidate in root.rglob("*"):
        relative = candidate.relative_to(root).as_posix()
        require("__pycache__" not in candidate.parts, f"SHA256SUMS: __pycache__ is present: {relative}")
        require(not candidate.is_symlink(), f"SHA256SUMS: symlink is present: {relative}")
        if candidate.is_file() and relative != "SHA256SUMS":
            actual.add(relative)
    require(actual == set(entries), "SHA256SUMS: regular-file inventory differs")
    for name in ("protocol.json", "summary.json"):
        require(name in entries, f"SHA256SUMS: required entry missing: {name}")
    return True


def verify_bundle(root: Path = ROOT, *, data_only: bool = False) -> dict[str, Any]:
    root = Path(root).resolve()
    require(root.is_dir() and not root.is_symlink(), "evidence root is not a directory")
    protocol, protocol_hash = check_protocol(root)
    check_environment(root, protocol)
    preliminary = check_preliminary(root, protocol)
    binaries = None if data_only else check_binaries(root, protocol)
    reports = check_capture_receipts(root, protocol, protocol_hash, binaries)
    check_summary(root, reports)
    if data_only:
        sealed = check_seal(root, required=False)
        return {"schema": "pptx-metadata-spool-verification-v1", "protocol_sha256": protocol_hash,
                "captures": len(reports), "samples": len(reports) * analyze.SAMPLES,
                "preliminary": preliminary, "data_only": True, "sealed": sealed,
                "status": "pass"}
    validation = check_validation_receipts(root, protocol)
    rust_validation = check_rust_validation(root, binaries, protocol)
    sealed = check_seal(root, required=True)
    return {"schema": "pptx-metadata-spool-verification-v1", "protocol_sha256": protocol_hash,
            "captures": len(reports), "samples": len(reports) * analyze.SAMPLES,
            "preliminary": preliminary, "data_only": False,
            "validation": validation, "rust_validation": rust_validation,
            "sealed": sealed, "status": "pass"}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--data-only", action="store_true",
                        help="verify captures and summary before the final validation ledger")
    args = parser.parse_args(argv)
    try:
        result = verify_bundle(args.root, data_only=args.data_only)
    except VerificationError as error:
        print(f"verify.py: FAIL: {error}", file=sys.stderr)
        return 1
    print(json.dumps(result, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
