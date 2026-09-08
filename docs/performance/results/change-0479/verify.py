#!/usr/bin/env python3
"""Fail-closed, portable verification for the 0479 evidence bundle.

The verifier follows only paths inside the bundle when checking retained
artifacts.  Absolute repository, temporary, and executable paths are recorded
as provenance but are never opened.  ``--data-only`` validates the frozen
protocol, reports, receipts, and derived summary before the final validation
ledger and seal exist; the default additionally checks the ledger and the
content-addressed seal.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import re
import sys
from pathlib import Path, PurePosixPath
from typing import Any, Mapping, NoReturn

import analyze


ROOT = Path(__file__).resolve().parent
SHA256 = re.compile(r"^[0-9a-fA-F]{64}$")
LABEL = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]*$")
RUST_VALIDATION_SCHEMA = "docx-plain-paragraph-tail-append-rust-validation-v1"
VERIFICATION_SCHEMA = "docx-plain-paragraph-tail-append-verification-v1"
EMBEDDED_INPUT_SCHEMA = "docx-embedded-inputs-v1"
EMBEDDED_INPUTS_PATH = "embedded-inputs.json"

REQUIRED_FINAL_LABELS = (
    "harness-tests-final", "harness-clippy-final", "harness-format-final", "harness-rustdoc",
    "docx-paragraph-copy", "docx-paragraph-removal", "boundaries",
    "registry-strict", "evidence-tests-final", "analyze-final", "build-normal",
    "build-allocator", "pilot-normal-total", "pilot-normal-phases",
    "pilot-allocator-total", "pilot-allocator-phases",
)


class VerificationError(ValueError):
    """A retained artifact failed an independent check."""


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
            parse_constant=lambda value: (_ for _ in ()).throw(ValueError(value)),
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
            f"{label}: expected SHA-256 digest")
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
    require(timestamp(finish, f"{label}.finished_utc") >= timestamp(start, f"{label}.started_utc"),
            f"{label}: finished before started")


def safe_relative(value: Any, label: str) -> str:
    require(isinstance(value, str) and value, f"{label}: expected a relative POSIX path")
    require("\\" not in value, f"{label}: backslashes are forbidden")
    path = PurePosixPath(value)
    require(not path.is_absolute() and value not in {".", ".."}, f"{label}: absolute/parent path is forbidden")
    require(all(part not in {"", ".", ".."} for part in path.parts), f"{label}: path escapes bundle root")
    return path.as_posix()


def bundle_file(root: Path, value: Any, label: str) -> Path:
    relative = safe_relative(value, label)
    current = root
    for part in PurePosixPath(relative).parts:
        current /= part
        require(not current.is_symlink(), f"{label}: symlink component is forbidden")
    try:
        resolved = current.resolve(strict=True)
    except OSError as error:
        fail(f"{label}: cannot resolve: {error}")
    base = root.resolve()
    require(resolved == base or base in resolved.parents, f"{label}: path escapes bundle root")
    require(current.is_file(), f"{label}: file is missing")
    return current


def metadata(path: Path, label: str = "file") -> dict[str, Any]:
    value, size = sha256_file(path, label)
    return {"bytes": size, "sha256": value}


def check_metadata(path: Path, expected: Any, label: str) -> None:
    require(isinstance(expected, dict) and set(expected) == {"bytes", "sha256"}, f"{label}: metadata fields differ")
    integer(expected.get("bytes"), f"{label}.bytes")
    expected_hash = digest(expected.get("sha256"), f"{label}.sha256")
    require(metadata(path, label) == {"bytes": expected["bytes"], "sha256": expected_hash}, f"{label}: metadata differs")


def _manifest(root: Path, reference: Any, label: str) -> tuple[str, dict[str, str]]:
    require(isinstance(reference, dict) and set(reference) == {"path", "sha256", "files"}, f"{label}: source manifest reference differs")
    relative = safe_relative(reference.get("path"), f"{label}.path")
    path = bundle_file(root, relative, f"{label}.path")
    actual_hash = sha256_file(path, label)[0]
    require(actual_hash == digest(reference.get("sha256"), f"{label}.sha256"), f"{label}: manifest digest differs")
    value = read_json(path, label)
    require(isinstance(value, dict), f"{label}: manifest must be an object")
    # gate.py snapshots are deliberately plain path -> SHA-256 maps.  The
    # reference carries the file count separately; there is no nested
    # ``files`` wrapper in the retained manifest itself.
    file_count = integer(reference.get("files"), f"{label}.files")
    require(file_count > 0, f"{label}.files: manifest must contain files")
    require(len(value) == file_count, f"{label}.files: manifest count differs")
    result: dict[str, str] = {}
    for source, source_hash in value.items():
        require(isinstance(source, str), f"{label}: manifest source path is malformed")
        safe_relative(source, f"{label}.{source}.path")
        result[source] = digest(source_hash, f"{label}.{source}.sha256")
    return actual_hash, result


def _check_protocol_scripts(root: Path, protocol: Mapping[str, Any]) -> None:
    scripts = protocol["scripts"]
    for name, expected in scripts.items():
        path = bundle_file(root, name, f"protocol.scripts[{name!r}]")
        require(sha256_file(path, name)[0] == digest(expected, f"protocol.scripts[{name!r}]"), f"protocol script digest differs: {name}")


def check_protocol(root: Path) -> tuple[dict[str, Any], str]:
    path = bundle_file(root, "protocol.json", "protocol.json")
    protocol = read_json(path, "protocol.json")
    require(isinstance(protocol, dict), "protocol.json: expected object")
    try:
        rows = analyze.protocol_rows(protocol)
    except analyze.AnalysisError as error:
        fail(str(error))
    _check_protocol_scripts(root, protocol)
    for spec in rows:
        _check_argv(spec["argv"], spec, f"protocol.captures[{spec['label']}]")
    return protocol, sha256_file(path, "protocol.json")[0]


def _check_argv(argv: Any, spec: Mapping[str, Any], label: str) -> None:
    require(isinstance(argv, list) and len(argv) == 18 and all(isinstance(value, str) for value in argv), f"{label}.argv: expected 18 strings")
    fixed = {0: "/usr/bin/time", 1: "-v", 2: "-o", 4: "/usr/bin/taskset", 5: "-c", 6: "2", 8: "--mode", 10: "--counts", 12: "--samples", 13: "30", 14: "--warmups", 15: "3", 16: "--json"}
    for index, expected in fixed.items():
        require(argv[index] == expected, f"{label}.argv[{index}]: expected {expected!r}")
    require(argv[9] == spec["mode"] and argv[11] == str(spec["count"]), f"{label}: mode/count are not bound")
    require(Path(argv[7]).is_absolute(), f"{label}.argv[7]: executable path must be absolute provenance")
    require(Path(argv[3]).is_absolute() and argv[3].endswith(f"/captures/{spec['label']}.resource"), f"{label}.argv[3]: resource path does not bind label")
    require(Path(argv[17]).is_absolute() and argv[17].endswith(f"/captures/{spec['label']}.report.json"), f"{label}.argv[17]: report path does not bind label")


def check_environment(root: Path, protocol: Mapping[str, Any]) -> None:
    path = bundle_file(root, "environment.json", "environment.json")
    value = read_json(path, "environment.json")
    require(isinstance(value, dict), "environment.json: expected object")
    filesystem = value.get("filesystem")
    require(isinstance(filesystem, dict), "environment.filesystem: missing")
    require(filesystem.get("exit_code") == 0 and filesystem.get("stderr") == "", "environment filesystem probe failed")
    stdout = filesystem.get("stdout")
    require(isinstance(stdout, str) and any(line.split()[:2] == ["tmpfs", "tmpfs"] for line in stdout.splitlines()[1:]), "environment filesystem is not tmpfs")
    for key, expected in protocol["environment"].items():
        # The environment artifact does not duplicate the capture environment;
        # this loop documents that the frozen protocol remains authoritative.
        require(isinstance(expected, str), f"protocol.environment.{key}: expected text")
    profiler = bundle_file(root, "profiler-capabilities.json", "profiler-capabilities.json")
    profiler_value = read_json(profiler, "profiler-capabilities.json")
    require(isinstance(profiler_value, dict), "profiler-capabilities.json: expected object")


def _check_embedded_inputs(root: Path, reference: Any, label: str) -> dict[str, Any]:
    require(isinstance(reference, dict) and set(reference) == {"path", "bytes", "sha256"}, f"{label}: reference differs")
    relative = safe_relative(reference["path"], f"{label}.path")
    require(relative == EMBEDDED_INPUTS_PATH, f"{label}.path: canonical path differs")
    path = bundle_file(root, relative, f"{label}.path")
    check_metadata(path, {"bytes": reference["bytes"], "sha256": reference["sha256"]}, label)
    value = read_json(path, label)
    require(isinstance(value, dict) and value.get("schema") == EMBEDDED_INPUT_SCHEMA and isinstance(value.get("files"), dict), f"{label}: manifest schema/files differ")
    files: dict[str, Any] = {}
    for source, item in value["files"].items():
        require(isinstance(source, str) and isinstance(item, dict) and set(item) == {"bytes", "path", "sha256"}, f"{label}.{source}: entry differs")
        destination = safe_relative(item["path"], f"{label}.{source}.path")
        destination_path = bundle_file(root, destination, f"{label}.{source}.path")
        check_metadata(destination_path, {"bytes": item["bytes"], "sha256": item["sha256"]}, f"{label}.{source}")
        files[source] = {"path": destination, "bytes": item["bytes"], "sha256": item["sha256"]}
    require(len(files) == 8, f"{label}: expected eight embedded DOCX inputs")
    return {"path": relative, "bytes": reference["bytes"], "sha256": digest(reference["sha256"], f"{label}.sha256"), "files": files}


def check_binaries(root: Path, protocol: Mapping[str, Any]) -> dict[str, dict[str, Any]]:
    value = read_json(root / "binaries.json", "binaries.json")
    require(isinstance(value, dict) and set(value) == set(analyze.INSTRUMENTATIONS), "binaries: normal/allocator identities required")
    result: dict[str, dict[str, Any]] = {}
    for instrumentation in analyze.INSTRUMENTATIONS:
        raw = value[instrumentation]
        required = {"path", "bytes", "sha256", "embedded_inputs", "embedded_inputs_unchanged", "build_gate", "build_gate_sha256", "source_manifest_sha256"}
        require(isinstance(raw, dict) and set(raw) == required, f"binaries.{instrumentation}: fields differ")
        require(isinstance(raw["path"], str) and Path(raw["path"]).is_absolute(), f"binaries.{instrumentation}.path: provenance path required")
        integer(raw["bytes"], f"binaries.{instrumentation}.bytes")
        require(raw["bytes"] > 0, f"binaries.{instrumentation}.bytes: must be positive")
        digest(raw["sha256"], f"binaries.{instrumentation}.sha256")
        digest(raw["build_gate_sha256"], f"binaries.{instrumentation}.build_gate_sha256")
        digest(raw["source_manifest_sha256"], f"binaries.{instrumentation}.source_manifest_sha256")
        require(raw["embedded_inputs_unchanged"] is True, f"binaries.{instrumentation}: embedded inputs changed")
        _check_embedded_inputs(root, raw["embedded_inputs"], f"binaries.{instrumentation}.embedded_inputs")
        gate_path = bundle_file(root, raw["build_gate"], f"binaries.{instrumentation}.build_gate")
        require(sha256_file(gate_path, "build gate")[0] == raw["build_gate_sha256"], f"binaries.{instrumentation}: build gate digest differs")
        gate = read_json(gate_path, "build gate")
        require(isinstance(gate, dict) and gate.get("exit_code") == 0 and gate.get("source_unchanged") is True, f"binaries.{instrumentation}: build gate failed")
        expected_binary = {"path": raw["path"], "bytes": raw["bytes"], "sha256": raw["sha256"]}
        copied_binary = gate.get("binary")
        require(isinstance(copied_binary, dict) and set(copied_binary) == {"path", "bytes", "sha256"}, f"binaries.{instrumentation}: copied binary metadata is missing")
        require(copied_binary == expected_binary, f"binaries.{instrumentation}: copied binary metadata differs")
        original_binary = gate.get("original_binary")
        require(isinstance(original_binary, dict) and set(original_binary) == {"path", "bytes", "sha256"}, f"binaries.{instrumentation}: original binary metadata is missing")
        require(isinstance(original_binary["path"], str) and Path(original_binary["path"]).is_absolute(), f"binaries.{instrumentation}: original binary path is not provenance")
        require(original_binary["bytes"] == raw["bytes"] and original_binary["sha256"] == raw["sha256"], f"binaries.{instrumentation}: copied/original binary digest differs")
        require(gate.get("embedded_inputs") == raw["embedded_inputs"] and gate.get("embedded_inputs_unchanged") is True, f"binaries.{instrumentation}: embedded input binding differs")
        validation_gate = gate.get("validation_gate")
        require(isinstance(validation_gate, dict) and set(validation_gate) == {"path", "sha256"}, f"binaries.{instrumentation}: validation gate binding is missing")
        validation_gate_path = bundle_file(root, validation_gate["path"], f"binaries.{instrumentation}.validation_gate.path")
        require(sha256_file(validation_gate_path, f"binaries.{instrumentation}.validation_gate")[0] == digest(validation_gate["sha256"], f"binaries.{instrumentation}.validation_gate.sha256"), f"binaries.{instrumentation}: validation gate digest differs")
        gate_driver_sha = sha256_file(bundle_file(root, "gate.py", "gate.py"), "gate.py")[0]
        gate_common_sha = sha256_file(bundle_file(root, "common.py", "common.py"), "common.py")[0]
        require(gate.get("driver_sha256") == gate_driver_sha and gate.get("common_sha256") == gate_common_sha, f"binaries.{instrumentation}: gate driver/common binding differs")
        before_hash, _ = _manifest(root, gate.get("source_before"), f"{instrumentation}.source_before")
        after_hash, _ = _manifest(root, gate.get("source_after"), f"{instrumentation}.source_after")
        require(before_hash == after_hash == raw["source_manifest_sha256"], f"binaries.{instrumentation}: source manifest differs")
        # Preserve the exact object loaded from binaries.json because capture
        # receipts bind their ``binary`` object byte-for-byte to that record.
        result[instrumentation] = dict(raw)
    require(result["normal"]["source_manifest_sha256"] == result["allocator"]["source_manifest_sha256"], "normal/allocator sources differ")
    require(result["normal"]["embedded_inputs"] == result["allocator"]["embedded_inputs"], "normal/allocator embedded inputs differ")
    return result


def _artifact_check(root: Path, directory: Path, artifacts: Any, names: set[str], label: str) -> None:
    require(isinstance(artifacts, dict) and set(artifacts) == names, f"{label}: artifact inventory differs")
    for name, expected in artifacts.items():
        require(name == Path(name).name and "\\" not in name, f"{label}.{name}: unsafe artifact name")
        check_metadata(directory / name, expected, f"{label}/{name}")


def check_capture_receipts(root: Path, protocol: Mapping[str, Any], protocol_hash: str, binaries: Mapping[str, Mapping[str, Any]] | None) -> dict[str, dict[str, Any]]:
    directory = root / "captures"
    require(directory.is_dir() and not directory.is_symlink(), "captures directory is missing")
    reports: dict[str, dict[str, Any]] = {}
    for spec in analyze.protocol_rows(protocol):
        label = spec["label"]
        started = read_json(directory / f"{label}.started.json", f"captures/{label}.started.json")
        finished = read_json(directory / f"{label}.json", f"captures/{label}.json")
        require(isinstance(started, dict) and set(started) == {"capture", "cwd", "started_utc", "protocol_sha256", "binary", "environment"}, f"{label}: started envelope fields differ")
        require(isinstance(finished, dict) and set(finished) == {"capture", "cwd", "started_utc", "protocol_sha256", "binary", "environment", "exit_code", "finished_utc", "artifacts"}, f"{label}: finished envelope fields differ")
        require(started.get("capture") == spec, f"{label}: started capture differs")
        require(started.get("protocol_sha256") == protocol_hash, f"{label}: protocol binding differs")
        require(started.get("environment") == protocol["environment"], f"{label}: environment differs")
        require(isinstance(started.get("cwd"), str) and Path(started["cwd"]).is_absolute(), f"{label}.cwd: provenance path required")
        if binaries is not None:
            require(started.get("binary") == binaries[spec["instrumentation"]], f"{label}: binary binding differs")
        else:
            require(isinstance(started.get("binary"), dict), f"{label}: binary binding missing")
        _check_argv(spec["argv"], spec, f"captures/{label}")
        require(spec["argv"][7] == started["binary"].get("path"), f"{label}: capture executable does not bind recorded binary")
        for key in ("capture", "cwd", "environment", "protocol_sha256", "binary"):
            require(finished.get(key) == started.get(key), f"{label}: started/final {key} differs")
        require(finished.get("exit_code") == 0, f"{label}: capture failed")
        interval(started.get("started_utc"), finished.get("finished_utc"), f"capture {label}")
        names = {f"{label}.stdout", f"{label}.stderr", f"{label}.resource", f"{label}.report.json"}
        _artifact_check(root, directory, finished.get("artifacts"), names, f"capture {label}")
        resource = directory / f"{label}.resource"
        require(re.search(r"Maximum resident set size \(kbytes\):\s*\d+", resource.read_text(encoding="utf-8")), f"{label}: resource RSS missing")
        report = read_json(directory / f"{label}.report.json", f"{label}.report.json")
        require(isinstance(report, dict), f"{label}: report is malformed")
        try:
            reports[label] = analyze.validate_report(report, spec, str(directory / f"{label}.report.json"))
        except analyze.AnalysisError as error:
            fail(str(error))
    return reports


def required_validation_argv(protocol: Mapping[str, Any], binaries: Mapping[str, Mapping[str, Any]] | None = None, root: Path = ROOT) -> dict[str, list[str]]:
    """Return the fixed command contract used by ``run-gates.py`` and pilots."""
    manifest = ["--manifest-path", "tools/perf-baseline/Cargo.toml"]
    targets = ["--lib", "--bin", "docx_plain_paragraph_tail_append"]
    expected: dict[str, list[str]] = {
        "harness-tests-final": ["cargo", "test", "--release", "--locked", *manifest, "--features", "allocator-metrics", *targets],
        "harness-clippy-final": ["cargo", "clippy", "--release", "--locked", *manifest, "--features", "allocator-metrics", *targets, "--no-deps", "--", "-D", "warnings"],
        "harness-format-final": ["cargo", "fmt", *manifest, "--", "--check"],
        "harness-rustdoc": ["env", "RUSTDOCFLAGS=-D warnings", "cargo", "doc", "--locked", *manifest, "--lib", "--no-deps"],
        "docx-paragraph-copy": ["cargo", "test", "--release", "--locked", "-p", "litchi-docx", "--test", "source_backed_paragraph_copy"],
        "docx-paragraph-removal": ["cargo", "test", "--release", "--locked", "-p", "litchi-docx", "--test", "source_backed_paragraph_removal"],
        "boundaries": ["python3", "-B", "tools/check_crate_boundaries.py"],
        "registry-strict": ["python3", "-B", "tools/check_perf_claims.py", "--registry", "docs/performance/claim-registry-v1.json", "--repo-root", ".", "--evidence-root", ".", "--mode", "strict"],
        "evidence-tests-final": ["python3", "-B", "-m", "unittest", "discover", "-s", "docs/performance/results/change-0479", "-p", "test_evidence.py"],
        "analyze-final": ["python3", "-B", "docs/performance/results/change-0479/analyze.py"],
        "build-normal": ["cargo", "build", "--release", "--locked", *manifest, "--bin", "docx_plain_paragraph_tail_append"],
        "build-allocator": ["cargo", "build", "--release", "--locked", *manifest, "--bin", "docx_plain_paragraph_tail_append", "--features", "allocator-metrics"],
    }
    if binaries is not None:
        for instrumentation in analyze.INSTRUMENTATIONS:
            for mode in analyze.MODES:
                label = f"pilot-{instrumentation}-{mode}"
                expected[label] = [binaries[instrumentation]["path"], "--counts", "64,8192,131072", "--mode", mode, "--samples", "1", "--warmups", "1", "--json", f"{root}/{label}.report.json"]
    return expected


def _argv_matches(label: str, actual: Any, expected: list[str]) -> bool:
    """Compare a retained command while ignoring only portable pilot output roots."""
    if not isinstance(actual, list) or not all(isinstance(item, str) for item in actual):
        return False
    if not label.startswith("pilot-"):
        return actual == expected
    if len(actual) != len(expected) or actual[:-1] != expected[:-1]:
        return False
    return Path(actual[-1]).is_absolute() and actual[-1].endswith(f"/{label}.report.json")


def check_validation_receipts(root: Path, protocol: Mapping[str, Any], binaries: Mapping[str, Mapping[str, Any]] | None = None) -> dict[str, int]:
    directory = root / "validation"
    require(directory.is_dir() and not directory.is_symlink(), "validation directory is missing")
    finished = sorted(path for path in directory.glob("*.json") if not path.name.endswith(".started.json"))
    require(finished, "validation has no receipts")
    started = {path.name for path in directory.glob("*.started.json")}
    require({f"{path.stem}.started.json" for path in finished} <= started, "validation finished receipt has no started record")
    expected_argv = required_validation_argv(protocol, binaries, root)
    gate_driver_sha = sha256_file(bundle_file(root, "gate.py", "gate.py"), "gate.py")[0]
    gate_common_sha = sha256_file(bundle_file(root, "common.py", "common.py"), "common.py")[0]
    counts = {"total": 0, "successful": 0, "failed": 0}
    for path in finished:
        label = path.stem
        require(LABEL.fullmatch(label) is not None, f"validation label is unsafe: {label!r}")
        begin = read_json(directory / f"{label}.started.json", f"validation/{label}.started.json")
        receipt = read_json(path, f"validation/{label}.json")
        require(isinstance(begin, dict) and set(begin) == {"argv", "cwd", "environment", "driver_sha256", "common_sha256", "started_utc", "source_before"}, f"validation/{label}.started.json: envelope fields differ")
        require(isinstance(receipt, dict) and set(receipt) == {"argv", "cwd", "environment", "driver_sha256", "common_sha256", "started_utc", "source_before", "exit_code", "finished_utc", "source_after", "artifacts", "source_unchanged"}, f"validation/{label}.json: envelope fields differ")
        for key in ("argv", "cwd", "environment", "driver_sha256", "common_sha256", "source_before", "started_utc"):
            require(key in begin, f"validation/{label}.started.json: {key} missing")
            require(receipt.get(key) == begin.get(key), f"validation/{label}: {key} differs")
        require(isinstance(begin["argv"], list) and begin["argv"] and all(isinstance(x, str) for x in begin["argv"]), f"validation/{label}.argv malformed")
        require(isinstance(begin["cwd"], str) and Path(begin["cwd"]).is_absolute(), f"validation/{label}.cwd malformed")
        require(digest(begin["driver_sha256"], f"validation/{label}.driver_sha256"), f"validation/{label}: driver digest malformed")
        require(digest(begin["common_sha256"], f"validation/{label}.common_sha256"), f"validation/{label}: common digest malformed")
        require(begin["driver_sha256"] == gate_driver_sha and begin["common_sha256"] == gate_common_sha, f"validation/{label}: gate driver/common binding differs")
        require(begin["environment"] == protocol["environment"], f"validation/{label}: environment differs")
        before_hash, _ = _manifest(root, begin["source_before"], f"validation/{label}.source_before")
        after_hash, _ = _manifest(root, receipt.get("source_after"), f"validation/{label}.source_after")
        require(receipt.get("source_unchanged") is (before_hash == after_hash), f"validation/{label}: source_unchanged is inconsistent")
        require(isinstance(receipt.get("exit_code"), int) and not isinstance(receipt["exit_code"], bool), f"validation/{label}: exit code malformed")
        interval(begin["started_utc"], receipt.get("finished_utc"), f"validation/{label}")
        _artifact_check(root, directory, receipt.get("artifacts"), {f"{label}.stdout", f"{label}.stderr"}, f"validation {label}")
        if label in REQUIRED_FINAL_LABELS:
            require(receipt["exit_code"] == 0 and receipt.get("source_unchanged") is True, f"validation required gate {label} failed")
            require(label in expected_argv and _argv_matches(label, begin["argv"], expected_argv[label]), f"validation required gate {label} argv differs")
        counts["total"] += 1
        counts["successful" if receipt["exit_code"] == 0 else "failed"] += 1
    return counts


def check_rust_validation(root: Path, binaries: Mapping[str, Mapping[str, Any]], protocol: Mapping[str, Any]) -> dict[str, int]:
    ledger = read_json(root / "rust-validation.json", "rust-validation.json")
    require(isinstance(ledger, dict) and ledger.get("schema") == RUST_VALIDATION_SCHEMA, "rust-validation schema differs")
    required = ledger.get("required")
    require(required == list(REQUIRED_FINAL_LABELS), "rust-validation required labels differ")
    final_source = digest(ledger.get("final_source_sha256"), "rust-validation.final_source_sha256")
    require(final_source == binaries["normal"]["source_manifest_sha256"] == binaries["allocator"]["source_manifest_sha256"], "rust-validation final source differs")
    attempts = ledger.get("attempts")
    require(isinstance(attempts, dict) and attempts, "rust-validation.attempts missing")
    expected_argv = required_validation_argv(protocol, binaries)
    actual_receipts = {path.stem for path in (root / "validation").glob("*.json") if not path.name.endswith(".started.json")}
    require(set(attempts) == actual_receipts, "rust-validation attempts do not cover validation receipts")
    for label, item in attempts.items():
        require(LABEL.fullmatch(label) is not None and isinstance(item, dict), f"rust-validation attempt {label!r} is malformed")
        require(item.get("path") == f"validation/{label}.json", f"rust-validation attempt path does not bind label: {label}")
        receipt_path = bundle_file(root, item.get("path"), f"rust-validation.{label}.path")
        require(sha256_file(receipt_path, f"rust-validation.{label}")[0] == digest(item.get("sha256"), f"rust-validation.{label}.sha256"), f"rust-validation receipt digest differs: {label}")
        receipt = read_json(receipt_path, f"validation/{label}.json")
        require(receipt.get("exit_code") == item.get("exit_code") and receipt.get("source_unchanged") == item.get("source_unchanged"), f"rust-validation receipt status differs: {label}")
        require(receipt.get("argv") == item.get("argv"), f"rust-validation receipt argv differs from ledger: {label}")
    for label in REQUIRED_FINAL_LABELS:
        item = attempts.get(label)
        require(isinstance(item, dict), f"rust-validation attempt missing: {label}")
        require(item.get("exit_code") == 0 and item.get("source_unchanged") is True, f"rust-validation required attempt failed: {label}")
        require(_argv_matches(label, item.get("argv"), expected_argv[label]), f"rust-validation required argv differs: {label}")
        require(item.get("source_exclusions", []) == [], f"rust-validation required source exclusions: {label}")
        receipt = read_json(bundle_file(root, item.get("path"), f"rust-validation.{label}.path"), f"validation/{label}.json")
        require(_argv_matches(label, receipt.get("argv"), expected_argv[label]), f"rust-validation receipt argv differs: {label}")
        require(receipt.get("source_unchanged") is True, f"rust-validation receipt source changed: {label}")
        source_before = receipt.get("source_before")
        source_after = receipt.get("source_after")
        require(isinstance(source_before, dict) and isinstance(source_after, dict) and source_before.get("sha256") == final_source and source_after.get("sha256") == final_source, f"rust-validation required manifest hash differs: {label}")
    return {"required_success": len(REQUIRED_FINAL_LABELS), "receipts": len(attempts)}


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


def check_summary(root: Path) -> None:
    summary = read_json(root / "summary.json", "summary.json")
    require(isinstance(summary, dict), "summary.json: expected object")
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
    for number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
        if not line.strip():
            continue
        fields = line.split(maxsplit=1)
        require(len(fields) == 2, f"SHA256SUMS:{number}: malformed line")
        name = fields[1][1:] if fields[1].startswith("*") else fields[1]
        name = safe_relative(name, f"SHA256SUMS:{number}.path")
        require(name != "SHA256SUMS" and name not in entries, f"SHA256SUMS:{number}: duplicate/self reference")
        entries[name] = digest(fields[0], f"SHA256SUMS:{number}.sha256")
    require(entries, "SHA256SUMS: no entries")
    for name, expected in entries.items():
        require(sha256_file(bundle_file(root, name, f"SHA256SUMS.{name}"), name)[0] == expected, f"SHA256SUMS: digest differs for {name}")
    actual = {candidate.relative_to(root).as_posix() for candidate in root.rglob("*") if candidate.is_file() and candidate.name != "SHA256SUMS"}
    require(actual == set(entries), "SHA256SUMS: file inventory differs")
    require({"protocol.json", "summary.json"} <= set(entries), "SHA256SUMS: protocol/summary missing")
    return True


def verify_bundle(root: Path = ROOT, *, data_only: bool = False) -> dict[str, Any]:
    root = Path(root).resolve()
    require(root.is_dir() and not root.is_symlink(), "evidence root is not a directory")
    protocol, protocol_hash = check_protocol(root)
    check_environment(root, protocol)
    # ``--data-only`` omits only the final validation ledger and seal.  Build,
    # embedded-input, pilot, and capture custody are already available before
    # formal analysis and remain part of the portable evidence boundary.
    binaries = check_binaries(root, protocol)
    reports = check_capture_receipts(root, protocol, protocol_hash, binaries)
    require(len(reports) == 24, "capture count differs")
    check_summary(root)
    if data_only:
        import closure
        closure_result = closure.check(root, complete=False)
        return {"schema": VERIFICATION_SCHEMA, "protocol_sha256": protocol_hash, "captures": len(reports), "samples": len(reports) * analyze.SAMPLES, "data_only": True, "closure": closure_result, "sealed": check_seal(root, required=False), "status": "pass"}
    validation = check_validation_receipts(root, protocol, binaries)
    rust = check_rust_validation(root, binaries, protocol)
    sealed = check_seal(root, required=True)
    import closure
    closure_result = closure.check(root, complete=True)
    return {"schema": VERIFICATION_SCHEMA, "protocol_sha256": protocol_hash, "captures": len(reports), "samples": len(reports) * analyze.SAMPLES, "data_only": False, "validation": validation, "rust_validation": rust, "sealed": sealed, "closure": closure_result, "status": "pass"}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--data-only", action="store_true")
    args = parser.parse_args(argv)
    try:
        print(json.dumps(verify_bundle(args.root, data_only=args.data_only), sort_keys=True))
    except VerificationError as error:
        print(f"verify.py: FAIL: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
