#!/usr/bin/env python3
"""Create or verify the immutable 0487 evidence-bundle seal.

The default command is ``verify``.  ``seal`` is intentionally explicit and
requires the final gate, comparison, profile, and fuzz selections.  Semantic
checks only read retained receipts; this helper never captures, builds, or
writes a comparison result.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import importlib
import json
from pathlib import Path
import re
import sys
from typing import Any, Mapping


# Imports of compare.py and its validators must not create a new __pycache__
# while the bundle is being sealed.  Bytecode already present in the bundle is
# rejected by _inventory().
sys.dont_write_bytecode = True


SCHEMA = "docx-opc-splice-audit-evidence-seal-v1"
VERSION = 1
VERIFICATION_SCHEMA = "docx-opc-splice-audit-seal-verification-v1"
COMPARISON_SUMMARY_SCHEMA = "docx-opc-splice-audit-consumed-prefix-comparison-summary-v1"
SEAL_NAME = "seal.json"
VERIFICATION_NAME = "seal-verification.json"
EXCLUDED = {
    "self": SEAL_NAME,
    "verification_receipts": [VERIFICATION_NAME],
    "bytecode": "reject",
}
TOKEN_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]*$")
BYTECODE_SUFFIXES = {".pyc", ".pyo"}
PROFILE_EXPECTED_PASSES = 12
SCOPE = "0487 final gates, 144-process/4320-sample formal comparison, passing profile receipts, and completed fuzz smoke"
PROFILE_WORKLOADS = ("s131072-a64-short-c64", "s64-a16384-short-c64")
PROFILE_EXPECTED_LABELS = frozenset(
    f"{version}-{tool}-{input_mode}-{workload}"
    for version, tools in (("before", ("perf",)), ("after", ("perf", "strace")))
    for tool in tools
    for input_mode in ("owned", "file")
    for workload in PROFILE_WORKLOADS
)
FUZZ_INPUT_CAP = 64 * 1024
FUZZ_RUN_SEEDS = (484, 485)
FUZZ_RUN_COUNT = len(FUZZ_RUN_SEEDS)
FUZZ_SEED_COUNT = 90
FUZZ_POSITIVE_CASES = 27
DEFAULT_ROOT = Path(__file__).resolve().parent


class SealError(RuntimeError):
    """A fail-closed bundle or receipt error."""


def _fail(message: str) -> None:
    raise SealError(message)


def _require(condition: bool, message: str) -> None:
    if not condition:
        _fail(message)


def _now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def _sha(path: Path) -> str:
    try:
        with path.open("rb") as stream:
            return hashlib.file_digest(stream, "sha256").hexdigest()
    except OSError as error:
        _fail(f"cannot hash {path}: {error}")
    raise AssertionError("unreachable")


def _meta(path: Path, label: str) -> dict[str, int | str]:
    _require(not path.is_symlink(), f"{label}: symlink is not a regular file: {path}")
    _require(path.is_file(), f"{label}: regular file is missing: {path}")
    try:
        size = path.stat().st_size
    except OSError as error:
        _fail(f"{label}: cannot stat {path}: {error}")
    return {"bytes": size, "sha256": _sha(path)}


def _read_json(path: Path, label: str) -> Any:
    _meta(path, label)
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as error:
        _fail(f"{label}: invalid JSON in {path}: {error}")
    raise AssertionError("unreachable")


def _rooted(root: Path) -> Path:
    _require(root.is_dir() and not root.is_symlink(), f"root is not a regular directory: {root}")
    try:
        resolved = root.resolve(strict=True)
    except OSError as error:
        _fail(f"cannot resolve root {root}: {error}")
    _require(resolved.is_dir() and not resolved.is_symlink(), f"resolved root is not a directory: {resolved}")
    return resolved


def _relative(root: Path, path: Path, label: str) -> str:
    try:
        return path.resolve(strict=True).relative_to(root).as_posix()
    except (OSError, ValueError) as error:
        _fail(f"{label}: path is outside root: {path} ({error})")
    raise AssertionError("unreachable")


def _inside(root: Path, value: str | Path, label: str) -> tuple[Path, str]:
    candidate = Path(value)
    if not candidate.is_absolute():
        candidate = root / candidate
    try:
        resolved = candidate.resolve(strict=True)
    except OSError as error:
        _fail(f"{label}: missing path {candidate}: {error}")
    relative = _relative(root, resolved, label)

    # Reject a symlink in any component even if it resolves back into ROOT.
    current = root
    for component in Path(relative).parts:
        current /= component
        _require(not current.is_symlink(), f"{label}: symlink path component: {current}")
    _meta(resolved, label)
    return resolved, relative


def _any_file(value: str | Path, label: str) -> Path:
    path = Path(value)
    _require(not path.is_symlink(), f"{label}: symlink is not allowed: {path}")
    try:
        resolved = path.resolve(strict=True)
    except OSError as error:
        _fail(f"{label}: missing path {path}: {error}")
    _meta(resolved, label)
    return resolved


def _is_bytecode(path: Path) -> bool:
    return "__pycache__" in path.parts or path.suffix.lower() in BYTECODE_SUFFIXES


def _inventory(root: Path) -> dict[str, dict[str, int | str]]:
    """Hash all regular files while rejecting links, bytecode, and specials."""

    result: dict[str, dict[str, int | str]] = {}
    try:
        paths = sorted(root.rglob("*"), key=lambda item: item.relative_to(root).as_posix())
    except OSError as error:
        _fail(f"cannot enumerate bundle: {error}")
    for path in paths:
        if path.is_symlink():
            _fail(f"bundle contains a symlink: {path}")
        if path.is_dir():
            if "__pycache__" in path.parts:
                _fail(f"remove generated bytecode directory before sealing: {path}")
            continue
        if not path.is_file():
            _fail(f"bundle contains a non-regular file: {path}")
        if _is_bytecode(path):
            _fail(f"remove generated bytecode before sealing: {path}")
        relative = path.relative_to(root).as_posix()
        if relative == SEAL_NAME or relative in EXCLUDED["verification_receipts"]:
            continue
        result[relative] = _meta(path, f"bundle file {relative}")
    return result


def _tree_inventory(directory: Path, label: str) -> dict[str, dict[str, int | str]]:
    """Rehash one retained tree using paths relative to its directory."""

    _require(directory.is_dir() and not directory.is_symlink(), f"{label}: directory is missing: {directory}")
    result: dict[str, dict[str, int | str]] = {}
    try:
        paths = sorted(directory.rglob("*"), key=lambda item: item.relative_to(directory).as_posix())
    except OSError as error:
        _fail(f"{label}: cannot enumerate {directory}: {error}")
    for path in paths:
        if path.is_symlink():
            _fail(f"{label}: symlink is not allowed: {path}")
        if path.is_dir():
            continue
        if not path.is_file():
            _fail(f"{label}: non-regular file: {path}")
        _require(not _is_bytecode(path), f"{label}: bytecode is not allowed: {path}")
        relative = path.relative_to(directory).as_posix()
        result[relative] = _meta(path, f"{label}/{relative}")
    return result


def _descriptor(root: Path, path: Path, label: str, *, allow_external: bool = False) -> dict[str, int | str]:
    resolved = path.resolve(strict=True)
    try:
        relative = resolved.relative_to(root).as_posix()
    except ValueError:
        _require(allow_external, f"{label}: path is outside root: {path}")
        relative = str(resolved)
    return {"path": relative, **_meta(resolved, label)}


def _descriptor_path(root: Path, value: Mapping[str, Any], label: str, *, allow_external: bool = False) -> Path:
    _require(isinstance(value, Mapping), f"{label}: descriptor must be an object")
    path_value = value.get("path")
    _require(isinstance(path_value, str) and path_value, f"{label}.path: missing path")
    candidate = Path(path_value)
    if candidate.is_absolute():
        try:
            resolved = candidate.resolve(strict=True)
            resolved.relative_to(root)
        except ValueError:
            _require(allow_external, f"{label}: external path is not allowed: {candidate}")
        except OSError as error:
            _fail(f"{label}: missing path {candidate}: {error}")
    else:
        resolved, _ = _inside(root, candidate, label)
    _meta(resolved, label)
    actual = _meta(resolved, label)
    _require(value.get("bytes") == actual["bytes"], f"{label}: byte length changed")
    _require(value.get("sha256") == actual["sha256"], f"{label}: SHA-256 changed")
    return resolved


def _write_exclusive(path: Path, value: Mapping[str, Any]) -> None:
    _require(not path.exists(), f"refusing to replace existing artifact: {path}")
    try:
        with path.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
            stream.write("\n")
    except OSError as error:
        _fail(f"cannot write {path}: {error}")


def _safe_token(value: Any, label: str) -> str:
    _require(isinstance(value, str) and TOKEN_RE.fullmatch(value) is not None, f"{label}: invalid path-safe token")
    return value


def _gate_record(root: Path, value: str | Path) -> dict[str, Any]:
    path, relative = _inside(root, value, "gate receipt")
    receipt = _read_json(path, f"gate receipt {relative}")
    _require(isinstance(receipt, Mapping), f"gate receipt {relative}: expected an object")
    _require(receipt.get("schema") == "docx-stream-append-gate-v1", f"gate receipt {relative}: schema differs")
    _require(isinstance(receipt.get("label"), str) and receipt["label"], f"gate receipt {relative}: label missing")
    _require(type(receipt.get("exit_code")) is int and receipt["exit_code"] == 0, f"gate receipt {relative}: exit_code is not zero")
    _require(receipt.get("source_unchanged") is True, f"gate receipt {relative}: source is not stable")
    _require(receipt.get("source_before") == receipt.get("source_after"), f"gate receipt {relative}: source snapshots differ")
    return {
        "path": relative,
        **_meta(path, f"gate receipt {relative}"),
        "schema": receipt["schema"],
        "label": receipt["label"],
        "attempt": receipt.get("attempt"),
        "exit_code": 0,
        "source_unchanged": True,
        "stable": True,
    }


def _gates_from_paths(root: Path, values: list[str]) -> list[dict[str, Any]]:
    _require(values, "seal requires at least one --gate final receipt")
    records = [_gate_record(root, value) for value in values]
    paths = [record["path"] for record in records]
    _require(len(paths) == len(set(paths)), "seal gate selection contains duplicates")
    return records


def _verify_gates(root: Path, records: Any) -> list[dict[str, Any]]:
    _require(isinstance(records, list) and records, "seal manifest has no selected final gates")
    actual: list[dict[str, Any]] = []
    for index, record in enumerate(records):
        _require(isinstance(record, Mapping), f"seal.gates[{index}]: expected an object")
        path = record.get("path")
        _require(isinstance(path, str) and path, f"seal.gates[{index}].path: missing path")
        current = _gate_record(root, path)
        for key in ("path", "bytes", "sha256", "schema", "label", "attempt", "exit_code", "source_unchanged", "stable"):
            _require(record.get(key) == current.get(key), f"seal.gates[{index}].{key}: binding changed")
        actual.append(current)
    paths = [record["path"] for record in actual]
    _require(len(paths) == len(set(paths)), "seal manifest selects duplicate gates")
    return actual


def _check_gate_sources(root: Path, gates: list[Mapping[str, Any]], expected_source: Any) -> None:
    for index, record in enumerate(gates):
        path, relative = _inside(root, record["path"], f"selected gate {index}")
        receipt = _read_json(path, f"selected gate {relative}")
        _require(isinstance(receipt, Mapping) and receipt.get("source_after") == expected_source, f"selected gate {relative}: source_after differs from the after build source")


def _profile_record(root: Path, value: str | Path, expected_pass_count: int | None = None) -> dict[str, Any]:
    result_path, result_relative = _inside(root, value, "profile result")
    result = _read_json(result_path, f"profile result {result_relative}")
    _require(isinstance(result, Mapping), f"profile result {result_relative}: expected an object")
    _require(result.get("status") == "pass", f"profile result {result_relative}: terminal status is not pass")
    records = result.get("records")
    _require(isinstance(records, list) and records, f"profile result {result_relative}: receipt inventory is empty")
    pass_count = 0
    receipts: list[dict[str, Any]] = []
    labels: set[str] = set()
    for index, item in enumerate(records):
        label = f"profile result {result_relative}.records[{index}]"
        _require(isinstance(item, Mapping), f"{label}: expected an object")
        name = item.get("label")
        _require(isinstance(name, str) and name and name not in labels, f"{label}.label: duplicate or missing label")
        labels.add(name)
        _require(item.get("status") == "pass", f"{label}: profile did not pass")
        pass_count += 1
        reference = item.get("receipt")
        receipt_path = _descriptor_path(root, reference, f"{label}.receipt")
        receipt = _read_json(receipt_path, f"profile receipt {name}")
        _require(isinstance(receipt, Mapping) and receipt.get("status") == "pass", f"profile receipt {name}: status is not pass")
        process = receipt.get("process")
        _require(isinstance(process, Mapping) and process.get("returncode") == 0 and not process.get("timed_out"), f"profile receipt {name}: process did not pass")
        binary = receipt.get("binary")
        _require(isinstance(binary, Mapping) and isinstance(binary.get("path"), str), f"profile receipt {name}: binary binding is missing")
        binary_path = _any_file(binary["path"], f"profile binary {name}")
        binary_meta = _meta(binary_path, f"profile binary {name}")
        _require(binary.get("bytes") == binary_meta["bytes"] and binary.get("sha256") == binary_meta["sha256"], f"profile binary changed: {name}")
        artifacts = receipt.get("artifacts")
        _require(isinstance(artifacts, Mapping) and artifacts, f"profile receipt {name}: artifacts are missing")
        for artifact_name, artifact in artifacts.items():
            _require(isinstance(artifact_name, str) and Path(artifact_name).name == artifact_name, f"profile receipt {name}: invalid artifact name")
            artifact_path = receipt_path.parent / artifact_name
            _require(isinstance(artifact, Mapping), f"profile receipt {name}: malformed artifact {artifact_name}")
            actual = _meta(artifact_path, f"profile artifact {name}/{artifact_name}")
            _require(artifact.get("bytes") == actual["bytes"] and artifact.get("sha256") == actual["sha256"], f"profile artifact changed: {name}/{artifact_name}")
        receipts.append(_descriptor(root, receipt_path, f"profile receipt {name}"))
    _require(expected_pass_count is None or pass_count == expected_pass_count, f"profile result {result_relative}: expected {expected_pass_count} passing receipts, found {pass_count}")
    _require(set(labels) == PROFILE_EXPECTED_LABELS, f"profile result {result_relative}: label set differs from the fixed 12-case lane")
    return {
        "result": _descriptor(root, result_path, f"profile result {result_relative}"),
        "record_count": len(records),
        "pass_count": pass_count,
        "receipts": receipts,
    }


def _verify_profile(root: Path, record: Any) -> dict[str, Any]:
    _require(isinstance(record, Mapping), "seal.profile: expected an object")
    result_path = _descriptor_path(root, record.get("result"), "seal.profile.result")
    actual = _profile_record(root, result_path, expected_pass_count=PROFILE_EXPECTED_PASSES)
    _require(actual["pass_count"] == PROFILE_EXPECTED_PASSES, f"seal.profile: expected {PROFILE_EXPECTED_PASSES} passing receipts")
    for key in ("result", "record_count", "pass_count", "receipts"):
        _require(record.get(key) == actual.get(key), f"seal.profile.{key}: binding changed")
    return actual


def _fuzz_record(root: Path, value: str | Path) -> dict[str, Any]:
    smoke_path, smoke_relative = _inside(root, value, "fuzz smoke result")
    smoke = _read_json(smoke_path, f"fuzz smoke result {smoke_relative}")
    _require(isinstance(smoke, Mapping) and smoke.get("schema") == "docx-stream-fuzz-smoke-v1", f"fuzz smoke result {smoke_relative}: schema differs")
    _require(smoke.get("input_cap_bytes") == FUZZ_INPUT_CAP, f"fuzz smoke result {smoke_relative}: input cap differs")

    data_dir = smoke_path.parent
    build_path = data_dir / "build.json"
    prepared_path = data_dir / "prepared.json"
    build = _read_json(build_path, "fuzz build result")
    prepared = _read_json(prepared_path, "fuzz prepared result")
    _require(isinstance(build, Mapping) and build.get("schema") == "docx-stream-fuzz-build-v1", "fuzz build result: schema differs")
    _require(isinstance(prepared, Mapping) and prepared.get("schema") == "docx-stream-fuzz-prepared-v1", "fuzz prepared result: schema differs")
    _require(build.get("source_snapshot_before") == build.get("source_snapshot_after"), "fuzz build result: source changed during build")
    seed_records = prepared.get("seed_records")
    _require(isinstance(seed_records, Mapping) and len(seed_records) == FUZZ_SEED_COUNT, "fuzz prepared result: seed inventory cardinality differs")
    expected_seeds: dict[str, dict[str, int | str]] = {}
    expected_positive_names: list[str] = []
    for name, record in sorted(seed_records.items()):
        _require(isinstance(name, str) and name and Path(name).name == name, f"fuzz prepared result: invalid seed name {name!r}")
        _require(isinstance(record, Mapping), f"fuzz prepared result: malformed seed record {name}")
        _require(type(record.get("bytes")) is int and record["bytes"] >= 0, f"fuzz prepared result: malformed byte count for {name}")
        _require(isinstance(record.get("sha256"), str), f"fuzz prepared result: malformed SHA-256 for {name}")
        _require(record.get("input_cap_bytes") == FUZZ_INPUT_CAP, f"fuzz prepared result: input cap differs for {name}")
        expected_seeds[name] = {"bytes": record["bytes"], "sha256": record["sha256"]}
        if record.get("expected_prepare_success") is True:
            expected_positive_names.append(name)
    expected_positive_names.sort()
    _require(len(expected_positive_names) == FUZZ_POSITIVE_CASES, "fuzz prepared result: positive seed cardinality differs")
    actual_seeds = _tree_inventory(data_dir / "seeds", "fuzz seed tree")
    _require(actual_seeds == expected_seeds, "fuzz seed bytes or hashes differ from prepared seed_records")

    positive = smoke.get("positive_check")
    _require(isinstance(positive, Mapping) and positive.get("schema") == "docx-stream-fuzz-positive-v1", f"fuzz smoke result {smoke_relative}: positive check is missing")
    _require(positive.get("input_cap_bytes") == FUZZ_INPUT_CAP, f"fuzz smoke result {smoke_relative}: positive input cap differs")
    cases = positive.get("cases")
    _require(isinstance(cases, list) and len(cases) == FUZZ_POSITIVE_CASES and positive.get("case_count") == FUZZ_POSITIVE_CASES and positive.get("exit_code") == 0, f"fuzz smoke result {smoke_relative}: positive corpus did not complete")
    binary = smoke.get("binary")
    _require(isinstance(binary, Mapping) and isinstance(binary.get("path"), str) and binary["path"], "fuzz smoke result: binary binding is missing")
    positive_path = data_dir / "positive.json"
    embedded_positive = _read_json(positive_path, "fuzz positive receipt")
    _require(embedded_positive == positive, "fuzz positive receipt differs from the smoke embedding")
    positive_root = data_dir / "positive-check"
    expected_case_directories: set[str] = set()
    for index, case in enumerate(cases):
        _require(isinstance(case, Mapping) and case.get("name") == expected_positive_names[index] and case.get("exit_code") == 0, f"fuzz positive case {index} did not pass or is out of order")
        case_argv = case.get("argv")
        _require(isinstance(case_argv, list) and f"-max_len={FUZZ_INPUT_CAP}" in case_argv, f"fuzz positive case {index}: input cap binding differs")
        case_name = case["name"]
        directory_name = f"{index + 1:03d}-{case_name}"
        expected_case_directories.add(directory_name)
        case_root = positive_root / directory_name
        actual_artifacts = _tree_inventory(case_root, f"fuzz positive case {index} artifacts")
        _require(case.get("artifacts") == actual_artifacts, f"fuzz positive case {index}: artifact inventory changed")
    _require(positive_root.is_dir() and {path.name for path in positive_root.iterdir()} == expected_case_directories, "fuzz positive-check directory inventory differs")

    runs = smoke.get("runs")
    _require(isinstance(runs, list) and len(runs) == FUZZ_RUN_COUNT, f"fuzz smoke result {smoke_relative}: expected exactly {FUZZ_RUN_COUNT} runs")
    post_run_root = data_dir / "post-run"
    expected_run_directories = {f"run-{index}" for index in range(1, FUZZ_RUN_COUNT + 1)}
    _require(post_run_root.is_dir() and {path.name for path in post_run_root.iterdir()} == expected_run_directories, "fuzz post-run directory inventory differs")
    for index, (seed, run) in enumerate(zip(FUZZ_RUN_SEEDS, runs), start=1):
        _require(isinstance(run, Mapping) and run.get("exit_code") == 0 and isinstance(run.get("retained"), Mapping) and run["retained"], f"fuzz run {index}: incomplete retained result")
        run_argv = run.get("argv")
        _require(isinstance(run_argv, list) and len(run_argv) == 7, f"fuzz run {index}: argv shape differs")
        _require(run_argv[2] == "-runs=10000" and run_argv[3] == f"-seed={seed}" and run_argv[4] == f"-max_len={FUZZ_INPUT_CAP}" and run_argv[5] == "-timeout=10", f"fuzz run {index}: run count, seed, or input cap binding differs")
        _require(isinstance(run_argv[0], str) and isinstance(run_argv[1], str) and run_argv[0] == binary["path"] and Path(run_argv[1]).name == "corpus", f"fuzz run {index}: executable or corpus binding differs")
        _require(isinstance(run_argv[6], str) and run_argv[6].endswith(f"/run-{index}/"), f"fuzz run {index}: artifact prefix binding differs")
        _require(isinstance(run.get("corpus_before"), Mapping) and run["corpus_before"], f"fuzz run {index}: corpus-before inventory is missing")
        retained_root = post_run_root / f"run-{index}"
        _require({path.name for path in retained_root.iterdir()} == {"corpus", "artifacts"}, f"fuzz run {index}: retained tree roots differ")
        actual_retained = _tree_inventory(retained_root, f"fuzz retained run {index}")
        _require(run["retained"] == actual_retained, f"fuzz run {index}: retained artifact inventory changed")

    binary_path = _any_file(binary["path"], "fuzz binary")
    actual_binary = {"path": str(binary_path), **_meta(binary_path, "fuzz binary")}
    for key in ("bytes", "sha256"):
        _require(binary.get(key) == actual_binary[key], f"fuzz binary {key} changed")
    build_binary = build.get("binary")
    _require(isinstance(build_binary, Mapping), "fuzz build result: binary binding is missing")
    _require(build_binary.get("path") == binary.get("path") and build_binary.get("bytes") == binary.get("bytes") and build_binary.get("sha256") == binary.get("sha256"), "fuzz smoke/build binary bindings differ")
    return {
        "smoke": _descriptor(root, smoke_path, f"fuzz smoke result {smoke_relative}"),
        "build": _descriptor(root, build_path, "fuzz build result"),
        "prepared": _descriptor(root, prepared_path, "fuzz prepared result"),
        "binary": actual_binary,
        "run_count": len(runs),
        "positive_case_count": len(cases),
    }


def _verify_fuzz(root: Path, record: Any) -> dict[str, Any]:
    _require(isinstance(record, Mapping), "seal.fuzz: expected an object")
    smoke_path = _descriptor_path(root, record.get("smoke"), "seal.fuzz.smoke")
    actual = _fuzz_record(root, smoke_path)
    for key in ("smoke", "build", "prepared", "binary", "run_count", "positive_case_count"):
        _require(record.get(key) == actual.get(key), f"seal.fuzz.{key}: binding changed")
    return actual


def _load_compare(root: Path) -> Any:
    path = str(root)
    if path not in sys.path:
        sys.path.insert(0, path)
    try:
        return importlib.import_module("compare")
    except Exception as error:
        _fail(f"comparison helper cannot be imported read-only: {error}")
    raise AssertionError("unreachable")


def _comparison_record(root: Path, value: str | Path) -> dict[str, Any]:
    summary_path, summary_relative = _inside(root, value, "comparison summary")
    summary = _read_json(summary_path, f"comparison summary {summary_relative}")
    _require(isinstance(summary, Mapping), f"comparison summary {summary_relative}: expected an object")
    _require(summary.get("schema") == COMPARISON_SUMMARY_SCHEMA, f"comparison summary {summary_relative}: schema differs")
    attempts = summary.get("attempts")
    _require(isinstance(attempts, Mapping), f"comparison summary {summary_relative}: attempts are missing")
    before_attempt = _safe_token(attempts.get("before"), "comparison attempts.before")
    after_attempt = _safe_token(attempts.get("after"), "comparison attempts.after")
    expected_inventory = {
        "arms": 18,
        "formal_processes": 144,
        "formal_samples": 4320,
        "samples_per_process": 30,
        "warmups_per_process": 3,
        "roles": ["normal", "allocator"],
        "repeats": [1, 2],
        "phases": ["before", "after"],
        "pilots_included": False,
    }
    _require(summary.get("inventory") == expected_inventory, f"comparison summary {summary_relative}: formal inventory is not 144 processes/4320 samples")

    compare = _load_compare(root)
    protocol_value = getattr(compare, "PROTOCOL_FILE", None)
    _require(isinstance(protocol_value, str) and protocol_value, "comparison helper protocol path is missing")
    protocol_candidate = Path(protocol_value)
    if not protocol_candidate.is_absolute():
        protocol_candidate = root / protocol_candidate
    protocol_path, protocol_relative = _inside(root, protocol_candidate, "comparison protocol")
    protocol_bytes = _sha(protocol_path)
    try:
        protocol, protocol_hash = compare._load_protocol()
        _require(protocol == _read_json(protocol_path, f"comparison protocol {protocol_relative}"), "comparison helper loaded a different protocol payload")
        _require(protocol_hash == protocol_bytes, "comparison helper protocol hash differs from the retained protocol file")
        builds = compare._load_builds(protocol)
        source = compare._source_summary(builds)
        capture_gates = {
            "before": compare._validate_capture_gate("before", before_attempt, builds["before"]["normal"]["source"]),
            "after": compare._validate_capture_gate("after", after_attempt, builds["after"]["normal"]["source"]),
        }
        before_rows = compare._validate_formal_phase(protocol, builds["before"], "before", before_attempt)
        after_rows = compare._validate_formal_phase(protocol, builds["after"], "after", after_attempt)
        rows = before_rows + after_rows
        compare.require(len(rows) == 144, "comparison formal process cardinality differs")
        compare._check_content_identity(rows)
        archive_changes = compare._candidate_archive_changes(rows)
        comparisons = compare._comparisons(rows)
    except Exception as error:
        _fail(f"comparison retained artifacts failed read-only validation: {error}")

    expected_analyzer = {"path": "compare.py", "sha256": _sha(root / "compare.py")}
    _require(summary.get("analyzer") == expected_analyzer, "comparison summary analyzer binding differs")
    _require(summary.get("protocol") == {"path": compare.PROTOCOL_FILE, "sha256": protocol_hash}, "comparison summary protocol binding differs")
    _require(summary.get("source_comparison") == source, "comparison summary source comparison differs")
    _require(summary.get("capture_gates") == capture_gates, "comparison summary capture-gate bindings differ")
    _require(summary.get("processes") == rows, "comparison summary process rows differ from retained captures")
    _require(summary.get("candidate_archive_changes") == archive_changes, "comparison summary archive comparison differs")
    _require(summary.get("comparisons") == comparisons, "comparison summary metric comparisons differ")

    expected_builds: dict[str, Any] = {}
    build_records: dict[str, Any] = {}
    external_binaries: dict[str, Any] = {}
    for phase in compare.PHASES:
        expected_builds[phase] = {}
        for role in compare.ROLE_NAMES:
            build = builds[phase][role]
            expected_builds[phase][role] = {
                "record_path": str(build["path"]),
                "record_sha256": build["sha256"],
                "binary": build["binary"],
                "source": build["source"],
                "gate": {"path": str(build["gate"]["path"]), "sha256": build["gate"]["sha256"]},
            }
            key = f"{phase}/{role}"
            build_records[key] = _descriptor(root, build["path"], f"comparison build record {key}", allow_external=True)
            external_binaries[key] = dict(build["binary"])
    _require(summary.get("builds") == expected_builds, "comparison summary build bindings differ")

    markdown_path = summary_path.with_name("comparison-summary.md")
    _inside(root, markdown_path, "comparison markdown")
    capture_descriptors = [
        _descriptor(root, root / gate["path"], f"comparison capture gate {phase}")
        for phase, gate in sorted(capture_gates.items())
    ]
    return {
        "summary": _descriptor(root, summary_path, f"comparison summary {summary_relative}"),
        "markdown": _descriptor(root, markdown_path, "comparison markdown"),
        "protocol": _descriptor(root, protocol_path, "comparison protocol"),
        "attempts": {"before": before_attempt, "after": after_attempt},
        "after_source": builds["after"]["normal"]["source"],
        "inventory": expected_inventory,
        "build_records": build_records,
        "capture_gates": capture_descriptors,
        "external_binaries": external_binaries,
    }


def _verify_comparison(root: Path, record: Any) -> dict[str, Any]:
    _require(isinstance(record, Mapping), "seal.comparison: expected an object")
    summary_path = _descriptor_path(root, record.get("summary"), "seal.comparison.summary")
    actual = _comparison_record(root, summary_path)
    for key in ("summary", "markdown", "protocol", "attempts", "after_source", "inventory", "build_records", "capture_gates", "external_binaries"):
        _require(record.get(key) == actual.get(key), f"seal.comparison.{key}: binding changed")
    return actual


def _seal_manifest(root: Path, gates: list[str], comparison: str, profile: str, fuzz: str) -> dict[str, Any]:
    gate_records = _gates_from_paths(root, gates)
    comparison_record = _comparison_record(root, comparison)
    _check_gate_sources(root, gate_records, comparison_record["after_source"])
    profile_record = _profile_record(root, profile, expected_pass_count=PROFILE_EXPECTED_PASSES)
    fuzz_record = _fuzz_record(root, fuzz)
    files = _inventory(root)
    return {
        "schema": SCHEMA,
        "version": VERSION,
        "sealed_utc": _now(),
        "root": ".",
        "scope": SCOPE,
        "excluded": EXCLUDED,
        "gates": gate_records,
        "comparison": comparison_record,
        "profile": profile_record,
        "fuzz": fuzz_record,
        "files": files,
    }


def _verify_manifest(root: Path, manifest: Mapping[str, Any]) -> dict[str, Any]:
    _require(manifest.get("schema") == SCHEMA and manifest.get("version") == VERSION, "seal schema/version differs")
    _require(isinstance(manifest.get("sealed_utc"), str) and manifest["sealed_utc"], "seal timestamp is missing")
    _require(manifest.get("root") == ".", "seal root binding differs")
    _require(manifest.get("scope") == SCOPE, "seal scope differs")
    _require(manifest.get("excluded") == EXCLUDED, "seal exclusion policy differs")
    files = _inventory(root)
    _require(manifest.get("files") == files, "sealed regular-file inventory or hash differs")
    gates = _verify_gates(root, manifest.get("gates"))
    comparison = _verify_comparison(root, manifest.get("comparison"))
    _check_gate_sources(root, gates, comparison["after_source"])
    profile = _verify_profile(root, manifest.get("profile"))
    fuzz = _verify_fuzz(root, manifest.get("fuzz"))
    return {"files": files, "gates": gates, "comparison": comparison, "profile": profile, "fuzz": fuzz}


def _verification_result(root: Path, checked: Mapping[str, Any]) -> dict[str, Any]:
    manifest_path = root / SEAL_NAME
    verifier_path = Path(__file__).resolve()
    return {
        "schema": VERIFICATION_SCHEMA,
        "status": "pass",
        "verified_utc": _now(),
        "manifest": _descriptor(root, manifest_path, "seal manifest"),
        "verifier": _descriptor(root, verifier_path, "seal verifier", allow_external=True),
        "files": len(checked["files"]),
        "final_gates": len(checked["gates"]),
        "formal_processes": checked["comparison"]["inventory"]["formal_processes"],
        "formal_samples": checked["comparison"]["inventory"]["formal_samples"],
        "profile_records": checked["profile"]["record_count"],
        "profile_passes": checked["profile"]["pass_count"],
        "fuzz_runs": checked["fuzz"]["run_count"],
    }


def _parse_args(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", nargs="?", choices=("seal", "verify"), default="verify")
    parser.add_argument("--root", type=Path, default=DEFAULT_ROOT)
    parser.add_argument("--gate", action="append", default=[], help="final validation receipt; repeat for every selected gate")
    parser.add_argument("--comparison", help="retained comparison-summary.json for seal")
    parser.add_argument("--profile", help="retained profiles/<attempt>/result.json for seal")
    parser.add_argument("--fuzz", help="retained fuzz smoke.json for seal")
    parser.add_argument("--receipt", action="store_true", help="write the excluded seal-verification.json receipt")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    try:
        root = _rooted(args.root)
        seal_path = root / SEAL_NAME
        receipt_path = root / VERIFICATION_NAME
        if args.command == "seal":
            _require(not seal_path.exists(), f"refusing to replace existing seal: {seal_path}")
            _require(args.comparison is not None and args.profile is not None and args.fuzz is not None, "seal requires --comparison, --profile, and --fuzz")
            manifest = _seal_manifest(root, args.gate, args.comparison, args.profile, args.fuzz)
            _write_exclusive(seal_path, manifest)
            print(json.dumps({"schema": SCHEMA, "status": "sealed", "files": len(manifest["files"]), "final_gates": len(manifest["gates"])}, sort_keys=True))
            return 0

        _require(not args.gate and args.comparison is None and args.profile is None and args.fuzz is None, "selection options are only valid with seal")
        _require(seal_path.is_file() and not seal_path.is_symlink(), f"seal manifest is missing: {seal_path}")
        manifest = _read_json(seal_path, "seal manifest")
        _require(isinstance(manifest, Mapping), "seal manifest: expected an object")
        checked = _verify_manifest(root, manifest)
        result = _verification_result(root, checked)
        if args.receipt:
            _write_exclusive(receipt_path, result)
        print(json.dumps(result, sort_keys=True))
        return 0
    except (SealError, OSError, TypeError, ValueError, KeyError) as error:
        print(f"verify_seal.py: error: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
