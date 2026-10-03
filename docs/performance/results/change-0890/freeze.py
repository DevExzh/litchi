#!/usr/bin/env python3
"""Freeze and verify the reviewed 0890 matched experiment inputs.

The command has no build or workload side effects.  Without an option it
validates the final v2 source freeze, retained 0885 release evidence, the
candidate build and quality receipts, and the four managed executable-copy
records before creating ``plan.json``.  ``--verify`` reopens that plan.  With
``--retained-only`` it accepts removed managed copies and owned build roots,
while still requiring the retained source, binary, receipt, and helper
evidence needed for replay.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import platform
import stat
import sys
from typing import Any

import experiment


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
BASELINE_PACKET = PACKET.parent / "change-0885"
BASELINE_QUALIFICATION = PACKET.parent / "change-0887" / "baseline-qualification"
FREEZE = PACKET / "candidate-freeze-v2.json"
PATCH = PACKET / "candidate-v2.patch"
SOURCE_REVIEW = PACKET / "source-review-v2.json"
METHODOLOGY_REVIEW = PACKET / "methodology-review-v3.json"
BUILD = PACKET / "build.json"
BUILD_START = PACKET / "build-start.json"
BUILD_EXECUTED = PACKET / "build-executed.py"
RELEASE_BUILD_LAUNCH = PACKET / "release-build-launch.json"
RELEASE_BUILD_COMPLETION = PACKET / "release-build-completion.json"
EXECUTION = PACKET / "execution-files.json"
QUALITY = PACKET / "quality-v2"
QUALITY_QUALIFICATION = PACKET / "quality-qualification.json"
OWNER_GATES_ATTEMPT = PACKET / "attempts/owner-gates-v2"
OWNER_GATES_RECEIPT = OWNER_GATES_ATTEMPT / "owner-gates-receipt.json"
PLAN = PACKET / "plan.json"
BASELINE_REVISION = "335a40350bf06f49691745c14d310fd7a3ca9117"
CANDIDATE_BASE = "f790c2228cbd9e89c2892ba3ffaeb49802fac43b"
QUALITY_GATES = ("fmt", "owner-tests", "clippy", "rustdoc", "shared-format-tests", "boundaries")
QUALITY_QUALIFICATION_SCHEMA = "litchi.mce-0890-quality-qualification.v1"
OWNER_GATE_COMMANDS = (
    "mce_owner_tests",
    "xlsx_shared_tests",
    "docx_shared_tests",
    "pptx_shared_tests",
    "fmt_check",
    "mce_owner_clippy",
    "mce_owner_rustdoc",
    "crate_boundary_check",
)
OWNER_CONTINUATION_COMMANDS = (
    "isolated_workspace_fmt",
    "root_owner_fmt",
    "mce_owner_unfiltered_tests",
)
ROLES = ("before", "after")
MODES = ("native", "allocator")
BUILD_ENV_KEYS = (
    "CARGO_BUILD_JOBS",
    "CARGO_INCREMENTAL",
    "CARGO_NET_OFFLINE",
    "CARGO_PROFILE_RELEASE_OPT_LEVEL",
    "CARGO_PROFILE_RELEASE_DEBUG",
    "CARGO_PROFILE_RELEASE_LTO",
    "CARGO_PROFILE_RELEASE_CODEGEN_UNITS",
    "CARGO_HOME",
    "HOME",
    "PATH",
    "LANG",
    "LC_ALL",
    "TZ",
)
BUILD_NAMES = {"native": "litchi-perf-baseline", "allocator": "litchi-perf-baseline-alloc"}
IMPORTED_HELPERS = {
    "candidate_build": PACKET / "build.py",
    "candidate_build_executed": BUILD_EXECUTED,
    "candidate_experiment": PACKET / "experiment.py",
    "candidate_reducer": PACKET / "reduce.py",
    "candidate_packet_verifier": PACKET / "verify_packet.py",
    "candidate_managed_copy_preparer": PACKET / "prepare_exec.py",
    "candidate_baseline_runner": PACKET / "baseline_runner.py",
    "candidate_baseline_source_verifier": PACKET / "verify_baseline_source.py",
    "candidate_diagnostic_overlay": PACKET / "overlay_candidate_diagnostic.py",
    "candidate_diagnostic_runner": PACKET / "candidate_diagnostic_runner.py",
    "candidate_diagnostic_comparator": PACKET / "compare_diagnostic_reports.py",
    "candidate_source_verifier": PACKET / "verify_candidate_source.py",
    "owner_gates_runner": PACKET / "owner_gates_runner.py",
    "owner_continuation_runner": PACKET / "owner_continuation_runner.py",
    "candidate_freeze_builder": PACKET / "freeze.py",
    "imported_0887_experiment": PACKET.parent / "change-0887" / "experiment.py",
    "imported_0887_reducer": PACKET.parent / "change-0887" / "reduce.py",
    "imported_0887_packet_verifier": PACKET.parent / "change-0887" / "verify_packet.py",
    "imported_0885_capture": BASELINE_PACKET / "capture.py",
    "imported_0885_verify": BASELINE_PACKET / "verify.py",
    "imported_0885_test_helpers": BASELINE_PACKET / "test_helpers.py",
}


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def read(path: Path) -> Any:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    info = path.stat()
    require(stat.S_ISREG(info.st_mode), f"expected regular file: {path}")
    return {
        "path": str(path),
        "bytes": info.st_size,
        "sha256": sha256(path),
        "mode": stat.S_IMODE(info.st_mode),
        "executable": bool(info.st_mode & stat.S_IXUSR),
    }


def path_of(raw: Any, *, base: Path = PACKET) -> Path:
    require(isinstance(raw, str), "descriptor path is missing")
    value = Path(raw)
    return (value if value.is_absolute() else base / value).resolve()


def check_descriptor(raw: Any, label: str, *, base: Path = PACKET, live: bool = True) -> dict[str, Any]:
    require(isinstance(raw, dict), f"{label} descriptor is not an object")
    path = path_of(raw.get("path"), base=base)
    if not live:
        return {
            "path": str(path),
            "bytes": raw.get("bytes"),
            "sha256": raw.get("sha256"),
            "mode": raw.get("mode"),
            "executable": raw.get("executable"),
        }
    actual = descriptor(path)
    for field in ("bytes", "sha256", "mode", "executable"):
        if field in raw:
            require(raw[field] == actual[field], f"{label} {field} does not match the retained file")
    return actual


def same_identity(actual: dict[str, Any], expected: dict[str, Any], label: str) -> None:
    require(actual.get("bytes") == expected.get("bytes"), f"{label} byte count differs")
    require(actual.get("sha256") == expected.get("sha256"), f"{label} SHA-256 differs")


def _json_descriptor(path: Path, label: str, *, live: bool = True) -> tuple[dict[str, Any], dict[str, Any]]:
    require(path.is_file(), f"{label} is missing: {path}")
    value = read(path)
    require(isinstance(value, dict), f"{label} must be a JSON object")
    return value, descriptor(path) if live else {"path": str(path.resolve())}


def _owner_continuation_receipt() -> Path | None:
    """Find a sealed unfiltered owner-gate continuation, if one exists."""

    candidates: list[Path] = []
    paths = list((PACKET / "attempts").glob("*/**/*receipt*.json")) + list(PACKET.glob("*owner*receipt*.json"))
    for path in sorted(set(paths)):
        if path == OWNER_GATES_RECEIPT:
            continue
        try:
            value = read(path)
        except (OSError, ValueError, TypeError):
            continue
        if value.get("schema") != "litchi.mce-0890-owner-continuation.v1" or value.get("status") != "pass":
            continue
        candidates.append(path)
    return candidates[0] if candidates else None


def _receipt_reference(raw: Any, label: str, expected_path: Path) -> dict[str, Any]:
    """Validate a receipt-owned path/hash reference and return its live identity."""

    if isinstance(raw, str):
        path = path_of(raw)
        raw = {}
    else:
        require(isinstance(raw, dict), f"{label} reference is missing")
        path = path_of(raw.get("path"))
    require(path == expected_path.resolve(), f"{label} path is not the expected retained receipt")
    actual = descriptor(path)
    if "bytes" in raw:
        require(raw["bytes"] == actual["bytes"], f"{label} byte count differs")
    if "sha256" in raw:
        require(raw["sha256"] == actual["sha256"], f"{label} SHA-256 differs")
    return actual


def _validate_initial_owner_gates() -> tuple[dict[str, Any], dict[str, Any]]:
    require(OWNER_GATES_RECEIPT.is_file(), f"initial owner-gates receipt is missing: {OWNER_GATES_RECEIPT}")
    receipt = read(OWNER_GATES_RECEIPT)
    require(receipt.get("schema") == "litchi.mce-0890-owner-gates.v1", "initial owner-gates receipt schema mismatch")
    require(receipt.get("status") == "fail", "initial owner-gates receipt must retain its known fmt failure")
    require(receipt.get("performance_claim") is False, "initial owner-gates receipt claims performance")
    commands = receipt.get("commands")
    require(isinstance(commands, list) and [item.get("name") for item in commands] == list(OWNER_GATE_COMMANDS), "initial owner-gates command set is not exact")
    for item in commands:
        require(item.get("expected") == "pass", f"initial owner-gates command expectation is not pass: {item.get('name')}")
        if item.get("name") == "fmt_check":
            require(item.get("returncode") == 1 and item.get("pass") is False, "initial owner-gates fmt failure changed")
        else:
            require(item.get("returncode") == 0 and item.get("pass") is True, f"initial owner-gates command failed: {item.get('name')}")
    require(receipt.get("failures") == ["fmt_check"], "initial owner-gates has a failure besides the known fmt check")
    require(receipt.get("source_before") == receipt.get("source_after"), "initial owner-gates changed source while running")
    return receipt, descriptor(OWNER_GATES_RECEIPT)


def _validate_owner_continuation(initial: dict[str, Any], *, require_live: bool) -> tuple[Path, dict[str, Any], dict[str, Any]]:
    continuation = _owner_continuation_receipt()
    require(continuation is not None, "unfiltered owner-gates continuation receipt is required before plan freeze")
    receipt = read(continuation)
    require(receipt.get("schema") == "litchi.mce-0890-owner-continuation.v1", "owner-gates continuation schema mismatch")
    require(receipt.get("status") == "pass" and receipt.get("performance_claim") is False, "owner-gates continuation is not passing or claims performance")
    require(receipt.get("failures", []) == [], "owner-gates continuation retained a failure")
    commands = receipt.get("commands")
    require(isinstance(commands, list) and [item.get("name") for item in commands] == list(OWNER_CONTINUATION_COMMANDS), "owner-gates continuation command set is not exact")
    for item in commands:
        require(item.get("expected") == "pass" and item.get("returncode") == 0 and item.get("pass") is True, f"owner-gates continuation command failed: {item.get('name')}")

    source_before = receipt.get("source_before")
    source_after = receipt.get("source_after")
    require(isinstance(source_before, dict) and isinstance(source_after, dict), "owner-gates continuation source manifests are missing")
    require(source_before == source_after, "owner-gates continuation changed source while running")
    require(source_before.get("root") == receipt.get("source_root"), "owner-gates continuation source root binding is inconsistent")
    require(source_before.get("unexpected_files") == [] and source_before.get("missing_files") == [], "owner-gates continuation source inventory is not exact")

    candidate_receipt_path = PACKET / "attempts/candidate-diagnostic-f790c2228c-v2-retry1/candidate-receipt.json"
    candidate_receipt = read(candidate_receipt_path)
    candidate_binary_before = receipt.get("candidate_binary_before")
    candidate_binary_after = receipt.get("candidate_binary_after")
    require(isinstance(candidate_binary_before, dict) and candidate_binary_before == candidate_binary_after, "owner-gates continuation changed the diagnostic binary")
    require(candidate_binary_before == candidate_receipt.get("candidate_binary"), "owner-gates continuation diagnostic binary binding differs")

    _receipt_reference(receipt.get("prior_owner_attempt"), "owner-gates continuation prior receipt", OWNER_GATES_RECEIPT)
    analysis = receipt.get("prior_global_fmt_analysis")
    require(isinstance(analysis, dict), "owner-gates continuation global fmt analysis is missing")
    _receipt_reference(analysis.get("receipt"), "owner-gates continuation fmt analysis receipt", OWNER_GATES_RECEIPT)
    prior_fmt_log = OWNER_GATES_ATTEMPT / "logs/fmt_check.log"
    require(path_of(analysis.get("log")) == prior_fmt_log.resolve(), "owner-gates continuation fmt analysis log path differs")
    require(analysis.get("log_sha256") == sha256(prior_fmt_log), "owner-gates continuation fmt analysis log hash differs")
    unrelated = ROOT / "crates/litchi/examples/bench_word_parse.rs"
    unrelated_ref = analysis.get("unrelated_untracked_file")
    require(isinstance(unrelated_ref, dict), "owner-gates continuation unrelated fmt file reference is missing")
    require(path_of(unrelated_ref.get("path"), base=ROOT) == unrelated.resolve(), "owner-gates continuation unrelated fmt file path differs")
    if require_live:
        unrelated_actual = descriptor(unrelated)
        same_identity(unrelated_actual, unrelated_ref, "owner-gates continuation unrelated fmt file")
    else:
        require(isinstance(unrelated_ref.get("bytes"), int) and unrelated_ref["bytes"] >= 0, "retained unrelated fmt file byte count is missing")
        require(isinstance(unrelated_ref.get("sha256"), str) and len(unrelated_ref["sha256"]) == 64, "retained unrelated fmt file hash is missing")
    require("unrelated" in str(analysis.get("cause", "")).lower() and "fmt" in str(analysis.get("cause", "")).lower(), "owner-gates continuation does not record the known unrelated fmt cause")
    return continuation, receipt, descriptor(continuation)


def patch_paths(path: Path) -> set[str]:
    paths: set[str] = set()
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.startswith("diff --git a/"):
            continue
        parts = line.split(" ", 3)
        require(len(parts) == 4, f"malformed patch header: {line}")
        left = parts[2][2:]
        right = parts[3][2:]
        require(left == right or left == "/dev/null", f"candidate patch path mismatch: {line}")
        name = right if left == "/dev/null" else left
        safe = Path(name)
        require(not safe.is_absolute() and ".." not in safe.parts, f"unsafe candidate patch path: {name}")
        paths.add(name)
    require(paths, "candidate patch has no paths")
    return paths


def _load_freeze(*, require_live: bool) -> tuple[dict[str, Any], dict[str, dict[str, Any]], dict[str, Any]]:
    freeze = read(FREEZE)
    require(freeze.get("schema") == "litchi.mce-candidate-freeze-0890.v1", "candidate freeze schema mismatch")
    require(freeze.get("base_commit") == CANDIDATE_BASE, "candidate freeze base commit mismatch")
    paths = freeze.get("candidate_paths")
    entries = freeze.get("files")
    require(isinstance(paths, list) and paths and len(paths) == len(set(paths)) and all(isinstance(item, str) for item in paths), "candidate freeze paths are invalid")
    require(isinstance(entries, list) and len(entries) == len(paths), "candidate freeze file inventory is incomplete")
    files: dict[str, dict[str, Any]] = {}
    for entry in entries:
        require(isinstance(entry, dict) and isinstance(entry.get("path"), str), "candidate freeze file descriptor is invalid")
        name = entry["path"]
        require(name in paths and name not in files, f"candidate freeze has an unexpected or duplicate file: {name}")
        require(isinstance(entry.get("bytes"), int) and entry["bytes"] >= 0, f"candidate freeze byte count is invalid: {name}")
        require(isinstance(entry.get("sha256"), str) and len(entry["sha256"]) == 64, f"candidate freeze hash is invalid: {name}")
        files[name] = entry
    require(set(files) == set(paths), "candidate freeze path and file inventories differ")
    patch_record = freeze.get("patch")
    require(isinstance(patch_record, dict), "candidate freeze patch record is missing")
    require(patch_record.get("path") == "docs/performance/results/change-0890/candidate-v2.patch", "candidate freeze does not bind candidate-v2.patch")
    patch = ROOT / "docs/performance/results/change-0890/candidate-v2.patch"
    require(patch.is_file(), "candidate-v2.patch is missing")
    patch_identity = descriptor(patch)
    same_identity(patch_identity, patch_record, "candidate-v2.patch")
    require(patch_paths(patch) == set(paths), "candidate patch paths differ from candidate freeze")
    if require_live:
        for name, expected in files.items():
            current = descriptor(ROOT / name)
            same_identity(current, expected, f"candidate overlay {name}")
    return freeze, files, patch_identity


def _validate_source_review(freeze: dict[str, Any], files: dict[str, dict[str, Any]], patch_identity: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    review, review_binding = _json_descriptor(SOURCE_REVIEW, "source review")
    require(review.get("status") == "approved_for_measurement", "source review is not approved")
    require(review.get("base_head") == CANDIDATE_BASE, "source review base head differs from candidate freeze")
    require(review.get("blocking_findings") == [], "source review has blocking findings")
    require(review.get("scope", {}).get("candidate_paths") == freeze["candidate_paths"], "source review candidate paths differ from freeze")
    review_freeze = review.get("freeze", {})
    require(review_freeze.get("path") == "docs/performance/results/change-0890/candidate-freeze-v2.json", "source review binds the wrong freeze")
    patch = review.get("patch", {})
    require(patch.get("path") == "docs/performance/results/change-0890/candidate-v2.patch", "source review binds the wrong patch")
    same_identity(patch_identity, patch, "source review candidate patch")
    for name, expected in files.items():
        after = review.get("files", {}).get(name, {}).get("after", {})
        same_identity(after, expected, f"source review overlay {name}")
    return review, review_binding


def _validate_methodology() -> tuple[dict[str, Any], dict[str, Any]]:
    review, binding = _json_descriptor(METHODOLOGY_REVIEW, "methodology review")
    require(review.get("status") == "pass", "methodology review is not passing")
    require(review.get("reviewed_without_build_or_workload") is True, "methodology review scope is not static")
    require(review.get("scope", {}).get("base_head") == CANDIDATE_BASE, "methodology review base head mismatch")
    for key, path in IMPORTED_HELPERS.items():
        recorded = review.get("helpers", {}).get({
            "candidate_build": "build",
            "candidate_experiment": "experiment",
            "candidate_reducer": "reducer",
            "candidate_managed_copy_preparer": "managed_executable_preparation",
            "candidate_packet_verifier": "packet_verifier",
            "candidate_baseline_source_verifier": "baseline_source_verifier",
            "candidate_baseline_runner": "correctness_baseline_runner",
            "candidate_freeze_builder": "freeze",
            "candidate_diagnostic_overlay": "diagnostic_overlay",
            "candidate_diagnostic_runner": "diagnostic_runner",
            "candidate_diagnostic_comparator": "diagnostic_comparator",
            "candidate_source_verifier": "candidate_source_verifier",
            "owner_gates_runner": "owner_gates_runner",
            "owner_continuation_runner": "owner_continuation_runner",
        }.get(key, ""), {})
        if not recorded:
            continue
        actual = descriptor(path)
        require(recorded.get("bytes") == actual["bytes"] and recorded.get("sha256") == actual["sha256"], f"methodology helper hash is stale for {key}")
    return review, binding


def _validate_quality(*, require_live: bool) -> dict[str, Any]:
    initial_receipt, initial_binding = _validate_initial_owner_gates()
    continuation, continuation_receipt, continuation_binding = _validate_owner_continuation(initial_receipt, require_live=require_live)
    require(QUALITY_QUALIFICATION.is_file(), "aggregate quality-qualification.json is required before plan freeze")
    if QUALITY_QUALIFICATION.is_file():
        qualification = read(QUALITY_QUALIFICATION)
        require(isinstance(qualification, dict), "quality qualification must be an object")
        require(qualification.get("schema") == QUALITY_QUALIFICATION_SCHEMA, "quality qualification schema mismatch")
        require(qualification.get("status") == "pass" and qualification.get("performance_claim") is False, "quality qualification is not passing or claims performance")
        source = qualification.get("source")
        require(isinstance(source, dict) and source.get("base_commit") == CANDIDATE_BASE, "quality qualification source binding is incomplete")
        _receipt_reference(source.get("candidate_freeze"), "quality qualification candidate freeze", FREEZE)
        _receipt_reference(source.get("candidate_build"), "quality qualification candidate build", BUILD)
        correctness_review = PACKET / "correctness-review.json"
        _receipt_reference(source.get("correctness_review"), "quality qualification correctness review", correctness_review)
        owner = qualification.get("owner_quality_gates")
        require(isinstance(owner, dict), "quality qualification owner-gate section is missing")
        aggregate_initial = owner.get("initial_receipt")
        _receipt_reference(aggregate_initial, "quality qualification initial owner receipt", OWNER_GATES_RECEIPT)
        require(aggregate_initial.get("status") == "fail" and aggregate_initial.get("performance_claim") is False, "quality qualification does not retain the initial owner failure")
        require(owner.get("initial_receipt_status") == "fail" and owner.get("initial_receipt_failures") == ["fmt_check"], "quality qualification initial failure summary changed")
        aggregate_continuation = owner.get("continuation", {}).get("receipt")
        _receipt_reference(aggregate_continuation, "quality qualification owner continuation", continuation)
        require(aggregate_continuation.get("status") == "pass" and aggregate_continuation.get("performance_claim") is False, "quality qualification continuation is not passing")
        continuation_summary = owner.get("continuation", {})
        require(continuation_summary.get("status") == "pass" and continuation_summary.get("isolated_workspace_fmt") is True and continuation_summary.get("root_owner_fmt") is True, "quality qualification continuation summary is incomplete")
        tests = continuation_summary.get("mce_owner_unfiltered_tests", {})
        require(tests.get("passed") == 399 and tests.get("failed") == 0, "quality qualification does not record all 399 owner tests")
        resolution = owner.get("fmt_failure_resolution")
        require(isinstance(resolution, dict) and resolution.get("scope") == "only unrelated untracked crates/litchi/examples/bench_word_parse.rs", "quality qualification fmt resolution scope changed")
        prior_fmt_log = OWNER_GATES_ATTEMPT / "logs/fmt_check.log"
        _receipt_reference(resolution.get("failed_log"), "quality qualification failed fmt log", prior_fmt_log)
        expected_resolution_logs = (
            PACKET / "attempts/owner-continuation-v2/logs/isolated_workspace_fmt.log",
            PACKET / "attempts/owner-continuation-v2/logs/root_owner_fmt.log",
        )
        resolution_logs = resolution.get("resolution_logs")
        require(isinstance(resolution_logs, list) and len(resolution_logs) == len(expected_resolution_logs), "quality qualification fmt resolution logs are incomplete")
        for index, path in enumerate(expected_resolution_logs):
            _receipt_reference(resolution_logs[index], f"quality qualification resolution log {index}", path)
        result: dict[str, Any] = {
            "quality_qualification": descriptor(QUALITY_QUALIFICATION),
            "correctness_review": descriptor(correctness_review),
            "owner_gates": {
                "receipt": initial_binding,
                "runner": descriptor(OWNER_GATES_ATTEMPT / "owner_gates_runner.py"),
                "commands": len(initial_receipt["commands"]),
                "status": initial_receipt["status"],
                "failures": list(initial_receipt["failures"]),
            },
            "owner_gates_continuation": {
                "receipt": continuation_binding,
                "runner": descriptor(continuation.parent / "owner_continuation_runner.py"),
                "commands": len(continuation_receipt["commands"]),
            },
        }
        return result
    if QUALITY.is_dir():
        result = {}
        for gate in QUALITY_GATES:
            receipt_path = QUALITY / f"{gate}.json"
            receipt = read(receipt_path)
            require(receipt.get("exit_code") == 0 and receipt.get("status") in (None, "pass"), f"quality gate {gate} is not passing")
            log_path = QUALITY / f"{gate}.log"
            require(log_path.is_file(), f"quality gate {gate} log is missing")
            start_path = QUALITY / f"{gate}-start.json"
            if not start_path.is_file():
                start_path = receipt_path
            start = read(start_path)
            require(isinstance(start.get("started_utc"), str) and isinstance(receipt.get("finished_utc"), str), f"quality gate {gate} timing receipt is incomplete")
            if start_path != receipt_path:
                require(start.get("argv") == receipt.get("argv"), f"quality gate {gate} start/finish command differs")
            result[gate] = {"receipt": descriptor(receipt_path), "start": descriptor(start_path), "log": descriptor(log_path), "exit_code": receipt["exit_code"]}
        result["owner_gates"] = {
            "receipt": initial_binding,
            "runner": descriptor(OWNER_GATES_ATTEMPT / "owner_gates_runner.py"),
            "commands": len(initial_receipt["commands"]),
            "status": initial_receipt["status"],
            "failures": list(initial_receipt["failures"]),
        }
        result["owner_gates_continuation"] = {"receipt": continuation_binding, "runner": descriptor(continuation.parent / "owner_continuation_runner.py"), "commands": len(continuation_receipt["commands"])}
        return result
    result: dict[str, Any] = {}
    result["owner_gates"] = {
        "receipt": initial_binding,
        "runner": descriptor(OWNER_GATES_ATTEMPT / "owner_gates_runner.py"),
        "commands": len(initial_receipt["commands"]),
        "status": initial_receipt["status"],
        "failures": list(initial_receipt["failures"]),
    }
    result["owner_gates_continuation"] = {"receipt": continuation_binding, "runner": descriptor(continuation.parent / "owner_continuation_runner.py"), "commands": len(continuation_receipt["commands"])}
    return result


def _build_receipt(path: Path, label: str) -> tuple[dict[str, Any], dict[str, Any]]:
    value, item = _json_descriptor(path, label)
    require(value.get("exit_code") == 0, f"{label} did not exit zero")
    return value, item


def _validate_build(path: Path, label: str, *, candidate: bool) -> tuple[dict[str, Any], dict[str, Any]]:
    build, binding = _json_descriptor(path, label)
    require(build.get("status") == "pass", f"{label} is not passing")
    require(isinstance(build.get("binaries"), dict), f"{label} binary map is missing")
    if candidate:
        require(build.get("schema") == "litchi.mce-candidate-build-0890.v1", "candidate build schema mismatch")
        require(build.get("baseline_revision") == BASELINE_REVISION, "candidate retained baseline revision mismatch")
        source = build.get("source")
        require(isinstance(source, dict), "candidate build freeze binding is missing")
        require(path_of(source.get("path")) == FREEZE.resolve(), "candidate build is bound to the wrong freeze path")
        same_identity(descriptor(FREEZE), source, "candidate build freeze")
    else:
        require(build.get("revision") == BASELINE_REVISION, "retained baseline revision mismatch")
        same_identity(descriptor(BASELINE_PACKET / "source.json"), build.get("source_manifest", {}), "retained baseline manifest")
        same_identity(descriptor(BASELINE_PACKET / "source.tar.gz"), build.get("source_archive", {}), "retained baseline archive")
    for role in MODES:
        item = build["binaries"].get(role)
        require(isinstance(item, dict), f"{label} {role} binary entry is missing")
        binary_path = path.parent / "binaries" / role if candidate else BASELINE_PACKET / "binaries" / role
        actual = descriptor(binary_path)
        same_identity(actual, item, f"{label} {role} binary")
        require(item.get("executable") is True and actual["executable"], f"{label} {role} binary is not executable")
        require(item.get("features") == ([] if role == "native" else ["allocator-metrics"]), f"{label} {role} feature set differs")
        receipt_path = path.parent / f"{role}-build.json" if candidate else BASELINE_PACKET / f"{role}-build.json"
        receipt, receipt_binding = _build_receipt(receipt_path, f"{label} {role} build receipt")
        require(isinstance(receipt.get("argv"), list) and BUILD_NAMES[role] in receipt["argv"], f"{label} {role} build command names the wrong binary")
        if isinstance(item.get("build_receipt"), dict):
            same_identity(receipt_binding, item["build_receipt"], f"{label} {role} build receipt binding")
        depfile_path = path.parent / f"{role}.d" if candidate else BASELINE_PACKET / f"{role}.d"
        depfile = descriptor(depfile_path)
        if isinstance(item.get("depfile"), dict):
            same_identity(depfile, item["depfile"], f"{label} {role} depfile binding")
    return build, binding


def _validate_build_context(baseline: dict[str, Any], candidate: dict[str, Any]) -> None:
    candidate_start = read(BUILD_START)
    baseline_start = read(BASELINE_PACKET / "build-start.json")
    require(candidate_start.get("compiler") == baseline_start.get("compiler"), "baseline/candidate compiler strings differ")
    require(candidate.get("compiler") == baseline.get("compiler"), "baseline/candidate compiler strings differ")
    require(candidate.get("cargo") == baseline.get("cargo"), "baseline/candidate Cargo strings differ")
    def effective(environment: dict[str, Any], key: str) -> Any:
        value = environment.get(key)
        if key == "CARGO_HOME" and not value:
            return str(Path(environment.get("HOME", "/home/zhuhe")) / ".cargo")
        return value
    for key in BUILD_ENV_KEYS:
        require(effective(candidate.get("environment", {}), key) == effective(baseline.get("environment", {}), key), f"baseline/candidate build environment differs for {key}")
    for role in MODES:
        before = read(BASELINE_PACKET / f"{role}-build.json")
        after = read(PACKET / f"{role}-build.json")
        require(after.get("argv") == before.get("argv"), f"baseline/candidate {role} build argv differs")


def _validate_execution(candidate: dict[str, Any], baseline: dict[str, Any], *, require_live: bool) -> tuple[dict[str, Any], dict[str, Any]]:
    receipt, binding = _json_descriptor(EXECUTION, "managed executable custody")
    require(receipt.get("schema") == "litchi.mce-managed-executables-0890.v1" and receipt.get("status") == "pass", "managed executable custody receipt is not passing")
    records = receipt.get("records")
    require(isinstance(records, dict), "managed executable records are missing")
    copies: dict[str, dict[str, dict[str, Any]]] = {}
    for side in ROLES:
        copies[side] = {}
        for mode in MODES:
            item = records.get(side, {}).get(mode)
            require(isinstance(item, dict) and isinstance(item.get("source"), dict) and isinstance(item.get("copy"), dict), f"managed executable record is incomplete: {side}/{mode}")
            source = item["source"]
            copy = item["copy"]
            expected = (baseline if side == "before" else candidate)["binaries"][mode]
            same_identity(source, expected, f"managed source {side}/{mode}")
            require(source.get("path") == expected.get("path"), f"managed source path differs for {side}/{mode}")
            require(copy.get("executable") is True, f"managed copy is not executable: {side}/{mode}")
            require(copy.get("path") != source.get("path"), f"managed copy aliases source path: {side}/{mode}")
            copy_path = path_of(copy.get("path"))
            require(copy_path.parent == (PACKET / "executables").resolve(), f"managed copy escapes the owned executable directory: {side}/{mode}")
            if require_live:
                live_copy = descriptor(copy_path)
                same_identity(live_copy, copy, f"managed copy {side}/{mode}")
                require(live_copy.get("mode") == copy.get("mode"), f"managed copy mode differs: {side}/{mode}")
                if "device" in copy and "inode" in copy:
                    info = copy_path.stat()
                    require(info.st_dev == copy["device"] and info.st_ino == copy["inode"], f"managed copy inode changed: {side}/{mode}")
            copies[side][mode] = copy
    return copies, binding


def _validate_tools() -> dict[str, dict[str, Any]]:
    result = {}
    for name, path in {"time": Path(experiment.TIME), "taskset": Path(experiment.TASKSET), "python": Path(sys.executable)}.items():
        require(path.is_file() and os.access(path, os.X_OK), f"required tool is unavailable: {path}")
        result[name] = descriptor(path)
    result["python"]["version"] = platform.python_version()
    return result


def _validate_hardware() -> dict[str, Any]:
    environment_path = BASELINE_PACKET / "environment.json"
    environment = read(environment_path)
    require(environment.get("capture_cpu") == 12, "retained experiment CPU differs from CPU 12")
    require(12 in os.sched_getaffinity(0), "CPU 12 is unavailable")
    return {
        "environment": descriptor(environment_path),
        "capture_cpu": environment["capture_cpu"],
        "platform": environment.get("platform"),
        "uname": environment.get("uname"),
        "exclusive_host": False,
        "claim_scope": "CPU affinity only; no exclusive-host or physical-cold guarantee",
    }


def _validate_correctness() -> dict[str, dict[str, Any]]:
    paths = {
        "baseline_receipt": PACKET / "attempts/baseline-f790c2228c-retry1/baseline-receipt.json",
        "baseline_post_source": PACKET / "attempts/baseline-f790c2228c-retry1/post-baseline-source-verification.json",
        "baseline_lock": PACKET / "attempts/baseline-f790c2228c-retry1/root-lock-custody.json",
        "candidate_overlay": PACKET / "candidate-diagnostic-overlay.json",
        "candidate_receipt": PACKET / "attempts/candidate-diagnostic-f790c2228c-v2-retry1/candidate-receipt.json",
        "candidate_post_source": PACKET / "attempts/candidate-diagnostic-f790c2228c-v2-retry1/post-candidate-source-verification.json",
        "cross_version_comparison": PACKET / "cross-version-diagnostic-comparison.json",
        "root_oracle_comparison": PACKET / "root-oracle-comparison.json",
    }
    result = {}
    for name, path in paths.items():
        value, item = _json_descriptor(path, name)
        if name == "baseline_lock":
            require(value.get("byte_identical") is True, f"{name} does not prove byte-identical lock custody")
        else:
            require(value.get("status") == "pass", f"{name} is not passing")
        result[name] = item
    overlay = read(paths["candidate_overlay"])
    require(overlay.get("candidate_freeze", {}).get("path") == "candidate-freeze-v2.json", "diagnostic overlay is bound to the wrong freeze")
    require(overlay.get("candidate_patch", {}).get("path") == "candidate-v2.patch", "diagnostic overlay is bound to the wrong patch")
    candidate = read(paths["candidate_receipt"])
    require(candidate.get("performance_claim") is False, "candidate correctness receipt makes a performance claim")
    comparison = read(paths["cross_version_comparison"])
    require(comparison.get("cross_version_byte_identical") is True and comparison.get("repeat_byte_identical") is True, "candidate correctness reports are not byte-identical to the baseline")
    oracle = read(paths["root_oracle_comparison"])
    require(oracle.get("repeat_byte_equal") is True and oracle.get("missing_oracle_refused") is True and oracle.get("missing_controls_refused") is True, "oracle comparison did not pass its repeat and negative controls")
    return result


def _helper_descriptors() -> dict[str, dict[str, Any]]:
    return {name: descriptor(path) for name, path in IMPORTED_HELPERS.items()}


def _binding_paths() -> dict[str, Path]:
    files: dict[str, Path] = {
        "candidate_freeze": FREEZE,
        "candidate_patch": PATCH,
        "candidate_source_review": SOURCE_REVIEW,
        "candidate_methodology_review": METHODOLOGY_REVIEW,
        "candidate_build": BUILD,
        "candidate_build_start": BUILD_START,
        "candidate_build_executed": BUILD_EXECUTED,
        "release_build_launch": RELEASE_BUILD_LAUNCH,
        "release_build_completion": RELEASE_BUILD_COMPLETION,
        "candidate_execution_files": EXECUTION,
        "candidate_build_native": PACKET / "native-build.json",
        "candidate_build_allocator": PACKET / "allocator-build.json",
        "candidate_build_native_start": PACKET / "native-build-start.json",
        "candidate_build_allocator_start": PACKET / "allocator-build-start.json",
        "candidate_native_log": PACKET / "native-build.log",
        "candidate_allocator_log": PACKET / "allocator-build.log",
        "candidate_native_depfile": PACKET / "native.d",
        "candidate_allocator_depfile": PACKET / "allocator.d",
        "baseline_build": BASELINE_PACKET / "build.json",
        "baseline_build_start": BASELINE_PACKET / "build-start.json",
        "baseline_source": BASELINE_PACKET / "source.json",
        "baseline_archive": BASELINE_PACKET / "source.tar.gz",
        "baseline_environment": BASELINE_PACKET / "environment.json",
        "baseline_native_binary": BASELINE_PACKET / "binaries/native",
        "baseline_allocator_binary": BASELINE_PACKET / "binaries/allocator",
        "baseline_native_build": BASELINE_PACKET / "native-build.json",
        "baseline_allocator_build": BASELINE_PACKET / "allocator-build.json",
        "baseline_native_start": BASELINE_PACKET / "native-build-start.json",
        "baseline_allocator_start": BASELINE_PACKET / "allocator-build-start.json",
        "baseline_native_log": BASELINE_PACKET / "native-build.log",
        "baseline_allocator_log": BASELINE_PACKET / "allocator-build.log",
        "baseline_native_depfile": BASELINE_PACKET / "native.d",
        "baseline_allocator_depfile": BASELINE_PACKET / "allocator.d",
        "correctness_baseline": PACKET / "attempts/baseline-f790c2228c-retry1/baseline-receipt.json",
        "correctness_post_source": PACKET / "attempts/baseline-f790c2228c-retry1/post-baseline-source-verification.json",
        "correctness_lock": PACKET / "attempts/baseline-f790c2228c-retry1/root-lock-custody.json",
        "candidate_diagnostic_overlay": PACKET / "candidate-diagnostic-overlay.json",
        "candidate_correctness": PACKET / "attempts/candidate-diagnostic-f790c2228c-v2-retry1/candidate-receipt.json",
        "candidate_post_source": PACKET / "attempts/candidate-diagnostic-f790c2228c-v2-retry1/post-candidate-source-verification.json",
        "cross_version_comparison": PACKET / "cross-version-diagnostic-comparison.json",
        "root_oracle_comparison": PACKET / "root-oracle-comparison.json",
    }
    for name, path in IMPORTED_HELPERS.items():
        files[f"helper/{name}"] = path
    if QUALITY_QUALIFICATION.is_file():
        files["quality/qualification"] = QUALITY_QUALIFICATION
        files["quality/correctness-review"] = PACKET / "correctness-review.json"
        files["quality/owner-gates-receipt"] = OWNER_GATES_RECEIPT
        files["quality/owner-gates-runner"] = OWNER_GATES_ATTEMPT / "owner_gates_runner.py"
        for path in sorted((OWNER_GATES_ATTEMPT / "logs").glob("*.log")):
            files[f"quality/owner-gates-log/{path.name}"] = path
        continuation = _owner_continuation_receipt()
        require(continuation is not None, "quality binding continuation is missing")
        files["quality/owner-continuation-receipt"] = continuation
        files["quality/owner-continuation-runner"] = continuation.parent / "owner_continuation_runner.py"
        for path in sorted(continuation.parent.glob("logs/*.log")):
            files[f"quality/owner-continuation-log/{path.name}"] = path
    elif QUALITY.is_dir():
        for gate in QUALITY_GATES:
            files[f"quality/{gate}"] = QUALITY / f"{gate}.json"
            files[f"quality/{gate}-log"] = QUALITY / f"{gate}.log"
            start = QUALITY / f"{gate}-start.json"
            files[f"quality/{gate}-start"] = start if start.is_file() else QUALITY / f"{gate}.json"
        continuation = _owner_continuation_receipt()
        if continuation is not None:
            files["quality/owner-gates-continuation"] = continuation
            files["quality/owner-continuation-runner"] = continuation.parent / "owner_continuation_runner.py"
            for path in sorted(continuation.parent.glob("logs/*.log")):
                files[f"quality/owner-continuation-log/{path.name}"] = path
    else:
        files["quality/owner-gates-receipt"] = OWNER_GATES_RECEIPT
        files["quality/owner-gates-runner"] = OWNER_GATES_ATTEMPT / "owner_gates_runner.py"
        continuation = _owner_continuation_receipt()
        if continuation is not None:
            files["quality/owner-gates-continuation"] = continuation
            files["quality/owner-continuation-runner"] = continuation.parent / "owner_continuation_runner.py"
            for path in sorted(continuation.parent.glob("logs/*.log")):
                files[f"quality/owner-continuation-log/{path.name}"] = path
        for path in sorted((OWNER_GATES_ATTEMPT / "logs").glob("*.log")):
            files[f"quality/owner-gates-log/{path.name}"] = path
    for path in sorted(BASELINE_QUALIFICATION.rglob("*")):
        if path.is_file():
            files[f"baseline_qualification/{path.relative_to(BASELINE_QUALIFICATION).as_posix()}"] = path
    for name in ("codec.rs", "tests.rs"):
        relative = f"crates/litchi-ooxml-common/src/mce/{name}"
        files[f"candidate_overlay/{relative}"] = ROOT / relative
    return files


def _make_bindings() -> dict[str, dict[str, Any]]:
    bindings = {}
    for key, path in _binding_paths().items():
        require(path.is_file(), f"required freeze binding is missing: {path}")
        bindings[key] = descriptor(path)
    return bindings


def _validate_bindings(plan: dict[str, Any], *, retained_only: bool) -> None:
    bindings = plan.get("bindings")
    require(isinstance(bindings, dict), "plan bindings are missing")
    expected = _binding_paths()
    for key, path in expected.items():
        require(key in bindings, f"plan binding is missing: {key}")
        require(path_of(bindings[key].get("path")) == path.resolve(), f"plan binding path mismatch: {key}")
    for key, item in bindings.items():
        if retained_only and key.startswith("candidate_overlay/"):
            require(isinstance(item, dict) and isinstance(item.get("bytes"), int) and isinstance(item.get("sha256"), str), f"retained candidate overlay binding is incomplete: {key}")
        else:
            check_descriptor(item, f"plan binding {key}")
    require(path_of(plan.get("candidate_build_json")) == BUILD.resolve(), "candidate build JSON path is not frozen")
    require(path_of(plan.get("baseline_build_json")) == (BASELINE_PACKET / "build.json").resolve(), "baseline build JSON path is not frozen")


def _check_binary_plan(plan: dict[str, Any], baseline: dict[str, Any], candidate: dict[str, Any], managed: dict[str, dict[str, dict[str, Any]]], *, retained_only: bool) -> None:
    require(plan.get("binaries", {}).get("before") == baseline["binaries"], "before binary plan differs from retained 0885 build")
    require(plan.get("binaries", {}).get("after") == candidate["binaries"], "after binary plan differs from candidate build")
    require(plan.get("managed_binaries") == managed, "managed executable records differ from execution-files.json")
    plan["retained_only"] = retained_only
    experiment.validate_plan(plan)


def _prerequisites(*, require_live: bool) -> tuple[dict[str, Any], dict[str, Any], dict[str, dict[str, Any]], dict[str, Any], dict[str, Any], dict[str, Any], dict[str, Any]]:
    freeze, files, patch_identity = _load_freeze(require_live=require_live)
    review, review_binding = _validate_source_review(freeze, files, patch_identity)
    methodology, methodology_binding = _validate_methodology()
    quality = _validate_quality(require_live=require_live)
    baseline, _ = _validate_build(BASELINE_PACKET / "build.json", "retained baseline build", candidate=False)
    candidate, _ = _validate_build(BUILD, "candidate build", candidate=True)
    _validate_build_context(baseline, candidate)
    managed, execution_binding = _validate_execution(candidate, baseline, require_live=require_live)
    correctness = _validate_correctness()
    return baseline, candidate, managed, quality, {"source": review, "source_binding": review_binding, "methodology": methodology, "methodology_binding": methodology_binding}, execution_binding, correctness


def make_plan() -> dict[str, Any]:
    baseline, candidate, managed, quality, reviews, _, correctness = _prerequisites(require_live=True)
    freeze, files, patch_identity = _load_freeze(require_live=True)
    bindings = _make_bindings()
    plan = {
        "schema": experiment.SCHEMA,
        "frozen_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
        "profile": experiment.PROFILE,
        "cpu_affinity": experiment.CPU_AFFINITY,
        "cache": "warm",
        "xlsx_shape": "medium",
        "cases": list(experiment.CASES),
        "native_blocks": experiment.NATIVE_BLOCKS,
        "allocator_blocks": experiment.ALLOCATOR_BLOCKS,
        "native_warmup": experiment.NATIVE_WARMUP,
        "native_samples": experiment.NATIVE_SAMPLES,
        "allocator_warmup": experiment.ALLOCATOR_WARMUP,
        "allocator_samples": experiment.ALLOCATOR_SAMPLES,
        "fail_fast": True,
        "fresh_child": True,
        "binaries": {"before": baseline["binaries"], "after": candidate["binaries"]},
        "managed_binaries": managed,
        "candidate_build_json": str(BUILD.resolve()),
        "baseline_build_json": str((BASELINE_PACKET / "build.json").resolve()),
        "revision": BASELINE_REVISION,
        "base_commit": CANDIDATE_BASE,
        "bindings": bindings,
        "source": {
            "baseline_manifest": descriptor(BASELINE_PACKET / "source.json"),
            "baseline_archive": descriptor(BASELINE_PACKET / "source.tar.gz"),
            "candidate_freeze": descriptor(FREEZE),
            "candidate_patch": patch_identity,
            "candidate_paths": list(freeze["candidate_paths"]),
            "overlays": {name: {**descriptor(ROOT / name), "freeze": files[name]} for name in freeze["candidate_paths"]},
        },
        "quality_gates": quality,
        "reviews": {
            "source": descriptor(SOURCE_REVIEW),
            "methodology": descriptor(METHODOLOGY_REVIEW),
            "source_status": reviews["source"].get("status"),
            "methodology_status": reviews["methodology"].get("status"),
        },
        "helpers": _helper_descriptors(),
        "release_build": {
            "runner": descriptor(BUILD_EXECUTED),
            "launch": descriptor(RELEASE_BUILD_LAUNCH),
            "completion": descriptor(RELEASE_BUILD_COMPLETION),
        },
        "execution": {"receipt": descriptor(EXECUTION), "copy_scope": "one content-identical managed copy per role/mode, reused by scheduled children"},
        "correctness": correctness,
        "tools": _validate_tools(),
        "hardware": _validate_hardware(),
        "environment": {"child": experiment.CHILD_ENV, "build": read(BUILD_START).get("environment"), "hardware_binding": "baseline_environment"},
        "protocol": {
            "schedule": "native lane completes before allocator lane",
            "counterbalance": "AB/BA alternates by (block + case_index) % 2; allocator by block parity",
            "total_children": experiment.TOTAL_CHILDREN,
            "native_children": experiment.TOTAL_NATIVE_CHILDREN,
            "allocator_children": experiment.TOTAL_ALLOCATOR_CHILDREN,
            "no_retries_or_replacements": True,
            "fresh_process_per_child": True,
        },
        "analysis": {
            "unit": "whole-child aggregate statistics",
            "native_metrics": ["p50_ns", "p95_ns", "p99_ns", "peak_rss_kib"],
            "pairing": "after/before ratios within scheduled pairs; samples never pooled",
            "primary_gate": {"native_p50_median_at_most": 0.98, "bootstrap_ci95_upper_below": 1.0},
            "control_review_ratio_above": 1.05,
            "native_p50_spread_ratio_above": 1.20,
            "primary_review_metrics": ["p50_ns", "p95_ns", "p99_ns", "peak_rss_kib"],
            "allocator_adoption": "not_evaluated",
            "bootstrap": {"seed": experiment.BOOTSTRAP_SEED, "draws": 10000, "ci95_indices": [249, 9749], "unit": "paired child ratios"},
            "performance_claim": None,
        },
        "host_scope": "CPU affinity only; no exclusive-host or physical-cold guarantee",
    }
    _check_binary_plan(plan, baseline, candidate, managed, retained_only=False)
    _validate_bindings(plan, retained_only=False)
    return plan


def verify_plan(*, retained_only: bool) -> dict[str, Any]:
    require(PLAN.is_file(), f"plan is missing: {PLAN}")
    plan = read(PLAN)
    require(isinstance(plan, dict), "plan must be an object")
    baseline, candidate, managed, quality, reviews, _, correctness = _prerequisites(require_live=not retained_only)
    freeze, files, patch_identity = _load_freeze(require_live=not retained_only)
    require(plan.get("source", {}).get("candidate_freeze", {}).get("sha256") == descriptor(FREEZE)["sha256"], "plan candidate freeze binding changed")
    require(plan.get("source", {}).get("candidate_patch", {}).get("sha256") == patch_identity["sha256"], "plan candidate patch binding changed")
    _check_binary_plan(plan, baseline, candidate, managed, retained_only=retained_only)
    _validate_bindings(plan, retained_only=retained_only)
    quality = plan.get("quality_gates", {})
    if QUALITY_QUALIFICATION.is_file():
        require("quality_qualification" in quality, "plan omits the aggregate quality qualification")
        require("correctness_review" in quality, "plan omits the correctness review bound by quality qualification")
        require("owner_gates" in quality and "owner_gates_continuation" in quality, "plan omits retained owner-gate custody")
        require(quality["owner_gates"].get("status") == "fail" and quality["owner_gates"].get("failures") == ["fmt_check"], "plan does not retain the initial owner-gate fmt failure")
    elif QUALITY.is_dir():
        require(quality.keys() >= set(QUALITY_GATES), "plan omits a quality gate")
    else:
        require("owner_gates" in quality, "plan omits the owner-gates quality receipt")
    require(plan.get("helpers") == _helper_descriptors(), "plan helper custody differs from current helper files")
    require(plan.get("correctness") == correctness, "plan correctness custody differs")
    release_build = plan.get("release_build", {})
    require(release_build.get("runner") == descriptor(BUILD_EXECUTED), "plan release-build runner binding changed")
    require(release_build.get("launch") == descriptor(RELEASE_BUILD_LAUNCH), "plan release-build launch binding changed")
    require(release_build.get("completion") == descriptor(RELEASE_BUILD_COMPLETION), "plan release-build completion binding changed")
    return {"status": "pass", "plan": str(PLAN.resolve()), "plan_sha256": sha256(PLAN), "retained_only": retained_only, "mutated": False}


def write_create_only(path: Path, value: Any) -> None:
    with path.open("x", encoding="utf-8") as handle:
        json.dump(value, handle, indent=2, sort_keys=True)
        handle.write("\n")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--verify", action="store_true", help="verify an existing plan without writing")
    parser.add_argument("--retained-only", action="store_true", help="with --verify, accept removed managed copies")
    args = parser.parse_args()
    require(not args.retained_only or args.verify, "--retained-only requires --verify")
    if args.verify:
        print(json.dumps(verify_plan(retained_only=args.retained_only), indent=2, sort_keys=True))
        return 0
    require(not PLAN.exists(), f"refusing to overwrite existing plan: {PLAN}")
    plan = make_plan()
    write_create_only(PLAN, plan)
    print(json.dumps({"status": "pass", "plan": str(PLAN.resolve()), "children": experiment.TOTAL_CHILDREN, "mutated": True}, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
