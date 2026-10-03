#!/usr/bin/env python3
"""Freeze and verify the reviewed 0887 paired experiment inputs.

The command has two deliberately small modes.  With no option it validates all
retained inputs and creates ``plan.json`` with exclusive creation.  ``--verify``
reopens that plan and checks every descriptor, source overlay, review, quality
receipt, tool, and retained binary without writing anything.  It does not build
or launch a workload.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
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
BASELINE_PACKET = PACKET.parent / "change-0885"
ROOT = PACKET.parents[3]
BASELINE_QUALIFICATION = PACKET / "baseline-qualification"
QUALITY = PACKET / "quality-v2"
PLAN = PACKET / "plan.json"
TASKSET = Path(experiment.TASKSET)
TIME = Path(experiment.TIME)
TOOLS = ("time", "taskset", "python")
QUALITY_GATES = ("fmt", "owner-tests", "clippy", "rustdoc", "shared-format-tests", "boundaries")
OVERLAY_PATHS = (
    "crates/litchi-ooxml-common/src/mce/codec.rs",
    "crates/litchi-ooxml-common/src/mce/expanded_duplicates_tests.rs",
)
BUILD_ROLES = ("native", "allocator")
BUILD_BINARY_NAMES = {"native": "litchi-perf-baseline", "allocator": "litchi-perf-baseline-alloc"}
CHILD_ENV = dict(experiment.CHILD_ENV)
BUILD_ENV_KEYS = (
    "CARGO_BUILD_JOBS",
    "CARGO_INCREMENTAL",
    "CARGO_NET_OFFLINE",
    "CARGO_PROFILE_RELEASE_OPT_LEVEL",
    "CARGO_PROFILE_RELEASE_DEBUG",
    "CARGO_PROFILE_RELEASE_LTO",
    "CARGO_PROFILE_RELEASE_CODEGEN_UNITS",
    "LANG",
    "LC_ALL",
    "TZ",
)


def _read(path: Path) -> Any:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    """Describe one retained regular file, resolving symlinks first."""

    path = path.resolve()
    info = path.stat()
    _require(stat.S_ISREG(info.st_mode), f"expected a regular file: {path}")
    return {
        "path": str(path),
        "bytes": info.st_size,
        "sha256": _sha256(path),
        "mode": stat.S_IMODE(info.st_mode),
        "executable": bool(info.st_mode & stat.S_IXUSR),
    }


def _path(path: str | Path, *, base: Path = PACKET) -> Path:
    value = Path(path)
    if not value.is_absolute():
        value = base / value
    return value.resolve()


def _check_descriptor(raw: Any, label: str) -> dict[str, Any]:
    _require(isinstance(raw, dict), f"{label} descriptor is not an object")
    _require(isinstance(raw.get("path"), str), f"{label} descriptor has no path")
    _require(isinstance(raw.get("bytes"), int) and raw["bytes"] >= 0, f"{label} descriptor has no byte count")
    _require(isinstance(raw.get("sha256"), str) and len(raw["sha256"]) == 64, f"{label} descriptor has no sha256")
    actual = descriptor(_path(raw["path"]))
    for field in ("bytes", "sha256", "mode", "executable"):
        if field in raw:
            _require(raw[field] == actual[field], f"{label} {field} does not match retained file")
    return actual


def _same_descriptor(raw: Any, expected: dict[str, Any], label: str) -> None:
    _require(isinstance(raw, dict), f"{label} descriptor is not an object")
    _require(_path(raw.get("path", "")) == _path(expected["path"]), f"{label} path is not bound to the expected file")
    for field in ("bytes", "sha256"):
        _require(raw.get(field) == expected[field], f"{label} {field} is not bound to the expected file")


def _file_binding(path: Path) -> dict[str, Any]:
    return descriptor(path)


def _json_descriptor(path: Path, label: str) -> tuple[dict[str, Any], dict[str, Any]]:
    _require(path.is_file(), f"{label} is missing: {path}")
    value = _read(path)
    _require(isinstance(value, dict), f"{label} must be a JSON object")
    return value, descriptor(path)


def _build_json(label: str, path: Path) -> tuple[dict[str, Any], dict[str, Any]]:
    value, item = _json_descriptor(path, label)
    _require(value.get("status") == "pass", f"{label} status is not pass")
    return value, item


def _validate_quality() -> dict[str, dict[str, Any]]:
    """Require each v2 gate to have a complete, successful receipt."""

    _require(QUALITY.is_dir(), "quality-v2 directory is missing")
    result: dict[str, dict[str, Any]] = {}
    for gate in QUALITY_GATES:
        receipt, _ = _json_descriptor(QUALITY / f"{gate}.json", f"quality gate {gate}")
        start_path = QUALITY / f"{gate}-start.json"
        # The boundary checker emits one atomic receipt rather than a
        # separate start record; its started_utc field is retained in that
        # receipt and is checked below with the other gate results.
        if gate == "boundaries" and not start_path.is_file():
            start = receipt
        else:
            start = _json_descriptor(start_path, f"quality gate {gate} start")[0]
        log = QUALITY / f"{gate}.log"
        _require(log.is_file(), f"quality gate {gate} log is missing")
        _require(receipt.get("exit_code") == 0, f"quality gate {gate} did not exit zero")
        _require(isinstance(start.get("started_utc"), str) and isinstance(receipt.get("finished_utc"), str), f"quality gate {gate} is not a completed receipt")
        if "status" in receipt:
            _require(receipt["status"] == "pass", f"quality gate {gate} status is not pass")
        _require(start.get("argv") == receipt.get("argv"), f"quality gate {gate} start/finish argv mismatch")
        result[gate] = {"start": start, "receipt": receipt}
    return result


def _validate_reviews() -> tuple[dict[str, Any], dict[str, Any]]:
    semantic, _ = _json_descriptor(PACKET / "semantic-review.json", "semantic review")
    methodology, _ = _json_descriptor(PACKET / "methodology-review.json", "methodology review")
    _require(semantic.get("status") == "pass", "semantic review is not passing")
    _require(methodology.get("status") == "pass", "methodology review is not passing")
    scope = methodology.get("scope", {})
    _require(scope.get("reviewed_without_workload_or_Cargo") is True, "methodology review does not record static review scope")
    for key, filename in (("experiment_helper", "experiment.py"), ("reduction_helper", "reduce.py")):
        reviewed = scope.get(key, {})
        actual = descriptor(PACKET / filename)
        reviewed_path = str(reviewed.get("path", ""))
        allowed_paths = {filename, str(PACKET / filename), str(PACKET.relative_to(ROOT) / filename)}
        _require(reviewed_path in allowed_paths, f"methodology review does not name {filename}")
        _require(reviewed.get("bytes") == actual["bytes"] and reviewed.get("sha256") == actual["sha256"], f"methodology review hash mismatch for {filename}")
    return semantic, methodology


def _validate_helper_checks() -> dict[str, Any]:
    checks, _ = _json_descriptor(PACKET / "helper-checks.json", "helper checks")
    _require(checks.get("status") == "pass" and checks.get("exit_code") == 0, "helper checks did not pass")
    _require(checks.get("scope", {}).get("workloads_launched") is False and checks.get("scope", {}).get("cargo_commands") is False, "helper checks scope is not static")
    for name in ("experiment.py", "reduce.py", "freeze.py"):
        recorded = checks.get("files", {}).get(name, {})
        actual = descriptor(PACKET / name)
        _require(recorded.get("bytes") == actual["bytes"] and recorded.get("sha256") == actual["sha256"], f"helper check hash mismatch for {name}")
    required = ("ast_parse", "schedule", "bootstrap", "primary_review_flags")
    for name in required:
        item = checks.get("checks", {}).get(name)
        _require(isinstance(item, dict) and item.get("status") == "pass", f"helper check {name} is missing or failed")
    return checks


def _validate_source_overlays(candidate_source: dict[str, Any], *, require_scratch: bool, require_live: bool = True) -> dict[str, dict[str, Any]]:
    overlays = candidate_source.get("overlays")
    _require(isinstance(overlays, dict) and set(overlays) == set(OVERLAY_PATHS), "candidate source overlays are not the two reviewed files")
    source_root = Path(candidate_source.get("source_root", ""))
    result: dict[str, dict[str, Any]] = {}
    for relative in OVERLAY_PATHS:
        expected = overlays[relative]
        _require(isinstance(expected, dict), f"candidate overlay {relative} is not a descriptor")
        _require(isinstance(expected.get("sha256"), str) and isinstance(expected.get("bytes"), int), f"candidate overlay {relative} identity is incomplete")
        live = ROOT / relative
        if require_live:
            actual = descriptor(live)
            _require(actual["sha256"] == expected["sha256"] and actual["bytes"] == expected["bytes"], f"current source overlay changed: {relative}")
        else:
            actual = {"path": str(live.resolve()), "bytes": expected["bytes"], "sha256": expected["sha256"], "mode": expected.get("mode"), "executable": expected.get("executable", False)}
        scratch = source_root / relative
        if scratch.is_file():
            scratch_actual = descriptor(scratch)
            _require(scratch_actual["sha256"] == expected["sha256"] and scratch_actual["bytes"] == expected["bytes"], f"candidate source tree overlay changed: {relative}")
        else:
            _require(not require_scratch, f"candidate source tree overlay is missing before freeze: {scratch}")
        result[relative] = actual
    return result


def _validate_build(path: Path, label: str, *, candidate: bool) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    build, build_binding = _build_json(label, path)
    _require(isinstance(build.get("binaries"), dict), f"{label} has no binary map")
    if candidate:
        candidate_revision = _read(PACKET / "candidate-source.json").get("baseline_revision")
        if "revision" in build:
            _require(build.get("revision") == candidate_revision, "candidate revision/source revision mismatch")
        source = build.get("source")
        _require(isinstance(source, dict), "candidate build has no candidate-source binding")
        _same_descriptor(source, descriptor(PACKET / "candidate-source.json"), "candidate build source")
    else:
        _require(build.get("revision") == _read(BASELINE_PACKET / "source.json").get("revision"), "baseline build revision/source revision mismatch")
        _same_descriptor(build.get("source_manifest"), descriptor(BASELINE_PACKET / "source.json"), "baseline source manifest")
        _same_descriptor(build.get("source_archive"), descriptor(BASELINE_PACKET / "source.tar.gz"), "baseline source archive")
    for role in BUILD_ROLES:
        raw_item = build["binaries"].get(role)
        _require(isinstance(raw_item, dict), f"{label} {role} binary entry is missing")
        item = dict(raw_item)
        binary_path = _path(item.get("path", ""))
        expected_path = (path.parent / "binaries" / role).resolve()
        if not candidate:
            expected_path = (BASELINE_PACKET / "binaries" / role).resolve()
        _require(binary_path == expected_path, f"{label} {role} binary path is not retained in the packet")
        actual_binary = descriptor(binary_path)
        for field in ("bytes", "sha256"):
            _require(item.get(field) == actual_binary[field], f"{label} {role} binary {field} mismatch")
        _require(item.get("executable") is True and actual_binary["executable"], f"{label} {role} binary is not executable")
        _require(item.get("features") == ([] if role == "native" else ["allocator-metrics"]), f"{label} {role} feature identity mismatch")
        expected_receipt = (path.parent / f"{role}-build.json").resolve()
        if not candidate:
            expected_receipt = (BASELINE_PACKET / f"{role}-build.json").resolve()
        receipt_descriptor = item.get("build_receipt", descriptor(expected_receipt))
        receipt_actual = _check_descriptor(receipt_descriptor, f"{label} {role} build receipt")
        receipt_path = _path(receipt_descriptor["path"])
        _require(receipt_path == expected_receipt, f"{label} {role} build receipt is not the retained role receipt")
        receipt = _read(receipt_path)
        _require(receipt.get("exit_code") == 0, f"{label} {role} build receipt did not exit zero")
        argv = receipt.get("argv")
        _require(isinstance(argv, list) and BUILD_BINARY_NAMES[role] in argv, f"{label} {role} build receipt names the wrong binary")
        if isinstance(receipt.get("log"), dict):
            _check_descriptor(receipt["log"], f"{label} {role} build log")
        receipt_binary = receipt.get("binary")
        if isinstance(receipt_binary, dict):
            _require(isinstance(receipt_binary.get("sha256"), str) and isinstance(receipt_binary.get("bytes"), int), f"{label} {role} receipt binary identity is incomplete")
            _require(receipt_binary.get("sha256") == item["sha256"], f"{label} {role} receipt binary sha256 mismatch")
            _require(receipt_binary["bytes"] == item["bytes"], f"{label} {role} receipt binary byte count mismatch")
        depfile_path = (path.parent / f"{role}.d").resolve()
        if not candidate:
            depfile_path = (BASELINE_PACKET / f"{role}.d").resolve()
        depfile_descriptor = item.get("depfile", descriptor(depfile_path))
        _check_descriptor(depfile_descriptor, f"{label} {role} depfile")
        _require(_path(depfile_descriptor["path"]) == depfile_path, f"{label} {role} depfile is not retained in the packet")
        # Older 0887 build manifests omit these two links even though their
        # retained files exist.  Normalize the manifest in memory so the plan
        # always carries the command receipt, depfile, and binary identity.
        item["build_receipt"] = {**receipt_descriptor, "sha256": receipt_actual["sha256"], "bytes": receipt_actual["bytes"]}
        item["depfile"] = descriptor(depfile_path)
        build["binaries"][role] = item
    return build, build_binding, {role: build["binaries"][role] for role in BUILD_ROLES}


def _validate_candidate_source(candidate_source: dict[str, Any], baseline_build: dict[str, Any], *, require_scratch: bool, require_live: bool) -> tuple[dict[str, Any], dict[str, Any]]:
    source, source_binding = _json_descriptor(PACKET / "candidate-source.json", "candidate source")
    _validate_source_overlays(source, require_scratch=require_scratch, require_live=require_live)
    patch = descriptor(PACKET / "candidate.patch")
    _same_descriptor(source.get("patch"), patch, "candidate patch")
    _same_descriptor(source.get("baseline_manifest"), descriptor(BASELINE_PACKET / "source.json"), "candidate baseline manifest")
    _same_descriptor(source.get("baseline_archive"), descriptor(BASELINE_PACKET / "source.tar.gz"), "candidate baseline archive")
    _require(source.get("baseline_revision") == baseline_build.get("revision"), "candidate source is based on a different baseline revision")
    _require(source.get("verified_source_files") == len(_read(BASELINE_PACKET / "source.json").get("files", {})), "candidate source file census mismatch")
    return source, source_binding


def _validate_build_context(baseline: dict[str, Any], candidate: dict[str, Any]) -> None:
    baseline_start = _read(BASELINE_PACKET / "build-start.json")
    candidate_start = _read(PACKET / "build-start.json")
    _require(candidate_start.get("compiler") == baseline_start.get("compiler"), "baseline/candidate compiler strings differ")
    for key in BUILD_ENV_KEYS:
        _require(baseline.get("environment", {}).get(key) == candidate.get("environment", {}).get(key), f"release build environment differs for {key}")
    for role in BUILD_ROLES:
        baseline_receipt = _read(BASELINE_PACKET / f"{role}-build.json")
        candidate_receipt = _read(PACKET / f"{role}-build.json")
        _require(candidate_receipt.get("argv") == baseline_receipt.get("argv"), f"baseline/candidate {role} build argv differs")


def _validate_baseline_qualification() -> None:
    _require(BASELINE_QUALIFICATION.is_dir(), "baseline qualification directory is missing")
    for case in experiment.CASES:
        folder = BASELINE_QUALIFICATION / case
        report = _read(folder / "report.json")
        _require(report.get("tool", {}).get("profile") == "release", f"baseline qualification {case} is not release")
        results = report.get("results")
        _require(type(results) is list and len(results) == 1 and results[0].get("case") == case, f"baseline qualification case mismatch: {case}")
        elapsed = results[0].get("elapsed_ns", {})
        _require(type(elapsed.get("samples")) is list and elapsed["samples"] == sorted(elapsed["samples"]), f"baseline qualification elapsed vector mismatch: {case}")
        if case in (experiment.PRIMARY_CASE, "docx_producer_source_selected_paragraph", "pptx_producer_source_selected_slide"):
            _require((folder / "producer.json").is_file(), f"baseline qualification producer sidecar missing: {case}")
    # This invokes the imported 0885 semantic verifier for the primary corpus;
    # it only parses sealed JSON and never launches a child.
    experiment._load_expected_corpora()


def _validate_tools() -> dict[str, dict[str, Any]]:
    paths = {"time": TIME, "taskset": TASKSET, "python": Path(sys.executable)}
    result = {}
    for name, path in paths.items():
        _require(path.is_file() and os.access(path, os.X_OK), f"{name} tool is unavailable: {path}")
        result[name] = descriptor(path)
    result["python"]["version"] = platform.python_version()
    return result


def _validate_hardware() -> dict[str, Any]:
    environment, _ = _json_descriptor(BASELINE_PACKET / "environment.json", "baseline environment")
    _require(environment.get("capture_cpu") == 12, "environment receipt was captured on a different CPU")
    _require(12 in os.sched_getaffinity(0), "CPU 12 is unavailable for the frozen experiment")
    return {
        "environment_file": str((BASELINE_PACKET / "environment.json").resolve()),
        "capture_cpu": environment["capture_cpu"],
        "platform": environment.get("platform"),
        "uname": environment.get("uname"),
        "exclusive_host": False,
        "claim_scope": "CPU affinity only; no exclusive-host or physical-cold guarantee",
    }


def _binding_paths() -> dict[str, Path]:
    files: dict[str, Path] = {
        "candidate_build": PACKET / "build.json",
        "candidate_build_start": PACKET / "build-start.json",
        "candidate_build_helper": PACKET / "build.py",
        "candidate_source": PACKET / "candidate-source.json",
        "candidate_patch": PACKET / "candidate.patch",
        "candidate_native_build_start": PACKET / "native-build-start.json",
        "candidate_native_build": PACKET / "native-build.json",
        "candidate_native_build_log": PACKET / "native-build.log",
        "candidate_native_depfile": PACKET / "native.d",
        "candidate_allocator_build_start": PACKET / "allocator-build-start.json",
        "candidate_allocator_build": PACKET / "allocator-build.json",
        "candidate_allocator_build_log": PACKET / "allocator-build.log",
        "candidate_allocator_depfile": PACKET / "allocator.d",
        "baseline_build": BASELINE_PACKET / "build.json",
        "baseline_build_start": BASELINE_PACKET / "build-start.json",
        "baseline_build_helper": BASELINE_PACKET / "build.py",
        "baseline_source": BASELINE_PACKET / "source.json",
        "baseline_archive": BASELINE_PACKET / "source.tar.gz",
        "baseline_native_build_start": BASELINE_PACKET / "native-build-start.json",
        "baseline_native_build": BASELINE_PACKET / "native-build.json",
        "baseline_native_build_log": BASELINE_PACKET / "native-build.log",
        "baseline_native_depfile": BASELINE_PACKET / "native.d",
        "baseline_allocator_build_start": BASELINE_PACKET / "allocator-build-start.json",
        "baseline_allocator_build": BASELINE_PACKET / "allocator-build.json",
        "baseline_allocator_build_log": BASELINE_PACKET / "allocator-build.log",
        "baseline_allocator_depfile": BASELINE_PACKET / "allocator.d",
        "baseline_capture_helper": BASELINE_PACKET / "capture.py",
        "baseline_verify_helper": BASELINE_PACKET / "verify.py",
        "baseline_test_helper": BASELINE_PACKET / "test_helpers.py",
        "baseline_freeze_helper": BASELINE_PACKET / "freeze.py",
        "baseline_environment": BASELINE_PACKET / "environment.json",
        "hardware_environment": BASELINE_PACKET / "environment.json",
        "experiment_helper": PACKET / "experiment.py",
        "reduction_helper": PACKET / "reduce.py",
        "freeze_helper": PACKET / "freeze.py",
        "helper_checks": PACKET / "helper-checks.json",
        "semantic_review": PACKET / "semantic-review.json",
        "methodology_review": PACKET / "methodology-review.json",
        "tool_time": TIME,
        "tool_taskset": TASKSET,
        "tool_python": Path(sys.executable),
    }
    for path in sorted(BASELINE_QUALIFICATION.rglob("*")):
        if path.is_file():
            files[f"baseline_qualification/{path.relative_to(BASELINE_QUALIFICATION).as_posix()}"] = path
    for path in sorted(QUALITY.rglob("*")):
        if path.is_file():
            files[f"quality_v2/{path.relative_to(QUALITY).as_posix()}"] = path
    for relative in OVERLAY_PATHS:
        files[f"candidate_overlay/{relative}"] = ROOT / relative
    return files


def _make_bindings() -> dict[str, dict[str, Any]]:
    files = _binding_paths()
    bindings = {}
    for key, path in files.items():
        _require(path.is_file(), f"required freeze binding is missing: {path}")
        bindings[key] = _file_binding(path)
    return bindings


def _validate_bindings(plan: dict[str, Any], *, retained_only: bool = False) -> None:
    bindings = plan.get("bindings")
    _require(isinstance(bindings, dict), "plan bindings are missing")
    expected = _binding_paths()
    for key, path in expected.items():
        _require(key in bindings, f"plan binding is missing: {key}")
        _require(_path(bindings[key].get("path", "")) == path.resolve(), f"plan binding path mismatch: {key}")
    # Validate every provided binding, including future additions, so an extra
    # custody item cannot silently evade verification.
    for key, item in bindings.items():
        if retained_only and key.startswith("candidate_overlay/"):
            _require(isinstance(item, dict) and isinstance(item.get("bytes"), int) and isinstance(item.get("sha256"), str), f"retained overlay binding {key} is incomplete")
        else:
            _check_descriptor(item, f"plan binding {key}")
    _require(_path(plan.get("candidate_build_json", "")) == (PACKET / "build.json").resolve(), "candidate build JSON path is not frozen")
    _require(_path(plan.get("baseline_build_json", "")) == (BASELINE_PACKET / "build.json").resolve(), "baseline build JSON path is not frozen")


def _validate_binary_plan(plan: dict[str, Any], before: dict[str, Any], after: dict[str, Any]) -> None:
    _require(plan.get("binaries", {}).get("before") == before["binaries"], "before binary plan entries differ from retained 0885 build")
    _require(plan.get("binaries", {}).get("after") == after["binaries"], "after binary plan entries differ from candidate build")
    experiment.validate_plan(plan)
    for role in BUILD_ROLES:
        for side, build in (("before", before), ("after", after)):
            item = plan["binaries"][side][role]
            _check_descriptor(item, f"plan binary {side}/{role}")
            _check_descriptor(item["build_receipt"], f"plan binary {side}/{role} build receipt")


def _validate_prerequisites(*, require_scratch: bool, require_live: bool) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any], dict[str, Any], dict[str, Any]]:
    baseline, _, _ = _validate_build(BASELINE_PACKET / "build.json", "baseline build", candidate=False)
    candidate, _, _ = _validate_build(PACKET / "build.json", "candidate build", candidate=True)
    _validate_build_context(baseline, candidate)
    candidate_source, _ = _validate_candidate_source(candidate_source=_read(PACKET / "candidate-source.json"), baseline_build=baseline, require_scratch=require_scratch, require_live=require_live)
    _validate_baseline_qualification()
    quality = _validate_quality()
    semantic, methodology = _validate_reviews()
    checks = _validate_helper_checks()
    return baseline, candidate, candidate_source, quality, {"semantic": semantic, "methodology": methodology, "checks": checks}


def _make_plan() -> dict[str, Any]:
    baseline, candidate, candidate_source, quality, reviews = _validate_prerequisites(require_scratch=True, require_live=True)
    overlays = _validate_source_overlays(candidate_source, require_scratch=True, require_live=True)
    tools = _validate_tools()
    hardware = _validate_hardware()
    bindings = _make_bindings()
    plan = {
        "schema": experiment.SCHEMA,
        "frozen_utc": _datetime.datetime.now(_datetime.timezone.utc).isoformat(),
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
        "candidate_build_json": str((PACKET / "build.json").resolve()),
        "baseline_build_json": str((BASELINE_PACKET / "build.json").resolve()),
        "revision": candidate_source["baseline_revision"],
        "bindings": bindings,
        "source": {
            "baseline_manifest": descriptor(BASELINE_PACKET / "source.json"),
            "baseline_archive": descriptor(BASELINE_PACKET / "source.tar.gz"),
            "candidate_source": descriptor(PACKET / "candidate-source.json"),
            "candidate_patch": descriptor(PACKET / "candidate.patch"),
            "overlays": overlays,
            "source_root_checked_before_cleanup": candidate_source.get("source_root"),
        },
        "quality_gates": {
            gate: {
                "start": descriptor(QUALITY / f"{gate}-start.json" if (QUALITY / f"{gate}-start.json").is_file() else QUALITY / f"{gate}.json"),
                "receipt": descriptor(QUALITY / f"{gate}.json"),
                "log": descriptor(QUALITY / f"{gate}.log"),
                "exit_code": quality[gate]["receipt"]["exit_code"],
            }
            for gate in QUALITY_GATES
        },
        "reviews": {
            "semantic": descriptor(PACKET / "semantic-review.json"),
            "methodology": descriptor(PACKET / "methodology-review.json"),
            "semantic_status": reviews["semantic"].get("status"),
            "methodology_status": reviews["methodology"].get("status"),
        },
        "tools": tools,
        "environment": {
            "child": CHILD_ENV,
            "build": _read(PACKET / "build-start.json").get("environment"),
            "hardware_binding": "hardware_environment",
        },
        "hardware": hardware,
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
            "primary_p95_p99_rss_review_ratio_above": 1.05,
            "primary_review_metrics": ["p50_ns", "p95_ns", "p99_ns", "peak_rss_kib"],
            "allocator_adoption": "not_evaluated",
            "bootstrap": {"seed": experiment.BOOTSTRAP_SEED, "draws": 10000, "ci95_indices": [249, 9749], "unit": "paired child ratios"},
            "performance_claim": None,
        },
        "host_scope": "CPU affinity only; no exclusive-host or physical-cold guarantee",
    }
    _validate_binary_plan(plan, baseline, candidate)
    _validate_bindings(plan)
    return plan


def _verify_plan(*, retained_only: bool) -> dict[str, Any]:
    _require(PLAN.is_file(), "plan.json is missing")
    plan = _read(PLAN)
    _require(isinstance(plan, dict), "plan.json must be an object")
    baseline, candidate, candidate_source, _, _ = _validate_prerequisites(require_scratch=False, require_live=not retained_only)
    _validate_source_overlays(candidate_source, require_scratch=False, require_live=not retained_only)
    _validate_binary_plan(plan, baseline, candidate)
    _validate_bindings(plan, retained_only=retained_only)
    for relative in OVERLAY_PATHS:
        frozen = plan.get("source", {}).get("overlays", {}).get(relative, {})
        expected = candidate_source["overlays"][relative]
        _require(frozen.get("bytes") == expected.get("bytes") and frozen.get("sha256") == expected.get("sha256"), f"frozen overlay identity changed: {relative}")
    if not retained_only:
        _require(plan.get("source", {}).get("overlays") == {relative: descriptor(ROOT / relative) for relative in OVERLAY_PATHS}, "plan source overlay identities changed")
    _require(plan.get("quality_gates", {}).keys() >= set(QUALITY_GATES), "plan omits a quality gate")
    return {"status": "pass", "plan": str(PLAN.resolve()), "plan_sha256": _sha256(PLAN), "mutated": False}


def _write_create_only(path: Path, value: Any) -> None:
    with path.open("x", encoding="utf-8") as handle:
        json.dump(value, handle, indent=2, sort_keys=True)
        handle.write("\n")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--verify", action="store_true", help="verify an existing plan without writing")
    parser.add_argument("--retained-only", action="store_true", help="with --verify, allow the candidate source checkout to be gone or reverted")
    args = parser.parse_args()
    _require(not args.retained_only or args.verify, "--retained-only requires --verify")
    if args.verify:
        result = _verify_plan(retained_only=args.retained_only)
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0
    _require(not PLAN.exists(), f"refusing to overwrite existing plan: {PLAN}")
    plan = _make_plan()
    _write_create_only(PLAN, plan)
    print(json.dumps({"status": "pass", "plan": str(PLAN.resolve()), "children": experiment.TOTAL_CHILDREN, "mutated": True}, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
