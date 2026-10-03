#!/usr/bin/env python3
"""Freeze an 0894 experiment plan from sealed source/build receipts.

This helper only reads receipts and executable descriptors and creates one
plan. It never builds, copies, launches, or mutates a source tree. The build
agent may retain binaries wherever it owns them; the plan binds their exact
path, size, mode, and SHA-256 for the later capture helper.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import importlib.util
import json
from pathlib import Path
import re
import sys
from typing import Any


PACKET = Path(__file__).resolve().parent
_EXPERIMENT_PATH = PACKET / "experiment.py"
_EXPERIMENT_SPEC = importlib.util.spec_from_file_location("litchi_selected_metadata_experiment_0894_freeze", _EXPERIMENT_PATH)
if _EXPERIMENT_SPEC is None or _EXPERIMENT_SPEC.loader is None:
    raise ImportError(f"cannot load local experiment helper: {_EXPERIMENT_PATH}")
experiment = importlib.util.module_from_spec(_EXPERIMENT_SPEC)
sys.modules[_EXPERIMENT_SPEC.name] = experiment
_EXPERIMENT_SPEC.loader.exec_module(experiment)


PLAN_SCHEMA = experiment.SCHEMA


def _read(path: Path) -> Any:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def _descriptor(path: Path) -> dict[str, Any]:
    return experiment.descriptor(path)


def _ref(path: Path) -> dict[str, Any]:
    return experiment.descriptor(path)


def _binary_from(value: Any, label: str) -> dict[str, Any]:
    """Find a native/allocator descriptor in common sealed build layouts."""
    metadata: dict[str, Any] = {}
    if isinstance(value, str):
        path = Path(value)
        _require(path.is_file() and not path.is_symlink(), f"{label} binary is missing or symlinked: {path}")
        item = _descriptor(path)
    elif isinstance(value, dict):
        if isinstance(value.get("identity"), dict):
            item = dict(value["identity"])
            item.update({key: val for key, val in value.items() if key != "identity"})
        elif isinstance(value.get("descriptor"), dict):
            item = dict(value["descriptor"])
            item.update({key: val for key, val in value.items() if key != "descriptor"})
        else:
            item = value
        metadata = {
            key: val for key, val in item.items()
            if key not in {"path", "bytes", "sha256", "mode", "executable", "identity", "descriptor"}
        }
        path = Path(str(item.get("path", "")))
        _require(path.is_file() and not path.is_symlink(), f"{label} binary is missing or symlinked: {path}")
        actual = _descriptor(path)
        for key in ("bytes", "sha256", "mode", "executable"):
            if key in item:
                _require(item[key] == actual[key], f"{label} {key} does not match the file")
        item = actual
    else:
        raise ValueError(f"{label} descriptor is invalid")
    _require(item["executable"] is True, f"{label} is not executable")
    # Preserve the build-owned metadata.  In particular, role/name/features,
    # the exact role build receipt and depfile, and the ELF build ID must travel
    # into the frozen plan alongside the file identity.
    result = dict(item)
    result.update(metadata)
    result.update({key: item[key] for key in ("path", "bytes", "sha256", "mode", "executable")})
    return result


def _build_binary(build: dict[str, Any], mode: str, label: str) -> dict[str, Any]:
    candidates = []
    binaries = build.get("binaries")
    if isinstance(binaries, dict):
        candidates.append(binaries.get(mode))
    candidates.append(build.get(mode))
    candidates.append(build.get(f"{mode}_binary"))
    for candidate in candidates:
        if candidate is not None:
            return _binary_from(candidate, label)
    raise ValueError(f"{label} has no {mode} binary descriptor")


def _build_configuration(build: dict[str, Any], label: str) -> dict[str, Any]:
    raw = build.get("flags") or build.get("build_flags") or build.get("configuration")
    if raw is None:
        environment = build.get("environment") or {}
        compiler = str(build.get("compiler", ""))
        match = re.search(r"rustc\s+(1\.95\.0[^\s]*)", compiler)
        raw = {
            "toolchain": match.group(1) if match else None,
            "profile": "release" if "--release" in build.get("argv", []) or "CARGO_PROFILE_RELEASE_OPT_LEVEL" in environment else None,
            "opt_level": environment.get("CARGO_PROFILE_RELEASE_OPT_LEVEL"),
            "debug": environment.get("CARGO_PROFILE_RELEASE_DEBUG"),
            "lto": environment.get("CARGO_PROFILE_RELEASE_LTO"),
            "codegen_units": environment.get("CARGO_PROFILE_RELEASE_CODEGEN_UNITS"),
            "jobs": environment.get("CARGO_BUILD_JOBS"),
        }
    _require(isinstance(raw, dict), f"{label} build flags are not an object")

    def pick(*names: str) -> Any:
        for name in names:
            if name in raw:
                return raw[name]
            if name in build:
                return build[name]
        return None

    toolchain = pick("rust_toolchain", "toolchain")
    if isinstance(toolchain, str):
        _require(toolchain.startswith("1.95.0"), f"{label} toolchain is not Rust 1.95")
    else:
        raise ValueError(f"{label} toolchain is missing")
    values = {
        "toolchain": toolchain,
        "profile": pick("profile"),
        "opt_level": pick("opt_level", "optimization_level"),
        "debug": pick("debug", "debuginfo", "debug_info"),
        "lto": pick("lto"),
        "codegen_units": pick("codegen_units", "codegen-units"),
        "jobs": pick("jobs", "build_jobs"),
    }
    _require(values["profile"] == "release", f"{label} profile is not release")
    _require(values["opt_level"] in (3, "3"), f"{label} opt-level is not 3")
    _require(values["debug"] in (1, "1", True), f"{label} debug info is not enabled")
    _require(values["lto"] == "fat", f"{label} LTO is not fat")
    _require(values["codegen_units"] in (1, "1"), f"{label} codegen-units is not 1")
    _require(values["jobs"] in (2, "2"), f"{label} build jobs are not 2")
    return values


def _source_contract(source: dict[str, Any], source_path: Path, candidate_path: Path, candidate: dict[str, Any]) -> dict[str, Any]:
    _require(source_path.resolve() == (experiment.PACKET / "source.json").resolve(), "source receipt is not the packet manifest")
    _require(source.get("schema") == experiment.SOURCE_SCHEMA, "source receipt schema mismatch")
    _require(source.get("status") == "pass", "source receipt does not pass")
    for key in ("revision", "baseline_commit", "source_revision"):
        _require(source.get(key) == experiment.BASE_COMMIT, f"source receipt {key} mismatch")
    _require(source.get("timed_harness_guard_contract") == experiment.GUARD_CONTRACT, "source receipt guard mismatch")
    _require(source.get("selection") == experiment.SOURCE_SELECTION, "source receipt selection changed")
    for key in ("file_count", "selected_file_count", "source_file_count", "live_files_hash_checked"):
        _require(source.get(key) == experiment.SOURCE_FILE_COUNT, f"source receipt {key} census mismatch")
    files = source.get("files")
    _require(isinstance(files, dict) and len(files) == experiment.SOURCE_FILE_COUNT, "source receipt file census is incomplete")
    for relative_path, expected_sha in experiment.BASELINE_QUALITY_SOURCE_HASHES.items():
        _require(files.get(relative_path, {}).get("sha256") == expected_sha, f"source archive {relative_path} does not match 0893 v3 quality")
    archive = source.get("archive")
    _require(isinstance(archive, dict), "source archive identity is missing")
    archive_path = Path(str(archive.get("path", "")))
    _require(archive_path.is_file() and not archive_path.is_symlink(), "source archive is not a regular non-symlink file")
    _require(archive_path.resolve() == (experiment.PACKET / "source.tar.gz").resolve(), "source archive path is not the packet archive")
    actual_archive = _descriptor(archive_path)
    _require(archive == actual_archive, "source archive descriptor does not match retained archive")
    _require(actual_archive["bytes"] == experiment.SOURCE_ARCHIVE_BYTES and actual_archive["sha256"] == experiment.SOURCE_ARCHIVE_SHA256, "source archive identity is not the exact 10771-file archive")
    candidate_receipt = experiment._candidate_freeze_contract(candidate_path, source, source_path, actual_archive)
    return {
        "schema": source["schema"],
        "status": source["status"],
        "receipt": _ref(source_path),
        "revision": source["revision"],
        "baseline_commit": source["baseline_commit"],
        "source_revision": source["source_revision"],
        "timed_harness_guard_contract": source["timed_harness_guard_contract"],
        "selection": list(source["selection"]),
        "file_count": source["file_count"],
        "selected_file_count": source["selected_file_count"],
        "source_file_count": source["source_file_count"],
        "live_files_hash_checked": source["live_files_hash_checked"],
        "archive_bytes": actual_archive["bytes"],
        "archive_sha256": actual_archive["sha256"],
        "candidate_freeze": _ref(candidate_path),
        "candidate_overlay": candidate_receipt["candidate_patch"],
    }


def _load_bindings(path: Path) -> dict[str, Any]:
    raw = _read(path)
    _require(isinstance(raw, dict), "binding manifest must be an object")
    _require(set(raw) == set(experiment.REQUIRED_BINDINGS), "binding manifest set is incomplete or unexpected")
    bindings = {}
    for name in experiment.REQUIRED_BINDINGS:
        value = raw[name]
        _require(isinstance(value, dict), f"binding manifest entry must include metadata: {name}")
        expected = experiment.BINDING_SPECS[name]
        for key, expected_value in expected.items():
            _require(value.get(key) == expected_value, f"binding manifest {name} {key} is not {expected_value!r}")
        candidate = Path(str(value.get("path", "")))
        if not candidate.is_absolute():
            candidate = path.parent / candidate
        _require(candidate.is_file() and not candidate.is_symlink(), f"binding manifest file is missing: {name}")
        candidate = candidate.resolve()
        actual = _descriptor(candidate)
        for key in ("bytes", "sha256", "mode", "executable"):
            if key in value:
                _require(value[key] == actual[key], f"binding manifest {name} {key} mismatch")
        bindings[name] = {**actual, **expected}
    return bindings


def freeze(source_path: Path, baseline_build_path: Path, candidate_build_path: Path, bindings_path: Path, output: Path) -> dict[str, Any]:
    for path in (source_path, baseline_build_path, candidate_build_path, bindings_path):
        _require(path.is_file(), f"receipt is missing: {path}")
    source = _read(source_path)
    baseline_build = _read(baseline_build_path)
    candidate_build = _read(candidate_build_path)
    _require(isinstance(source, dict) and isinstance(baseline_build, dict) and isinstance(candidate_build, dict), "receipts must be objects")
    bindings = _load_bindings(bindings_path)
    candidate_freeze_path = Path(bindings["candidate_freeze"]["path"])
    source_contract = _source_contract(source, source_path, candidate_freeze_path, _read(candidate_freeze_path))
    experiment._same_descriptor(bindings["source_identity"], _ref(source_path), "source identity binding")
    archive = source.get("archive")
    _require(isinstance(archive, dict), "source archive is missing")
    experiment._same_descriptor(bindings["source_archive"], archive, "source archive binding")
    experiment._same_descriptor(bindings["baseline_build"], _ref(baseline_build_path), "baseline build binding")
    experiment._same_descriptor(bindings["candidate_build"], _ref(candidate_build_path), "candidate build binding")
    for name, build in (("baseline", baseline_build), ("candidate", candidate_build)):
        _require(build.get("status") == "pass", f"{name} build receipt does not pass")
        _require(build.get("profile") == "release", f"{name} build is not release")
        guard = build.get("timed_harness_guard_contract") or build.get("guard_contract")
        _require(guard == experiment.GUARD_CONTRACT, f"{name} build guard contract mismatch")
        for key in ("base_commit", "baseline_commit", "source_revision"):
            _require(build.get(key) == experiment.BASE_COMMIT, f"{name} build {key} mismatch")
    baseline_configuration = _build_configuration(baseline_build, "baseline")
    candidate_configuration = _build_configuration(candidate_build, "candidate")
    _require(candidate_configuration == baseline_configuration == experiment.BUILD_CONFIGURATION, "baseline/candidate build flags differ")
    before = {mode: _build_binary(baseline_build, mode, f"baseline/{mode}") for mode in ("native", "allocator")}
    after = {mode: _build_binary(candidate_build, mode, f"candidate/{mode}") for mode in ("native", "allocator")}
    for mode in ("native", "allocator"):
        before[mode]["evidence_bindings"] = experiment.binary_evidence_bindings("before", mode)
        after[mode]["evidence_bindings"] = experiment.binary_evidence_bindings("after", mode)
    plan = {
        "schema": PLAN_SCHEMA,
        "status": "frozen",
        "frozen_utc": _datetime.datetime.now(_datetime.timezone.utc).isoformat(),
        "base_commit": experiment.BASE_COMMIT,
        "profile": experiment.PROFILE,
        "cpu_affinity": experiment.CPU_AFFINITY,
        "cache": "warm",
        "cases": list(experiment.CASES),
        "xlsx_shape": "medium",
        "native_blocks": experiment.NATIVE_BLOCKS,
        "native_warmup": experiment.NATIVE_WARMUP,
        "native_samples": experiment.NATIVE_SAMPLES,
        "allocator_blocks": experiment.ALLOCATOR_BLOCKS,
        "allocator_warmup": experiment.ALLOCATOR_WARMUP,
        "allocator_samples": experiment.ALLOCATOR_SAMPLES,
        "bootstrap_seed": experiment.BOOTSTRAP_SEED,
        "bootstrap_draws": experiment.BOOTSTRAP_DRAWS,
        "bootstrap_ci_indices": list(experiment.BOOTSTRAP_CI_INDICES),
        "total_children": experiment.TOTAL_CHILDREN,
        "fresh_child": True,
        "fail_fast": True,
        "guard_contract": experiment.GUARD_CONTRACT,
        "source": source_contract,
        "builds": {"before": _ref(baseline_build_path), "after": _ref(candidate_build_path)},
        "build_configuration": baseline_configuration,
        "bindings": bindings,
        "fixtures": experiment.fixture_descriptors(),
        "helper_hashes": experiment.helper_descriptors(),
        "binaries": {"before": before, "after": after},
        "primary_gate": {"metric": "native primary p50 median ratio", "median_ratio_at_most": 0.98, "ci95_upper_below": 1.0},
        "review_flags": {"pair_ratio_above": experiment.PAIR_RATIO_REVIEW_THRESHOLD, "native_p50_spread_above": experiment.P50_SPREAD_REVIEW_THRESHOLD, "rss_causal_claim": False},
        "performance_claim": None,
    }
    experiment.validate_plan(plan)
    experiment._validate_bindings(plan)
    _write = output.resolve()
    _require(not _write.exists(), f"plan output already exists: {_write}")
    _write.parent.mkdir(parents=True, exist_ok=True)
    with _write.open("x", encoding="utf-8") as handle:
        json.dump(plan, handle, indent=2, sort_keys=True)
        handle.write("\n")
    return plan


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--baseline-build", type=Path, required=True)
    parser.add_argument("--candidate-build", type=Path, required=True)
    parser.add_argument("--bindings", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    print(json.dumps(freeze(args.source.resolve(), args.baseline_build.resolve(), args.candidate_build.resolve(), args.bindings.resolve(), args.output.resolve()), indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
