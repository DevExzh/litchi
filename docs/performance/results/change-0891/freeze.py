#!/usr/bin/env python3
"""Freeze the 0891 measurement plan and verify retained packet custody.

This is a small packet-level gate. It binds the sealed source patch, source
and methodology reviews, passing quality receipt, candidate release build,
retained 0885 release binaries, corpora, tools, and managed executable copies
into one create-only plan. It performs no build and launches no workload.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
from pathlib import Path
import platform
import stat
import subprocess
import sys
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))
import experiment


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
BASELINE_PACKET = PACKET.parent / "change-0885"
CORPUS_PACKET = PACKET.parent / "change-0887" / "baseline-qualification"
FREEZE = PACKET / "candidate-freeze-v3.json"
PATCH = PACKET / "candidate-v3.patch"
SOURCE_REVIEW = PACKET / "source-review-v3.json"
METHODOLOGY_REVIEW = PACKET / "methodology-review-v4.json"
QUALITY = PACKET / "quality.json"
BUILD = PACKET / "build.json"
BUILD_INVOCATION = PACKET / "build-invocation.json"
BASELINE_BUILD = BASELINE_PACKET / "build.json"
EXECUTION = PACKET / "execution-files.json"
FIXTURE_CENSUS = PACKET / "fixture-census.json"
PLAN = PACKET / "plan.json"
CURRENT_BASE = "3348479699f248223184c7bdbc005879b0153307"
BASELINE_REVISION = "335a40350bf06f49691745c14d310fd7a3ca9117"
CASES = tuple(experiment.CASES)
REQUIRED_HELPERS = (
    "build.py",
    "build-executed.py",
    "prepare_exec.py",
    "experiment.py",
    "reduce.py",
    "verify_packet.py",
    "freeze.py",
    "quality.py",
    "quality_runner.py",
)
IMPORTED_HELPERS = (
    PACKET.parent / "change-0887" / "experiment.py",
    PACKET.parent / "change-0887" / "reduce.py",
    PACKET.parent / "change-0887" / "verify_packet.py",
    BASELINE_PACKET / "capture.py",
    BASELINE_PACKET / "verify.py",
)


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
    require(stat.S_ISREG(info.st_mode) and not path.is_symlink(), f"expected regular packet artifact: {path}")
    return {"path": str(path), "bytes": info.st_size, "sha256": sha256(path), "mode": stat.S_IMODE(info.st_mode), "executable": bool(info.st_mode & stat.S_IXUSR)}


def path_from(raw: Any, base: Path = PACKET) -> Path:
    require(isinstance(raw, str), "descriptor path is missing")
    value = Path(raw)
    return (value if value.is_absolute() else base / value).resolve()


def packet_or_root_path(raw: Any) -> Path:
    require(isinstance(raw, str), "packet descriptor path is missing")
    value = Path(raw)
    if value.is_absolute():
        return value.resolve()
    if value.parts[:2] == ("docs", "performance"):
        return (ROOT / value).resolve()
    return (PACKET / value).resolve()


def check_descriptor(raw: Any, label: str, *, base: Path = PACKET) -> dict[str, Any]:
    require(isinstance(raw, dict) and isinstance(raw.get("path"), str), f"{label} descriptor is invalid")
    actual = descriptor(path_from(raw["path"], base))
    for key in ("bytes", "sha256", "mode", "executable"):
        if key in raw:
            require(raw[key] == actual[key], f"{label} {key} differs from its retained file")
    return actual


def write_once(path: Path, value: Any) -> None:
    with path.open("x", encoding="utf-8") as handle:
        json.dump(value, handle, indent=2, sort_keys=True)
        handle.write("\n")


def check_status(path: Path, label: str, accepted: set[str]) -> tuple[dict[str, Any], dict[str, Any]]:
    require(path.is_file(), f"{label} is missing: {path}")
    value = read(path)
    require(isinstance(value, dict) and value.get("status") in accepted, f"{label} status is not passing")
    require(value.get("performance_claim") in (None, False), f"{label} makes a performance claim")
    return value, descriptor(path)


def candidate_freeze() -> tuple[dict[str, Any], dict[str, dict[str, Any]], dict[str, Any]]:
    require(FREEZE.is_file(), f"candidate freeze is missing: {FREEZE}")
    freeze = read(FREEZE)
    require(isinstance(freeze, dict), "candidate freeze must be an object")
    require(freeze.get("performance_claim") in (None, False), "candidate freeze makes a performance claim")
    require(freeze.get("schema") == "litchi.xlsx-x14ac-candidate-freeze-0891.v1", "candidate freeze schema mismatch")
    predecessor = freeze.get("supersedes")
    require(isinstance(predecessor, dict), "candidate freeze v3 predecessor descriptor is missing")
    predecessor_path = packet_or_root_path(predecessor.get("path"))
    predecessor_file = PACKET / "candidate-freeze-v2.json"
    require(predecessor_path == predecessor_file.resolve(), "candidate freeze v3 predecessor path mismatch")
    predecessor_actual = descriptor(predecessor_file)
    require(predecessor.get("bytes") == predecessor_actual["bytes"] and predecessor.get("sha256") == predecessor_actual["sha256"], "candidate freeze v2 predecessor identity mismatch")
    require(freeze.get("base_commit") == CURRENT_BASE, "candidate freeze base commit mismatch")
    paths = freeze.get("candidate_paths")
    expected_paths = {
        "crates/litchi-xlsx/src/raw/worksheet/x14ac.rs",
        "crates/litchi-xlsx/src/raw/worksheet/x14ac/reuse_tests.rs",
    }
    require(isinstance(paths, list) and set(paths) == expected_paths, "candidate freeze path set mismatch")
    records = freeze.get("files")
    require(isinstance(records, list) and len(records) == len(expected_paths), "candidate freeze files are incomplete")
    files: dict[str, dict[str, Any]] = {}
    for record in records:
        require(isinstance(record, dict) and isinstance(record.get("path"), str), "candidate freeze file descriptor is invalid")
        name = record["path"]
        require(name in expected_paths and name not in files, f"candidate freeze has unexpected file: {name}")
        files[name] = record
    require(set(files) == expected_paths, "candidate freeze files differ from candidate paths")
    patch_record = freeze.get("patch")
    require(isinstance(patch_record, dict), "candidate patch descriptor is missing")
    require(packet_or_root_path(patch_record.get("path")) == PATCH.resolve(), "candidate patch path mismatch")
    check_descriptor({**patch_record, "path": str(PATCH)}, "candidate patch")
    require(freeze.get("immutable") is True or freeze.get("create_only") is True, "candidate freeze does not declare immutable custody")
    for name, record in files.items():
        current = ROOT / name
        # The source may be reverted or removed after measurement. During plan
        # creation it must still agree with the sealed overlay record.
        if current.exists():
            actual = descriptor(current)
            require(actual["bytes"] == record.get("bytes") and actual["sha256"] == record.get("sha256"), f"candidate overlay changed: {name}")
    return freeze, files, {"candidate_freeze": descriptor(FREEZE), "candidate_patch": descriptor(PATCH)}


def _path_hash_references(value: Any):
    if isinstance(value, dict):
        path = value.get("path")
        digest = value.get("sha256")
        if isinstance(path, str) and isinstance(digest, str):
            yield path, digest
        for child in value.values():
            yield from _path_hash_references(child)
    elif isinstance(value, list):
        for child in value:
            yield from _path_hash_references(child)


def require_final_candidate_references(value: Any, label: str, *, patch_required: bool) -> None:
    freeze_hash = sha256(FREEZE)
    patch_hash = sha256(PATCH)
    references = []
    for raw_path, digest in _path_hash_references(value):
        try:
            path = packet_or_root_path(raw_path)
        except (TypeError, ValueError):
            continue
        references.append((path, digest))
    require(any(path == FREEZE.resolve() and digest == freeze_hash for path, digest in references), f"{label} does not bind candidate-freeze-v3.json and its hash")
    if patch_required:
        require(any(path == PATCH.resolve() and digest == patch_hash for path, digest in references), f"{label} does not bind candidate-v3.patch and its hash")


def binary_descriptor(build: dict[str, Any], role: str) -> dict[str, Any]:
    raw = build.get("binaries", {}).get(role)
    require(isinstance(raw, dict), f"{role} binary descriptor is missing")
    check_descriptor(raw, f"{role} binary")
    require(raw.get("executable") is True, f"{role} binary is not executable")
    require(isinstance(raw.get("bytes"), int) and isinstance(raw.get("sha256"), str), f"{role} binary identity is incomplete")
    return raw


def managed_descriptor(execution: dict[str, Any], role: str, mode: str) -> dict[str, Any]:
    raw = execution.get("records", {}).get(role, {}).get(mode)
    require(isinstance(raw, dict) and isinstance(raw.get("copy"), dict), f"managed {role}/{mode} descriptor is missing")
    copy = raw["copy"]
    check_descriptor(copy, f"managed {role}/{mode}")
    require(copy.get("executable") is True, f"managed {role}/{mode} is not executable")
    source = raw.get("source")
    require(isinstance(source, dict) and source.get("sha256") == copy.get("sha256") and source.get("bytes") == copy.get("bytes"), f"managed {role}/{mode} source identity differs")
    return copy


def corpus_bindings() -> dict[str, Any]:
    result: dict[str, Any] = {}
    for case in CASES:
        folder = CORPUS_PACKET / case
        report = folder / "report.json"
        require(report.is_file(), f"sealed corpus report is missing: {case}")
        report_value = read(report)
        require(report_value.get("tool", {}).get("profile") == "release", f"corpus is not release: {case}")
        case_rows = report_value.get("results")
        require(isinstance(case_rows, list) and len(case_rows) == 1 and case_rows[0].get("case") == case, f"corpus case mismatch: {case}")
        item: dict[str, Any] = {"report": descriptor(report)}
        producer = folder / "producer.json"
        if producer.is_file():
            item["producer"] = descriptor(producer)
        elif case in (CASES[0], CASES[2], CASES[3]):
            raise ValueError(f"producer sidecar is required: {case}")
        result[case] = item
    return result


def fixture_binding() -> dict[str, Any] | None:
    if not FIXTURE_CENSUS.is_file():
        return None
    value = read(FIXTURE_CENSUS)
    require(value.get("status") == "pass" and value.get("performance_claim") in (None, False), "fixture census is not passing")
    return descriptor(FIXTURE_CENSUS)


def helper_bindings() -> dict[str, Any]:
    values: dict[str, Any] = {}
    for name in REQUIRED_HELPERS:
        values[name] = descriptor(PACKET / name)
    for path in IMPORTED_HELPERS:
        values[f"imported/{path.parent.parent.name}/{path.name}"] = descriptor(path)
    for path in (Path("/usr/bin/taskset"), Path("/usr/bin/time")):
        require(path.is_file(), f"required capture tool is missing: {path}")
        values[f"tool/{path.name}"] = descriptor(path)
    return values


def build_execution_binding() -> dict[str, Any]:
    require(BUILD_INVOCATION.is_file(), f"executed build invocation is missing: {BUILD_INVOCATION}")
    invocation = read(BUILD_INVOCATION)
    require(invocation.get("schema") == "litchi.xlsx-x14ac-build-invocation-0891.v1", "build invocation schema mismatch")
    helper = invocation.get("helper")
    require(isinstance(helper, dict), "build invocation helper descriptor is missing")
    executed = PACKET / "build-executed.py"
    require(packet_or_root_path(helper.get("path")) == executed.resolve(), "build invocation helper path mismatch")
    helper_actual = check_descriptor({**helper, "path": str(executed)}, "executed build helper")
    require(helper_actual["sha256"] == helper.get("sha256"), "build invocation helper hash mismatch")
    return {"invocation": descriptor(BUILD_INVOCATION), "helper": helper_actual}


def equivalent_environment(baseline: dict[str, Any], candidate: dict[str, Any]) -> dict[str, Any]:
    keys = ("CARGO_BUILD_JOBS", "CARGO_INCREMENTAL", "CARGO_NET_OFFLINE", "CARGO_PROFILE_RELEASE_OPT_LEVEL", "CARGO_PROFILE_RELEASE_DEBUG", "CARGO_PROFILE_RELEASE_LTO", "CARGO_PROFILE_RELEASE_CODEGEN_UNITS", "LANG", "LC_ALL", "TZ")
    for key in keys:
        require(baseline.get("environment", {}).get(key) == candidate.get("environment", {}).get(key), f"build environment differs for {key}")
    require(baseline.get("compiler") == candidate.get("compiler"), "candidate compiler differs from retained 0885 compiler")
    return {key: candidate.get("environment", {}).get(key) for key in keys}


def make_plan() -> dict[str, Any]:
    require(not PLAN.exists(), f"plan already exists: {PLAN}")
    freeze, overlays, source_records = candidate_freeze()
    source_review, source_review_descriptor = check_status(SOURCE_REVIEW, "source review", {"approved_for_measurement"})
    methodology, methodology_descriptor = check_status(METHODOLOGY_REVIEW, "methodology review", {"pass", "approved_for_measurement"})
    quality, quality_descriptor = check_status(QUALITY, "quality receipt", {"pass"})
    require_final_candidate_references(source_review, "source review", patch_required=True)
    require_final_candidate_references(quality, "quality receipt", patch_required=False)
    build, build_descriptor = check_status(BUILD, "candidate build", {"pass"})
    baseline, baseline_descriptor = check_status(BASELINE_BUILD, "retained baseline build", {"pass"})
    execution, execution_descriptor = check_status(EXECUTION, "managed executable custody", {"pass"})
    require(build.get("base_commit") == CURRENT_BASE and build.get("baseline_revision") == BASELINE_REVISION, "candidate build revision binding mismatch")
    require(baseline.get("revision") == BASELINE_REVISION, "retained baseline revision mismatch")
    environment = equivalent_environment(baseline, build)
    before = {mode: binary_descriptor(baseline, mode) for mode in ("native", "allocator")}
    after = {mode: binary_descriptor(build, mode) for mode in ("native", "allocator")}
    managed = {role: {mode: managed_descriptor(execution, role, mode) for mode in ("native", "allocator")} for role in ("before", "after")}
    corpora = corpus_bindings()
    fixture = fixture_binding()
    helpers = helper_bindings()
    build_execution = build_execution_binding()
    bindings: dict[str, Any] = {
        "candidate_freeze": source_records["candidate_freeze"],
        "candidate_patch": source_records["candidate_patch"],
        "baseline_source_archive": descriptor(BASELINE_PACKET / "source.tar.gz"),
        "baseline_source_manifest": descriptor(BASELINE_PACKET / "source.json"),
        "source_review": source_review_descriptor,
        "methodology_review": methodology_descriptor,
        "quality": quality_descriptor,
        "candidate_build": build_descriptor,
        "build_execution": build_execution,
        "baseline_build": baseline_descriptor,
        "execution_files": execution_descriptor,
        "corpora": corpora,
        "fixture_census": fixture,
        "helpers": helpers,
        "candidate_overlay": {name: {"path": str(ROOT / name), "bytes": item["bytes"], "sha256": item["sha256"]} for name, item in overlays.items()},
    }
    plan = {
        "schema": "litchi.xlsx-x14ac-experiment-0891.v1",
        "status": "frozen",
        "base_commit": CURRENT_BASE,
        "revision": BASELINE_REVISION,
        "profile": "release",
        "cpu_affinity": "12",
        "cache": "warm",
        "xlsx_shape": "medium",
        "cases": list(CASES),
        "native_blocks": 8,
        "allocator_blocks": 4,
        "native_warmup": 3,
        "native_samples": 50,
        "allocator_warmup": 3,
        "allocator_samples": 15,
        "total_children": 72,
        "fail_fast": True,
        "fresh_child": True,
        "bootstrap_seed": 891,
        "bootstrap_draws": 10_000,
        "bootstrap_ci_indices": [249, 9749],
        "primary_gate": {"metric": "native primary p50 median ratio", "median_ratio_at_most": 0.98, "ci95_upper_below": 1.0},
        "review_flags": {"pair_ratio_above": 1.05, "native_p50_spread_above": 1.20, "rss_causal_claim": False},
        "host_scope": "CPU affinity only; no exclusive-host or physical-cold-cache guarantee",
        "environment": environment,
        "binaries": {"before": before, "after": after},
        "managed_binaries": managed,
        "candidate_build_json": str(BUILD.resolve()),
        "baseline_build_json": str(BASELINE_BUILD.resolve()),
        "source": {
            "freeze": source_records["candidate_freeze"],
            "patch": source_records["candidate_patch"],
            "overlay_paths": sorted(overlays),
            "overlays": {name: {"path": str(ROOT / name), "bytes": item["bytes"], "sha256": item["sha256"]} for name, item in overlays.items()},
        },
        "reviews": {"source": source_review_descriptor, "methodology": methodology_descriptor, "quality": quality_descriptor},
        "corpora": corpora,
        "fixture_census": fixture,
        "tools": {name: value for name, value in helpers.items() if name.startswith("tool/")},
        "helpers": helpers,
        "bindings": bindings,
        "quality_summary": {key: quality[key] for key in ("schema", "status") if key in quality},
        "source_review_summary": {key: source_review[key] for key in ("schema", "status") if key in source_review},
        "methodology_review_summary": {key: methodology[key] for key in ("schema", "status") if key in methodology},
        "quality_gates": {"receipt": quality_descriptor, "status": quality.get("status")},
        "frozen_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
        "hardware": {"platform": platform.platform(), "machine": platform.machine(), "processor": platform.processor(), "cpu_count": __import__("os").cpu_count()},
        "performance_claim": None,
    }
    # This validates the exact plan consumed by the capture adapter before the
    # create-only write. It checks the 72-child protocol and all binary/copy
    # identities without starting a workload.
    experiment.validate_plan(plan)
    write_once(PLAN, plan)
    return plan


def verify_plan(retained_only: bool) -> dict[str, Any]:
    require(PLAN.is_file(), "plan.json is missing")
    plan = read(PLAN)
    require(isinstance(plan, dict) and plan.get("schema") == "litchi.xlsx-x14ac-experiment-0891.v1", "plan schema mismatch")
    require(Path(plan.get("candidate_build_json", "")).resolve() == BUILD.resolve(), "candidate_build_json binding mismatch")
    require(Path(plan.get("baseline_build_json", "")).resolve() == BASELINE_BUILD.resolve(), "baseline_build_json binding mismatch")
    managed_paths = [
        Path(str(plan["managed_binaries"][role][mode]["path"]))
        for role in ("before", "after")
        for mode in ("native", "allocator")
    ]
    present = [path.exists() for path in managed_paths]
    require(not any(present) or all(present), "managed executable custody is partial")
    plan["retained_only"] = retained_only or not any(present)
    experiment.validate_plan(plan)
    checked = 0
    skipped_managed = 0
    for name, record in _iter_records(plan.get("bindings", {})):
        if "candidate_overlay" in name:
            skipped_managed += 1
            continue
        path = path_from(record["path"], PLAN.parent)
        if retained_only and ("managed" in name or path.parent == PACKET / "executables"):
            if path.exists():
                check_descriptor(record, name, base=PLAN.parent)
            else:
                skipped_managed += 1
            continue
        check_descriptor(record, name, base=PLAN.parent)
        checked += 1
    return {"status": "pass", "plan": str(PLAN.resolve()), "plan_sha256": sha256(PLAN), "retained_only": retained_only, "bindings_checked": checked, "bindings_skipped": skipped_managed, "workload_launched": False}


def _iter_records(value: Any, prefix: str = ""):
    if isinstance(value, dict) and isinstance(value.get("path"), str) and isinstance(value.get("sha256"), str):
        yield prefix, value
    elif isinstance(value, dict):
        for key, child in value.items():
            name = f"{prefix}.{key}" if prefix else key
            yield from _iter_records(child, name)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--verify", action="store_true", help="verify the existing plan without creating a new one")
    parser.add_argument("--retained-only", action="store_true", help="permit removed managed copies and build roots")
    args = parser.parse_args()
    result = verify_plan(args.retained_only) if args.verify else make_plan()
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
