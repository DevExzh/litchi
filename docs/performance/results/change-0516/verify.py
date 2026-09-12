#!/usr/bin/env python3
"""Fail-closed final evidence verifier for the 0516 XLSX experiment.

The verifier is deliberately a composition layer. Raw elapsed vectors and
operation/allocation invariants are checked by the retained 0514 verifier and
the 0516 metrics helper; Callgrind attribution is checked by profiles.py.
This file binds those checks to the 0516 source epochs, receipts, commands,
quality gates, and decision contract. It never builds, captures, or makes a
performance claim.
"""

from __future__ import annotations

import argparse
import copy
import datetime as _datetime
import hashlib
import importlib.util
import json
import math
import re
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Callable


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
SCRATCH = Path("/tmp/litchi-goal-0516")
BASE_REVISION = "c5abaef129f5ae0000a925b295cf6507b7474e3c"
SCHEMA = "litchi-0516-verifier-v1"

CASES = (
    "xlsx_one_cell_commit",
    "xlsx_one_percent_commit",
    "xlsx_one_cell_commit_save",
    "xlsx_one_percent_commit_save",
)
SHAPES = ("tiny", "medium", "dense-wide")
GUARD_SCENARIOS = (
    "cold-first-cell-read",
    "cold-same-one-cell",
    "cold-same-one-percent",
    "warm-same-one-cell",
    "warm-same-one-percent",
    "warm-changed-one-cell",
    "warm-changed-one-percent",
)
FALLBACK_SCENARIOS = ("warm-changed-one-cell", "warm-changed-one-percent")
COMMON_LANES = ("preflight", "pilot", "allocator-r1", "allocator-r2")
FORMAL_LANES = ("r1", "r2")
STATISTICS = ("p50", "mean", "p95", "p99")
ALLOC_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
    "region_peak_live_bytes",
)
TEST_FIXES = {
    "crates/litchi-xlsx/tests/source_backed_cell_values.rs",
    "crates/litchi-xlsx/tests/source_backed_row_visibility.rs",
}
EXPECTED_CHAIN = {
    "second-unit": "initial-unit",
    "third-unit": "second-unit",
    "fourth-unit": "third-unit",
    "fifth-unit": "baseline-tests",
    "after": "fifth-unit",
}
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
RSS_RE = re.compile(
    r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$", re.MULTILINE
)
TEST_RESULT_RE = re.compile(
    r"test result:\s*(?:ok|FAILED)\.\s*"
    r"([\d,]+) passed;\s*([\d,]+) failed;\s*"
    r"([\d,]+) ignored;.*?([\d,]+) filtered out;",
    re.IGNORECASE,
)
PATCH_HEADER_RE = re.compile(r"^diff --git a/(.+) b/(.+)$")


DECISION_SKELETON: dict[str, Any] = {
    "schema": "litchi-0516-verifier-decision-v1",
    "decision": "reject",
    "final_epoch": "baseline-tests",
    "candidate_epoch": "after",
    "baseline_epoch": "before",
    "baseline_test_epoch": "baseline-tests",
    "source": {
        "final_manifest": "baseline-tests/source-manifest.json",
        "candidate_manifest": "after/source-manifest.json",
        "baseline_test_manifest": "baseline-tests/source-manifest.json",
        "production_restored": True,
        "baseline_test_fixes_retained": True,
        "sealed_sha256sums": "SHA256SUMS",
    },
    "tests": {"required_receipts": [
        "after/boundaries-receipt.json",
        "after/clippy-receipt.json",
        "after/fmt-receipt.json",
        "after/owner-check-receipt.json",
        "after/rustdoc-receipt.json",
        "after/xlsx-features-receipt.json",
        "baseline-tests/boundaries-receipt.json",
        "baseline-tests/claims-receipt.json",
        "baseline-tests/clippy-receipt.json",
        "baseline-tests/fmt-receipt.json",
        "baseline-tests/owner-check-receipt.json",
        "baseline-tests/rustdoc-receipt.json",
        "baseline-tests/xlsx-features-receipt.json",
        "fifth-unit/xlsx-features-receipt.json",
        "fifth-unit/xlsx-unit-receipt.json",
    ]},
    "builds": {"required_receipts": [
        "after/allocator-build-receipt.json",
        "after/build-receipt.json",
        "after/fallback-build-allocator-receipt.json",
        "after/fallback-build-receipt.json",
        "after/guard-build-allocator-receipt.json",
        "after/guard-build-receipt.json",
        "before/allocator-build-receipt.json",
        "before/build-receipt.json",
        "before/fallback-build-allocator-receipt.json",
        "before/fallback-build-receipt.json",
        "before/guard-build-allocator-receipt.json",
        "before/guard-build-receipt.json",
    ]},
    "captures": {"required_lanes": [
        "after/allocator-r1-receipt.json",
        "after/allocator-r2-receipt.json",
        "after/fallback-allocator-r1-receipt.json",
        "after/fallback-allocator-r2-receipt.json",
        "after/fallback-pilot-receipt.json",
        "after/fallback-preflight-receipt.json",
        "after/guard-allocator-r1-receipt.json",
        "after/guard-allocator-r2-receipt.json",
        "after/guard-pilot-receipt.json",
        "after/guard-preflight-receipt.json",
        "after/pilot-receipt.json",
        "after/preflight-receipt.json",
        "after/profile-receipt.json",
        "before/allocator-r1-receipt.json",
        "before/allocator-r2-receipt.json",
        "before/fallback-allocator-r1-receipt.json",
        "before/fallback-allocator-r2-receipt.json",
        "before/fallback-pilot-receipt.json",
        "before/fallback-preflight-receipt.json",
        "before/guard-allocator-r1-receipt.json",
        "before/guard-allocator-r2-receipt.json",
        "before/guard-pilot-receipt.json",
        "before/guard-preflight-receipt.json",
        "before/pilot-receipt.json",
        "before/preflight-receipt.json",
        "before/profile-receipt.json",
    ], "formal_abba": False},
    "performance_claim": "none",
    "claim_authorized": False,
}


class VerificationError(ValueError):
    """A missing, malformed, or inconsistent retained evidence item."""


def fail(message: str) -> None:
    raise VerificationError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _reject_constant(value: str) -> None:
    fail(f"non-finite JSON number {value!r}")


def _no_duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        require(key not in result, f"duplicate JSON object key {key!r}")
        result[key] = value
    return result


def load_json(path: Path) -> Any:
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=_no_duplicate_pairs,
            parse_constant=_reject_constant,
        )
    except VerificationError:
        raise
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"cannot load {path}: {error}")


def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(),
            f"cannot hash missing or symlinked file {path}")
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot read {path}: {error}")
    return digest.hexdigest()


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="strict")
    except (OSError, UnicodeError) as error:
        fail(f"cannot read {path}: {error}")


def check_hash(value: Any, context: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
            f"{context} is not a lowercase SHA-256 digest")
    return value


def finite(value: Any, context: str, positive: bool = False) -> float:
    require(isinstance(value, (int, float)) and not isinstance(value, bool),
            f"{context} must be numeric")
    number = float(value)
    require(math.isfinite(number), f"{context} must be finite")
    require(not positive or number > 0, f"{context} must be positive")
    return number


def canonical(value: Any, context: str = "value") -> str:
    try:
        return json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False)
    except (TypeError, ValueError) as error:
        fail(f"{context} is not canonical JSON: {error}")


def safe_relative(value: Any, context: str) -> Path:
    require(isinstance(value, str) and value, f"{context} must be a path")
    path = Path(value)
    require(not path.is_absolute() and ".." not in path.parts
            and path.as_posix() == value, f"{context} is unsafe")
    return path


def bundle_path(root: Path, value: Any, context: str,
                require_file: bool = True) -> Path:
    relative = safe_relative(value, context)
    target = root / relative
    require(target.resolve().is_relative_to(root.resolve()),
            f"{context} escapes the evidence root")
    if require_file:
        require(target.is_file() and not target.is_symlink(),
                f"{context} is missing or not a regular file")
    return target


def repository_file(relative: str, context: str) -> Path:
    path = safe_relative(relative, context)
    target = (REPO / path).resolve()
    require(target.is_relative_to(REPO.resolve()), f"{context} escapes repository")
    require(target.is_file() and not target.is_symlink(),
            f"{context} is missing or not a regular file")
    return target


def parse_time(value: Any, context: str) -> _datetime.datetime:
    require(isinstance(value, str), f"{context} is not an ISO timestamp")
    try:
        result = _datetime.datetime.fromisoformat(value)
    except ValueError as error:
        fail(f"{context} is not an ISO timestamp: {error}")
    require(result.tzinfo is not None, f"{context} has no timezone")
    return result


def _load_module(path: Path, name: str) -> Any:
    require(path.is_file() and not path.is_symlink(), f"helper {path} is unavailable")
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None,
            f"cannot create import specification for {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    try:
        spec.loader.exec_module(module)
    except Exception as error:
        sys.modules.pop(name, None)
        fail(f"cannot load helper {path}: {error}")
    return module


def load_helpers() -> tuple[Any, Any, Any]:
    # Registering the modules is needed by dataclasses in the 0515/0514
    # dynamic imports. The helpers remain standalone; no verifier code is
    # copied into this module.
    metrics = _load_module(HERE / "metrics.py", "litchi_change0516_metrics_for_verify")
    profiles = _load_module(HERE / "profiles.py", "litchi_change0516_profiles_for_verify")
    prior = _load_module(HERE.parent / "change-0514" / "verify.py",
                         "litchi_change0514_verify_for_0516")
    prior.BASE_REVISION = BASE_REVISION
    return metrics, profiles, prior


def load_manifest(path: Path, context: str) -> dict[str, str]:
    value = load_json(path)
    require(isinstance(value, dict) and value, f"{context} is empty")
    result: dict[str, str] = {}
    for name, digest in value.items():
        safe_relative(name, f"{context} path")
        result[name] = check_hash(digest, f"{context}.{name}")
    return result


def manifest_diff(left: dict[str, str], right: dict[str, str]) -> set[str]:
    return {name for name in set(left) | set(right)
            if left.get(name) != right.get(name)}


def current_sources() -> dict[str, str]:
    try:
        names = subprocess.check_output(
            ["git", "ls-files", "--cached", "--others", "--exclude-standard", "-z",
             "crates", "tools/perf-baseline", "Cargo.toml", "Cargo.lock",
             "rust-toolchain.toml", ".cargo"],
            cwd=REPO,
        ).decode("utf-8").split("\0")
    except (OSError, UnicodeError, subprocess.CalledProcessError) as error:
        fail(f"cannot enumerate live source files: {error}")
    result: dict[str, str] = {}
    for name in sorted(item for item in names if item):
        if Path(name).suffix in {".rs", ".toml", ".lock"}:
            result[name] = sha(repository_file(name, f"live source {name}"))
    require(result, "live source manifest is empty")
    return result


def load_decision(root: Path) -> tuple[dict[str, Any], bool]:
    path = root / "decision.json"
    if not path.is_file():
        fail("decision.json is missing; required decision skeleton: "
             + json.dumps(DECISION_SKELETON, sort_keys=True))
    decision = load_json(path)
    require(isinstance(decision, dict), "decision.json must be an object")
    required = {
        "schema", "decision", "final_epoch", "candidate_epoch", "baseline_epoch",
        "baseline_test_epoch", "source", "tests", "builds", "captures",
        "performance_claim", "claim_authorized",
    }
    require(required <= set(decision), "decision.json omits required contract fields")
    require(decision["schema"] == "litchi-0516-verifier-decision-v1",
            "decision schema differs")
    status = decision["decision"]
    require(status in {"retain", "reject"}, "decision.decision must be retain or reject")
    require(all(isinstance(decision[key], str) and decision[key]
                for key in ("final_epoch", "candidate_epoch", "baseline_epoch",
                            "baseline_test_epoch")),
            "decision epoch fields must be nonempty strings")
    require(decision["candidate_epoch"] == "after",
            "candidate_epoch must identify the frozen after candidate")
    source = decision["source"]
    require(isinstance(source, dict), "decision.source must be an object")
    source_required = {"final_manifest", "candidate_manifest", "baseline_test_manifest",
                       "production_restored", "baseline_test_fixes_retained"}
    require(source_required <= set(source), "decision.source is incomplete")
    require(isinstance(source["production_restored"], bool)
            and isinstance(source["baseline_test_fixes_retained"], bool),
            "decision source booleans are malformed")
    require(isinstance(decision["performance_claim"], str)
            and decision["performance_claim"] == "none",
            "performance claims are not permitted by the 0516 contract")
    require(decision["claim_authorized"] is False,
            "claim_authorized must remain false")
    for name in ("tests", "builds", "captures"):
        require(isinstance(decision[name], dict), f"decision.{name} must be an object")
        listed = decision[name].get("required_receipts",
                                    decision[name].get("required_lanes", []))
        require(isinstance(listed, list), f"decision.{name} receipt disclosure is not a list")
        seen: set[str] = set()
        for item in listed:
            path_item = safe_relative(item, f"decision.{name} receipt disclosure")
            require(path_item.as_posix() not in seen,
                    f"decision.{name} repeats {item!r}")
            seen.add(path_item.as_posix())
    formal = decision["captures"].get("formal_abba")
    require(isinstance(formal, bool), "decision.captures.formal_abba must be boolean")
    kept = status == "retain"
    if kept:
        require(decision["final_epoch"] == "after",
                "retain decision must select after as final_epoch")
        require(source["production_restored"] is False,
                "retain decision cannot attest restored production")
        require(formal is True, "retain decision requires formal ABBA")
    else:
        require(source["production_restored"] is True,
                "reject decision must attest restored production")
        require(decision["final_epoch"] in {"before", "baseline-tests"},
                "reject decision final_epoch must be before or baseline-tests")
    return decision, kept


def _git_blob(relative: str) -> bytes | None:
    try:
        return subprocess.check_output(
            ["git", "show", f"{BASE_REVISION}:{relative}"],
            cwd=REPO,
            stderr=subprocess.DEVNULL,
        )
    except (OSError, subprocess.CalledProcessError):
        return None


def _patch_paths(text: str, context: str) -> list[str]:
    paths: list[str] = []
    for line in text.splitlines():
        match = PATCH_HEADER_RE.match(line)
        if match is None:
            continue
        left, right = match.groups()
        require(left == right, f"{context} contains a rename or copy")
        safe_relative(left, f"{context} path")
        paths.append(left)
    require(paths and len(paths) == len(set(paths)),
            f"{context} has no unique diff-section inventory")
    return paths


def _expected_previous(epoch: str) -> str | None:
    if epoch in EXPECTED_CHAIN:
        return EXPECTED_CHAIN[epoch]
    return None


def replay_patch(root: Path, epoch: str, before: dict[str, str],
                 manifest: dict[str, str], record: dict[str, Any]) -> dict[str, Any]:
    directory = root / epoch
    patch = bundle_path(root, f"{epoch}/candidate.patch", f"{epoch} candidate.patch")
    require(record.get("base_revision") == BASE_REVISION,
            f"{epoch} candidate patch base revision differs")
    require(sha(patch) == check_hash(record.get("patch_sha256"),
                                     f"{epoch}.patch_sha256"),
            f"{epoch} candidate.patch hash differs")
    require(record.get("source_manifest_sha256") == sha(directory / "source-manifest.json"),
            f"{epoch} candidate patch manifest hash differs")
    changed = manifest_diff(before, manifest)
    paths = record.get("paths")
    require(isinstance(paths, list) and paths and all(isinstance(item, str) for item in paths)
            and len(set(paths)) == len(paths), f"{epoch}.paths is malformed")
    path_set = {safe_relative(item, f"{epoch} candidate path").as_posix() for item in paths}
    require(path_set == changed, f"{epoch} candidate path inventory differs from manifests")
    hashes = record.get("candidate_file_sha256")
    require(isinstance(hashes, dict) and set(hashes) == changed,
            f"{epoch}.candidate_file_sha256 inventory differs")
    for name, digest in hashes.items():
        require(manifest.get(name) == check_hash(digest, f"{epoch}.{name}"),
                f"{epoch} candidate hash differs for {name}")
    patch_text = read_text(patch)
    require(sorted(_patch_paths(patch_text, f"{epoch}/candidate.patch")) == sorted(changed),
            f"{epoch} patch sections differ from changed inventory")

    expected_previous = _expected_previous(epoch)
    previous = record.get("previous_epoch")
    if expected_previous is not None:
        require(previous == expected_previous,
                f"{epoch}.previous_epoch does not preserve the custody chain")
    else:
        require(previous is None,
                f"{epoch} has an unexpected previous_epoch")
    if previous is not None:
        require(isinstance(previous, str) and previous.replace("-", "").isalnum(),
                f"{epoch}.previous_epoch is malformed")
        previous_manifest = load_manifest(root / previous / "source-manifest.json",
                                         f"{epoch} previous source manifest")
        require(manifest_diff(before, previous_manifest) <= changed,
                f"{epoch} drops a path retained by its previous epoch")

    temporary_name: str | None = None
    with tempfile.TemporaryDirectory(prefix="litchi-0516-patch-replay-") as temporary:
        temporary_name = temporary
        target = Path(temporary)
        for name in sorted(changed):
            blob = _git_blob(name)
            if name in before:
                require(blob is not None, f"{epoch} base blob is missing for {name}")
                require(hashlib.sha256(blob).hexdigest() == before[name],
                        f"{epoch} base blob differs for {name}")
                destination = target / name
                destination.parent.mkdir(parents=True, exist_ok=True)
                destination.write_bytes(blob)
            else:
                require(blob is None, f"{epoch} new path exists at base: {name}")
        checked = subprocess.run(
            ["git", "apply", "--check", "--whitespace=nowarn", str(patch)],
            cwd=target, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        )
        require(checked.returncode == 0,
                f"{epoch} candidate patch does not apply to base blobs")
        applied = subprocess.run(
            ["git", "apply", "--whitespace=nowarn", str(patch)],
            cwd=target, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        )
        require(applied.returncode == 0, f"{epoch} candidate patch replay failed")
        actual = {
            str(path.relative_to(target)): sha(path)
            for path in target.rglob("*")
            if path.is_file()
        }
        require(actual == hashes, f"{epoch} replay file inventory or hashes differ")
    require(temporary_name is not None and not Path(temporary_name).exists(),
            f"{epoch} patch replay temporary directory was retained")
    return {
        "epoch": epoch,
        "previous_epoch": previous,
        "changed": sorted(changed),
        "patch_sha256": sha(patch),
        "source_manifest_sha256": sha(directory / "source-manifest.json"),
        "exact_replay_passed": True,
        "temporary_tree_cleaned": True,
    }


def verify_source(root: Path, decision: dict[str, Any], kept: bool) -> dict[str, Any]:
    plan = load_json(root / "plan.json")
    require(isinstance(plan, dict) and plan.get("base_revision") == BASE_REVISION,
            "plan base_revision differs")
    before_path = root / "before" / "source-manifest.json"
    previous_path = root.parent / "change-0515" / "source-manifest.json"
    before = load_manifest(before_path, "before source manifest")
    if previous_path.is_file():
        require(before == load_manifest(previous_path, "0515 control manifest"),
                "before source manifest does not preserve the 0515 control")
    baseline_path = root / "baseline-tests" / "source-manifest.json"
    baseline = load_manifest(baseline_path, "baseline-tests source manifest")
    candidate_path = bundle_path(root, decision["source"]["candidate_manifest"],
                                 "decision candidate manifest")
    candidate = load_manifest(candidate_path, "after candidate source manifest")
    require(candidate_path == root / "after" / "source-manifest.json",
            "candidate manifest must be after/source-manifest.json")
    final_path = bundle_path(root, decision["source"]["final_manifest"],
                             "decision final manifest")
    final = load_manifest(final_path, "final source manifest")
    require(final_path == root / decision["final_epoch"] / "source-manifest.json",
            "decision final manifest does not match final_epoch")
    require(decision["baseline_epoch"] == "before"
            and decision["baseline_test_epoch"] == "baseline-tests",
            "decision baseline epoch names differ")
    require(manifest_diff(before, baseline) == TEST_FIXES,
            "baseline-tests must change exactly the two focused test files")
    require(decision["source"]["baseline_test_fixes_retained"] is True,
            "final decision must state whether the independent test fixes are retained")
    for name in TEST_FIXES:
        require(baseline[name] == candidate.get(name),
                f"candidate does not carry baseline test fix {name}")
        if name in final:
            require(final[name] == baseline[name], f"final test fix differs for {name}")
    production_changed = manifest_diff(baseline, candidate)
    require(production_changed, "after candidate has no production change")
    require(all(name.startswith("crates/litchi-xlsx/") and name.endswith(".rs")
                for name in production_changed),
            "candidate production changes leave the XLSX Rust source scope")
    require(not production_changed & TEST_FIXES,
            "baseline test repair is mixed into the after production delta")
    require("crates/litchi-xlsx/src/workbook/edit/semantic/transaction.rs"
            in production_changed, "candidate omits the XLSX transaction production change")
    all_changed = manifest_diff(before, candidate)
    require(all_changed <= set(candidate), "candidate changed inventory is incomplete")
    require(all(name.startswith("crates/litchi-xlsx/") for name in all_changed),
            "candidate source change set leaves the XLSX crate")
    current = current_sources()
    require(current == final,
            "live source tree does not equal the decision's final source manifest")
    if kept:
        require(final == candidate, "retained final source is not the after candidate")
    else:
        require(final == before or final == baseline,
                "rejected final source is neither before nor baseline-tests")
        restoration_path = root / "restoration.json"
        restoration = load_json(restoration_path)
        require(isinstance(restoration, dict), "restoration.json is not an object")
        require(restoration.get("decision") == "reject"
                and restoration.get("production_restored_from") == BASE_REVISION,
                "restoration record does not bind the rejected candidate to the base")
        require(restoration.get("candidate_manifest_sha256") == sha(candidate_path)
                and restoration.get("final_manifest_sha256") == sha(final_path),
                "restoration source manifest hashes differ")
        restored = restoration.get("restored_paths")
        removed = restoration.get("removed_candidate_paths")
        retained = restoration.get("retained_test_fixes")
        require(isinstance(restored, list) and isinstance(removed, list)
                and isinstance(retained, list),
                "restoration path inventories are malformed")
        restored_set = {safe_relative(item, "restored path").as_posix() for item in restored}
        removed_set = {safe_relative(item, "removed candidate path").as_posix() for item in removed}
        retained_set = {safe_relative(item, "retained test fix").as_posix() for item in retained}
        require(restored_set.isdisjoint(removed_set)
                and restored_set | removed_set == production_changed,
                "restoration production path inventory differs")
        require(retained_set == TEST_FIXES,
                "restoration does not retain exactly the two baseline test fixes")
        require(restoration.get("exact_final_source_match") is True,
                "restoration does not attest exact final source match")
        patch_info = restoration.get("candidate_patch_retained")
        require(isinstance(patch_info, dict)
                and patch_info.get("path") == "after/candidate.patch"
                and patch_info.get("sha256") == sha(root / "after/candidate.patch"),
                "restoration does not retain the after candidate patch")

    patch_results = []
    for record_path in sorted(root.glob("*/candidate-patch.json")):
        epoch = record_path.parent.name
        record = load_json(record_path)
        require(isinstance(record, dict), f"{epoch}/candidate-patch.json is not an object")
        manifest_path = record_path.parent / "source-manifest.json"
        manifest = load_manifest(manifest_path, f"{epoch} source manifest")
        if epoch == "baseline-tests":
            require(record.get("retained_test_fix") is True,
                    "baseline-tests patch does not attest retained test fix")
            require(set(record.get("paths", [])) == TEST_FIXES,
                    "baseline-tests patch path inventory differs")
        if epoch != "baseline-tests":
            require(not record.get("retained_test_fix", False),
                    f"{epoch} unexpectedly labels a production snapshot as test-only")
        patch_results.append(replay_patch(root, epoch, before, manifest, record))
    require(patch_results, "no candidate patch records are retained")

    return {
        "base_revision": BASE_REVISION,
        "before_manifest_sha256": sha(before_path),
        "baseline_test_manifest_sha256": sha(baseline_path),
        "candidate_manifest_sha256": sha(candidate_path),
        "final_manifest_sha256": sha(final_path),
        "live_source_matches_final": True,
        "changed_from_baseline_tests": sorted(production_changed),
        "patch_replays": patch_results,
    }



def _environment_check(environment: Any, context: str) -> None:
    if environment is None:
        return
    require(isinstance(environment, dict), f"{context}.environment is not an object")
    # These checks only constrain values the producer actually recorded.
    # Missing inherited variables are never fabricated by this verifier.
    expected = {
        "CARGO_BUILD_JOBS": "2",
        "CARGO_INCREMENTAL": "0",
        "CARGO_PROFILE_RELEASE_DEBUG": "0",
        "CARGO_PROFILE_DEV_DEBUG": "0",
        "CARGO_PROFILE_TEST_DEBUG": "0",
    }
    for key, value in expected.items():
        if key in environment:
            require(environment[key] == value, f"{context}.environment.{key} differs")
    if "RUSTDOCFLAGS" in environment:
        require(environment["RUSTDOCFLAGS"] == "-D warnings",
                f"{context}.environment.RUSTDOCFLAGS differs")
    for key in ("CARGO_TARGET_DIR", "TMPDIR"):
        if key in environment:
            require(isinstance(environment[key], str) and environment[key],
                    f"{context}.environment.{key} is malformed")


def _receipt_common(root: Path, path: Path, receipt: dict[str, Any],
                    expected_source: str, context: str, expected_exit: int = 0,
                    expected_binary: str | None = None,
                    expected_probe: str | None = None,
                    required_artifacts: set[str] | None = None,
                    post_cleanup: bool = False) -> dict[str, Any]:
    require(isinstance(receipt, dict), f"{context} receipt is not an object")
    require(receipt.get("exit_code") == expected_exit,
            f"{context} exit code differs")
    require(receipt.get("source_unchanged") is True,
            f"{context} source changed during command")
    require(receipt.get("source_manifest_sha256") == expected_source,
            f"{context} source manifest differs")
    check_hash(receipt.get("source_manifest_sha256"), f"{context}.source_manifest_sha256")
    parse_time(receipt.get("started_utc"), f"{context}.started_utc")
    elapsed = finite(receipt.get("elapsed_seconds"), f"{context}.elapsed_seconds", True)
    command = receipt.get("command")
    require(isinstance(command, list) and command and all(isinstance(item, str) for item in command),
            f"{context}.command is not an argv list")
    _environment_check(receipt.get("environment"), context)
    active = receipt.get("active_source_roles")
    if active is not None:
        require(isinstance(active, list) and active and all(isinstance(item, str) for item in active),
                f"{context}.active_source_roles is malformed")
        require(path.parent.name in active,
                f"{context}.active_source_roles omits its source epoch")
    if expected_probe is not None:
        require(receipt.get("probe") == expected_probe,
                f"{context}.probe identity differs")
        probe_manifest = bundle_path(root, f"{expected_probe}-source-manifest.json",
                                     f"{context} probe source manifest")
        require(receipt.get("guard_manifest_sha256") == sha(probe_manifest),
                f"{context} probe source manifest hash differs")
    artifacts = receipt.get("artifacts")
    if required_artifacts is not None:
        # An empty set is an explicit statement that this producer records no
        # artifact map (quality receipts use log_sha256 instead).
        if required_artifacts:
            require(isinstance(artifacts, dict), f"{context}.artifacts is missing")
            require(required_artifacts <= set(artifacts),
                    f"{context}.artifacts omits required files")
        elif artifacts is not None:
            require(isinstance(artifacts, dict), f"{context}.artifacts is not an object")
    elif artifacts is not None:
        require(isinstance(artifacts, dict), f"{context}.artifacts is not an object")
    checked_artifacts: dict[str, str] = {}
    if isinstance(artifacts, dict):
        for name, digest in artifacts.items():
            relative = safe_relative(name, f"{context} artifact path")
            target = path.parent / relative
            require(target.parent.resolve() == path.parent.resolve(),
                    f"{context} artifact path is not a stage basename")
            require(target.is_file() and not target.is_symlink(),
                    f"{context} artifact {name} is missing or symlinked")
            checked = check_hash(digest, f"{context}.artifacts.{name}")
            require(sha(target) == checked, f"{context} artifact {name} hash differs")
            checked_artifacts[name] = checked
    log_name = path.name.removesuffix("-receipt.json") + ".log"
    log_path = path.with_name(log_name)
    require(log_path.is_file() and not log_path.is_symlink(),
            f"{context} log is missing")
    log_digest = sha(log_path)
    if "log_sha256" in receipt:
        require(receipt["log_sha256"] == log_digest,
                f"{context}.log_sha256 differs")
    if isinstance(artifacts, dict) and log_name in artifacts:
        require(artifacts[log_name] == log_digest,
                f"{context} log artifact hash differs")
    binary_digest = None
    if expected_binary is not None:
        binary_digest = check_hash(receipt.get("binary_sha256"),
                                   f"{context}.binary_sha256")
        binary_path = SCRATCH / path.parent.name / expected_binary
        if binary_path.exists():
            require(binary_path.is_file() and not binary_path.is_symlink(),
                    f"{context} retained binary is not regular")
            require(sha(binary_path) == binary_digest,
                    f"{context} retained binary hash differs")
        elif SCRATCH.exists() and not post_cleanup:
            fail(f"{context} binary is absent while the owned scratch root remains")
    return {
        "path": path.relative_to(root).as_posix(),
        "command": command,
        "started_utc": receipt["started_utc"],
        "elapsed_seconds": elapsed,
        "binary_sha256": binary_digest,
        "log_sha256": log_digest,
        "artifacts": checked_artifacts,
    }


def _option(command: list[str], flag: str, context: str) -> str:
    positions = [index for index, value in enumerate(command) if value == flag]
    require(len(positions) == 1 and positions[0] + 1 < len(command),
            f"{context} command must contain one {flag}")
    return command[positions[0] + 1]


def _equals_option(command: list[str], flag: str, context: str) -> str:
    prefix = flag + "="
    values = [item[len(prefix):] for item in command if item.startswith(prefix)]
    require(flag not in command and len(values) == 1 and values[0],
            f"{context} command must contain one {prefix} value")
    return values[0]


def _pinned(command: list[str], context: str) -> None:
    require(command[:5] == ["/usr/bin/time", "-v", "taskset", "-c", "2"],
            f"{context} command is not pinned to CPU 2 with GNU time")


def validate_main_command(command: list[str], stem: str, samples: int,
                          warmups: int, allocator: bool, profile: bool) -> None:
    context = stem
    _pinned(command, context)
    binary = "litchi-perf-baseline-alloc" if allocator else "litchi-perf-baseline"
    require(sum(Path(item).name == binary for item in command) == 1,
            f"{context} command does not select one expected binary")
    require(_option(command, "--samples", context) == str(samples),
            f"{context} sample count differs")
    require(_option(command, "--warmup", context) == str(warmups),
            f"{context} warmup count differs")
    require(Path(_option(command, "--json", context)).name == f"{stem}-report.json",
            f"{context} report path differs")
    require(Path(_option(command, "--corpus-manifest", context)).name
            == f"{stem}-catalog.json", f"{context} catalog path differs")
    if profile:
        require("valgrind" in command and "--tool=callgrind" in command,
                f"{context} profile is not Callgrind")
        require("--collect-atstart=no" in command
                and "--toggle-collect=*litchi_perf_baseline::xlsx_commit_save_operation" in command,
                f"{context} profile boundary differs")
        require(Path(_equals_option(command, "--callgrind-out-file", context)).name
                == "profile.out", f"{context} raw profile path differs")
        require(_option(command, "--case", context) == "xlsx_one_percent_commit_save",
                f"{context} profile case differs")
        require(_option(command, "--xlsx-shape", context) == "dense-wide",
                f"{context} profile shape differs")
    else:
        require(_option(command, "--case", context) == ",".join(CASES),
                f"{context} case selection differs")
        require(_option(command, "--xlsx-shape", context) == ",".join(SHAPES),
                f"{context} shape selection differs")


def validate_probe_command(command: list[str], stem: str, probe: str,
                           samples: int, warmups: int, allocator: bool,
                           build: bool = False) -> None:
    if build:
        require(command[:3] == ["/usr/bin/time", "-v", "cargo"],
                f"{stem} build command does not use GNU time")
        require("--offline" in command and "--locked" in command,
                f"{stem} probe build is not locked/offline")
        manifest = _option(command, "--manifest-path", stem)
        require(Path(manifest).as_posix().endswith(f"{probe}-probe/Cargo.toml"),
                f"{stem} probe manifest differs")
        require(_option(command, "--bin", stem) == "litchi-xlsx-commit-guard",
                f"{stem} probe binary differs")
        if allocator:
            require(command[-2:] == ["--features", "allocator-metrics"],
                    f"{stem} allocator feature differs")
        else:
            require("--features" not in command,
                    f"{stem} normal probe build carries allocator feature")
        return
    _pinned(command, stem)
    binary = f"{probe}-alloc" if allocator else probe
    require(sum(Path(item).name == binary for item in command) == 1,
            f"{stem} command binary differs")
    require(_option(command, "--samples", stem) == str(samples)
            and _option(command, "--warmup", stem) == str(warmups),
            f"{stem} sample protocol differs")
    require(_option(command, "--shape", stem) == ",".join(SHAPES),
            f"{stem} shape selection differs")
    require(Path(_option(command, "--json", stem)).name == f"{stem}-report.json",
            f"{stem} report path differs")


QUALITY_COMMANDS = {
    "xlsx-unit": ["cargo", "test", "--locked", "-p", "litchi-xlsx", "--lib"],
    "xlsx-features": ["cargo", "test", "--locked", "-p", "litchi-xlsx", "--all-features"],
    "fmt": ["cargo", "fmt", "--all", "--", "--check"],
    "clippy": ["cargo", "clippy", "--locked", "-p", "litchi-xlsx", "--all-features",
               "--lib", "--", "-D", "warnings"],
    "rustdoc": ["cargo", "doc", "--locked", "-p", "litchi-xlsx", "--all-features", "--no-deps"],
    "owner-check": ["cargo", "check", "--locked", "-p", "litchi-xlsx", "--all-features"],
    "boundaries": ["python3", "-B", "tools/check_crate_boundaries.py"],
    "claims": ["python3", "-B", "tools/check_perf_claims.py", "--registry",
               "docs/performance/claim-registry-v1.json", "--repo-root", ".",
               "--evidence-root", ".", "--mode", "strict"],
}


def test_counts(log: str, context: str) -> dict[str, int]:
    matches = TEST_RESULT_RE.findall(log)
    require(matches, f"{context} log has no Rust test result")
    totals = [sum(int(match[index].replace(",", "")) for match in matches)
              for index in range(4)]
    return {"passed": totals[0], "failed": totals[1], "ignored": totals[2],
            "filtered": totals[3], "groups": len(matches)}


def validate_quality(root: Path, epoch: str, name: str, source_digest: str,
                     post_cleanup: bool = False) -> dict[str, Any]:
    require(name in QUALITY_COMMANDS, f"unknown quality receipt {name}")
    path = root / epoch / f"{name}-receipt.json"
    receipt = load_json(path)
    result = _receipt_common(root, path, receipt, source_digest,
                             f"{epoch}/{name}", required_artifacts=None,
                             post_cleanup=post_cleanup)
    require(receipt.get("command") == QUALITY_COMMANDS[name],
            f"{epoch}/{name} command differs from check.py")
    if name in {"xlsx-unit", "xlsx-features"}:
        result["test_counts"] = test_counts(
            read_text(path.with_name(f"{name}.log")), f"{epoch}/{name}"
        )
    return result


def verify_quality_groups(root: Path, decision: dict[str, Any], kept: bool,
                          candidate_digest: str, final_digest: str,
                          baseline_digest: str, post_cleanup: bool) -> dict[str, Any]:
    candidate_names = ("xlsx-features", "fmt", "clippy", "rustdoc", "owner-check", "boundaries")
    candidate = {name: validate_quality(root, "after", name, candidate_digest, post_cleanup)
                 for name in candidate_names}
    require(candidate["xlsx-features"]["test_counts"] == {
        "passed": 1284, "failed": 0, "ignored": 0, "filtered": 0,
        "groups": candidate["xlsx-features"]["test_counts"]["groups"],
    }, "after all-features suite is not the complete 1284-test result")
    if kept:
        candidate["claims"] = validate_quality(root, "after", "claims", candidate_digest,
                                                post_cleanup)

    baseline_names = ("xlsx-features", "fmt", "clippy", "rustdoc", "owner-check", "boundaries")
    baseline = {name: validate_quality(root, "baseline-tests", name, baseline_digest,
                                       post_cleanup) for name in baseline_names}
    require(baseline["xlsx-features"]["test_counts"] == {
        "passed": 1263, "failed": 0, "ignored": 0, "filtered": 0,
        "groups": baseline["xlsx-features"]["test_counts"]["groups"],
    }, "baseline-tests all-features suite is not the complete 1263-test result")

    final: dict[str, Any] = {}
    final_names = ("xlsx-features", "fmt", "clippy", "rustdoc", "owner-check", "boundaries", "claims")
    for name in final_names:
        final[name] = validate_quality(root, decision["final_epoch"], name,
                                        final_digest, post_cleanup)
    if decision["final_epoch"] == "baseline-tests":
        require(final["xlsx-features"]["test_counts"] == baseline["xlsx-features"]["test_counts"],
                "final baseline-tests suite does not retain the 1263-test result")

    historical = {
        "xlsx-unit": validate_quality(
            root, "fifth-unit", "xlsx-unit",
            sha(root / "fifth-unit" / "source-manifest.json"), post_cleanup
        ),
        "xlsx-features": validate_quality(
            root, "fifth-unit", "xlsx-features",
            sha(root / "fifth-unit" / "source-manifest.json"), post_cleanup
        ),
    }
    require(historical["xlsx-unit"]["test_counts"]["passed"] == 970
            and historical["xlsx-unit"]["test_counts"]["failed"] == 0,
            "fifth-unit historical unit result is not 970 passed")
    require(historical["xlsx-features"]["test_counts"]["passed"] == 1284
            and historical["xlsx-features"]["test_counts"]["failed"] == 0,
            "fifth-unit historical all-features result is not 1284 passed")
    return {"candidate": candidate, "baseline_tests": baseline,
            "final": final, "historical_fifth": historical}



def _check_log_and_artifacts(root: Path, rel: str, receipt: dict[str, Any],
                             source_digest: str, expected_exit: int) -> None:
    path = bundle_path(root, rel, f"retained receipt {rel}")
    _receipt_common(root, path, receipt, source_digest, rel, expected_exit=expected_exit)


def verify_expected_failures(root: Path) -> dict[str, Any]:
    expected_path = root / "expected-failures.json"
    expected = load_json(expected_path)
    require(isinstance(expected, dict) and expected, "expected-failures.json is empty")
    for relative, detail in expected.items():
        path = bundle_path(root, relative, "expected failure receipt")
        require(path.name.endswith("-receipt.json") and isinstance(detail, dict),
                f"{relative} expected failure entry is malformed")
        expected_exit = detail.get("exit_code")
        require(isinstance(expected_exit, int) and expected_exit != 0,
                f"{relative} expected failure exit code is malformed")
        source_path = path.parent / "source-manifest.json"
        source_digest = sha(source_path)
        receipt = load_json(path)
        _check_log_and_artifacts(root, relative, receipt, source_digest, expected_exit)
        require(isinstance(detail.get("reason"), str) and detail["reason"].strip(),
                f"{relative} expected failure reason is missing")
    retained_nonzero: set[str] = set()
    for path in sorted(root.glob("*/*-receipt.json")):
        receipt = load_json(path)
        if receipt.get("exit_code") != 0:
            retained_nonzero.add(path.relative_to(root).as_posix())
    require(retained_nonzero == set(expected),
            "a retained nonzero receipt is missing from or absent in expected-failures.json")
    return {"allowlist": expected, "retained_nonzero_receipts": sorted(retained_nonzero)}


def verify_baseline_failures(root: Path) -> dict[str, Any]:
    record = load_json(root / "baseline-failures.json")
    require(isinstance(record, dict), "baseline-failures.json is not an object")
    tests = record.get("tests")
    require(tests == [
        "managed_scalar_exact_noop_publishes_without_detaching_source",
        "managed_multi_sheet_exact_noop_publishes_without_detaching_sources",
    ], "baseline failure test inventory differs")
    reference = record.get("reference_receipt")
    candidate = record.get("candidate_receipt")
    require(isinstance(reference, str) and isinstance(candidate, str),
            "baseline failure receipt paths are missing")
    for relative in (reference, candidate):
        path = bundle_path(root, relative, "baseline failure receipt")
        receipt = load_json(path)
        require(receipt.get("exit_code") == 101, f"{relative} does not retain exit 101")
        log = read_text(path.with_name(path.name.replace("-receipt.json", ".log")))
        require(any(test in log for test in tests),
                f"{relative} log does not retain the named typed failures")
    require(record.get("same_typed_errors") is True,
            "baseline failure record does not attest identical typed errors")
    fixed = record.get("fixed_suite")
    require(isinstance(fixed, dict) and fixed.get("passed") == 1263
            and fixed.get("failed") == fixed.get("ignored") == fixed.get("filtered") == 0,
            "baseline failure fixed suite is not 1263 passed with no exclusions")
    fixed_rel = record.get("fixed_suite_receipt")
    require(fixed_rel == "baseline-tests/xlsx-features-receipt.json",
            "baseline failure fixed suite receipt differs")
    fixed_path = bundle_path(root, fixed_rel, "baseline fixed suite receipt")
    fixed_receipt = load_json(fixed_path)
    require(fixed_receipt.get("exit_code") == 0, "baseline fixed suite did not pass")
    return {"tests": tests, "reference": reference, "candidate": candidate,
            "fixed_suite": fixed, "same_typed_errors": True}


def verify_correctness_admission(root: Path, candidate_digest: str,
                                 candidate_patch_digest: str) -> dict[str, Any]:
    path = root / "correctness-admission.json"
    record = load_json(path)
    require(isinstance(record, dict), "correctness-admission.json is not an object")
    require(record.get("source_manifest_sha256") == candidate_digest,
            "correctness admission source manifest differs")
    require(record.get("candidate_patch_sha256") == candidate_patch_digest,
            "correctness admission patch differs")
    counts = record.get("test_counts")
    require(isinstance(counts, dict) and counts.get("passed") == 1284
            and counts.get("failed") == counts.get("ignored") == counts.get("filtered") == 0,
            "correctness admission does not retain the complete 1284 test result")
    gates = record.get("gates")
    require(isinstance(gates, dict), "correctness admission gates are missing")
    expected = {f"after/{name}-receipt.json" for name in
                ("xlsx-features", "clippy", "fmt", "rustdoc", "owner-check", "boundaries")}
    require(set(gates) == expected, "correctness admission gate inventory differs")
    for relative, digest in gates.items():
        check_hash(digest, f"correctness admission {relative}")
        require(sha(bundle_path(root, relative, "correctness admission gate")) == digest,
                f"correctness admission gate hash differs for {relative}")
    reviews = record.get("reviews")
    require(isinstance(reviews, dict) and set(reviews) == {
        "final-implementation-review.md", "final-resource-review.md",
    }, "correctness admission review inventory differs")
    for name, digest in reviews.items():
        check_hash(digest, f"correctness admission review {name}")
        require(sha(bundle_path(root, name, "correctness admission review")) == digest,
                f"correctness admission review hash differs for {name}")
    require(record.get("ready_for_measurement") is True
            and record.get("performance_admitted") is False,
            "correctness admission has an unsafe performance status")
    return {"source_manifest_sha256": candidate_digest, "gates": gates,
            "reviews": reviews, "ready_for_measurement": True,
            "performance_admitted": False}


def verify_adr(root: Path) -> dict[str, Any]:
    record = load_json(root / "adr-manifest.json")
    require(isinstance(record, dict) and record.get("revision") == BASE_REVISION,
            "ADR manifest revision differs from base")
    files = record.get("files")
    require(isinstance(files, dict) and len(files) == 30,
            "ADR manifest must retain 30 files")
    for name, digest in files.items():
        check_hash(digest, f"ADR {name}")
        require(sha(repository_file(name, f"ADR {name}")) == digest,
                f"ADR file changed: {name}")
    return {"revision": BASE_REVISION, "files": len(files)}


def _probe_manifest(root: Path, probe: str) -> tuple[dict[str, str], str]:
    path = bundle_path(root, f"{probe}-source-manifest.json", f"{probe} source manifest")
    manifest = load_manifest(path, f"{probe} source manifest")
    for name, digest in manifest.items():
        require(sha(repository_file(name, f"{probe} source {name}")) == digest,
                f"{probe} source changed: {name}")
    return manifest, sha(path)


def verify_builds(root: Path, stage: str, source_digest: str,
                  post_cleanup: bool) -> dict[str, Any]:
    result: dict[str, Any] = {}
    specs = {
        "main": ("build", "litchi-perf-baseline", False, None),
        "main-allocator": ("allocator-build", "litchi-perf-baseline-alloc", True, None),
        "guard": ("guard-build", "guard", False, "guard"),
        "guard-allocator": ("guard-build-allocator", "guard-alloc", True, "guard"),
        "fallback": ("fallback-build", "fallback", False, "fallback"),
        "fallback-allocator": ("fallback-build-allocator", "fallback-alloc", True, "fallback"),
    }
    for role, (name, binary, allocator, probe) in specs.items():
        path = root / stage / f"{name}-receipt.json"
        receipt = load_json(path)
        context = f"{stage}/{name}"
        probe_manifest_digest = None
        if probe is None:
            result[role] = _receipt_common(root, path, receipt, source_digest, context,
                                           expected_binary=binary, post_cleanup=post_cleanup)
            expected_command = [
                "cargo", "build", "--release", "--locked", "--manifest-path",
                "tools/perf-baseline/Cargo.toml", "--bin", binary,
            ]
            if allocator:
                expected_command += ["--features", "allocator-metrics"]
            require(receipt.get("command") == expected_command,
                    f"{context} command differs from run.py")
        else:
            _, probe_manifest_digest = _probe_manifest(root, probe)
            result[role] = _receipt_common(
                root, path, receipt, source_digest, context,
                expected_binary=binary, expected_probe=probe,
                required_artifacts={f"{name}.log"}, post_cleanup=post_cleanup,
            )
            validate_probe_command(receipt["command"], name, probe, 0, 0, allocator, build=True)
        result[role]["probe_manifest_sha256"] = probe_manifest_digest
    return result


def validate_main_identity(report: dict[str, Any], build: dict[str, Any],
                           samples: int, warmups: int, cases: list[str],
                           shapes: list[str], context: str, allocator: bool) -> None:
    require(report.get("schema_version") == 1, f"{context} schema version differs")
    tool = report.get("tool")
    require(isinstance(tool, dict), f"{context}.tool is missing")
    expected_binary = "litchi-perf-baseline-alloc" if allocator else "litchi-perf-baseline"
    expected_instrumentation = "system_allocator_operation_scoped" if allocator else "none"
    for key, value in {
        "name": "litchi-perf-baseline", "version": "0.1.0", "binary": expected_binary,
        "profile": "release", "target_os": "linux", "target_arch": "x86_64",
        "instrumentation": expected_instrumentation,
    }.items():
        require(tool.get(key) == value, f"{context}.tool.{key} differs")
    if allocator:
        require(tool.get("allocator_counter_revision") == "serialized_region_peak_v3",
                f"{context} allocator counter revision differs")
    else:
        require("allocator_counter_revision" not in tool,
                f"{context} normal report exposes allocator counters")
    identity = report.get("binary_identity")
    require(isinstance(identity, dict), f"{context}.binary_identity is missing")
    require(identity.get("binary_sha256") == build["binary_sha256"],
            f"{context} binary identity differs from build")
    check_hash(identity.get("binary_sha256"), f"{context}.binary_identity.binary_sha256")
    require(Path(str(identity.get("path"))).name == expected_binary
            and identity.get("executable") is True
            and identity.get("profile") == "release",
            f"{context}.binary_identity is malformed")
    require(isinstance(identity.get("binary_bytes"), int) and identity["binary_bytes"] > 0,
            f"{context}.binary_identity.binary_bytes is invalid")
    environment = report.get("environment")
    require(isinstance(environment, dict), f"{context}.environment is missing")
    require(environment.get("git_revision") == BASE_REVISION,
            f"{context} report git revision differs")
    require(environment.get("cpu_affinity") == "2", f"{context} CPU affinity differs")
    allocator_name = (
        "CountingSystemAllocator(std::alloc::System)" if allocator
        else "Rust system allocator"
    )
    require(environment.get("allocator") == allocator_name,
            f"{context} allocator identity differs")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{context}.configuration is missing")
    require(configuration.get("samples_per_case") == samples
            and configuration.get("warmup_iterations_per_case") == warmups,
            f"{context} sample protocol differs")
    require(configuration.get("cases") == cases
            and configuration.get("xlsx_shapes") == shapes,
            f"{context} case/shape selection differs")
    require(configuration.get("filesystem_process_isolated") is True,
            f"{context} process isolation is not recorded")
    # In-process XLSX captures do not carry a per-row fresh-child guarantee.
    # The verifier intentionally does not require or synthesize that field.


def _main_rows(report: dict[str, Any], prior: Any, metrics: Any,
               build: dict[str, Any], samples: int, warmups: int,
               allocator: bool, profile: bool, context: str,
               catalog: dict[str, Any] | None = None
               ) -> dict[tuple[str, str], dict[str, Any]]:
    cases = ["xlsx_one_percent_commit_save"] if profile else list(CASES)
    shapes = ["dense-wide"] if profile else list(SHAPES)
    validate_main_identity(report, build, samples, warmups, cases, shapes, context, allocator)
    if catalog is not None:
        try:
            prior.validate_binding(report, catalog)
        except Exception as error:
            fail(f"{context} report/catalog binding failed: {error}")
    expected = [(shape, case) for shape in shapes for case in cases]
    rows = report.get("results")
    require(isinstance(rows, list) and len(rows) == len(expected),
            f"{context} has the wrong complete row count")
    parsed: dict[tuple[str, str], dict[str, Any]] = {}
    for index, ((shape, case), row) in enumerate(zip(expected, rows)):
        row_context = f"{context}.results[{index}]"
        try:
            checked = prior.verify_row(row, case, shape, samples, row_context, allocator)
            summary = metrics._main_row_summary(
                row, (shape, case), samples, allocator, row_context
            )
        except Exception as error:
            fail(f"{row_context} failed canonical row validation: {error}")
        require((shape, case) not in parsed, f"{context} duplicate row key")
        parsed[shape, case] = {"row": row, "checked": checked, "summary": summary}
    require(set(parsed) == set(expected), f"{context} row keys are incomplete")
    return parsed



FALLBACK_TIMER_SCOPE = (
    "Edit::commit only; fresh Workbook open, complete public Store warming, edit preparation, "
    "output serialization, source-marker/readback/cell oracles, and Commit/View drop are outside the clock"
)


def _probe_rows(report: dict[str, Any], family: str, prior: Any, metrics: Any,
                samples: int, warmups: int, allocator: bool, context: str
                ) -> dict[tuple[str, str], dict[str, Any]]:
    fallback = family == "fallback"
    expected_scenarios = FALLBACK_SCENARIOS if fallback else GUARD_SCENARIOS
    expected_probe = (
        "litchi-xlsx-public-commit-fallback-guard-v1" if fallback
        else "litchi-xlsx-public-commit-guard-v1"
    )
    expected_timer = FALLBACK_TIMER_SCOPE if fallback else prior.GUARD_TIMER_SCOPE
    require(isinstance(report, dict), f"{context} report is not an object")
    expected_keys = {
        "schema_version", "probe", "timer_scope", "allocation", "samples",
        "warmups", "shapes",
    }
    if fallback:
        expected_keys.add("corpus_variant")
    require(set(report) == expected_keys, f"{context} report schema differs")
    require(report.get("schema_version") == 1 and report.get("probe") == expected_probe,
            f"{context} probe identity differs")
    require(report.get("timer_scope") == expected_timer,
            f"{context} timer scope differs")
    require(report.get("samples") == samples and report.get("warmups") == warmups,
            f"{context} sample protocol differs")
    if fallback:
        require(report.get("corpus_variant") == {
            "name": "numeric-two-sheet-public-default-descent-x14ac-v1",
            "source_generator": "public Workbook/Edit two-sheet numeric generator",
            "default_height": 15.0,
            "default_descent": 0.2,
            "metadata_application": (
                "public WorksheetEdit::defaults().height(15).descent(0.2) "
                "before initial commit"
            ),
        },
                f"{context} fallback corpus variant differs")
    allocation = report.get("allocation")
    require(
        isinstance(allocation, dict)
        and set(allocation) == {"binary", "allocator", "instrumentation", "counter_revision"},
        f"{context} allocation identity schema differs",
    )
    binary = "litchi-xlsx-commit-guard-alloc" if allocator else "litchi-xlsx-commit-guard"
    allocator_name = (
        "CountingSystemAllocator(std::alloc::System)" if allocator
        else "Rust system allocator"
    )
    require(
        allocation.get("binary") == binary
        and allocation.get("allocator") == allocator_name
        and allocation.get("instrumentation") ==
        ("system_allocator_operation_scoped" if allocator else "none")
        and allocation.get("counter_revision") ==
        ("serialized_region_peak_v3" if allocator else None),
        f"{context} allocation identity differs",
    )
    shapes = report.get("shapes")
    require(
        isinstance(shapes, list) and len(shapes) == len(SHAPES)
        and all(isinstance(item, dict) for item in shapes)
        and [item.get("shape") for item in shapes] == list(SHAPES),
        f"{context} shape order differs",
    )
    parsed: dict[tuple[str, str], dict[str, Any]] = {}
    for index, shape_report in enumerate(shapes):
        shape_context = f"{context}.shapes[{index}]"
        shape_name = shape_report.get("shape")
        require(shape_name in SHAPES, f"{shape_context}.shape is unknown")
        shape_keys = {
            "shape", "sheet_count", "rows_per_sheet", "columns_per_sheet", "cells",
            "one_percent_update_count", "corpus_bytes", "corpus_sha256", "scenarios",
        }
        if fallback:
            shape_keys |= {"corpus_variant", "source_markers", "descent_readback"}
            require(
                shape_report.get("corpus_variant") ==
                "numeric-two-sheet-public-default-descent-x14ac-v1",
                f"{shape_context} fallback corpus variant differs",
            )
            require(
                shape_report.get("source_markers") == dict.fromkeys(
                    [
                        "worksheet_parts_checked", "x14ac_namespace_bindings",
                        "mce_namespace_bindings", "mce_ignorable_attributes",
                        "dy_descent_attributes",
                    ],
                    2,
                ),
                f"{shape_context} fallback source markers differ",
            )
            require(
                shape_report.get("descent_readback") ==
                {"expected": 0.2, "sheets_checked": 2, "all_sheets_match": True},
                f"{shape_context} fallback descent readback differs",
            )
        require(set(shape_report) == shape_keys, f"{shape_context} schema differs")
        side = prior.GUARD_SHAPE_SIDES[shape_name]
        require(
            shape_report.get("sheet_count") == 2
            and shape_report.get("rows_per_sheet") == side
            and shape_report.get("columns_per_sheet") == side
            and shape_report.get("cells") == 2 * side * side
            and shape_report.get("one_percent_update_count") == (2 * side * side + 99) // 100,
            f"{shape_context} dimensions differ",
        )
        check_hash(shape_report.get("corpus_sha256"), f"{shape_context}.corpus_sha256")
        scenarios = shape_report.get("scenarios")
        require(
            isinstance(scenarios, list)
            and len(scenarios) == len(expected_scenarios)
            and [row.get("scenario") for row in scenarios] == list(expected_scenarios),
            f"{shape_context} scenario inventory differs",
        )
        for scenario_index, original in enumerate(scenarios):
            row_context = f"{shape_context}.scenarios[{scenario_index}]"
            canonical_row = copy.deepcopy(original)
            if fallback:
                oracle = canonical_row.get("oracle")
                require(isinstance(oracle, dict), f"{row_context}.oracle is missing")
                for field in (
                    "output_serialized", "untouched_data_equal",
                    "defaults_descent_readback", "source_markers",
                ):
                    require(oracle.pop(field, None) is True,
                            f"{row_context}.oracle.{field} is not proven")
            try:
                checked = prior._guard_scenario(
                    canonical_row, shape_name, samples, warmups, row_context, allocator
                )
                summary = metrics._guard_row_summary(
                    original, shape_name, original["scenario"], samples, warmups,
                    allocator, row_context, family,
                )
            except Exception as error:
                fail(f"{row_context} failed canonical guard validation: {error}")
            key = (shape_name, original["scenario"])
            require(key not in parsed, f"{context} duplicate row key {key}")
            parsed[key] = {
                "row": original, "checked": checked, "summary": summary,
                "shape": shape_report,
            }
    require(len(parsed) == len(SHAPES) * len(expected_scenarios),
            f"{context} row inventory is incomplete")
    return parsed


def capture_spec(family: str, lane: str) -> tuple[int, int, bool, bool, str]:
    allocator = lane.startswith("allocator-")
    profile = family == "main" and lane == "profile"
    if allocator:
        return 10, 1, True, False, "allocator-" + lane.removeprefix("allocator-")
    if profile:
        return 3, 0, False, True, lane
    if lane == "preflight":
        return 1, 0, False, profile, lane
    if lane == "pilot":
        return 20, 2, False, profile, lane
    return (
        (500, 5, False, profile, lane)
        if family == "main"
        else (100, 3, False, profile, lane)
    )


def capture_stem(family: str, lane: str) -> str:
    return (
        "guard-" if family == "guard"
        else "fallback-" if family == "fallback"
        else ""
    ) + lane


def verify_capture(root: Path, stage: str, family: str, lane: str,
                   builds: dict[str, Any], source_digest: str, prior: Any,
                   metrics: Any, post_cleanup: bool) -> dict[str, Any]:
    samples, warmups, allocator, profile, _ = capture_spec(family, lane)
    stem = capture_stem(family, lane)
    path = root / stage / f"{stem}-receipt.json"
    build_role = {
        "main": "main-allocator" if allocator else "main",
        "guard": "guard-allocator" if allocator else "guard",
        "fallback": "fallback-allocator" if allocator else "fallback",
    }[family]
    expected_binary = {
        "main": (
            "litchi-perf-baseline-alloc" if allocator
            else "litchi-perf-baseline"
        ),
        "guard": "guard-alloc" if allocator else "guard",
        "fallback": "fallback-alloc" if allocator else "fallback",
    }[family]
    required_artifacts = {f"{stem}-report.json", f"{stem}.log"}
    if family == "main":
        required_artifacts.add(f"{stem}-catalog.json")
        if profile:
            required_artifacts.add("profile.out")
    probe = family if family in {"guard", "fallback"} else None
    receipt = load_json(path)
    common = _receipt_common(
        root, path, receipt, source_digest, f"{stage}/{stem}",
        expected_binary=expected_binary, expected_probe=probe,
        required_artifacts=required_artifacts, post_cleanup=post_cleanup,
    )
    if family == "main":
        validate_main_command(receipt["command"], stem, samples, warmups, allocator, profile)
        report = load_json(root / stage / f"{stem}-report.json")
        catalog = load_json(root / stage / f"{stem}-catalog.json")
        rows = _main_rows(
            report, prior, metrics, builds[build_role], samples, warmups,
            allocator, profile, f"{stage}/{stem}", catalog,
        )
    else:
        validate_probe_command(
            receipt["command"], stem, family, samples, warmups, allocator
        )
        report = load_json(root / stage / f"{stem}-report.json")
        rows = _probe_rows(
            report, family, prior, metrics, samples, warmups, allocator,
            f"{stage}/{stem}",
        )
    log = read_text(root / stage / f"{stem}.log")
    rss = None
    if not allocator:
        values = RSS_RE.findall(log)
        require(
            len(values) == 1 and int(values[0]) > 0,
            f"{stage}/{stem} log must contain exactly one positive GNU time RSS line",
        )
        rss = int(values[0])
    return {
        "stage": stage, "family": family, "lane": lane, "stem": stem,
        "samples": samples, "warmups": warmups, "allocator": allocator,
        "profile": profile, "receipt": receipt, "receipt_info": common,
        "report": report, "rows": rows, "rss_kib": rss,
    }


def required_capture_keys(kept: bool) -> list[tuple[str, str, str]]:
    result: list[tuple[str, str, str]] = []
    for stage in ("before", "after"):
        result.extend((stage, "main", lane) for lane in (*COMMON_LANES, "profile"))
        result.extend((stage, "guard", lane) for lane in COMMON_LANES)
        result.extend((stage, "fallback", lane) for lane in COMMON_LANES)
        if kept:
            result.extend((stage, "main", lane) for lane in FORMAL_LANES)
            result.extend((stage, "guard", lane) for lane in FORMAL_LANES)
            result.extend((stage, "fallback", lane) for lane in FORMAL_LANES)
    return result


def _capture_present(root: Path, stage: str, family: str, lane: str) -> bool:
    stem = capture_stem(family, lane)
    directory = root / stage
    names = [f"{stem}-report.json", f"{stem}-receipt.json", f"{stem}.log"]
    if family == "main":
        names.append(f"{stem}-catalog.json")
        if lane == "profile":
            names += [
                "profile.out", "profile-inclusive.txt",
                "profile-exclusive.txt", "profile-annotations.json",
            ]
    return any(
        (directory / name).exists() or (directory / name).is_symlink()
        for name in names
    )


def verify_capture_inventory(root: Path, kept: bool,
                             builds_by_stage: dict[str, dict[str, Any]],
                             prior: Any, metrics: Any,
                             post_cleanup: bool
                             ) -> dict[tuple[str, str, str], dict[str, Any]]:
    required = required_capture_keys(kept)
    captures: dict[tuple[str, str, str], dict[str, Any]] = {}
    for stage, family, lane in required:
        source_digest = sha(root / stage / "source-manifest.json")
        captures[stage, family, lane] = verify_capture(
            root, stage, family, lane, builds_by_stage[stage], source_digest,
            prior, metrics, post_cleanup,
        )
    # Optional formal lanes on a rejection are still validated when present.
    all_known = {
        (family, lane)
        for family in ("main", "guard", "fallback")
        for lane in (*COMMON_LANES, *FORMAL_LANES, "profile")
        if not (family != "main" and lane == "profile")
    }
    for stage in ("before", "after"):
        for report in sorted((root / stage).glob("*-report.json")):
            stem = report.name.removesuffix("-report.json")
            family = (
                "fallback" if stem.startswith("fallback-")
                else "guard" if stem.startswith("guard-")
                else "main"
            )
            lane = stem.removeprefix(family + "-") if family != "main" else stem
            require((family, lane) in all_known,
                    f"{stage}/{report.name} is an unknown capture lane")
        for family, lane in sorted(all_known):
            if (stage, family, lane) in captures:
                continue
            if not _capture_present(root, stage, family, lane):
                continue
            require(
                not (family == "main" and lane == "profile"),
                f"{stage}/profile artifacts are incomplete",
            )
            source_digest = sha(root / stage / "source-manifest.json")
            captures[stage, family, lane] = verify_capture(
                root, stage, family, lane, builds_by_stage[stage], source_digest,
                prior, metrics, post_cleanup,
            )
    missing = [key for key in required if key not in captures]
    require(not missing, "required capture lanes are missing: " + repr(missing))
    return captures



def row_identity(family: str, key: tuple[str, str], item: dict[str, Any]) -> Any:
    row = item["row"]
    if family == "main":
        return {
            "corpus": row.get("corpus"),
            "sink": row.get("sink"),
            "output_sha256": row.get("output_sha256"),
        }
    shape = item.get("shape", {})
    identity = {
        name: shape.get(name)
        for name in (
            "shape", "sheet_count", "rows_per_sheet", "columns_per_sheet",
            "cells", "one_percent_update_count", "corpus_bytes", "corpus_sha256",
        )
    }
    identity["scenario"] = key[1]
    checked = item["checked"]
    identity.update({
        name: checked.get(name)
        for name in ("warm_store", "changed", "update_count")
    })
    if family == "fallback":
        identity["corpus_variant"] = shape.get("corpus_variant")
        identity["source_markers"] = shape.get("source_markers")
        identity["descent_readback"] = shape.get("descent_readback")
    return identity


def verify_identities(captures: dict[tuple[str, str, str], dict[str, Any]]) -> None:
    for family in ("main", "guard", "fallback"):
        reference: dict[tuple[str, str], Any] = {}
        for (stage, current_family, lane), capture in captures.items():
            del stage, lane
            if current_family != family:
                continue
            for key, item in capture["rows"].items():
                identity = row_identity(family, key, item)
                if key in reference:
                    require(canonical(reference[key]) == canonical(identity),
                            f"{family} corpus/output identity differs for {key}")
                else:
                    reference[key] = identity


def percent_change(after: float, before: float) -> float | None:
    if before <= 0:
        return None
    value = (after / before - 1.0) * 100.0
    require(math.isfinite(value), "percent change is not finite")
    return value


def comparison(before: Any, after: Any) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for name in STATISTICS:
        left = before[name]
        right = after[name]
        change = percent_change(float(right), float(left))
        result[name] = {
            "before": left,
            "after": right,
            "candidate_minus_control_percent": change,
            "adverse_over_5_percent": None if change is None else change > 5.0,
        }
    return result


def compare_captures(
    captures: dict[tuple[str, str, str], dict[str, Any]],
) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    for family in ("main", "guard", "fallback"):
        lanes = sorted({
            lane for _, current, lane in captures if current == family
        })
        for lane in lanes:
            before = captures.get(("before", family, lane))
            after = captures.get(("after", family, lane))
            if before is None or after is None:
                continue
            require(set(before["rows"]) == set(after["rows"]),
                    f"matched {family}/{lane} rows are incomplete")
            for key in sorted(before["rows"]):
                left = before["rows"][key]["summary"]
                right = after["rows"][key]["summary"]
                entry: dict[str, Any] = {
                    "family": family,
                    "lane": lane,
                    "row_key": f"{key[0]}/{key[1]}",
                    "identity": "matched",
                    "elapsed_excluded": bool(
                        left.get("elapsed_excluded") or right.get("elapsed_excluded")
                    ),
                    "latency": None,
                    "allocation": None,
                }
                if not entry["elapsed_excluded"]:
                    entry["latency"] = comparison(left["latency"], right["latency"])
                left_alloc = left.get("allocation", {"status": "unavailable"})
                right_alloc = right.get("allocation", {"status": "unavailable"})
                if left_alloc.get("status") == right_alloc.get("status") == "measured":
                    fields: dict[str, Any] = {}
                    for name in sorted(
                        set(left_alloc["fields"]) & set(right_alloc["fields"])
                    ):
                        fields[name] = comparison(
                            {
                                key_name: left_alloc["fields"][name][key_name]
                                for key_name in ("p50", "mean", "p95", "p99")
                            },
                            {
                                key_name: right_alloc["fields"][name][key_name]
                                for key_name in ("p50", "mean", "p95", "p99")
                            },
                        )
                    entry["allocation"] = {
                        "status": "compared",
                        "fields": fields,
                        "scope": "allocator counters and peaks only",
                    }
                else:
                    entry["allocation"] = {
                        "status": "unavailable_or_uncompared",
                        "before": left_alloc.get("status"),
                        "after": right_alloc.get("status"),
                    }
                output.append(entry)
    return output


def individual_rows(
    captures: dict[tuple[str, str, str], dict[str, Any]],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for key in sorted(captures):
        stage, family, lane = key
        capture = captures[key]
        for row_key in sorted(capture["rows"]):
            item = capture["rows"][row_key]
            summary = item["summary"]
            rows.append({
                "stage": stage,
                "family": family,
                "lane": lane,
                "row_key": f"{row_key[0]}/{row_key[1]}",
                "sample_count": capture["samples"],
                "warmups": capture["warmups"],
                "elapsed_excluded": capture["allocator"],
                "statistics": summary.get("latency"),
                "allocation": summary.get("allocation"),
                "rss_kib": capture["rss_kib"],
            })
    return rows


def verify_serial(
    captures: dict[tuple[str, str, str], dict[str, Any]], kept: bool
) -> dict[str, Any]:
    intervals: list[tuple[_datetime.datetime, _datetime.datetime, str]] = []
    for (stage, family, lane), capture in captures.items():
        start = parse_time(
            capture["receipt"]["started_utc"],
            f"{stage}/{family}/{lane}.started_utc",
        )
        end = start + _datetime.timedelta(
            seconds=float(capture["receipt"]["elapsed_seconds"])
        )
        intervals.append((start, end, f"{stage}/{family}/{lane}"))
    intervals.sort()
    for left, right in zip(intervals, intervals[1:]):
        require(left[1] <= right[0],
                f"capture operations overlap: {left[2]} and {right[2]}")
    formal_orders: dict[str, list[str]] = {}
    if kept:
        for family in ("main", "guard", "fallback"):
            order = [
                ("before", "r1"), ("after", "r1"),
                ("after", "r2"), ("before", "r2"),
            ]
            times = []
            for stage, lane in order:
                capture = captures[stage, family, lane]
                start = parse_time(capture["receipt"]["started_utc"], "formal start")
                end = start + _datetime.timedelta(
                    seconds=float(capture["receipt"]["elapsed_seconds"])
                )
                times.append((start, end))
            require(
                all(left[1] <= right[0] for left, right in zip(times, times[1:])),
                f"{family} formal lanes are not ABBA ordered",
            )
            formal_orders[family] = [f"{stage}/{lane}" for stage, lane in order]
    return {"all_capture_intervals_serial": True, "formal_abba": formal_orders}


def verify_profiles(root: Path, profiles: Any, post_cleanup: bool) -> dict[str, Any]:
    # profiles.py binds raw/annotation hashes and selected direct edges. Receipt
    # command/source/binary binding remains in this verifier.
    for stage in ("before", "after"):
        path = root / stage / "profile-receipt.json"
        receipt = load_json(path)
        source_digest = sha(root / stage / "source-manifest.json")
        _receipt_common(
            root, path, receipt, source_digest, f"{stage}/profile",
            expected_binary="litchi-perf-baseline",
            required_artifacts={
                "profile-report.json", "profile-catalog.json",
                "profile.log", "profile.out",
            },
            post_cleanup=post_cleanup,
        )
        validate_main_command(receipt["command"], "profile", 3, 0, False, True)
    result = profiles.collect(root)
    require(
        result.get("status") == "complete"
        and not result.get("issues")
        and set(result.get("profiles", {})) == {"before", "after"},
        "profile parser did not produce complete before/after evidence",
    )
    return result


def verify_cleanup(root: Path, post_cleanup: bool) -> dict[str, Any]:
    relocation = root / "target-relocation.json"
    result: dict[str, Any] = {
        "post_cleanup": post_cleanup,
        "same_path_replay": True,
    }
    if relocation.is_file():
        record = load_json(relocation)
        require(isinstance(record, dict), "target-relocation.json is not an object")
        require(record.get("logical_target") == str(SCRATCH / "target"),
                "logical target relocation differs")
        require(record.get("cargo_environment_unchanged") is True,
                "target relocation does not preserve cargo environment")
        require(isinstance(record.get("inventory_sha256"), str),
                "target relocation inventory hash is missing")
        required = record.get("cleanup_required")
        require(isinstance(required, list) and required,
                "target relocation cleanup inventory is missing")
        for value in required:
            require(isinstance(value, str) and Path(value).is_absolute(),
                    "target relocation cleanup path is malformed")
        result["relocation"] = {
            "logical_target": record["logical_target"],
            "physical_target": record.get("physical_target"),
            "cleanup_required": required,
            "inventory_sha256": record["inventory_sha256"],
        }
        if post_cleanup:
            require(all(not Path(value).exists() for value in required),
                    "owned cleanup path remains after post-cleanup replay")
    if post_cleanup:
        require(not SCRATCH.exists(),
                "owned scratch root remains after post-cleanup replay")
        result["owned_paths_absent"] = True
    return result


def verify_sealed_sha(root: Path, decision: dict[str, Any]) -> dict[str, Any]:
    source = decision["source"]
    value = source.get("sealed_sha256sums", decision.get("sealed_sha256sums"))
    if value is None:
        for candidate in (
            "SHA256SUMS", "sealed-sha256sums.json", "sealed-inventory.json"
        ):
            if (root / candidate).is_file():
                value = candidate
                break
    require(value is not None, "sealed SHA-256 inventory is missing")
    path = bundle_path(root, value, "sealed SHA-256 inventory")
    entries: dict[str, str] = {}
    if path.suffix == ".json":
        value_json = load_json(path)
        require(isinstance(value_json, dict), "sealed JSON inventory is not an object")
        value_json = value_json.get("files", value_json)
        require(isinstance(value_json, dict), "sealed JSON inventory files are missing")
        for name, digest in value_json.items():
            relative = safe_relative(name, "sealed inventory path").as_posix()
            require(relative not in entries, f"sealed inventory repeats {relative}")
            entries[relative] = check_hash(digest, f"sealed inventory {name}")
    else:
        for line_number, line in enumerate(read_text(path).splitlines(), 1):
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            fields = line.split(None, 1)
            require(len(fields) == 2, f"sealed inventory line {line_number} is malformed")
            digest, name = fields[0], fields[1].lstrip("*")
            relative = safe_relative(
                name, f"sealed inventory line {line_number}"
            ).as_posix()
            require(relative not in entries, f"sealed inventory repeats {relative}")
            entries[relative] = check_hash(digest, f"sealed inventory line {line_number}")
    require(entries, "sealed SHA-256 inventory is empty")
    declared_exclusions = decision.get("sealed_replay_outputs", source.get("sealed_replay_outputs", []))
    require(isinstance(declared_exclusions, list),
            "sealed replay-output exclusions are not a list")
    exclusions = {safe_relative(item, "sealed replay-output exclusion").as_posix()
                  for item in declared_exclusions}
    inventory_name = path.relative_to(root).as_posix()
    exclusions.add(inventory_name)
    retained_files = {
        item.relative_to(root).as_posix()
        for item in root.rglob("*")
        if item.is_file() and not item.is_symlink()
    }
    expected_files = retained_files - exclusions
    require(set(entries) == expected_files,
            "sealed SHA-256 inventory membership differs from retained evidence")
    for name, digest in entries.items():
        target = bundle_path(root, name, f"sealed inventory {name}")
        require(sha(target) == digest, f"sealed inventory hash differs for {name}")
    declared_hash = source.get(
        "sealed_sha256sums_sha256", decision.get("sealed_sha256sums_sha256")
    )
    if declared_hash is not None:
        require(
            sha(path) == check_hash(declared_hash, "sealed inventory digest"),
            "sealed inventory file digest differs",
        )
    return {"path": path.relative_to(root).as_posix(), "entries": len(entries),
            "sha256_checked": True}


def _expect_rejection(label: str, function: Callable[[], Any]) -> bool:
    try:
        function()
    except Exception:
        return True
    fail(f"negative vector was accepted: {label}")


def verify_negative_vectors(
    root: Path, captures: dict[tuple[str, str, str], dict[str, Any]],
    prior: Any, metrics: Any,
) -> dict[str, bool]:
    main = captures["before", "main", "pilot"]
    key = ("dense-wide", "xlsx_one_percent_commit_save")
    row = main["rows"][key]["row"]

    short = copy.deepcopy(row)
    short["elapsed_ns"]["samples"].pop()
    changed = copy.deepcopy(row)
    changed["elapsed_ns"]["samples"][0] += 1
    reordered = copy.deepcopy(row)
    indices = reordered["operation_metrics"]["sample_indices"]
    indices[0], indices[1] = indices[1], indices[0]

    allocator = captures["before", "main", "allocator-r1"]["rows"][key]["row"]
    alloc_short = copy.deepcopy(allocator)
    alloc_short["operation_metrics"]["allocation"]["allocated_bytes"]["values"].pop()
    alloc_bounds = copy.deepcopy(allocator)
    alloc_bounds["operation_metrics"]["allocation"]["region_peak_live_bytes"]["values"][0] = 0
    normal_alloc = copy.deepcopy(row)
    normal_alloc["operation_metrics"]["allocation"]["allocation_calls"]["status"] = "measured"

    report = copy.deepcopy(main["report"])
    missing_report = copy.deepcopy(report)
    missing_report["results"].pop()
    duplicate_report = copy.deepcopy(report)
    duplicate_report["results"][-1] = copy.deepcopy(duplicate_report["results"][0])
    unknown_report = copy.deepcopy(report)
    unknown_report["results"][0]["case"] = "unknown-case"

    guard = captures["before", "guard", "pilot"]["report"]
    missing_guard = copy.deepcopy(guard)
    missing_guard["shapes"][0]["scenarios"].pop()

    receipt = copy.deepcopy(main["receipt"])
    receipt["source_manifest_sha256"] = "0" * 64
    after_record = load_json(root / "after" / "candidate-patch.json")
    bad_record = copy.deepcopy(after_record)
    bad_record["previous_epoch"] = "before"

    def parse_main(value: dict[str, Any]) -> Any:
        return _main_rows(
            value, prior, metrics,
            {"binary_sha256": main["report"]["binary_identity"]["binary_sha256"]},
            20, 2, False, False, "negative-main", main["catalog"],
        )

    def parse_guard(value: dict[str, Any]) -> Any:
        return _probe_rows(value, "guard", prior, metrics, 20, 2, False,
                           "negative-guard")

    def check_receipt() -> Any:
        _receipt_common(
            root, root / "before/pilot-receipt.json", receipt,
            sha(root / "before/source-manifest.json"), "negative receipt",
        )

    def check_patch() -> Any:
        before = load_manifest(root / "before/source-manifest.json", "negative before")
        after = load_manifest(root / "after/source-manifest.json", "negative after")
        replay_patch(root, "after", before, after, bad_record)

    return {
        "short_elapsed_vector_rejected": _expect_rejection(
            "short elapsed vector",
            lambda: metrics._validate_elapsed(short, 20, "negative short"),
        ),
        "altered_elapsed_statistics_rejected": _expect_rejection(
            "altered elapsed vector",
            lambda: metrics._validate_elapsed(changed, 20, "negative changed"),
        ),
        "reordered_operation_vector_rejected": _expect_rejection(
            "reordered operation vector",
            lambda: prior.verify_row(
                reordered, "xlsx_one_percent_commit_save", "dense-wide", 20,
                "negative operation", False
            ),
        ),
        "short_allocator_vector_rejected": _expect_rejection(
            "short allocator vector",
            lambda: prior.verify_row(
                alloc_short, "xlsx_one_percent_commit_save", "dense-wide", 10,
                "negative allocator", True
            ),
        ),
        "allocator_peak_bounds_rejected": _expect_rejection(
            "allocator peak bounds",
            lambda: prior.verify_row(
                alloc_bounds, "xlsx_one_percent_commit_save", "dense-wide", 10,
                "negative peak", True
            ),
        ),
        "normal_allocator_values_rejected": _expect_rejection(
            "normal allocator values",
            lambda: prior.verify_row(
                normal_alloc, "xlsx_one_percent_commit_save", "dense-wide", 20,
                "negative normal allocation", False
            ),
        ),
        "missing_main_row_rejected": _expect_rejection(
            "missing main row", lambda: parse_main(missing_report)
        ),
        "duplicate_main_row_rejected": _expect_rejection(
            "duplicate main row", lambda: parse_main(duplicate_report)
        ),
        "unknown_main_row_rejected": _expect_rejection(
            "unknown main row", lambda: parse_main(unknown_report)
        ),
        "missing_guard_row_rejected": _expect_rejection(
            "missing guard row", lambda: parse_guard(missing_guard)
        ),
        "receipt_source_mutation_rejected": _expect_rejection(
            "receipt source mutation", check_receipt
        ),
        "patch_chain_mutation_rejected": _expect_rejection(
            "patch chain mutation", check_patch
        ),
    }



def verify(root: Path = HERE, post_cleanup: bool = False) -> dict[str, Any]:
    root = Path(root).resolve()
    require(root.is_dir(), f"evidence root is not a directory: {root}")
    decision, kept = load_decision(root)
    metrics, profiles, prior = load_helpers()

    source = verify_source(root, decision, kept)
    verify_adr(root)
    candidate_digest = source["candidate_manifest_sha256"]
    final_digest = source["final_manifest_sha256"]
    baseline_digest = source["baseline_test_manifest_sha256"]
    patch_record = load_json(root / "after" / "candidate-patch.json")
    candidate_patch_digest = check_hash(
        patch_record.get("patch_sha256"), "after.patch_sha256"
    )
    correctness = verify_correctness_admission(
        root, candidate_digest, candidate_patch_digest
    )
    failures = verify_expected_failures(root)
    baseline_failures = verify_baseline_failures(root)
    quality = verify_quality_groups(
        root, decision, kept, candidate_digest, final_digest,
        baseline_digest, post_cleanup,
    )

    builds = {
        stage: verify_builds(
            root, stage, sha(root / stage / "source-manifest.json"), post_cleanup
        )
        for stage in ("before", "after")
    }
    captures = verify_capture_inventory(
        root, kept, builds, prior, metrics, post_cleanup
    )
    verify_identities(captures)
    profile_result = verify_profiles(root, profiles, post_cleanup)
    comparisons = compare_captures(captures)
    serial = verify_serial(captures, kept)
    negative = verify_negative_vectors(root, captures, prior, metrics)
    require(all(negative.values()), "one or more negative vectors were not rejected")
    cleanup = verify_cleanup(root, post_cleanup)
    sealed = verify_sealed_sha(root, decision)

    return {
        "schema": SCHEMA,
        "decision": {
            "value": decision["decision"],
            "candidate_epoch": decision["candidate_epoch"],
            "final_epoch": decision["final_epoch"],
            "formal_abba_required": kept,
        },
        "performance_claim": "none",
        "claim_authorized": False,
        "source": source,
        "adr": {"revision": BASE_REVISION, "checked": True},
        "correctness_admission": correctness,
        "expected_failures": failures,
        "baseline_failures": baseline_failures,
        "quality": quality,
        "builds": builds,
        "captures": {
            "required_lanes": [list(item) for item in required_capture_keys(kept)],
            "individual_rows": individual_rows(captures),
            "comparisons": comparisons,
            "rss_scope": (
                "GNU time RSS from native non-allocator logs only; "
                "allocator elapsed/RSS excluded"
            ),
            "serial": serial,
        },
        "profiles": profile_result,
        "negative_vectors": negative,
        "cleanup": cleanup,
        "sealed_sha256": sealed,
        "scope": "OLE2/OOXML XLSX evidence; ODF is deferred",
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Verify retained 0516 XLSX evidence.")
    parser.add_argument(
        "--root", type=Path, default=HERE,
        help="0516 evidence root (default: this script's directory)",
    )
    parser.add_argument(
        "--post-cleanup", action="store_true",
        help="verify the same evidence root after owned scratch cleanup",
    )
    parser.add_argument(
        "--skeleton", action="store_true",
        help="print the exact decision contract skeleton and exit",
    )
    args = parser.parse_args(argv)
    if args.skeleton:
        print(json.dumps(DECISION_SKELETON, indent=2, sort_keys=True))
        return 0
    try:
        result = verify(args.root, args.post_cleanup)
    except (OSError, VerificationError, AssertionError) as error:
        print(json.dumps({
            "schema": SCHEMA,
            "status": "incomplete",
            "performance_claim": "none",
            "claim_authorized": False,
            "decision_skeleton": DECISION_SKELETON,
            "error": str(error),
        }, indent=2, sort_keys=True))
        return 1
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
