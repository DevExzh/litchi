#!/usr/bin/env python3
"""Freeze, run, replay, and reduce the prospective 0933 paired cohort.

The runner is deliberately inert until ``capture`` is given a manifest whose
root authorization flag is true and ``LITCHI_PERF_0933_CAPTURE_APPROVED=1``.
``freeze`` only verifies custody and writes an unauthorized manifest.  Raw
reports and process logs belong in an external packet; repository records keep
descriptors and scalar validation summaries.

The frozen cohort is a 48-child normal lane plus a 24-child allocator lane
(72 children total).  Each child is terminally recorded before its report is
opened.  An observation timeout polls the same process; only a confirmed
process deadline terminates it, and no failed child is retried.
"""

from __future__ import annotations

import argparse
import copy
import difflib
import hashlib
import importlib.util
import json
import math
import os
import random
import re
import signal
import shutil
import statistics
import subprocess
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


REPO_ROOT = Path(__file__).resolve().parents[4]
COMPARISON_PATH = REPO_ROOT / "docs/performance/results/change-0925/comparison.py"
PLAN_DEFAULT = REPO_ROOT / "docs/performance/results/change-0933/comparison-plan-v1.json"
CORRECTION_DEFAULT = REPO_ROOT / "docs/performance/results/change-0933/comparison-formal-gate-correction-v1.json"
BASE_REVISION = "a17d87da48504e8d1c1a6240a6c5d1b11219031f"
CASE = "docx_source_backed_existing_section_layout_edit_save"
ROUTES = ("generic_read_at", "owned_vec", "owned_reader", "range_short")
PHASES = ("preparation_ns", "open_ns", "query_ns", "edit_ns", "commit_ns", "publication_ns")
ALLOCATOR_SECTIONS = (16, 8192)
ALLOCATOR_OWNERSHIP = ("shared_snapshot_clones", "independent_sources")
ALLOCATOR_REPETITIONS = 3
CPU = 12
WARMUPS = 5
SAMPLES = 50
BLOCKS = 6
NORMAL_CHILDREN = BLOCKS * len(ROUTES) * 2
ALLOCATOR_CHILDREN = len(ALLOCATOR_SECTIONS) * len(ALLOCATOR_OWNERSHIP) * ALLOCATOR_REPETITIONS * 2
CHILDREN = NORMAL_CHILDREN + ALLOCATOR_CHILDREN
NORMAL_DEADLINE_SECONDS = 900.0
ALLOCATOR_DEADLINE_SECONDS = 600.0
OBSERVE_SECONDS = 1.0
MAX_ARTIFACT_BYTES = 1 << 20
MAX_SOURCE_LISTING_BYTES = 64 << 20
CHILD_ENV_BASE = {
    "PATH": "/usr/bin:/bin",
    "LC_ALL": "C",
    "LANG": "C",
    "TZ": "UTC",
    "RUST_BACKTRACE": "0",
}
EXPECTED_CHANGED_PATHS = (
    "crates/litchi-docx/src/section/inventory.rs",
    "crates/litchi-docx/src/source_backed.rs",
)
COMMON_RUNTIME_PATHS = (
    "tools/perf-baseline/src/lib.rs",
    "tools/perf-baseline/src/docx_section_retention.rs",
    "tools/perf-baseline/src/bin/litchi-perf-baseline-alloc.rs",
    "tools/perf-baseline/tests/docx_section_retention.rs",
    "crates/litchi-docx/src/section/tests.rs",
)


class EvidenceError(RuntimeError):
    """A custody, schedule, report, or replay contract failed closed."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def parse_utc(value: Any, label: str) -> datetime:
    require(isinstance(value, str), f"{label}: timestamp missing")
    try:
        parsed = datetime.fromisoformat(value)
    except ValueError as error:
        raise EvidenceError(f"{label}: invalid timestamp") from error
    require(parsed.tzinfo is not None, f"{label}: timestamp has no timezone")
    return parsed.astimezone(timezone.utc)


def sha256_path(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def descriptor(path: Path, *, root: Path | None = None) -> dict[str, Any]:
    path = path.resolve()
    value = str(path)
    if root is not None:
        try:
            value = path.relative_to(root.resolve()).as_posix()
        except ValueError as error:
            raise EvidenceError(f"path is outside packet root: {path}") from error
    return {"path": value, "bytes": path.stat().st_size, "sha256": sha256_path(path)}


def read_json(path: Path, *, limit: int | None = MAX_ARTIFACT_BYTES) -> dict[str, Any]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    if limit is not None:
        require(path.stat().st_size < limit,
                f"JSON exceeds repository evidence limit: {path} ({path.stat().st_size} bytes)")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as error:
        raise EvidenceError(f"cannot parse JSON {path}: {error}") from error
    require(isinstance(value, dict), f"JSON is not an object: {path}")
    return value


def write_new_json(path: Path, value: Any, *, limit: int | None = None) -> None:
    data = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    if limit is not None:
        require(len(data) < limit, f"output exceeds bound: {path} ({len(data)} bytes)")
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("xb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError as error:
        raise EvidenceError(f"refusing to overwrite receipt: {path}") from error


def safe_relative(value: str, label: str) -> Path:
    path = Path(value)
    require(not path.is_absolute(), f"{label}: absolute path is not portable")
    require(path.parts and all(part not in ("", ".", "..") for part in path.parts),
            f"{label}: unsafe relative path {value!r}")
    return path


def load_comparison(path: Path = COMPARISON_PATH) -> Any:
    path = path.resolve()
    spec = importlib.util.spec_from_file_location("litchi_perf_0933_comparison", path)
    require(spec is not None and spec.loader is not None, f"cannot load validator: {COMPARISON_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def schedule_rows() -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for block in range(1, BLOCKS + 1):
        routes = ROUTES if block % 2 else tuple(reversed(ROUTES))
        for position, route in enumerate(routes, start=1):
            versions = ("baseline", "candidate") if block % 2 else ("candidate", "baseline")
            for version_position, version in enumerate(versions, start=1):
                rows.append({
                    "id": f"normal-block-{block:02d}-{position:02d}-{route}-{version}",
                    "pair_id": f"normal-block-{block:02d}-{position:02d}-{route}",
                    "lane": "normal",
                    "block": block,
                    "position": position,
                    "version_position": version_position,
                    "route": route,
                    "version": version,
                    "cpu": CPU,
                    "warmups": WARMUPS,
                    "samples": SAMPLES,
                    "deadline_seconds": int(NORMAL_DEADLINE_SECONDS),
                })
    allocator_pair = 0
    for section_count in ALLOCATOR_SECTIONS:
        for ownership in ALLOCATOR_OWNERSHIP:
            for repetition in range(1, ALLOCATOR_REPETITIONS + 1):
                allocator_pair += 1
                versions = ("baseline", "candidate") if repetition % 2 else ("candidate", "baseline")
                for version_position, version in enumerate(versions, start=1):
                    rows.append({
                        "id": f"allocator-section-{section_count:05d}-{ownership}-rep-{repetition:02d}-{version}",
                        "pair_id": f"allocator-section-{section_count:05d}-{ownership}-rep-{repetition:02d}",
                        "lane": "allocator",
                        "block": allocator_pair,
                        "position": 1,
                        "version_position": version_position,
                        "route": None,
                        "version": version,
                        "cpu": CPU,
                        "warmups": 0,
                        "samples": 1,
                        "deadline_seconds": int(ALLOCATOR_DEADLINE_SECONDS),
                        "sections": section_count,
                        "concurrent_commits": 8,
                        "policy": "exact",
                        "ownership": ownership,
                        "repetition": repetition,
                    })
    require(len(rows) == CHILDREN, "formal schedule child count changed")
    normal_rows = [row for row in rows if row["lane"] == "normal"]
    allocator_rows = [row for row in rows if row["lane"] == "allocator"]
    require(len(normal_rows) == NORMAL_CHILDREN and len(allocator_rows) == ALLOCATOR_CHILDREN,
            "formal lane child counts changed")
    for route in ROUTES:
        require(sum(row["lane"] == "normal" and row["route"] == route for row in rows) == BLOCKS * 2,
                f"formal schedule route count changed: {route}")
        for version in ("baseline", "candidate"):
            require(sum(row["lane"] == "normal" and row["route"] == route and row["version"] == version for row in rows) == BLOCKS,
                    f"formal schedule version count changed: {route}/{version}")
    for section_count in ALLOCATOR_SECTIONS:
        for ownership in ALLOCATOR_OWNERSHIP:
            require(sum(row["lane"] == "allocator" and row["sections"] == section_count
                        and row["ownership"] == ownership for row in rows) == ALLOCATOR_REPETITIONS * 2,
                    f"allocator schedule cell count changed: {section_count}/{ownership}")
    return rows


def schedule_hash(rows: list[dict[str, Any]]) -> str:
    data = json.dumps(rows, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(data).hexdigest()


def verify_descriptor(value: dict[str, Any], *, label: str, root: Path | None = None) -> Path:
    require(isinstance(value, dict), f"{label}: descriptor is not an object")
    raw = value.get("path")
    require(isinstance(raw, str), f"{label}: descriptor path missing")
    path = (root / safe_relative(raw, label)).resolve() if root is not None else Path(raw).resolve()
    actual = descriptor(path)
    require(actual["bytes"] == value.get("bytes") and actual["sha256"] == value.get("sha256"),
            f"{label}: bytes or SHA256 changed")
    return path


def validate_source_receipt(receipt_path: Path, *, candidate: bool | None = None) -> dict[str, Any]:
    """Fail closed on the complete isolated source tree and its listing."""
    receipt = read_json(receipt_path)
    schema = str(receipt.get("schema", ""))
    require(schema.startswith("litchi.performance.0933") or "source" in schema,
            f"unsupported source receipt schema: {receipt.get('schema')!r}")
    require(receipt.get("base_revision") == BASE_REVISION,
            "source base revision changed")
    source_revision = receipt.get("source_revision")
    require(isinstance(source_revision, str) and len(source_revision) == 40
            and all(char in "0123456789abcdefABCDEF" for char in source_revision),
            "source revision is missing or is not a 40-hex identifier")
    source_root = Path(str(receipt.get("source_root", ""))).resolve()
    require(source_root.is_dir(), f"source root missing: {source_root}")
    listing_raw = receipt.get("full_source_listing")
    require(isinstance(listing_raw, str), "full source listing path missing")
    listing_path = Path(listing_raw).resolve()
    require(listing_path.is_file() and listing_path.stat().st_size <= MAX_SOURCE_LISTING_BYTES,
            "full source listing missing or too large")
    listing = descriptor(listing_path)
    require(listing["bytes"] == receipt.get("full_source_listing_bytes")
            and listing["sha256"] == receipt.get("full_source_listing_sha256"),
            "full source listing descriptor changed")
    if candidate is None:
        candidate = receipt.get("role") == "candidate" or "candidate" in schema
    patch: dict[str, Any] | None = None
    if candidate:
        patch_value = receipt.get("candidate_patch")
        require(isinstance(patch_value, dict), "candidate patch descriptor missing")
        verify_descriptor(patch_value, label="candidate patch")
        patch = dict(patch_value)
        changed = receipt.get("changed_paths")
        require(changed == list(EXPECTED_CHANGED_PATHS),
                "candidate changed-file allowlist changed")
        require(receipt.get("changed_path_count", len(changed)) == len(EXPECTED_CHANGED_PATHS),
                "candidate changed-file count changed")
    require(receipt.get("generated_targets_included") is False,
            "generated target entered source closure")

    rows: list[dict[str, Any]] = []
    seen: set[str] = set()
    for line in listing_path.read_text(encoding="utf-8").splitlines():
        try:
            row = json.loads(line)
        except ValueError as error:
            raise EvidenceError(f"invalid source listing row: {error}") from error
        require(isinstance(row, dict), "source listing row is not an object")
        rel = row.get("path")
        require(isinstance(rel, str), "source listing path missing")
        safe_relative(rel, "source listing path")
        require(rel not in seen, f"duplicate source listing path: {rel}")
        seen.add(rel)
        kind = row.get("kind", "file")
        require(kind == "file", f"unsupported source listing row kind: {rel}")
        size = row.get("size_bytes", row.get("bytes"))
        digest = row.get("sha256")
        require(type(size) is int and size >= 0, f"invalid source listing size: {rel}")
        require(isinstance(digest, str) and len(digest) == 64,
                f"invalid source listing hash: {rel}")
        rows.append({"path": rel, "size_bytes": size, "sha256": digest})
    actual_files = {
        str(path.relative_to(source_root)).replace(os.sep, "/"): path
        for path in source_root.rglob("*")
        if path.is_file() and not path.is_symlink() and ".git" not in path.parts
        and "target" not in path.parts
    }
    require(set(actual_files) == seen, "source tree file set differs from full listing")
    total = 0
    file_descriptors: dict[str, dict[str, Any]] = {}
    for row in rows:
        path = actual_files[row["path"]]
        actual = descriptor(path)
        require(actual["bytes"] == row["size_bytes"] and actual["sha256"] == row["sha256"],
                f"source tree file changed: {row['path']}")
        total += actual["bytes"]
        file_descriptors[row["path"]] = actual
    expected_count = receipt.get("source_file_count", receipt.get("file_count"))
    expected_bytes = receipt.get("source_bytes")
    require(len(rows) == expected_count, "source file count changed")
    require(expected_bytes is None or total == expected_bytes, "source byte total changed")
    if candidate:
        for path in EXPECTED_CHANGED_PATHS:
            require(path in actual_files, f"changed source file missing: {path}")
        baseline_root_value = receipt.get("baseline_source_root")
        require(isinstance(baseline_root_value, str), "candidate baseline source root missing")
        baseline_root = Path(baseline_root_value).resolve()
        require(baseline_root.is_dir(), f"candidate baseline source root missing: {baseline_root}")
        expected_patch: list[str] = []
        for path in EXPECTED_CHANGED_PATHS:
            left_path = baseline_root / path
            right_path = actual_files[path]
            require(left_path.is_file(), f"candidate patch baseline path missing: {path}")
            left = left_path.read_text(encoding="utf-8").splitlines(keepends=True)
            right = right_path.read_text(encoding="utf-8").splitlines(keepends=True)
            expected_patch.extend(difflib.unified_diff(left, right, fromfile=f"a/{path}",
                                                        tofile=f"b/{path}", n=3))
        require(Path(str(patch["path"])).read_text(encoding="utf-8")
                == "".join(expected_patch),
                "candidate patch bytes do not equal the declared baseline-to-candidate diff")
    return {
        "receipt": descriptor(receipt_path),
        "source_root": str(source_root),
        "base_revision": receipt.get("base_revision", BASE_REVISION),
        "full_source_listing": listing,
        "candidate_patch": patch,
        "source_file_count": len(rows),
        "source_bytes": total,
        "changed_paths": list(EXPECTED_CHANGED_PATHS) if candidate else [],
        "source_files": file_descriptors,
        "role": "candidate" if candidate else "baseline",
        "source_revision": source_revision,
    }


def _command_argv(receipt: dict[str, Any]) -> list[str]:
    argv = receipt.get("command_argv")
    if isinstance(argv, list) and all(isinstance(value, str) for value in argv):
        return list(argv)
    command = receipt.get("command")
    require(isinstance(command, str), "build command is missing")
    import shlex
    return shlex.split(command)


def _receipt_binary(receipt: dict[str, Any], label: str) -> dict[str, Any]:
    value = receipt.get("binary") or receipt.get("artifact")
    require(isinstance(value, dict), f"{label}: binary descriptor missing")
    normalized = dict(value)
    if "bytes" not in normalized and "size_bytes" in normalized:
        normalized["bytes"] = normalized["size_bytes"]
    path = verify_descriptor(normalized, label=f"{label} binary")
    require(os.access(path, os.X_OK), f"{label}: binary is not executable")
    return {"path": str(path), "bytes": normalized["bytes"], "sha256": normalized["sha256"]}


def _receipt_argv(receipt: dict[str, Any]) -> list[str]:
    value = receipt.get("argv") or receipt.get("command_argv") or receipt.get("command")
    if isinstance(value, list) and all(isinstance(item, str) for item in value):
        return list(value)
    return _command_argv(receipt)


def _bound_hash(container: Any, *names: str) -> str | None:
    if not isinstance(container, dict):
        return None
    for name in names:
        value = container.get(name)
        if isinstance(value, str):
            return value
        if isinstance(value, dict) and isinstance(value.get("sha256"), str):
            return value["sha256"]
    return None


def validate_build_receipt(receipt_path: Path, source: dict[str, Any], *, lane: str) -> dict[str, Any]:
    receipt = read_json(receipt_path)
    terminal = receipt.get("terminal") if isinstance(receipt.get("terminal"), dict) else {}
    exit_code = receipt.get("exit_code", terminal.get("exit_code"))
    if exit_code is None:
        exit_code = terminal.get("tool_return_exit_code")
    require(exit_code == 0, f"{lane} build receipt does not record successful completion")
    argv = _receipt_argv(receipt)
    for flag in ("--offline", "--locked", "--release", "-j2"):
        require(flag in argv, f"{lane} build command lacks {flag}")
    require(any(value == "+1.95.0" or "1.95.0" in value for value in argv),
            f"{lane} build command is not bound to +1.95.0")
    require("--features" not in argv or "allocator-metrics" in argv,
            f"{lane} build feature selection is not explicit")
    binary = _receipt_binary(receipt, lane)
    binary_name = Path(binary["path"]).name
    if lane == "normal":
        require(binary_name == "litchi-perf-baseline" and "--bin" in argv
                and "litchi-perf-baseline" in argv,
                "normal build binary identity changed")
    else:
        require(binary_name == "litchi-perf-baseline-alloc" and "allocator-metrics" in argv
                and "litchi-perf-baseline-alloc" in argv,
                "allocator build binary/feature identity changed")
    source_binding = receipt.get("source_custody") or receipt.get("source")
    require(isinstance(source_binding, dict), f"{lane} build source custody missing")
    bound_root = source_binding.get("source_root") or source_binding.get("root") or receipt.get("cwd")
    require(bound_root == source["source_root"], f"{lane} build source root binding differs")
    bound_receipt_hash = _bound_hash(source_binding, "source_receipt_sha256", "source_receipt")
    bound_listing_hash = _bound_hash(source_binding, "full_source_listing_sha256", "full_source_listing")
    require(bound_receipt_hash == source["receipt"]["sha256"],
            f"{lane} build/source receipt binding differs")
    require(bound_listing_hash == source["full_source_listing"]["sha256"],
            f"{lane} build/source listing binding differs")
    require(receipt.get("source_revision") == source["source_revision"],
            f"{lane} build/source revision binding differs")
    require(isinstance(receipt.get("env_policy") or receipt.get("environment_scope"), str),
            f"{lane} build environment policy is missing")
    profile = receipt.get("profile", receipt.get("build_profile", "release"))
    require(profile == "release", f"{lane} build profile changed")
    locks = receipt.get("locks") or receipt.get("lock_bindings")
    require(isinstance(locks, dict) and locks, f"{lane} build lock identities missing")
    terminal_value = receipt.get("terminal_receipt")
    if isinstance(terminal_value, dict):
        verify_descriptor(terminal_value, label=f"{lane} build terminal receipt")
    log_value = receipt.get("build_log") or receipt.get("stderr")
    if isinstance(log_value, dict):
        verify_descriptor(log_value, label=f"{lane} build log")
    return {
        "receipt": descriptor(receipt_path), "binary": binary,
        "source_root": source["source_root"], "command_argv": argv,
        "profile": profile, "environment_scope": receipt.get("env_policy", receipt.get("environment_scope")),
        "toolchain": receipt.get("toolchain", {"cargo": "+1.95.0"}),
        "locks": locks, "terminal_receipt": terminal_value, "build_log": log_value,
        "source_revision": source["source_revision"], "lane": lane,
    }


def validate_retained_allocator_receipt(receipt_path: Path, equivalence_path: Path) -> dict[str, Any]:
    receipt = read_json(receipt_path)
    terminal = receipt.get("terminal") if isinstance(receipt.get("terminal"), dict) else {}
    exit_code = receipt.get("exit_code", terminal.get("tool_return_exit_code"))
    require(exit_code == 0, "retained allocator build did not record successful completion")
    binary = _receipt_binary(receipt, "retained allocator")
    require(binary["sha256"] == "ad4e6a43c81af7057a961abb5c5893ddde61c7823620a22327f672290d2208cf",
            "retained allocator binary identity changed")
    equivalence = descriptor(equivalence_path)
    return {
        "receipt": descriptor(receipt_path), "binary": binary,
        "source_root": None, "profile": "release", "lane": "allocator",
        "equivalence": equivalence, "source_revision": BASE_REVISION,
        "toolchain": receipt.get("toolchain", {"cargo": "+1.95.0"}),
        "locks": receipt.get("locks", {"perf_baseline": "db62bdd9174484613acae1f49f8187cf105d0b4976e2fea4de91df7372a893d5"}),
    }


def validate_plan(plan_path: Path, correction_path: Path) -> tuple[dict[str, Any], dict[str, Any]]:
    plan = read_json(plan_path)
    correction = read_json(correction_path)
    require(plan.get("routes") == list(ROUTES), "plan route set changed")
    formal = plan.get("formal_paired_cohort")
    require(isinstance(formal, dict), "formal cohort section missing")
    require(formal.get("children") == NORMAL_CHILDREN
            and formal.get("blocks") == BLOCKS
            and formal.get("children_per_route_per_version") == BLOCKS
            and formal.get("measured_samples_per_route_per_version") == BLOCKS * SAMPLES,
            "formal cohort cardinality changed")
    require(formal.get("odd_block_route_order") == list(ROUTES), "odd route order changed")
    require(formal.get("even_block_route_order") == list(reversed(ROUTES)),
            "even route order changed")
    allocator = plan.get("allocator_paired_cohort")
    require(isinstance(allocator, dict) and allocator.get("children") == ALLOCATOR_CHILDREN
            and allocator.get("section_counts") == list(ALLOCATOR_SECTIONS)
            and allocator.get("ownership") == list(ALLOCATOR_OWNERSHIP)
            and allocator.get("repetitions_per_cell_per_version") == ALLOCATOR_REPETITIONS,
            "allocator cohort cardinality changed")
    primary = correction.get("normal_primary", {})
    require(primary.get("statistic")
            == "median of paired baseline_p50/candidate_p50 ratios; ratio is baseline p50 divided by candidate p50, so >1 means candidate speedup",
            "formal ratio orientation is not corrected")
    require(primary.get("median_ratio_minimum") == 1.05
            and primary.get("bootstrap_lower_95_strictly_greater_than") == 1.0
            and primary.get("draws") == 10000 and primary.get("seed") == 930,
            "formal statistical gate changed")
    secondary = correction.get("allocator_secondary", {})
    require(secondary.get("minimum_requested_bytes_reduction_fraction") == 0.20
            and secondary.get("minimum_allocation_calls_reduction_fraction") is None
            and secondary.get("calls_must_strictly_reduce") is True
            and secondary.get("correctness_required") is True,
            "allocator reduction gate changed")
    return plan, correction


def script_bindings(plan_path: Path, correction_path: Path) -> dict[str, dict[str, Any]]:
    values = {
        "runner": Path(__file__).resolve(),
        "validator": COMPARISON_PATH.resolve(),
        "plan": plan_path.resolve(),
        "gate_correction": correction_path.resolve(),
    }
    return {name: descriptor(path) for name, path in values.items()}


def make_manifest(args: argparse.Namespace) -> dict[str, Any]:
    plan, correction = validate_plan(args.plan, args.gate_correction)
    baseline_source = validate_source_receipt(args.baseline_source_receipt, candidate=False)
    candidate_source = validate_source_receipt(args.candidate_source_receipt, candidate=True)
    require(baseline_source["base_revision"] == candidate_source["base_revision"],
            "baseline/candidate source revisions differ")
    for path in COMMON_RUNTIME_PATHS:
        left = baseline_source["source_files"].get(path)
        right = candidate_source["source_files"].get(path)
        require(isinstance(left, dict) and isinstance(right, dict)
                and left.get("bytes") == right.get("bytes")
                and left.get("sha256") == right.get("sha256"),
                f"common harness/source path differs: {path}")
    baseline_normal = validate_build_receipt(args.baseline_build_receipt, baseline_source, lane="normal")
    candidate_normal = validate_build_receipt(args.candidate_build_receipt, candidate_source, lane="normal")
    candidate_allocator = validate_build_receipt(args.candidate_allocator_build_receipt,
                                                  candidate_source, lane="allocator")
    allocator_equivalence = descriptor(args.allocator_equivalence)
    baseline_allocator = validate_retained_allocator_receipt(
        args.baseline_allocator_build_receipt, args.allocator_equivalence
    )
    require(baseline_normal["toolchain"] == candidate_normal["toolchain"],
            "normal baseline/candidate toolchain differs")
    require(candidate_normal["toolchain"] == candidate_allocator["toolchain"],
            "candidate normal/allocator toolchain differs")
    rows = schedule_rows()
    sources = {
        "baseline": {key: value for key, value in baseline_source.items() if key != "source_files"},
        "candidate": {key: value for key, value in candidate_source.items() if key != "source_files"},
    }
    builds = {
        "normal": {"baseline": baseline_normal, "candidate": candidate_normal},
        "allocator": {"baseline": baseline_allocator, "candidate": candidate_allocator},
    }
    custody = {
        "baseline_source_receipt": baseline_source["receipt"],
        "candidate_source_receipt": candidate_source["receipt"],
        "baseline_normal_build_receipt": baseline_normal["receipt"],
        "candidate_normal_build_receipt": candidate_normal["receipt"],
        "baseline_allocator_build_receipt": baseline_allocator["receipt"],
        "candidate_allocator_build_receipt": candidate_allocator["receipt"],
        "allocator_source_equivalence": allocator_equivalence,
        "plan": descriptor(args.plan),
        "gate_correction": descriptor(args.gate_correction),
    }
    return {
        "schema": "litchi.performance.0933.formal-comparison-manifest.v2",
        "created_utc": utc_now(), "capture_authorized": False, "formal_children_run": 0,
        "base_revision": BASE_REVISION,
        "claim_scope": "prospective 48-child paired normal lane plus 24-child allocator lane; no claim until all 72 terminal records, semantic validation, and moved-packet replay pass",
        "failure_policy": "retain every terminal child; any failed report/control invalidates its complete cohort; never retry or selectively recapture",
        "source": sources, "build": builds, "custody_sources": custody,
        "scripts": script_bindings(args.plan, args.gate_correction),
        "schedule": {
            "blocks": BLOCKS, "normal_children": NORMAL_CHILDREN,
            "allocator_children": ALLOCATOR_CHILDREN, "children": CHILDREN,
            "routes": list(ROUTES), "children_per_route_per_version": BLOCKS,
            "warmups": WARMUPS, "samples_per_child": SAMPLES, "cpu": CPU,
            "normal_deadline_seconds": int(NORMAL_DEADLINE_SECONDS),
            "allocator_deadline_seconds": int(ALLOCATOR_DEADLINE_SECONDS),
            "rows": rows, "sha256": schedule_hash(rows),
        },
        "argv_template": {
            "normal": ["/usr/bin/time", "-v", "taskset", "-c", str(CPU), "<binary>",
                       "--case", CASE, "--samples", str(SAMPLES), "--warmup", str(WARMUPS),
                       "--json", "<report.json>"],
            "allocator": ["/usr/bin/time", "-v", "taskset", "-c", str(CPU),
                          "<allocator-binary>", "docx-section-retention", "--sections", "<16|8192>",
                          "--concurrent-commits", "8", "--policy", "exact",
                          "--ownership", "<shared_snapshot_clones|independent_sources>",
                          "--source-revision", "<40-hex>", "--output", "<report.json>"],
        },
        "route_selector": {"name": "LITCHI_DOCX_INGRESS_ROUTE", "values": list(ROUTES)},
        "child_environment": dict(CHILD_ENV_BASE),
        "corpus": {
            "archive_bytes": 2_101_427,
            "archive_sha256": "183204a832a3518bf1d4474e9b7ff347ad7d0dc95ad2e9b1dbdde405837a29e8",
            "expected_output_bytes": 2_101_446,
            "expected_output_sha256": "5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec",
        },
        "formal_gate": {
            "normal_primary_route": "generic_read_at",
            "normal_ratio": "baseline_p50 / candidate_p50",
            "normal_median_ratio_minimum": 1.05,
            "normal_bootstrap_lower_95_strictly_greater_than": 1.0,
            "normal_bootstrap_draws": 10000, "normal_bootstrap_seed": 930,
            "allocator_cells": [
                {"sections": sections, "ownership": ownership}
                for sections in ALLOCATOR_SECTIONS for ownership in ALLOCATOR_OWNERSHIP
            ],
            "allocator_min_requested_bytes_reduction_fraction": 0.20,
            "allocator_calls_must_strictly_reduce": True,
            "allocator_correctness_required": True,
            "allocator_latency_safety_candidate_over_baseline_p50_maximum": 1.05,
            "review_flags": {"tails_or_rss_increase": ">5%", "semantic_or_counter_failure": "invalidates cohort"},
            "source_plan": descriptor(args.plan), "correction": descriptor(args.gate_correction),
        },
        "status": "prospective_unfinished",
        "raw_evidence": "external packet only; terminal receipts contain descriptors and scalar validation summaries, never duplicated raw vectors",
    }


def freeze(args: argparse.Namespace) -> dict[str, Any]:
    manifest = make_manifest(args)
    output = args.manifest.resolve()
    require(not output.exists(), f"refusing to replace manifest: {output}")
    write_new_json(output, manifest, limit=MAX_ARTIFACT_BYTES)
    return {"status": manifest["status"], "manifest": str(output),
            "bytes": output.stat().st_size, "sha256": sha256_path(output),
            "children": CHILDREN, "capture_authorized": False}


def verify_scripts(manifest: dict[str, Any], packet: Path | None = None) -> None:
    scripts = manifest.get("scripts")
    require(isinstance(scripts, dict), "manifest scripts missing")
    local = manifest.get("packet_scripts") if packet is not None else None
    for name, value in scripts.items():
        require(isinstance(value, dict), f"script descriptor malformed: {name}")
        if packet is None:
            verify_descriptor(value, label=f"manifest script {name}")
        else:
            require(isinstance(local, dict) and name in local, f"packet script missing: {name}")
            local_value = local[name]
            path = verify_descriptor(local_value, label=f"packet script {name}", root=packet)
            require(local_value.get("bytes") == value.get("bytes")
                    and local_value.get("sha256") == value.get("sha256"),
                    f"packet script differs from manifest: {name}")
            require(path.name == Path(str(value["path"])).name, f"packet script basename changed: {name}")


def verify_manifest(path: Path, *, authorized: bool = False,
                    allow_authorized: bool = False, packet: Path | None = None) -> dict[str, Any]:
    manifest_path = path.resolve()
    manifest = read_json(manifest_path)
    require(manifest.get("schema") == "litchi.performance.0933.formal-comparison-manifest.v2",
            "manifest schema changed")
    require(manifest.get("base_revision") == BASE_REVISION, "manifest base revision changed")
    require(manifest.get("schedule", {}).get("sha256") == schedule_hash(schedule_rows()),
            "manifest schedule hash changed")
    require(manifest.get("schedule", {}).get("rows") == schedule_rows(), "manifest schedule changed")
    require(manifest.get("schedule", {}).get("children") == CHILDREN
            and manifest.get("schedule", {}).get("normal_children") == NORMAL_CHILDREN
            and manifest.get("schedule", {}).get("allocator_children") == ALLOCATOR_CHILDREN,
            "manifest child count changed")
    require(manifest.get("child_environment") == CHILD_ENV_BASE,
            "child environment allowlist changed")
    if authorized:
        require(manifest.get("capture_authorized") is True,
                "capture authorization state changed")
    elif not allow_authorized:
        require(manifest.get("capture_authorized") is False,
                "unauthorized freeze manifest was changed")
    if authorized:
        require(os.environ.get("LITCHI_PERF_0933_CAPTURE_APPROVED") == "1",
                "LITCHI_PERF_0933_CAPTURE_APPROVED=1 is required")
    verify_scripts(manifest, packet)
    def manifest_ref(value: dict[str, Any], label: str) -> Path:
        raw = value.get("path")
        require(isinstance(raw, str), f"{label}: path missing")
        path = Path(raw)
        return path.resolve() if path.is_absolute() else (manifest_path.parent / safe_relative(raw, label)).resolve()

    validate_plan(manifest_ref(manifest["formal_gate"]["source_plan"], "source plan"),
                  manifest_ref(manifest["formal_gate"]["correction"], "gate correction"))
    source_values: dict[str, dict[str, Any]] = {}
    for version in ("baseline", "candidate"):
        source = manifest["source"][version]
        # Source/build/binary are rehashed at capture start and at replay.
        source_now = validate_source_receipt(
            manifest_ref(source["receipt"], f"{version} source receipt"), candidate=(version == "candidate")
        )
        require(source_now["source_root"] == source["source_root"],
                f"{version} source root changed")
        source_values[version] = source_now
    for path in COMMON_RUNTIME_PATHS:
        left = source_values["baseline"]["source_files"].get(path)
        right = source_values["candidate"]["source_files"].get(path)
        require(isinstance(left, dict) and isinstance(right, dict)
                and left.get("bytes") == right.get("bytes")
                and left.get("sha256") == right.get("sha256"),
                f"common harness/source path differs at capture: {path}")
    for version in ("baseline", "candidate"):
        build = manifest["build"]["normal"][version]
        build_now = validate_build_receipt(manifest_ref(build["receipt"], f"normal {version} build receipt"),
                                           source_values[version], lane="normal")
        require(build_now["binary"] == build["binary"],
                f"normal {version} binary binding changed")
    allocator_candidate = manifest["build"]["allocator"]["candidate"]
    candidate_alloc_now = validate_build_receipt(
        manifest_ref(allocator_candidate["receipt"], "allocator candidate build receipt"),
        source_values["candidate"], lane="allocator")
    require(candidate_alloc_now["binary"] == allocator_candidate["binary"],
            "allocator candidate binary binding changed")
    allocator_baseline = manifest["build"]["allocator"]["baseline"]
    for label, value in (("retained allocator receipt", allocator_baseline["receipt"]),
                         ("allocator source equivalence", allocator_baseline["equivalence"]),
                         ("retained allocator binary", allocator_baseline["binary"])):
        actual_path = manifest_ref(value, label)
        actual = descriptor(actual_path)
        require(actual["bytes"] == value.get("bytes") and actual["sha256"] == value.get("sha256"),
                f"{label}: descriptor changed")
    return manifest


def rss_bytes(time_path: Path) -> int | None:
    text = time_path.read_text(encoding="utf-8", errors="replace")
    match = re.search(r"Maximum resident set size \(kbytes\):\s*(\d+)", text)
    return int(match.group(1)) * 1024 if match else None


def nearest_percentile(values: list[float], fraction: float) -> float:
    ordered = sorted(values)
    index = min(max(math.ceil(fraction * len(ordered)) - 1, 0), len(ordered) - 1)
    return float(ordered[index])


def paired_gate(pair_values: dict[str, dict[str, dict[str, float | int]]]) -> dict[str, Any]:
    """Compute the frozen baseline/candidate paired p50 gate from child summaries."""
    per_route: dict[str, Any] = {}
    for route in ROUTES:
        pairs = [value for (pair_route, _), value in pair_values.items() if pair_route == route]
        require(len(pairs) == BLOCKS, f"paired gate pair count changed: {route}")
        require(all(set(value) == {"baseline", "candidate"} for value in pairs),
                f"paired gate version set changed: {route}")
        ratios = [float(value["baseline"]["p50"]) / float(value["candidate"]["p50"])
                  for value in pairs]
        randomizer = random.Random(930)
        bootstraps = [statistics.median(randomizer.choice(ratios) for _ in ratios)
                      for _ in range(10000)]
        per_route[route] = {
            "pairs": len(pairs),
            "baseline_p50_ns": [value["baseline"]["p50"] for value in pairs],
            "candidate_p50_ns": [value["candidate"]["p50"] for value in pairs],
            "paired_baseline_over_candidate_p50": ratios,
            "median_ratio": float(statistics.median(ratios)),
            "bootstrap_lower_95": nearest_percentile(bootstraps, 0.025),
            "bootstrap_upper_95": nearest_percentile(bootstraps, 0.975),
            "bootstrap_seed": 930,
            "bootstrap_draws": 10000,
        }
    primary = per_route["generic_read_at"]
    return {
        "schema": "litchi.performance.0933.paired-gate.v1",
        "ratio": "baseline_p50 / candidate_p50",
        "primary_route": "generic_read_at",
        "median_ratio_minimum": 1.05,
        "bootstrap_lower_95_strictly_greater_than": 1.0,
        "primary_pass": primary["median_ratio"] >= 1.05 and primary["bootstrap_lower_95"] > 1.0,
        "per_route": per_route,
        "claim": "descriptive paired gate only; adoption requires independent review and semantic controls",
    }


def source_revision_for(manifest: dict[str, Any], version: str) -> str:
    value = manifest["source"][version].get("source_revision")
    require(isinstance(value, str) and len(value) == 40
            and all(char in "0123456789abcdefABCDEF" for char in value),
            f"{version} source revision is not a 40-hex identifier")
    return value


def binary_for_row(manifest: dict[str, Any], row: dict[str, Any]) -> dict[str, Any]:
    lane = row["lane"]
    require(lane in manifest["build"], f"build lane missing: {lane}")
    return manifest["build"][lane][row["version"]]["binary"]


def child_environment(row: dict[str, Any], manifest: dict[str, Any]) -> dict[str, str]:
    environment = dict(manifest["child_environment"])
    if row["lane"] == "normal":
        environment["LITCHI_DOCX_INGRESS_ROUTE"] = row["route"]
    return environment


def child_command(row: dict[str, Any], binary: Path, report: Path,
                  source_revision: str) -> list[str]:
    prefix = ["/usr/bin/time", "-v", "taskset", "-c", str(CPU), str(binary)]
    if row["lane"] == "normal":
        return prefix + ["--case", CASE, "--samples", str(SAMPLES), "--warmup", str(WARMUPS),
                         "--json", str(report)]
    command = prefix + [
        "docx-section-retention", "--sections", str(row["sections"]),
        "--concurrent-commits", str(row["concurrent_commits"]),
        "--policy", row["policy"], "--ownership", row["ownership"],
        "--source-revision", source_revision, "--output", str(report),
    ]
    revision_index = command.index("--source-revision")
    require(command[revision_index + 1] == source_revision,
            "allocator command lost the bound source revision")
    return command


def packet_path(packet: Path, child_id: str, name: str) -> Path:
    require(child_id and all(ch.isalnum() or ch in "-_" for ch in child_id),
            f"unsafe child id: {child_id!r}")
    return packet / "children" / child_id / name


def process_group_alive(pgid: int) -> bool:
    try:
        os.killpg(pgid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


def close_process_group(pgid: int, *, force: bool = False) -> tuple[bool, int]:
    """Poll the process group until absent; terminate leftovers once."""
    probes = 0
    deadline = time.monotonic() + (5.0 if force else 2.0)
    while process_group_alive(pgid) and time.monotonic() < deadline:
        probes += 1
        time.sleep(0.05)
    if not process_group_alive(pgid):
        return True, probes + 1
    probes += 1
    try:
        os.killpg(pgid, signal.SIGTERM)
    except ProcessLookupError:
        return True, probes
    deadline = time.monotonic() + 2.0
    while process_group_alive(pgid) and time.monotonic() < deadline:
        probes += 1
        time.sleep(0.05)
    if process_group_alive(pgid):
        probes += 1
        try:
            os.killpg(pgid, signal.SIGKILL)
        except ProcessLookupError:
            return True, probes
        deadline = time.monotonic() + 2.0
        while process_group_alive(pgid) and time.monotonic() < deadline:
            probes += 1
            time.sleep(0.05)
    return not process_group_alive(pgid), probes


def scalar_validation(report: dict[str, Any], validated: dict[str, Any], route: str,
                      report_desc: dict[str, Any]) -> dict[str, Any]:
    output = report["results"][0].get("output_sha256")
    phases = validated["phase_ns"]
    return {
        "schema": "litchi.performance.0933.scalar-validation.v1",
        "status": "passed",
        "route": route,
        "samples": SAMPLES,
        "warmups": WARMUPS,
        "report": report_desc,
        "output_sha256": output,
        "phase_sample_counts": {name: len(values) for name, values in phases.items()},
    }


def _u64(value: Any, label: str) -> int:
    require(type(value) is int and value >= 0, f"{label} must be a non-negative integer")
    return value


def _allocator_gauge(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} gauge missing")
    for field in ("allocation_calls", "deallocation_calls", "reallocation_calls",
                  "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
                  "live_bytes", "peak_live_bytes"):
        _u64(value.get(field), f"{label}.{field}")
    require(value.get("overflowed") is False, f"{label}.overflowed is true")
    require(value.get("observer_invalid") is False, f"{label}.observer_invalid is true")
    require(value["live_bytes"] == value["allocated_bytes"] - value["deallocated_bytes"],
            f"{label} live-byte balance changed")
    require(value["peak_live_bytes"] >= value["live_bytes"],
            f"{label} peak does not cover live bytes")
    return value


def validate_allocator_report(report: dict[str, Any], row: dict[str, Any],
                              report_desc: dict[str, Any], source_revision: str) -> dict[str, Any]:
    require(report.get("schema") == "docx_section_retention_v1", "allocator report schema changed")
    require(report.get("section_count") == row["sections"], "allocator section count changed")
    require(report.get("concurrent_commits") == 8, "allocator commit count changed")
    require(report.get("retention_policy") == "exact", "allocator policy changed")
    require(report.get("ownership_mode") == row["ownership"], "allocator ownership changed")
    require(report.get("source_revision") == source_revision, "allocator source revision changed")
    require(type(report.get("source_archive_bytes")) is int and report["source_archive_bytes"] > 0,
            "allocator source archive size missing")
    require(isinstance(report.get("source_archive_sha256"), str)
            and len(report["source_archive_sha256"]) == 64, "allocator source archive hash missing")
    for field in ("expected_output_sha256", "expected_output_bytes"):
        values = report.get(field)
        require(isinstance(values, list) and len(values) == 8,
                f"allocator {field} vector changed")
    admission = report.get("retention_admission")
    require(isinstance(admission, dict) and admission.get("boundary_checks_verified") is True,
            "allocator cap boundary checks failed")
    required = _u64(admission.get("required_retained_base_bytes"), "required retained base bytes")
    require(_u64(admission.get("zero_retained_base_bytes"), "zero retained base bytes") == 0,
            "zero cap retained bytes changed")
    require(_u64(admission.get("one_under_retained_base_bytes"), "one-under retained base bytes") == 0,
            "one-under cap retained bytes changed")
    require(_u64(admission.get("exact_retained_base_bytes"), "exact retained base bytes") == required,
            "exact cap retained bytes changed")
    scenarios = report.get("scenarios")
    require(isinstance(scenarios, list) and len(scenarios) == 1, "allocator scenario count changed")
    scenario = scenarios[0]
    require(scenario.get("ownership_mode") == row["ownership"]
            and scenario.get("section_count") == row["sections"]
            and scenario.get("concurrent_commits") == 8,
            "allocator scenario identity changed")
    pair_count = 8 * 7 // 2
    require(scenario.get("source_inventory_pair_count") == pair_count,
            "allocator source pair count changed")
    expected_shared = pair_count if row["ownership"] == "shared_snapshot_clones" else 0
    require(scenario.get("source_inventory_shared_pairs") == expected_shared,
            "allocator source sharing count changed")
    require(scenario.get("source_inventory_sharing_verified") is True,
            "allocator source sharing verification failed")
    require(scenario.get("measured_commit_shape_verified") is True
            and scenario.get("measured_output_verified") is True
            and scenario.get("measured_inverse_verified") is True
            and scenario.get("preflight_output_verified") is True
            and scenario.get("preflight_inverse_verified") is True,
            "allocator correctness flags changed")
    expected_validations = 1 if row["ownership"] == "shared_snapshot_clones" else 8
    require(scenario.get("measured_commit_validation_count") == expected_validations,
            "allocator measured validation count changed")
    retained = scenario.get("retained_base_bytes_each")
    require(isinstance(retained, list) and len(retained) == 8
            and all(value == required for value in retained),
            "allocator retained-byte vector changed")
    gauges = scenario.get("gauges")
    require(isinstance(gauges, dict), "allocator gauges missing")
    names = ("baseline", "after_snapshot_drop", "commits_retained",
             "after_release_all_but_last", "after_release_last", "after_commit_drop")
    absolute = {}
    for name in names:
        value = gauges[name] if name == "baseline" else gauges[name].get("absolute")
        absolute[name] = _allocator_gauge(value, f"gauges.{name}")
    for before, after in zip(names, names[1:]):
        left, right = absolute[before], absolute[after]
        for field in ("allocation_calls", "deallocation_calls", "reallocation_calls",
                      "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
                      "peak_live_bytes"):
            require(right[field] >= left[field], f"allocator {field} regressed at {after}")
        expected_live = left["live_bytes"] + (right["allocated_bytes"] - left["allocated_bytes"]) \
            - (right["deallocated_bytes"] - left["deallocated_bytes"])
        require(right["live_bytes"] == expected_live, f"allocator live transition changed at {after}")
    require(gauges["commits_retained"]["absolute"].get("region_peak_live_bytes") is None,
            "held checkpoint unexpectedly claims a region peak")
    for name in ("after_snapshot_drop", "after_release_all_but_last", "after_release_last", "after_commit_drop"):
        peak = gauges[name]["absolute"].get("region_peak_live_bytes")
        require(type(peak) is int and peak >= absolute[name]["live_bytes"],
                f"allocator region peak invalid at {name}")
    require(gauges.get("returned_to_baseline") is True
            and absolute["after_commit_drop"]["live_bytes"] == absolute["baseline"]["live_bytes"],
            "allocator did not return to baseline")
    after = absolute["after_snapshot_drop"]
    baseline = absolute["baseline"]
    return {
        "schema": "litchi.performance.0933.allocator-validation.v1",
        "status": "passed",
        "report": report_desc,
        "sections": row["sections"], "ownership": row["ownership"], "repetition": row["repetition"],
        "source_revision": source_revision,
        "after_snapshot_drop": {
            "allocated_bytes_delta": after["allocated_bytes"] - baseline["allocated_bytes"],
            "allocation_calls_delta": after["allocation_calls"] - baseline["allocation_calls"],
            "region_peak_live_bytes": gauges["after_snapshot_drop"]["absolute"]["region_peak_live_bytes"],
            "absolute_allocated_bytes": after["allocated_bytes"],
            "absolute_allocation_calls": after["allocation_calls"],
        },
        "retained_base_bytes_each": retained,
        "source_inventory_shared_pairs": scenario["source_inventory_shared_pairs"],
        "measured_commit_validation_count": scenario["measured_commit_validation_count"],
    }


def run_child(row: dict[str, Any], manifest: dict[str, Any], packet: Path,
              comparison: Any, source_root: Path, binary: dict[str, Any]) -> dict[str, Any]:
    child = packet_path(packet, row["id"], "report.json").parent
    child.mkdir(parents=True, exist_ok=False)
    report_path = child / "report.json"
    stdout_path = child / "stdout.log"
    stderr_path = child / "stderr.log"
    time_path = child / "time.log"
    source_revision = source_revision_for(manifest, row["version"])
    command = child_command(row, Path(binary["path"]), report_path, source_revision)
    environment = child_environment(row, manifest)
    started = utc_now()
    monotonic_start = time.monotonic()
    deadline_seconds = float(row["deadline_seconds"])
    exit_code: int | None = None
    observation_timeouts = 0
    deadline_exceeded = False
    runner_error: str | None = None
    group_closed = False
    group_probe_count = 0
    process_group_id: int | None = None
    with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
        try:
            process = subprocess.Popen(command, cwd=str(source_root), env=dict(environment),
                                       stdout=stdout, stderr=stderr, start_new_session=True)
            process_group_id = process.pid
            while True:
                try:
                    exit_code = process.wait(timeout=OBSERVE_SECONDS)
                    break
                except subprocess.TimeoutExpired:
                    observation_timeouts += 1
                    if time.monotonic() - monotonic_start <= deadline_seconds:
                        continue
                    deadline_exceeded = True
                    try:
                        os.killpg(process.pid, signal.SIGTERM)
                    except ProcessLookupError:
                        pass
                    try:
                        exit_code = process.wait(timeout=10)
                    except subprocess.TimeoutExpired:
                        try:
                            os.killpg(process.pid, signal.SIGKILL)
                        except ProcessLookupError:
                            pass
                        exit_code = process.wait(timeout=10)
                    break
        except Exception as error:
            runner_error = f"{type(error).__name__}: {error}"
    ended = utc_now()
    if process_group_id is not None:
        group_closed, group_probe_count = close_process_group(
            process_group_id, force=deadline_exceeded or runner_error is not None
        )
    with time_path.open("xb") as stream:
        stream.write(stderr_path.read_bytes())
    report_desc = descriptor(report_path, root=packet) if report_path.is_file() else None
    terminal = {
        "schema": "litchi.performance.0933.formal-terminal.v2",
        "status": "terminal", "schedule": row, "argv": command,
        "environment": environment, "cwd": str(source_root),
        "started_utc": started, "ended_utc": ended, "exit_code": exit_code,
        "observation_timeout_count": observation_timeouts,
        "deadline_seconds": deadline_seconds, "deadline_exceeded": deadline_exceeded,
        "runner_error": runner_error, "process_group_id": process_group_id,
        "process_group_closed": group_closed, "process_group_probe_count": group_probe_count,
        "stdout": descriptor(stdout_path, root=packet),
        "stderr": descriptor(stderr_path, root=packet),
        "time": descriptor(time_path, root=packet),
        "rss_bytes": rss_bytes(time_path), "report": report_desc, "binary": binary,
        "source_revision": source_revision,
        "source_receipt_sha256": manifest["source"][row["version"]]["receipt"]["sha256"],
        "runner_sha256": manifest["scripts"]["runner"]["sha256"],
        "manifest_sha256": sha256_path(packet / "manifest.json"),
    }
    terminal_path = child / "terminal.json"
    write_new_json(terminal_path, terminal, limit=MAX_ARTIFACT_BYTES)
    valid = (exit_code == 0 and not deadline_exceeded and runner_error is None
             and group_closed and terminal["rss_bytes"] is not None
             and report_desc is not None and report_desc["bytes"] < MAX_ARTIFACT_BYTES)
    validation: dict[str, Any]
    if valid:
        try:
            report = read_json(report_path)
            if row["lane"] == "normal":
                result = comparison.validate_report(report, row["route"], warmups=WARMUPS,
                                                    samples=SAMPLES, label=str(report_path),
                                                    binary=binary)
                validation = scalar_validation(report, result, row["route"], report_desc)
            else:
                validation = validate_allocator_report(report, row, report_desc, source_revision)
        except Exception as error:
            validation = {"schema": "litchi.performance.0933.validation.v2", "status": "failed",
                          "error": f"{type(error).__name__}: {error}", "report": report_desc}
            valid = False
    else:
        validation = {"schema": "litchi.performance.0933.validation.v2", "status": "failed",
                      "error": "terminal process failed, RSS missing, or report missing",
                      "report": report_desc}
    validation_path = child / "validation.json"
    write_new_json(validation_path, validation, limit=MAX_ARTIFACT_BYTES)
    return {
        "id": row["id"], "pair_id": row["pair_id"], "lane": row["lane"],
        "block": row["block"], "position": row["position"], "route": row["route"],
        "version": row["version"], "status": "passed" if valid else "failed",
        "exit_code": exit_code, "deadline_exceeded": deadline_exceeded,
        "rss_bytes": terminal["rss_bytes"], "terminal": descriptor(terminal_path, root=packet),
        "report": report_desc, "validation": descriptor(validation_path, root=packet),
        "validation_status": validation.get("status"),
    }


def capture(args: argparse.Namespace) -> dict[str, Any]:
    manifest_path = args.manifest.resolve()
    manifest = verify_manifest(manifest_path, authorized=True)
    packet = args.packet.resolve()
    require(not packet.exists(), f"refusing to replace packet: {packet}")
    packet.mkdir(parents=True)
    (packet / "children").mkdir()
    (packet / "scripts").mkdir()
    for name, value in manifest["scripts"].items():
        source = Path(str(value["path"]))
        target = packet / "scripts" / source.name
        with target.open("xb") as stream:
            stream.write(source.read_bytes())
    (packet / "custody").mkdir()
    packet_custody: dict[str, dict[str, Any]] = {}
    for name, value in manifest["custody_sources"].items():
        source = Path(str(value["path"]))
        target = packet / "custody" / f"{name}.json"
        with target.open("xb") as stream:
            stream.write(source.read_bytes())
        copied = descriptor(target, root=packet)
        require(copied["bytes"] == value["bytes"] and copied["sha256"] == value["sha256"],
                f"packet custody copy changed: {name}")
        packet_custody[name] = copied
    packet_manifest = dict(manifest)
    packet_manifest["packet_scripts"] = {
        name: descriptor(packet / "scripts" / Path(str(value["path"])).name, root=packet)
        for name, value in manifest["scripts"].items()
    }
    packet_manifest["packet_custody"] = packet_custody
    # Replay resolves the copied receipts and plans locally.  The copied
    # receipts retain their external source/listing/binary descriptors, whose
    # hashes are rechecked by the same fail-closed validators.
    packet_manifest["source"] = {
        version: dict(value, receipt=packet_custody[f"{version}_source_receipt"])
        for version, value in manifest["source"].items()
    }
    packet_manifest["build"] = {
        lane: {
            version: dict(value, receipt=packet_custody[f"{version}_{lane}_build_receipt"])
            for version, value in versions.items()
        }
        for lane, versions in manifest["build"].items()
    }
    packet_manifest["build"]["allocator"]["baseline"]["equivalence"] = packet_custody[
        "allocator_source_equivalence"
    ]
    packet_manifest["formal_gate"] = dict(
        manifest["formal_gate"],
        source_plan=packet_custody["plan"],
        correction=packet_custody["gate_correction"],
    )
    write_new_json(packet / "manifest.json", packet_manifest, limit=MAX_ARTIFACT_BYTES)
    comparison = load_comparison(packet / "scripts" / Path(str(manifest["scripts"]["validator"]["path"])).name)
    results: list[dict[str, Any]] = []
    started = utc_now()
    for row in manifest["schedule"]["rows"]:
        version = row["version"]
        source_root = Path(manifest["source"][version]["source_root"])
        binary = binary_for_row(manifest, row)
        results.append(run_child(row, manifest, packet, comparison, source_root, binary))
        # Keep progress scalar and durable; raw vectors remain only in reports.
        latest = results[-1]
        write_new_json(packet / f"capture-state-{len(results):03d}.json", {
            "schema": "litchi.performance.0933.formal-state-checkpoint.v2",
            "status": "running", "started_utc": started,
            "children_run": len(results), "children_expected": CHILDREN,
            "last_child": {key: latest.get(key) for key in
                           ("id", "lane", "pair_id", "version", "status", "exit_code",
                            "deadline_exceeded", "rss_bytes")},
            "passed": sum(item["status"] == "passed" for item in results),
            "failed": sum(item["status"] != "passed" for item in results),
            "recapture_count": 0,
        }, limit=MAX_ARTIFACT_BYTES)
    passed = len(results) == CHILDREN and all(row["status"] == "passed" for row in results)
    state = {
        "schema": "litchi.performance.0933.formal-state.v2",
        "status": "completed" if passed else "invalid",
        "started_utc": started, "ended_utc": utc_now(),
        "children_run": len(results), "children_expected": CHILDREN,
        "children": results, "recapture_count": 0,
        "failure_policy": "invalid cohort retained; no retry or selective recapture",
    }
    write_new_json(packet / "capture-state.json", state, limit=MAX_ARTIFACT_BYTES)
    return {"status": state["status"], "packet": str(packet), "children_run": len(results),
            "passed": sum(row["status"] == "passed" for row in results),
            "failed": sum(row["status"] != "passed" for row in results)}


def packet_file(packet: Path, value: dict[str, Any], label: str) -> Path:
    return verify_descriptor(value, label=label, root=packet)


def verify_packet_custody(manifest: dict[str, Any], packet: Path) -> None:
    sources = manifest.get("custody_sources")
    copies = manifest.get("packet_custody")
    require(isinstance(sources, dict) and isinstance(copies, dict),
            "copied packet custody descriptors are missing")
    require(set(sources) == set(copies), "copied packet custody set changed")
    for name, original in sources.items():
        local = copies[name]
        path = packet_file(packet, local, f"packet custody {name}")
        require(local.get("bytes") == original.get("bytes")
                and local.get("sha256") == original.get("sha256"),
                f"packet custody descriptor changed: {name}")
        # These are copied receipts/manifests, so parse each local JSON before
        # consulting any external path in the semantic revalidation below.
        read_json(path)


def replay_packet(packet: Path, *, manifest_path: Path | None = None) -> dict[str, Any]:
    packet = packet.resolve()
    manifest_file = manifest_path.resolve() if manifest_path else packet / "manifest.json"
    manifest = read_json(manifest_file)
    require(manifest.get("schema") == "litchi.performance.0933.formal-comparison-manifest.v2",
            "copied packet manifest schema changed")
    verify_scripts(manifest, packet if manifest_path is None else None)
    if manifest_path is None:
        verify_packet_custody(manifest, packet)
    verify_manifest(manifest_file, allow_authorized=True,
                    packet=packet if manifest_path is None else None)
    validator_path = (packet / safe_relative(
        manifest["packet_scripts"]["validator"]["path"], "packet validator"
    )) if manifest_path is None else Path(manifest["scripts"]["validator"]["path"])
    comparison = load_comparison(validator_path)
    rows = manifest["schedule"]["rows"]
    state = read_json(packet / "capture-state.json")
    require(state.get("status") in ("completed", "invalid"), "packet is not terminal")
    children = state.get("children")
    require(isinstance(children, list) and len(children) == CHILDREN,
            "packet child count is incomplete")
    by_id = {row.get("id"): row for row in children}
    require(set(by_id) == {row["id"] for row in rows},
            "packet child IDs differ from frozen schedule")
    reduced_rows: dict[tuple[str, str], list[dict[str, Any]]] = {}
    pair_values: dict[tuple[str, str], dict[str, dict[str, float | int]]] = {}
    allocator_values: dict[tuple[int, str, str], dict[str, list[dict[str, Any]]]] = {}
    for row in rows:
        child = by_id[row["id"]]
        require(child.get("status") == "passed", f"child is not a passed terminal: {row['id']}")
        terminal_path = packet_file(packet, child["terminal"], f"terminal {row['id']}")
        terminal = read_json(terminal_path)
        require(terminal.get("schedule") == row, f"terminal schedule changed: {row['id']}")
        require(terminal.get("exit_code") == 0 and terminal.get("deadline_exceeded") is False
                and terminal.get("runner_error") is None and terminal.get("process_group_closed") is True,
                f"terminal status failed: {row['id']}")
        expected_environment = child_environment(row, manifest)
        require(terminal.get("environment") == expected_environment,
                f"child environment changed: {row['id']}")
        require(terminal.get("manifest_sha256") == sha256_path(packet / "manifest.json"),
                f"terminal manifest binding changed: {row['id']}")
        require(terminal.get("runner_sha256") == manifest["scripts"]["runner"]["sha256"],
                f"terminal runner binding changed: {row['id']}")
        require(terminal.get("source_receipt_sha256")
                == manifest["source"][row["version"]]["receipt"]["sha256"],
                f"terminal source binding changed: {row['id']}")
        for name in ("stdout", "stderr", "time"):
            packet_file(packet, terminal[name], f"{row['id']} {name}")
        report_path = packet_file(packet, terminal["report"], f"report {row['id']}")
        require(terminal["report"] == child.get("report"), f"report descriptor changed: {row['id']}")
        binary = binary_for_row(manifest, row)
        require(terminal.get("binary") == binary, f"terminal binary binding changed: {row['id']}")
        report = read_json(report_path)
        source_revision = source_revision_for(manifest, row["version"])
        if row["lane"] == "normal":
            validated = comparison.validate_report(report, row["route"], warmups=WARMUPS,
                                                   samples=SAMPLES, label=str(report_path), binary=binary)
        else:
            validated = validate_allocator_report(report, row, child["report"], source_revision)
        validation_path = packet_file(packet, child["validation"], f"validation {row['id']}")
        validation = read_json(validation_path)
        require(validation.get("status") == "passed" and validation.get("report") == child["report"],
                f"validation changed: {row['id']}")
        if row["lane"] == "normal":
            reduced_rows.setdefault((row["version"], row["route"]), []).append(validated)
            elapsed = list(validated["elapsed_ns"])
            pair_values.setdefault((row["route"], row["pair_id"]), {})[row["version"]] = {
                "p50": float(statistics.median(elapsed)),
                "p95": nearest_percentile([float(value) for value in elapsed], 0.95),
                "p99": nearest_percentile([float(value) for value in elapsed], 0.99),
                "rss": int(child["rss_bytes"]),
            }
        else:
            key = (row["sections"], row["ownership"], row["pair_id"])
            allocator_values.setdefault(key, {}).setdefault(row["version"], []).append({
                "metrics": validated["after_snapshot_drop"], "rss_bytes": int(child["rss_bytes"]),
                "report": child["report"], "repetition": row["repetition"],
            })
    summary: dict[str, Any] = {
        "schema": "litchi.performance.0933.formal-replay.v2", "status": "passed",
        "packet": str(packet), "children": CHILDREN,
        "source": "all reports, terminal streams, and scalar validations rehashed from packet",
        "normal": {"per_version_route": {}}, "allocator": {"cells": {}},
    }
    for key, values in reduced_rows.items():
        version, route = key
        elapsed = [value for item in values for value in item["elapsed_ns"]]
        phases = {phase: [value for item in values for value in item["phase_ns"][phase]]
                  for phase in PHASES}
        summary["normal"]["per_version_route"][f"{version}/{route}"] = {
            "samples": len(elapsed), "children": len(values),
            "elapsed_ns": comparison._statistics(elapsed),
            "phase_statistics_ns": {phase: comparison._statistics(vector)
                                     for phase, vector in phases.items()},
            "rss_bytes": [int(child["rss_bytes"]) for child in children
                           if child.get("lane") == "normal" and child["version"] == version
                           and child["route"] == route],
        }
    summary["normal"]["paired_gate"] = paired_gate(pair_values)
    for (sections, ownership, pair_id), versions in allocator_values.items():
        require(set(versions) == {"baseline", "candidate"}, f"allocator pair incomplete: {pair_id}")
        require(len(versions["baseline"]) == len(versions["candidate"]) == 1,
                f"allocator pair duplicate: {pair_id}")
        base = versions["baseline"][0]; cand = versions["candidate"][0]
        bm, cm = base["metrics"], cand["metrics"]
        require(bm["allocated_bytes_delta"] >= 0 and cm["allocated_bytes_delta"] >= 0,
                f"allocator byte delta negative: {pair_id}")
        cell = summary["allocator"]["cells"].setdefault(f"sections={sections}/{ownership}", {
            "sections": sections, "ownership": ownership, "pairs": [],
        })
        cell["pairs"].append({
            "pair_id": pair_id,
            "baseline_requested_bytes": bm["allocated_bytes_delta"],
            "candidate_requested_bytes": cm["allocated_bytes_delta"],
            "requested_bytes_reduction_fraction": (1.0 - cm["allocated_bytes_delta"] / bm["allocated_bytes_delta"])
                if bm["allocated_bytes_delta"] else None,
            "baseline_allocation_calls": bm["allocation_calls_delta"],
            "candidate_allocation_calls": cm["allocation_calls_delta"],
            "allocation_calls_strictly_reduced": cm["allocation_calls_delta"] < bm["allocation_calls_delta"],
            "baseline_rss_bytes": base["rss_bytes"], "candidate_rss_bytes": cand["rss_bytes"],
        })
    for cell in summary["allocator"]["cells"].values():
        pairs = cell["pairs"]
        summary_values = {
            "pairs": len(pairs),
            "requested_bytes_reduction_fractions": [p["requested_bytes_reduction_fraction"] for p in pairs],
            "allocation_calls_strictly_reduced_all": all(p["allocation_calls_strictly_reduced"] for p in pairs),
            "requested_bytes_reduction_all_ge_20_percent": all(
                p["requested_bytes_reduction_fraction"] is not None
                and p["requested_bytes_reduction_fraction"] >= 0.20 for p in pairs),
        }
        cell["gate_review"] = summary_values
    return summary


def reduce_packet(args: argparse.Namespace) -> dict[str, Any]:
    summary = replay_packet(args.packet, manifest_path=args.manifest)
    output = args.output.resolve()
    require(output.is_relative_to(REPO_ROOT) is False,
            "reduced external evidence must be outside the checkout")
    write_new_json(output, summary)
    summary["output"] = {"path": str(output), "bytes": output.stat().st_size,
                          "sha256": sha256_path(output)}
    return summary


def check_schedule() -> dict[str, Any]:
    rows = schedule_rows()
    return {"schema": "litchi.performance.0933.formal-runner-controls.v1",
            "status": "passed", "children": len(rows), "normal_children": NORMAL_CHILDREN,
            "allocator_children": ALLOCATOR_CHILDREN, "routes": list(ROUTES),
            "children_per_route_per_version": BLOCKS, "warmups": WARMUPS,
            "samples": SAMPLES, "cpu": CPU,
            "normal_deadline_seconds": NORMAL_DEADLINE_SECONDS,
            "allocator_deadline_seconds": ALLOCATOR_DEADLINE_SECONDS,
            "schedule_sha256": schedule_hash(rows), "source_mutation": False,
            "capture_started": False}


def controls(args: argparse.Namespace) -> dict[str, Any]:
    """Exercise custody-only failure paths without starting a document child."""
    manifest = read_json(args.manifest.resolve())
    scratch = Path(tempfile.mkdtemp(prefix="formal-runner-controls-",
                                    dir="/tmp/litchi-performance-0933"))
    result: dict[str, Any] = {
        "schema": "litchi.performance.0933.formal-runner-controls.v1",
        "status": "failed",
        "capture_started": False,
        "document_children_started": 0,
        "runner": descriptor(Path(__file__).resolve()),
        "manifest": descriptor(args.manifest.resolve()),
        "schedule_sha256": schedule_hash(schedule_rows()),
        "tests": {},
    }
    try:
        receipt = scratch / "receipt.json"
        write_new_json(receipt, {"first": True})
        first_hash = sha256_path(receipt)
        rejected = False
        try:
            write_new_json(receipt, {"second": True})
        except EvidenceError:
            rejected = True
        require(rejected, "exclusive receipt control did not reject overwrite")
        require(sha256_path(receipt) == first_hash, "exclusive receipt changed original bytes")
        result["tests"]["exclusive_receipt"] = {"passed": True, "sha256": first_hash}

        packet_scripts = scratch / "scripts"
        packet_scripts.mkdir()
        original = Path(str(manifest["scripts"]["validator"]["path"]))
        local = packet_scripts / original.name
        shutil.copy2(original, local)
        local_manifest = {
            "scripts": {"validator": manifest["scripts"]["validator"]},
            "packet_scripts": {"validator": descriptor(local, root=scratch)},
        }
        verify_scripts(local_manifest, scratch)
        module = load_comparison(local)
        require(callable(getattr(module, "validate_report", None)),
                "copied packet validator did not load")
        result["tests"]["copied_validator_resolution"] = {
            "passed": True, "sha256": sha256_path(local), "bytes": local.stat().st_size,
        }

        allocator_report_path = Path(
            "/home/zhuhe/litchi-performance-0931/formal-packet-v1/children/"
            "s8192-c08-exact-shared_snapshot_clones-r01/report.json"
        )
        allocator_report = read_json(allocator_report_path)
        allocator_row = {
            "lane": "allocator", "sections": 8192, "ownership": "shared_snapshot_clones",
            "repetition": 1, "concurrent_commits": 8, "policy": "exact",
        }
        allocator_desc = descriptor(allocator_report_path)
        validate_allocator_report(allocator_report, allocator_row, allocator_desc, BASE_REVISION)
        tampered = copy.deepcopy(allocator_report)
        tampered["scenarios"][0]["gauges"]["baseline"]["observer_invalid"] = True
        rejected = False
        try:
            validate_allocator_report(tampered, allocator_row, allocator_desc, BASE_REVISION)
        except EvidenceError:
            rejected = True
        require(rejected, "allocator observer-invalid negative control was accepted")
        result["tests"]["allocator_validation_and_negative"] = {
            "passed": True, "report": allocator_desc,
            "rejected_observer_invalid": True,
        }

        candidate_revision = source_revision_for(manifest, "candidate")
        allocator_command = child_command(
            allocator_row, Path("/bin/true"), scratch / "allocator-command-report.json",
            candidate_revision
        )
        revision_index = allocator_command.index("--source-revision")
        require(allocator_command[revision_index + 1] == candidate_revision,
                "allocator command control lost candidate source revision")
        result["tests"]["allocator_source_revision_binding"] = {
            "passed": True, "source_revision": candidate_revision,
        }

        child = subprocess.Popen(["/bin/true"], start_new_session=True)
        child.wait(timeout=10)
        closed, probes = close_process_group(child.pid)
        require(closed, "process-group closure control failed")
        result["tests"]["process_group_closure"] = {"passed": True, "probes": probes}
        result["status"] = "passed"
        return result
    finally:
        shutil.rmtree(scratch, ignore_errors=True)
        write_new_json(args.output.resolve(), result, limit=MAX_ARTIFACT_BYTES)


def parser() -> argparse.ArgumentParser:
    value = argparse.ArgumentParser(description=__doc__)
    sub = value.add_subparsers(dest="command", required=True)
    check = sub.add_parser("check-schedule")
    check.set_defaults(function=lambda args: check_schedule())
    controls_parser = sub.add_parser("controls")
    controls_parser.add_argument("--manifest", type=Path, required=True)
    controls_parser.add_argument("--output", type=Path, required=True)
    controls_parser.set_defaults(function=controls)
    freeze_parser = sub.add_parser("freeze")
    freeze_parser.add_argument("--manifest", type=Path, required=True)
    freeze_parser.add_argument("--plan", type=Path, default=PLAN_DEFAULT)
    freeze_parser.add_argument("--gate-correction", type=Path, default=CORRECTION_DEFAULT)
    freeze_parser.add_argument("--baseline-source-receipt", type=Path, required=True)
    freeze_parser.add_argument("--candidate-source-receipt", type=Path, required=True)
    freeze_parser.add_argument("--baseline-build-receipt", type=Path, required=True)
    freeze_parser.add_argument("--candidate-build-receipt", type=Path, required=True)
    freeze_parser.add_argument("--baseline-allocator-build-receipt", type=Path, required=True)
    freeze_parser.add_argument("--candidate-allocator-build-receipt", type=Path, required=True)
    freeze_parser.add_argument("--allocator-equivalence", type=Path, required=True)
    freeze_parser.set_defaults(function=freeze)
    capture_parser = sub.add_parser("capture")
    capture_parser.add_argument("--manifest", type=Path, required=True)
    capture_parser.add_argument("--packet", type=Path, required=True)
    capture_parser.set_defaults(function=capture)
    replay_parser = sub.add_parser("replay")
    replay_parser.add_argument("--packet", type=Path, required=True)
    replay_parser.add_argument("--manifest", type=Path)
    replay_parser.set_defaults(function=lambda args: replay_packet(args.packet, manifest_path=args.manifest))
    reduce_parser = sub.add_parser("reduce")
    reduce_parser.add_argument("--packet", type=Path, required=True)
    reduce_parser.add_argument("--manifest", type=Path)
    reduce_parser.add_argument("--output", type=Path, required=True)
    reduce_parser.set_defaults(function=reduce_packet)
    return value


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    value = args.function(args)
    print(json.dumps(value, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"comparison formal runner failed: {type(error).__name__}: {error}", file=sys.stderr)
        raise
