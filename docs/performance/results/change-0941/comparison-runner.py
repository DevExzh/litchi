#!/usr/bin/env python3
"""Freeze, capture, and reduce the prospective 0941 DOCX ingress cohort.

The runner is intentionally inert until ``capture`` receives an authorized
manifest and ``LITCHI_PERF_0941_CAPTURE_APPROVED=1``.  Synthetic preflight
uses the existing child-validation boundary and report fixtures only.  A real
child is run once, with its terminal receipt written by
``tools/perf_child_validation.py`` before this runner parses the RSS sidecar
or interprets the report.  No child is retried or replaced.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
import random
import re
import shutil
import statistics
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Mapping, Sequence


REPO_ROOT = Path("/home/zhuhe/code/litchi").resolve()
ARTIFACT_ROOT = Path("/home/zhuhe/litchi-performance-0941").resolve()
PLAN_DEFAULT = ARTIFACT_ROOT / "plans/docx-plan-v2.json"
VALIDATOR_DEFAULT = Path(
    "/home/zhuhe/litchi-performance-0940/docx_owned_attribution_validator_v3.py"
).resolve()
SYNTHETIC_DEFAULT = (
    REPO_ROOT / "docs/performance/results/change-0940/comparison.py"
).resolve()
HELPER_DEFAULT = (REPO_ROOT / "tools/perf_child_validation.py").resolve()
TIME = Path("/usr/bin/time").resolve()
TASKSET = Path("/usr/bin/taskset").resolve()
SOURCE_REVISION = "3ae856d76424f2f9b58d26dcd20fd074b8e97541"
BASELINE_SOURCE = Path("/home/zhuhe/litchi-performance-0940/source-candidate-v1").resolve()
CANDIDATE_SOURCE = Path("/home/zhuhe/litchi-performance-0941/source-candidate-v3").resolve()
BASELINE_LISTING = Path(
    "/home/zhuhe/litchi-performance-0940/source-candidate-manifest-v1.jsonl"
).resolve()
CANDIDATE_LISTING = Path(
    "/home/zhuhe/litchi-performance-0941/source-candidate-manifest-v3.jsonl"
).resolve()
BASELINE_RECEIPT = Path(
    "/home/zhuhe/litchi-performance-0940/receipts/source-candidate-v1.json"
).resolve()
CANDIDATE_RECEIPT = Path(
    "/home/zhuhe/litchi-performance-0941/receipts/source-candidate-v3.json"
).resolve()
QUALITY_SOURCE = Path("/home/zhuhe/litchi-performance-0941/source-quality-v4").resolve()
QUALITY_RECEIPT = Path("/home/zhuhe/litchi-performance-0941/receipts/source-quality-v4.json").resolve()
QUALITY_SUMMARY = (REPO_ROOT / "docs/performance/results/change-0941/new-quality.json").resolve()
EQUIVALENCE_RECEIPT = Path(
    "/home/zhuhe/litchi-performance-0941/receipts/baseline-runtime-equivalence-v1.json"
).resolve()
QUALITY_CHECK_ROOT = Path("/home/zhuhe/litchi-performance-0941/checks").resolve()
QUALITY_CHECKS = (
    "opc-clippy-default-v4", "opc-clippy-all-features-v4",
    "fingerprint-test-default-v4", "fingerprint-test-all-features-v4",
)
BASE_REVISION_LENGTH = 40
WARMUPS = 5
SAMPLES = 50
CPU = 12
BLOCKS = 12
ALLOCATOR_PAIRS = 3
DEADLINE_SECONDS = 900.0
OBSERVATION_SECONDS = 1.0
MAX_REPORT_BYTES = 1 << 20
APPROVAL_ENV = "LITCHI_PERF_0941_CAPTURE_APPROVED"
APPROVAL_VALUE = "1"
NORMAL_BINARY_IDENTITY = "litchi-perf-baseline"
ALLOCATOR_BINARY_IDENTITY = "litchi-perf-baseline-alloc"
NORMAL_BIN_NAME = "litchi-perf-baseline"
ALLOCATOR_BIN_NAME = "litchi-perf-baseline-alloc"
DOCX_SOURCE_BYTES = 2_101_427
DOCX_SOURCE_SHA256 = "183204a832a3518bf1d4474e9b7ff347ad7d0dc95ad2e9b1dbdde405837a29e8"
DOCX_OUTPUT_BYTES = 2_101_446
DOCX_OUTPUT_SHA256 = "5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec"
TIMING_SCOPE = (
    "operation_elapsed_ns includes prep, open, and workload phases; "
    "wall_elapsed_ns is the measured outer bracket from before-prep observation "
    "through the selected body end (publication for workload, post-open observation for calibration)"
)
OBSERVER_GAP_SCOPE = (
    "observer_gap_ns starts after prep's first Region::split and ends after the second split; "
    "allocation observer_gap contains both procfs reads, allocation combined ends before "
    "after-open observation, process combined includes after-open observation, and wall includes "
    "the full outer bracket"
)
CHILD_ENVIRONMENT = {
    "PATH": "/usr/bin:/bin",
    "LC_ALL": "C",
    "LANG": "C",
    "TZ": "UTC",
    "RUST_BACKTRACE": "0",
}
ALLOWED_RUNTIME_DIFF = {"crates/litchi-opc/src/phys_pkg.rs"}
ALLOWED_ADDED_PATHS = {
    "crates/litchi-opc/tests/reader_ingress.rs",
    "docs/performance/results/change-0416/corpus/opc-local-only-signed.zip",
}
EXPECTED_BASELINE_PHYS_SHA256 = "3dfa9ec92c0d090b0bc753ecacc63f263975a470c267b79b03ae1c5898ac3ab8"
EXPECTED_CANDIDATE_PHYS_SHA256 = "899ce1567755d8b4d2c3128a016ccc85c968874750e426d0ffd7da1ff5dc147c"
EXPECTED_READER_TEST_SHA256 = "70de65fe8de228216f51cf84b72806faba12d43153472f9502ff73515eced46f"
EXPECTED_SOURCE_LISTING = {
    "baseline": {"files": 10_780, "bytes": 260_009_611},
    "candidate": {"files": 10_782, "bytes": 260_028_642},
}
EXPECTED_SCRIPT_HASHES: dict[str, str] = {}


class RunnerError(RuntimeError):
    """A custody, schedule, report, or gate contract failed closed."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RunnerError(message)


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def sha256_path(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def descriptor(path: Path, *, root: Path | None = None) -> dict[str, Any]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    rendered = str(path)
    if root is not None:
        try:
            rendered = path.relative_to(root.resolve()).as_posix()
        except ValueError as error:
            raise RunnerError(f"path outside packet root: {path}") from error
    return {"path": rendered, "bytes": path.stat().st_size, "sha256": sha256_path(path)}


def read_json(path: Path, *, max_bytes: int | None = MAX_REPORT_BYTES) -> dict[str, Any]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    if max_bytes is not None:
        require(path.stat().st_size < max_bytes, f"JSON exceeds limit: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as error:
        raise RunnerError(f"cannot parse JSON {path}: {error}") from error
    require(isinstance(value, dict), f"JSON is not an object: {path}")
    return value


def write_json_exclusive(path: Path, value: Any, *, max_bytes: int | None = MAX_REPORT_BYTES) -> None:
    data = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    if max_bytes is not None:
        require(len(data) < max_bytes, f"JSON output exceeds limit: {path}")
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("xb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError as error:
        raise RunnerError(f"refusing to overwrite: {path}") from error


def write_bytes_exclusive(path: Path, data: bytes) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("xb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError as error:
        raise RunnerError(f"refusing to overwrite: {path}") from error


def load_module(path: Path, expected_sha256: str, name: str) -> Any:
    actual = sha256_path(path)
    require(actual == expected_sha256, f"{name} hash changed: {actual} != {expected_sha256}")
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None, f"cannot load {name}: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def schedule_rows() -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for block in range(1, BLOCKS + 1):
        routes = ("owned_vec", "owned_reader") if block % 2 else ("owned_reader", "owned_vec")
        versions = ("baseline", "candidate") if block % 2 else ("candidate", "baseline")
        for position, route in enumerate(routes, start=1):
            pair_id = f"normal-block-{block:02d}-{position:02d}-{route}"
            for version_position, version in enumerate(versions, start=1):
                rows.append({
                    "id": f"{pair_id}-{version}",
                    "pair_id": pair_id,
                    "lane": "normal",
                    "block": block,
                    "position": position,
                    "version_position": version_position,
                    "route": route,
                    "version": version,
                    "binary_lane": "normal",
                    "diagnostic": False,
                    "warmups": WARMUPS,
                    "samples": SAMPLES,
                    "cpu": CPU,
                    "deadline_seconds": int(DEADLINE_SECONDS),
                })
    for repetition in range(1, ALLOCATOR_PAIRS + 1):
        routes = ("owned_vec", "owned_reader") if repetition % 2 else ("owned_reader", "owned_vec")
        versions = ("baseline", "candidate") if repetition % 2 else ("candidate", "baseline")
        for position, route in enumerate(routes, start=1):
            pair_id = f"allocator-pair-{repetition:02d}-{position:02d}-{route}"
            for version_position, version in enumerate(versions, start=1):
                rows.append({
                    "id": f"{pair_id}-{version}",
                    "pair_id": pair_id,
                    "lane": "allocator",
                    "repetition": repetition,
                    "position": position,
                    "version_position": version_position,
                    "route": route,
                    "version": version,
                    "binary_lane": "allocator",
                    "diagnostic": True,
                    "warmups": WARMUPS,
                    "samples": SAMPLES,
                    "cpu": CPU,
                    "deadline_seconds": int(DEADLINE_SECONDS),
                })
    require(len(rows) == 60, f"schedule row count changed: {len(rows)}")
    require(sum(row["lane"] == "normal" for row in rows) == 48, "normal count changed")
    require(sum(row["lane"] == "allocator" for row in rows) == 12, "allocator count changed")
    require(len({row["id"] for row in rows}) == len(rows), "schedule IDs are not unique")
    for lane, expected in (("normal", 12), ("allocator", 3)):
        for route in ("owned_vec", "owned_reader"):
            for version in ("baseline", "candidate"):
                count = sum(row["lane"] == lane and row["route"] == route and row["version"] == version for row in rows)
                require(count == expected, f"schedule count changed: {lane}/{route}/{version}")
    return rows


def schedule_sha256(rows: Sequence[Mapping[str, Any]] | None = None) -> str:
    rows = schedule_rows() if rows is None else rows
    encoded = json.dumps(list(rows), sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def bound_descriptor(value: Any, label: str) -> tuple[Path, dict[str, Any]]:
    require(isinstance(value, Mapping), f"{label} descriptor missing")
    path = Path(str(value.get("path", ""))).resolve()
    actual = descriptor(path)
    require(value.get("bytes") == actual["bytes"], f"{label} size changed")
    require(value.get("sha256") == actual["sha256"], f"{label} hash changed")
    return path, actual


def safe_relative(value: Any, label: str) -> Path:
    require(isinstance(value, str) and value, f"{label} path missing")
    path = Path(value)
    require(not path.is_absolute() and ".." not in path.parts and path != Path("."), f"{label} path escapes root")
    return path


def source_listing(root: Path, listing: Path) -> tuple[int, int, dict[str, dict[str, Any]]]:
    rows: list[dict[str, Any]] = []
    seen: set[str] = set()
    with listing.open("r", encoding="utf-8") as stream:
        for line_number, line in enumerate(stream, start=1):
            try:
                row = json.loads(line)
            except json.JSONDecodeError as error:
                raise RunnerError(f"source listing line {line_number} is invalid JSON") from error
            require(isinstance(row, Mapping), f"source listing line {line_number} is not an object")
            relative = safe_relative(row.get("path"), f"source listing line {line_number}")
            key = relative.as_posix()
            require(key not in seen, f"duplicate source listing path: {key}")
            seen.add(key)
            require(type(row.get("bytes")) is int and row["bytes"] >= 0, f"invalid source size: {key}")
            digest = row.get("sha256")
            require(isinstance(digest, str) and re.fullmatch(r"[0-9a-f]{64}", digest) is not None, f"invalid source digest: {key}")
            rows.append({"path": key, "bytes": row["bytes"], "sha256": digest})
    actual: dict[str, Path] = {}
    for directory, directories, files in os.walk(root, followlinks=False):
        directory_path = Path(directory)
        for name in directories + files:
            path = directory_path / name
            require(not path.is_symlink(), f"source tree contains symlink: {path}")
        for name in files:
            path = directory_path / name
            require(path.is_file(), f"source tree entry is not regular: {path}")
            relative = path.relative_to(root).as_posix()
            require(relative not in actual, f"duplicate source tree path: {relative}")
            actual[relative] = path
    expected = {row["path"] for row in rows}
    require(set(actual) == expected, f"source listing path set mismatch: {len(actual)} != {len(expected)}")
    for row in rows:
        path = actual[row["path"]]
        require(path.stat().st_size == row["bytes"], f"source size changed: {path}")
        require(sha256_path(path) == row["sha256"], f"source digest changed: {path}")
    total = sum(row["bytes"] for row in rows)
    return len(rows), total, {row["path"]: row for row in rows}


def validate_plan(plan_path: Path) -> dict[str, Any]:
    plan = read_json(plan_path)
    require(plan.get("schema") == "litchi.performance.0941.docx-owned-ingress-plan.v2", "plan schema changed")
    require(plan.get("status") == "prospective_only_no_build_no_workload", "plan is not the frozen prospective plan")
    source = plan.get("source_binding")
    require(isinstance(source, Mapping), "source binding missing")
    require(source.get("report_source_revision") == SOURCE_REVISION, "report source revision changed")
    constants = plan.get("source_constants")
    require(isinstance(constants, Mapping), "source constants missing")
    require(constants.get("input_read_chunk_bytes") == 8192
            and constants.get("input_growth_step_bytes") == 65536
            and constants.get("input_growth_fraction_denominator") == 32,
            "source growth constants changed")
    require(constants.get("expected_capacity_expression") ==
            "current_capacity.saturating_add((current_capacity / INPUT_GROWTH_FRACTION_DENOMINATOR).min(INPUT_GROWTH_STEP_BYTES)).max(new_len).min(maximum)",
            "source growth expression changed")
    schedule = plan.get("schedule")
    require(isinstance(schedule, Mapping), "schedule missing")
    require(schedule.get("cpu") == CPU and schedule.get("warmups") == WARMUPS
            and schedule.get("samples") == SAMPLES and schedule.get("deadline_seconds") == int(DEADLINE_SECONDS)
            and schedule.get("total_children") == 60,
            "schedule scalar changed")
    normal = schedule.get("normal")
    allocator = schedule.get("allocator")
    require(isinstance(normal, Mapping) and normal.get("blocks") == BLOCKS and normal.get("children") == 48,
            "normal schedule changed")
    require(isinstance(allocator, Mapping) and allocator.get("pairs_per_route") == ALLOCATOR_PAIRS
            and allocator.get("children") == 12, "allocator schedule changed")
    gates = plan.get("gates")
    require(isinstance(gates, Mapping), "gates missing")
    latency = gates.get("latency_adoption")
    require(isinstance(latency, Mapping) and latency.get("primary_route") == "owned_reader"
            and latency.get("bootstrap_draws") == 10000 and latency.get("bootstrap_seed") == 941,
            "latency gate changed")
    allocation = gates.get("allocation_adoption")
    require(isinstance(allocation, Mapping) and allocation.get("target_route") == "owned_reader"
            and allocation.get("target_phase") == "allocation.open",
            "allocation gate changed")
    return plan


def _plan_descriptor(plan: Mapping[str, Any], side: str, key: str) -> Mapping[str, Any]:
    source = plan["source_binding"][side]
    value = source[key]
    require(isinstance(value, Mapping), f"plan {side} {key} descriptor missing")
    return value


def validate_sources(plan: Mapping[str, Any]) -> dict[str, Any]:
    roots = {"baseline": BASELINE_SOURCE, "candidate": CANDIDATE_SOURCE}
    listings = {"baseline": BASELINE_LISTING, "candidate": CANDIDATE_LISTING}
    receipts = {"baseline": BASELINE_RECEIPT, "candidate": CANDIDATE_RECEIPT}
    source_rows: dict[str, dict[str, dict[str, Any]]] = {}
    source_summary: dict[str, Any] = {}
    for side in ("baseline", "candidate"):
        expected = EXPECTED_SOURCE_LISTING[side]
        source = plan["source_binding"][side]
        require(Path(str(source["source_root"])).resolve() == roots[side], f"{side} source root changed")
        bound_listing, listing_desc = bound_descriptor(source["source_listing"], f"{side} source listing")
        require(bound_listing == listings[side], f"{side} source listing path changed")
        bound_receipt, receipt_desc = bound_descriptor(source["source_receipt"], f"{side} source receipt")
        require(bound_receipt == receipts[side], f"{side} source receipt path changed")
        require(listing_desc["bytes"] > 0, f"{side} source listing is empty")
        count, total, rows = source_listing(roots[side], listings[side])
        require(count == expected["files"] and total == expected["bytes"], f"{side} source closure changed")
        require(source.get("source_file_count") == count and source.get("source_bytes") == total,
                f"{side} source receipt totals changed")
        receipt = read_json(receipts[side])
        require(receipt.get("baseline_revision", receipt.get("source_revision")) == SOURCE_REVISION,
                f"{side} source revision changed")
        phys = roots[side] / "crates/litchi-opc/src/phys_pkg.rs"
        expected_phys = EXPECTED_BASELINE_PHYS_SHA256 if side == "baseline" else EXPECTED_CANDIDATE_PHYS_SHA256
        require(sha256_path(phys) == expected_phys, f"{side} phys_pkg source changed")
        source_rows[side] = rows
        source_summary[side] = {"root": str(roots[side]), "listing": listing_desc, "receipt": receipt_desc,
                                "file_count": count, "bytes": total, "phys_pkg_sha256": expected_phys}
    baseline = source_rows["baseline"]
    candidate = source_rows["candidate"]
    require(set(baseline) <= set(candidate), "candidate source lost baseline path")
    changed: set[str] = set()
    for path, row in baseline.items():
        if candidate[path] != row:
            changed.add(path)
    added = set(candidate) - set(baseline)
    require(changed == ALLOWED_RUNTIME_DIFF, f"unexpected source changes: {sorted(changed)}")
    require(added == ALLOWED_ADDED_PATHS, f"unexpected source additions: {sorted(added)}")
    require(not (set(baseline) - set(candidate)), "candidate removed a source path")
    require(sha256_path(CANDIDATE_SOURCE / "crates/litchi-opc/tests/reader_ingress.rs") == EXPECTED_READER_TEST_SHA256,
            "candidate reader ingress test changed")
    _validate_source_constants(CANDIDATE_SOURCE / "crates/litchi-opc/src/phys_pkg.rs")
    return source_summary


def _validate_source_constants(path: Path) -> None:
    text = path.read_text(encoding="utf-8")
    for fragment in (
        "const INPUT_READ_CHUNK_BYTES: usize = 8 * 1024;",
        "const INPUT_GROWTH_STEP_BYTES: usize = 64 * 1024;",
        "const INPUT_GROWTH_FRACTION_DENOMINATOR: usize = 32;",
        "current_capacity / INPUT_GROWTH_FRACTION_DENOMINATOR",
        ".min(INPUT_GROWTH_STEP_BYTES)",
        ".max(new_len)",
        ".min(maximum)",
        "if new_len > data.capacity()",
        "data.try_reserve_exact(additional)",
    ):
        require(fragment in text, f"candidate source constant/policy missing: {fragment}")


def quality_binding() -> dict[str, Any]:
    require(QUALITY_SOURCE.is_dir(), f"quality source root missing: {QUALITY_SOURCE}")
    quality_summary = descriptor(QUALITY_SUMMARY)
    summary = read_json(QUALITY_SUMMARY)
    require(summary.get("$schema") == "litchi.performance.0941.quality.v1"
            and summary.get("change") == "0941", "quality summary schema changed")
    quality_receipt = descriptor(QUALITY_RECEIPT)
    equivalence_receipt = descriptor(EQUIVALENCE_RECEIPT)
    receipt = read_json(QUALITY_RECEIPT)
    require(receipt.get("source") == str(QUALITY_SOURCE), "quality source receipt root changed")
    require(receipt.get("baseline_source") == str(CANDIDATE_SOURCE), "quality baseline source changed")
    require(receipt.get("scope", "").startswith("Quality-only overlay:"), "quality scope changed")
    changed = receipt.get("changed_file")
    require(isinstance(changed, Mapping), "quality changed file missing")
    changed_path, changed_desc = bound_descriptor(changed, "quality changed file")
    require(changed_path == QUALITY_SOURCE / "crates/litchi-opc/tests/source_fingerprint_quality.rs",
            "quality changed file path changed")
    require(changed_desc == {"path": str(changed_path), "bytes": changed["bytes"], "sha256": changed["sha256"]},
            "quality changed file descriptor changed")
    checks: dict[str, Any] = {}
    for name in QUALITY_CHECKS:
        terminal_path = QUALITY_CHECK_ROOT / f"{name}.terminal.json"
        start_path = QUALITY_CHECK_ROOT / f"{name}.started.json"
        stdout_path = QUALITY_CHECK_ROOT / f"{name}.stdout.log"
        stderr_path = QUALITY_CHECK_ROOT / f"{name}.stderr.log"
        terminal = read_json(terminal_path)
        start = read_json(start_path)
        require(terminal.get("exit_code") == 0 and start.get("cwd") == str(QUALITY_SOURCE)
                and terminal.get("cwd") == str(QUALITY_SOURCE), f"quality check failed: {name}")
        require(start.get("argv") == terminal.get("argv"), f"quality check argv binding changed: {name}")
        files = {
            "start": descriptor(start_path), "terminal": descriptor(terminal_path),
            "stdout": descriptor(stdout_path), "stderr": descriptor(stderr_path),
        }
        for stream in ("stdout", "stderr"):
            stream_value = terminal.get(stream)
            require(isinstance(stream_value, Mapping), f"quality {stream} descriptor missing: {name}")
            _, actual = bound_descriptor(stream_value, f"quality {stream} {name}")
            require(actual == stream_value, f"quality {stream} descriptor changed: {name}")
        checks[name] = files
    return {
        "summary": quality_summary, "source_root": str(QUALITY_SOURCE), "source_receipt": quality_receipt,
        "runtime_equivalence_receipt": equivalence_receipt, "checks": checks,
    }


def expected_report(binary_identity: str, diagnostic: bool) -> dict[str, Any]:
    if binary_identity == ALLOCATOR_BINARY_IDENTITY:
        require(diagnostic, "allocator report cannot disable diagnostics")
        arm = "allocator_on"
        allocator = "CountingSystemAllocator(std::alloc::System)"
        instrumentation = "system_allocator_operation_scoped"
        revision: str | None = "serialized_region_peak_v3"
    else:
        arm = "normal_on" if diagnostic else "normal_off"
        allocator = "Rust system allocator"
        instrumentation = "none"
        revision = None
    return {
        "source_revision": SOURCE_REVISION,
        "binary_identity": binary_identity,
        "diagnostic_arm": arm,
        "diagnostic_enabled": diagnostic,
        "allocator_identity": allocator,
        "instrumentation_identity": instrumentation,
        "allocator_counter_revision": revision,
        "timing_scope": TIMING_SCOPE,
        "observer_gap_scope": OBSERVER_GAP_SCOPE,
    }


def binary_descriptor_for_row(manifest: Mapping[str, Any], row: Mapping[str, Any]) -> dict[str, Any]:
    return _binary_descriptor_for_row(manifest, row, packet=None)


def _binary_descriptor_for_row(packet_manifest: Mapping[str, Any], row: Mapping[str, Any],
                               packet: Path | None) -> dict[str, Any]:
    lane = str(row["binary_lane"])
    version = str(row["version"])
    builds = packet_manifest.get("builds")
    require(isinstance(builds, Mapping), "manifest builds missing")
    lane_values = builds.get(lane)
    require(isinstance(lane_values, Mapping), f"manifest build lane missing: {lane}")
    entry = lane_values.get(version)
    require(isinstance(entry, Mapping), f"manifest build version missing: {lane}/{version}")
    binary_value = entry.get("binary")
    require(isinstance(binary_value, Mapping), f"{lane}/{version} binary descriptor missing")
    raw_path = binary_value.get("path")
    require(isinstance(raw_path, str) and raw_path, f"{lane}/{version} binary path missing")
    if packet is not None and not Path(raw_path).is_absolute():
        binary_value = dict(binary_value)
        binary_value["path"] = str(packet / safe_relative(raw_path, f"{lane}/{version} binary"))
    binary, _ = bound_descriptor(binary_value, f"{lane}/{version} binary")
    identity = ALLOCATOR_BINARY_IDENTITY if lane == "allocator" else NORMAL_BINARY_IDENTITY
    require(entry.get("binary_identity") == identity, f"binary identity changed: {lane}/{version}")
    return {
        "path": str(binary),
        "bytes": binary.stat().st_size,
        "sha256": sha256_path(binary),
        "identity": identity,
        "binary_identity": identity,
        "source_revision": SOURCE_REVISION,
        "expected_report": expected_report(identity, bool(row["diagnostic"])),
        "expected_allocation_present": lane == "allocator",
        "expected_process_present": lane == "allocator",
    }


def _expected_build_argv(*, source_root: Path, lane: str, binary_path: Path) -> list[str]:
    binary_name = ALLOCATOR_BINARY_IDENTITY if lane == "allocator" else NORMAL_BINARY_IDENTITY
    argv = [
        "/home/zhuhe/.cargo/bin/cargo", "+1.95.0", "build", "--offline", "--locked", "-j2", "--release",
    ]
    if lane == "allocator":
        argv += ["--features", "allocator-metrics"]
    argv += [
        "--bin", binary_name, "--manifest-path", str(source_root / "tools/perf-baseline/Cargo.toml"),
        "--target-dir", str(binary_path.parent.parent),
    ]
    return argv


def _validate_build_environment(value: Any, label: str) -> dict[str, str]:
    require(isinstance(value, Mapping), f"{label} environment missing")
    environment = dict(value)
    require(all(isinstance(key, str) and isinstance(item, str) for key, item in environment.items()),
            f"{label} environment is not string-valued")
    require("RUSTFLAGS" not in environment and "CARGO_ENCODED_RUSTFLAGS" not in environment,
            f"{label} carries forbidden Rust flags")
    allowed = {"HOME", "CARGO_HOME", "RUSTUP_HOME", "PATH", "LC_ALL", "LANG", "TZ", "CARGO_TERM_COLOR"}
    require(set(environment) <= allowed, f"{label} has an unallowlisted variable")
    require(environment.get("CARGO_TERM_COLOR") == "never", f"{label} cargo color policy changed")
    return environment


def validate_build_entry(entry: Any, *, version: str, lane: str, source_root: Path,
                         expected_binary_path: Path | None = None,
                         expected_terminal_path: Path | None = None) -> dict[str, Any]:
    require(isinstance(entry, Mapping), f"build entry missing: {lane}/{version}")
    identity = ALLOCATOR_BINARY_IDENTITY if lane == "allocator" else NORMAL_BINARY_IDENTITY
    require(entry.get("source_root") == str(source_root), f"build source root changed: {lane}/{version}")
    require(entry.get("source_revision") == SOURCE_REVISION, f"build source revision changed: {lane}/{version}")
    require(entry.get("binary_identity") == identity, f"build identity changed: {lane}/{version}")
    binary_path, binary_desc = bound_descriptor(entry.get("binary"), f"build binary {lane}/{version}")
    terminal_path, terminal_desc = bound_descriptor(entry.get("terminal"), f"build terminal {lane}/{version}")
    require(expected_binary_path is None or binary_path == expected_binary_path.resolve(),
            f"build binary path changed: {lane}/{version}")
    require(expected_terminal_path is None or terminal_path == expected_terminal_path.resolve(),
            f"build terminal path changed: {lane}/{version}")
    require(binary_path.parent.name == "release" and binary_path.parent.parent.name.startswith("target-"),
            f"build binary target layout changed: {lane}/{version}")
    terminal = read_json(terminal_path)
    require(terminal.get("exit_code") == 0, f"build failed: {lane}/{version}")
    require(isinstance(terminal.get("schema"), str)
            and terminal["schema"].endswith(".build-terminal.v1"),
            f"build terminal schema changed: {lane}/{version}")
    require(terminal.get("source_root") == str(source_root) and terminal.get("cwd") == str(source_root),
            f"build cwd/source root changed: {lane}/{version}")
    require(terminal.get("name") == f"{lane}-v1", f"build name changed: {lane}/{version}")
    require(terminal.get("target_dir") == str(binary_path.parent.parent),
            f"build target directory changed: {lane}/{version}")
    require(terminal.get("terminal_written_before_binary_descriptor") is True,
            f"build terminal ordering missing: {lane}/{version}")
    argv = terminal.get("argv")
    require(isinstance(argv, list) and all(isinstance(value, str) for value in argv),
            f"build argv missing: {lane}/{version}")
    require(argv == _expected_build_argv(source_root=source_root, lane=lane, binary_path=binary_path),
            f"build argv changed: {lane}/{version}")
    require(terminal.get("binary_name") == identity, f"build binary name changed: {lane}/{version}")
    environment_policy = terminal.get("environment_policy")
    require(environment_policy == "allowlisted build variables; RUSTFLAGS and CARGO_ENCODED_RUSTFLAGS explicitly unset",
            f"build environment policy changed: {lane}/{version}")
    terminal_environment = _validate_build_environment(terminal.get("environment"),
                                                       f"build terminal {lane}/{version}")
    expected_start_path = terminal_path.parent / "start.json"
    expected_receipt_path = terminal_path.parent / "binary.json"
    start_path, start_desc = bound_descriptor(entry.get("start"), f"build start {lane}/{version}")
    binary_receipt_path, binary_receipt_desc = bound_descriptor(
        entry.get("binary_receipt"), f"binary receipt {lane}/{version}"
    )
    require(start_path == expected_start_path, f"build start path changed: {lane}/{version}")
    require(binary_receipt_path == expected_receipt_path, f"binary receipt path changed: {lane}/{version}")
    start = read_json(start_path)
    require(isinstance(start.get("schema"), str) and start["schema"].endswith(".build-start.v1"),
            f"build start schema changed: {lane}/{version}")
    require(start.get("name") == f"{lane}-v1" and start.get("argv") == argv
            and start.get("cwd") == str(source_root)
            and start.get("source_root") == str(source_root)
            and start.get("target_dir") == str(binary_path.parent.parent),
            f"build start binding changed: {lane}/{version}")
    require(start.get("environment_policy") == environment_policy,
            f"build start environment policy changed: {lane}/{version}")
    start_environment = _validate_build_environment(start.get("environment"), f"build start {lane}/{version}")
    require(start_environment == terminal_environment and terminal.get("argv") == start.get("argv"),
            f"build start/terminal environment binding changed: {lane}/{version}")
    for stream in ("stdout", "stderr"):
        stream_path, stream_desc = bound_descriptor(terminal.get(stream), f"build {stream} {lane}/{version}")
        require(stream_desc == terminal[stream], f"build {stream} descriptor changed: {lane}/{version}")
        del stream_path
    receipt = read_json(binary_receipt_path)
    require(receipt == binary_desc, f"binary receipt mismatch: {lane}/{version}")
    return {"source_root": str(source_root), "binary": binary_desc, "binary_path": binary_path,
            "terminal": terminal_desc, "terminal_path": terminal_path, "start": start_desc,
            "binary_receipt": binary_receipt_desc, "binary_identity": identity,
            "expected_report": expected_report(identity, lane == "allocator"),
            "expected_allocation_present": lane == "allocator",
            "expected_process_present": lane == "allocator"}


def validate_manifest(manifest_path: Path, runner_path: Path, *, authorized: bool) -> tuple[dict[str, Any], dict[str, Any]]:
    manifest = read_json(manifest_path)
    require(manifest.get("schema") == "litchi.performance.0941.capture-manifest.v1", "manifest schema changed")
    plan_path, plan_desc = bound_descriptor(manifest.get("plan"), "manifest plan")
    plan = validate_plan(plan_path)
    require(plan_path == PLAN_DEFAULT, "manifest plan path changed")
    source_summary = validate_sources(plan)
    require(manifest.get("source") == source_summary, "manifest source custody changed")
    require(manifest.get("source_revision") == SOURCE_REVISION, "manifest source revision changed")
    require(manifest.get("quality_binding") == quality_binding(), "manifest quality custody changed")
    require(manifest.get("schedule_sha256") == schedule_sha256(), "manifest schedule hash changed")
    require(manifest.get("schedule") == schedule_rows(), "manifest schedule rows changed")
    require(manifest.get("plan_sha256") == plan_desc["sha256"], "manifest plan hash is not current")
    scripts = manifest.get("scripts")
    require(isinstance(scripts, Mapping), "manifest scripts missing")
    expected_scripts = {
        "runner": runner_path.resolve(),
        "child_validation": HELPER_DEFAULT,
        "report_validator": VALIDATOR_DEFAULT,
        "synthetic_fixture_provider": SYNTHETIC_DEFAULT,
        "time": TIME,
        "taskset": TASKSET,
    }
    script_descs: dict[str, dict[str, Any]] = {}
    for name, path in expected_scripts.items():
        bound, actual = bound_descriptor(scripts.get(name), f"script {name}")
        require(bound == path, f"script path changed: {name}")
        if name in {"time", "taskset"}:
            require(os.access(bound, os.X_OK), f"script is not executable: {name}")
        script_descs[name] = actual
    require(manifest.get("child_environment") == CHILD_ENVIRONMENT, "child environment changed")
    builds = manifest.get("builds")
    require(isinstance(builds, Mapping), "manifest builds missing")
    checked_builds: dict[str, dict[str, dict[str, Any]]] = {}
    roots = {"baseline": BASELINE_SOURCE, "candidate": CANDIDATE_SOURCE}
    for lane in ("normal", "allocator"):
        values = builds.get(lane)
        require(isinstance(values, Mapping), f"manifest lane missing: {lane}")
        checked_builds[lane] = {}
        for version in ("baseline", "candidate"):
            checked_builds[lane][version] = validate_build_entry(
                values.get(version), version=version, lane=lane, source_root=roots[version],
                expected_binary_path=Path(str(plan["binary_binding"][f"{version}_{lane}"]["path"])).resolve(),
                expected_terminal_path=Path(str(plan["binary_binding"][f"{version}_{lane}"]["build_terminal"])).resolve(),
            )
    auth = manifest.get("authorization")
    require(isinstance(auth, Mapping), "manifest authorization missing")
    if authorized:
        require(auth.get("capture_authorized") is True, "capture manifest is not root-authorized")
        require(os.environ.get(APPROVAL_ENV) == APPROVAL_VALUE,
                f"set {APPROVAL_ENV}=1 only for root-authorized capture")
    return manifest, {"plan": plan, "source": source_summary, "scripts": script_descs, "builds": checked_builds}


def build_entry_from_paths(*, version: str, lane: str, plan: Mapping[str, Any], terminal_path: Path,
                           binary_path: Path, start_path: Path, binary_receipt_path: Path) -> dict[str, Any]:
    identity = ALLOCATOR_BINARY_IDENTITY if lane == "allocator" else NORMAL_BINARY_IDENTITY
    source_root = Path(str(plan["source_binding"][version]["source_root"])).resolve()
    require(binary_path.resolve() == Path(str(plan["binary_binding"][f"{version}_{lane}"]["path"])).resolve(),
            f"{lane}/{version} binary path differs from frozen plan")
    require(terminal_path.resolve() == Path(str(plan["binary_binding"][f"{version}_{lane}"]["build_terminal"])).resolve(),
            f"{lane}/{version} terminal path differs from frozen plan")
    require(start_path.resolve() == terminal_path.resolve().parent / "start.json",
            f"{lane}/{version} start path must sit beside the build terminal")
    require(binary_receipt_path.resolve() == terminal_path.resolve().parent / "binary.json",
            f"{lane}/{version} binary receipt must sit beside the build terminal")
    entry: dict[str, Any] = {
        "source_root": str(source_root),
        "source_revision": SOURCE_REVISION,
        "binary_identity": identity,
        "binary": descriptor(binary_path),
        "terminal": descriptor(terminal_path),
    }
    entry["start"] = descriptor(start_path)
    entry["binary_receipt"] = descriptor(binary_receipt_path)
    return entry


def make_manifest(args: argparse.Namespace) -> dict[str, Any]:
    plan_path = args.plan.resolve()
    plan = validate_plan(plan_path)
    require(plan_path == PLAN_DEFAULT, "freeze plan path changed")
    source_summary = validate_sources(plan)
    quality = quality_binding()
    rows = schedule_rows()
    builds: dict[str, dict[str, dict[str, Any]]] = {"normal": {}, "allocator": {}}
    for lane in ("normal", "allocator"):
        for version in ("baseline", "candidate"):
            terminal = getattr(args, f"{version}_{lane}_terminal").resolve()
            binary = getattr(args, f"{version}_{lane}_binary").resolve()
            start_value = getattr(args, f"{version}_{lane}_start")
            receipt_value = getattr(args, f"{version}_{lane}_receipt")
            require(start_value is not None and receipt_value is not None,
                    f"freeze requires start and binary receipt: {lane}/{version}")
            builds[lane][version] = build_entry_from_paths(
                version=version, lane=lane, plan=plan, terminal_path=terminal, binary_path=binary,
                start_path=start_value.resolve(), binary_receipt_path=receipt_value.resolve(),
            )
            validate_build_entry(builds[lane][version], version=version, lane=lane,
                                 source_root=Path(str(plan["source_binding"][version]["source_root"])).resolve(),
                                 expected_binary_path=binary, expected_terminal_path=terminal)
    script_paths = {
        "runner": Path(__file__).resolve(),
        "child_validation": HELPER_DEFAULT,
        "report_validator": VALIDATOR_DEFAULT,
        "synthetic_fixture_provider": SYNTHETIC_DEFAULT,
        "time": TIME,
        "taskset": TASKSET,
    }
    scripts = {name: descriptor(path) for name, path in script_paths.items()}
    return {
        "schema": "litchi.performance.0941.capture-manifest.v1",
        "created_utc": utc_now(),
        "capture_authorized": False,
        "authorization": {"capture_authorized": False, "authorized_by": None, "authorized_utc": None},
        "plan": descriptor(plan_path),
        "plan_sha256": sha256_path(plan_path),
        "source_revision": SOURCE_REVISION,
        "source": source_summary,
        "quality_binding": quality,
        "scripts": scripts,
        "child_environment": dict(CHILD_ENVIRONMENT),
        "schedule": rows,
        "schedule_sha256": schedule_sha256(rows),
        "builds": builds,
        "binary_identity_rule": "Normal operation timing uses only normal probes-off binaries. Allocator binaries are an independent lane and never supply normal elapsed values.",
        "rss_contract": {"wrapper": [str(TIME), "-f", "%M", "-o", "CHILD/max-rss-kib.txt", "--", str(TASKSET), "--cpu-list", str(CPU)], "unit": "KiB", "parse_after_terminal": True},
        "claim_scope": "Prospective 48-child normal plus 12-child allocator comparison; no latency, allocation, RSS, or adoption claim until all terminal records, semantic validation, reduction, and review pass.",
        "failure_policy": "No retry or selective recapture. Retain every child terminal, report, stream, validation, and RSS sidecar; any failed child invalidates interpretation.",
        "status": "prospective_unfinished",
    }


def check_source_contract(plan_path: Path) -> dict[str, Any]:
    plan = validate_plan(plan_path)
    source = validate_sources(plan)
    rows = schedule_rows()
    return {"schema": "litchi.performance.0941.source-controls.v1", "status": "passed",
            "plan": descriptor(plan_path), "sources": source, "children": len(rows),
            "schedule_sha256": schedule_sha256(rows), "workloads_started": False}


def _synthetic_child(helper: Any, validator: Any, fixture: Any, scratch: Path, *, label: str,
                     route: str, diagnostic: bool, binary_identity: str, mutated: bool = False,
                     binary_metadata: Mapping[str, Any] | None = None) -> dict[str, Any]:
    child = scratch / label
    child.mkdir(parents=True, exist_ok=False)
    payload = fixture.synthetic_report(route=route, case="workload", diagnostic=diagnostic,
                                       binary_identity=binary_identity)
    if mutated:
        payload["output_sha256"][0] = "0" * 64
    payload_path = child / "payload.json"
    report_path = child / "report.json"
    write_bytes_exclusive(payload_path, json.dumps(payload, sort_keys=True).encode("utf-8"))
    code = "import pathlib,sys; pathlib.Path(sys.argv[2]).write_bytes(pathlib.Path(sys.argv[1]).read_bytes())"
    command = (sys.executable, "-c", code, str(payload_path), str(report_path))
    binary = dict(binary_metadata) if binary_metadata is not None else {
        "identity": binary_identity,
        "binary_identity": binary_identity,
        "source_revision": SOURCE_REVISION,
        "expected_report": expected_report(binary_identity, diagnostic),
        "expected_allocation_present": binary_identity == ALLOCATOR_BINARY_IDENTITY,
        "expected_process_present": binary_identity == ALLOCATOR_BINARY_IDENTITY,
    }
    spec = helper.ChildSpec(
        command=command,
        cwd=scratch,
        environment=dict(CHILD_ENVIRONMENT),
        paths=helper.ChildPaths(report=report_path, stdout=child / "stdout.log", stderr=child / "stderr.log",
                                terminal=child / "terminal.json", validation=child / "validation.json"),
        route=route,
        label=label,
        warmups=WARMUPS,
        samples=SAMPLES,
        expected_output_sha256=DOCX_OUTPUT_SHA256,
        binary=binary,
        report_contract=helper.REPORT_CONTRACT_DOCX_OWNED_WORKLOAD,
        deadline_seconds=30.0,
        observation_seconds=0.05,
        max_report_bytes=MAX_REPORT_BYTES,
    )
    outcome = helper.run_child_and_validate(spec, validator)
    return {"status": outcome["status"], "validation": outcome["validation"], "terminal": outcome["terminal"]}


def preflight(args: argparse.Namespace) -> dict[str, Any]:
    plan = validate_plan(args.plan.resolve())
    del plan
    helper_path = args.helper.resolve()
    validator_path = args.validator.resolve()
    fixture_path = args.synthetic.resolve()
    helper = load_module(helper_path, sha256_path(helper_path), "0941 child helper")
    validator = load_module(validator_path, sha256_path(validator_path), "0940 DOCX validator")
    fixture = load_module(fixture_path, sha256_path(fixture_path), "0940 synthetic fixture provider")
    scratch = args.scratch.resolve() if args.scratch else ARTIFACT_ROOT / "preflight" / f"synthetic-{os.getpid()}"
    scratch.mkdir(parents=True, exist_ok=False)
    positive_normal = _synthetic_child(helper, validator, fixture, scratch, label="normal-off", route="owned_vec",
                                       diagnostic=False, binary_identity=NORMAL_BINARY_IDENTITY)
    positive_alloc = _synthetic_child(helper, validator, fixture, scratch, label="allocator-on", route="owned_reader",
                                      diagnostic=True, binary_identity=ALLOCATOR_BINARY_IDENTITY)
    negative = _synthetic_child(helper, validator, fixture, scratch, label="mutated-output", route="owned_reader",
                                diagnostic=True, binary_identity=ALLOCATOR_BINARY_IDENTITY, mutated=True)
    rss = scratch / "rss-valid.txt"
    write_bytes_exclusive(rss, b"12345\n")
    parsed = parse_rss_sidecar(rss)
    require(parsed["max_rss_kib"] == 12345 and parsed["max_rss_bytes"] == 12345 * 1024,
            "synthetic RSS parse failed")
    invalid_rss = scratch / "rss-invalid.txt"
    write_bytes_exclusive(invalid_rss, b"123\n456\n")
    rejected = False
    try:
        parse_rss_sidecar(invalid_rss)
    except RunnerError:
        rejected = True
    require(rejected, "invalid RSS sidecar was accepted")
    output = {
        "schema": "litchi.performance.0941.synthetic-preflight.v1",
        "status": "passed",
        "real_workloads_started": False,
        "scratch": str(scratch),
        "helper": descriptor(helper_path),
        "validator": descriptor(validator_path),
        "synthetic_fixture": descriptor(fixture_path),
        "positive_normal": positive_normal["validation"],
        "positive_allocator": positive_alloc["validation"],
        "negative_mutated_output": negative["validation"],
        "terminal_before_report_validation": positive_normal["terminal"]["status"] == "terminal",
        "rss_parse": parsed,
        "invalid_rss_rejected": rejected,
    }
    if args.output:
        write_json_exclusive(args.output.resolve(), output)
    return output


def child_command(row: Mapping[str, Any], binary: Mapping[str, Any], report: Path, rss: Path) -> tuple[str, ...]:
    diagnostic = "on" if row["diagnostic"] else "off"
    return (
        str(TIME), "-f", "%M", "-o", str(rss), "--", str(TASKSET), "--cpu-list", str(CPU),
        str(binary["path"]), "docx-owned-attribution", "--case", "workload", "--diagnostic", diagnostic,
        "--warmups", str(WARMUPS), "--samples", str(SAMPLES), "--source-revision", SOURCE_REVISION,
        "--route", str(row["route"]), "--output", str(report),
    )


def parse_rss_sidecar(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing RSS sidecar: {path}")
    raw = path.read_bytes()
    try:
        text = raw.decode("ascii")
    except UnicodeDecodeError as error:
        raise RunnerError(f"RSS sidecar is not ASCII: {path}") from error
    lines = text.splitlines()
    require(len(lines) == 1 and re.fullmatch(r"[0-9]+", lines[0]) is not None,
            f"RSS sidecar must contain one decimal KiB line: {path}")
    kib = int(lines[0], 10)
    return {"path": str(path.resolve()), "bytes": len(raw), "sha256": hashlib.sha256(raw).hexdigest(),
            "unit": "KiB", "format": "%M", "max_rss_kib": kib, "max_rss_bytes": kib * 1024}


def copy_custody_file(source: Path, target: Path) -> dict[str, Any]:
    data = source.read_bytes()
    write_bytes_exclusive(target, data)
    value = descriptor(target)
    require(value["bytes"] == source.stat().st_size and value["sha256"] == sha256_path(source),
            f"custody copy changed: {source}")
    return value


def capture(args: argparse.Namespace) -> dict[str, Any]:
    manifest_path = args.manifest.resolve()
    manifest, binding = validate_manifest(manifest_path, Path(__file__).resolve(), authorized=True)
    packet = args.packet.resolve()
    require(packet != REPO_ROOT and REPO_ROOT not in packet.parents, "packet must be external to checkout")
    require(not packet.exists(), f"refusing to replace packet: {packet}")
    packet.mkdir(parents=True)
    (packet / "children").mkdir()
    (packet / "scripts").mkdir()
    script_copies: dict[str, dict[str, Any]] = {}
    for name, value in manifest["scripts"].items():
        source = Path(str(value["path"])).resolve()
        target = packet / "scripts" / source.name
        copy_custody_file(source, target)
        script_copies[name] = descriptor(target, root=packet)
    plan_target = packet / "plan.json"
    copy_custody_file(Path(str(manifest["plan"]["path"])).resolve(), plan_target)
    plan_copy = descriptor(plan_target, root=packet)
    packet_manifest = dict(manifest)
    packet_manifest["capture_origin"] = str(packet)
    packet_manifest["packet_scripts"] = script_copies
    packet_manifest["packet_plan"] = plan_copy
    packet_manifest["packet_original_manifest"] = descriptor(manifest_path)
    packet_manifest_path = packet / "manifest.json"
    write_json_exclusive(packet_manifest_path, packet_manifest)
    packet_manifest_sha = sha256_path(packet_manifest_path)
    write_json_exclusive(packet / "capture-start.json", {
        "schema": "litchi.performance.0941.capture-start.v1",
        "started_utc": utc_now(), "manifest_sha256": packet_manifest_sha,
        "schedule_sha256": manifest["schedule_sha256"], "children_expected": len(manifest["schedule"]),
        "terminal_first": True, "retry_policy": "none",
    })
    helper = load_module(Path(str(manifest["scripts"]["child_validation"]["path"])),
                         manifest["scripts"]["child_validation"]["sha256"], "0941 child helper")
    validator = load_module(Path(str(manifest["scripts"]["report_validator"]["path"])),
                            manifest["scripts"]["report_validator"]["sha256"], "0940 DOCX validator")
    results: list[dict[str, Any]] = []
    for index, row in enumerate(manifest["schedule"], start=1):
        child = packet / "children" / row["id"]
        child.mkdir(parents=True, exist_ok=False)
        report = child / "report.json"
        rss_path = child / "max-rss-kib.txt"
        binary = binary_descriptor_for_row(manifest, row)
        command = child_command(row, binary, report, rss_path)
        spec = helper.ChildSpec(
            command=command,
            cwd=Path(binding["source"][row["version"]]["root"]),
            environment=dict(manifest["child_environment"]),
            paths=helper.ChildPaths(report=report, stdout=child / "stdout.log", stderr=child / "stderr.log",
                                    terminal=child / "terminal.json", validation=child / "validation.json"),
            route=row["route"], label=row["id"], warmups=WARMUPS, samples=SAMPLES,
            expected_output_sha256=DOCX_OUTPUT_SHA256,
            binary=binary,
            report_contract=helper.REPORT_CONTRACT_DOCX_OWNED_WORKLOAD,
            deadline_seconds=DEADLINE_SECONDS, observation_seconds=OBSERVATION_SECONDS,
            max_report_bytes=MAX_REPORT_BYTES,
        )
        outcome = helper.run_child_and_validate(spec, validator)
        terminal_path = child / "terminal.json"
        require(terminal_path.is_file(), f"child helper did not publish terminal: {row['id']}")
        rss_result: dict[str, Any]
        try:
            rss_result = parse_rss_sidecar(rss_path)
            rss_sidecar_desc = descriptor(rss_path, root=packet)
            rss_result["path"] = rss_sidecar_desc["path"]
            rss_result["bytes"] = rss_sidecar_desc["bytes"]
            rss_result["sha256"] = rss_sidecar_desc["sha256"]
            terminal_desc = descriptor(terminal_path, root=packet)
            rss_result["terminal_path"] = terminal_desc["path"]
            rss_result["terminal_sha256"] = terminal_desc["sha256"]
            rss_status = "passed"
        except Exception as error:
            rss_result = {"path": str(rss_path.resolve()), "status": "failed", "error": f"{type(error).__name__}: {error}"}
            rss_status = "failed"
            rss_sidecar_desc = None
            terminal_desc = descriptor(terminal_path, root=packet)
        write_json_exclusive(child / "rss-receipt.json", {
            "schema": "litchi.performance.0941.rss-receipt.v1",
            "status": rss_status, "child_id": row["id"], "terminal_first": True,
            "terminal": terminal_desc, "sidecar": rss_sidecar_desc, "rss": rss_result,
        })
        passed = outcome["status"] == "passed" and rss_status == "passed"
        result = {
            "index": index, "id": row["id"], "pair_id": row["pair_id"], "lane": row["lane"],
            "route": row["route"], "version": row["version"], "status": "passed" if passed else "failed",
            "terminal": descriptor(terminal_path, root=packet),
            "report": descriptor(report, root=packet) if report.is_file() else None,
            "validation": descriptor(child / "validation.json", root=packet),
            "rss_receipt": descriptor(child / "rss-receipt.json", root=packet),
            "rss": rss_result,
        }
        results.append(result)
        write_json_exclusive(packet / f"capture-state-{index:03d}.json", {
            "schema": "litchi.performance.0941.capture-progress.v1",
            "manifest_sha256": packet_manifest_sha, "schedule_sha256": manifest["schedule_sha256"],
            "completed": index, "total": len(manifest["schedule"]),
            "last": {key: result[key] for key in ("index", "id", "lane", "route", "version", "status")},
            "passed": sum(item["status"] == "passed" for item in results),
            "failed": sum(item["status"] != "passed" for item in results),
            "retry_count": 0,
        })
    state = {
        "schema": "litchi.performance.0941.capture-state.v1",
        "status": "completed" if all(item["status"] == "passed" for item in results) else "invalid",
        "ended_utc": utc_now(), "manifest_sha256": packet_manifest_sha,
        "schedule_sha256": manifest["schedule_sha256"], "children_expected": len(manifest["schedule"]),
        "children_run": len(results), "passed": sum(item["status"] == "passed" for item in results),
        "failed": sum(item["status"] != "passed" for item in results), "retry_count": 0,
        "children": results,
        "failure_policy": "Retain every child and invalidate interpretation on any failure; no retry or selective recapture.",
    }
    write_json_exclusive(packet / "capture-state.json", state)
    return {"schema": "litchi.performance.0941.capture-result.v1", "status": state["status"],
            "packet": str(packet), "children": len(results), "passed": state["passed"], "failed": state["failed"]}


def real_preflight_rows() -> list[dict[str, Any]]:
    """Select one frozen-shape row per lane/version/route for an authorized smoke cohort."""
    frozen = schedule_rows()
    rows: list[dict[str, Any]] = []
    for lane in ("normal", "allocator"):
        for version in ("baseline", "candidate"):
            for route in ("owned_vec", "owned_reader"):
                matches = [row for row in frozen if row["lane"] == lane
                           and row["version"] == version and row["route"] == route]
                require(matches, f"preflight schedule row missing: {lane}/{version}/{route}")
                row = dict(matches[0])
                row["id"] = f"real-preflight-{lane}-{version}-{route}"
                row["pair_id"] = row["id"]
                row["preflight"] = True
                rows.append(row)
    require(len(rows) == 8 and len({row["id"] for row in rows}) == 8,
            "real preflight row count changed")
    return rows


def real_preflight(args: argparse.Namespace) -> dict[str, Any]:
    """Run eight real workload children after the same explicit root authorization as capture."""
    require(args.manifest is not None and args.packet is not None,
            "real preflight requires --manifest and --packet")
    manifest_path = args.manifest.resolve()
    manifest, binding = validate_manifest(manifest_path, Path(__file__).resolve(), authorized=True)
    rows = real_preflight_rows()
    packet = args.packet.resolve()
    require(packet != REPO_ROOT and REPO_ROOT not in packet.parents,
            "preflight packet must be external to checkout")
    require(not packet.exists(), f"refusing to replace packet: {packet}")
    packet.mkdir(parents=True)
    (packet / "children").mkdir()
    (packet / "scripts").mkdir()
    script_copies: dict[str, dict[str, Any]] = {}
    for name, value in manifest["scripts"].items():
        source = Path(str(value["path"])).resolve()
        target = packet / "scripts" / source.name
        copy_custody_file(source, target)
        script_copies[name] = descriptor(target, root=packet)
    plan_target = packet / "plan.json"
    copy_custody_file(Path(str(manifest["plan"]["path"])).resolve(), plan_target)
    packet_manifest = dict(manifest)
    packet_manifest["capture_origin"] = str(packet)
    packet_manifest["packet_scripts"] = script_copies
    packet_manifest["packet_plan"] = descriptor(plan_target, root=packet)
    packet_manifest["packet_original_manifest"] = descriptor(manifest_path)
    packet_manifest["preflight"] = True
    packet_manifest["schedule"] = rows
    packet_manifest["schedule_sha256"] = schedule_sha256(rows)
    packet_manifest_path = packet / "manifest.json"
    write_json_exclusive(packet_manifest_path, packet_manifest)
    packet_manifest_sha = sha256_path(packet_manifest_path)
    write_json_exclusive(packet / "preflight-start.json", {
        "schema": "litchi.performance.0941.real-preflight-start.v1", "started_utc": utc_now(),
        "manifest_sha256": packet_manifest_sha, "schedule_sha256": packet_manifest["schedule_sha256"],
        "children_expected": len(rows), "terminal_first": True, "retry_policy": "none",
    })
    helper = load_module(Path(str(manifest["scripts"]["child_validation"]["path"])),
                         manifest["scripts"]["child_validation"]["sha256"], "0941 child helper")
    validator = load_module(Path(str(manifest["scripts"]["report_validator"]["path"])),
                            manifest["scripts"]["report_validator"]["sha256"], "0940 DOCX validator")
    results: list[dict[str, Any]] = []
    for index, row in enumerate(rows, start=1):
        child = packet / "children" / row["id"]
        child.mkdir(parents=True, exist_ok=False)
        report = child / "report.json"
        rss_path = child / "max-rss-kib.txt"
        binary = binary_descriptor_for_row(manifest, row)
        spec = helper.ChildSpec(
            command=child_command(row, binary, report, rss_path),
            cwd=Path(binding["source"][row["version"]]["root"]),
            environment=dict(manifest["child_environment"]),
            paths=helper.ChildPaths(report=report, stdout=child / "stdout.log", stderr=child / "stderr.log",
                                    terminal=child / "terminal.json", validation=child / "validation.json"),
            route=row["route"], label=row["id"], warmups=WARMUPS, samples=SAMPLES,
            expected_output_sha256=DOCX_OUTPUT_SHA256, binary=binary,
            report_contract=helper.REPORT_CONTRACT_DOCX_OWNED_WORKLOAD,
            deadline_seconds=DEADLINE_SECONDS, observation_seconds=OBSERVATION_SECONDS,
            max_report_bytes=MAX_REPORT_BYTES,
        )
        outcome = helper.run_child_and_validate(spec, validator)
        terminal_path = child / "terminal.json"
        require(terminal_path.is_file(), f"preflight helper did not publish terminal: {row['id']}")
        try:
            rss_result = parse_rss_sidecar(rss_path)
            sidecar_desc = descriptor(rss_path, root=packet)
            terminal_desc = descriptor(terminal_path, root=packet)
            rss_result.update({"path": sidecar_desc["path"], "bytes": sidecar_desc["bytes"],
                               "sha256": sidecar_desc["sha256"], "terminal_path": terminal_desc["path"],
                               "terminal_sha256": terminal_desc["sha256"]})
            rss_status = "passed"
        except Exception as error:
            rss_result = {"path": str(rss_path.resolve()), "status": "failed",
                          "error": f"{type(error).__name__}: {error}"}
            sidecar_desc = None
            terminal_desc = descriptor(terminal_path, root=packet)
            rss_status = "failed"
        rss_receipt_path = child / "rss-receipt.json"
        write_json_exclusive(rss_receipt_path, {
            "schema": "litchi.performance.0941.rss-receipt.v1", "status": rss_status,
            "child_id": row["id"], "terminal_first": True, "terminal": terminal_desc,
            "sidecar": sidecar_desc, "rss": rss_result,
        })
        passed = outcome["status"] == "passed" and rss_status == "passed"
        results.append({
            "index": index, "id": row["id"], "pair_id": row["pair_id"], "lane": row["lane"],
            "route": row["route"], "version": row["version"], "status": "passed" if passed else "failed",
            "terminal": terminal_desc, "report": descriptor(report, root=packet) if report.is_file() else None,
            "validation": descriptor(child / "validation.json", root=packet),
            "rss_receipt": descriptor(rss_receipt_path, root=packet), "rss": rss_result,
        })
    state = {
        "schema": "litchi.performance.0941.real-preflight-state.v1",
        "status": "completed" if all(item["status"] == "passed" for item in results) else "invalid",
        "ended_utc": utc_now(), "manifest_sha256": packet_manifest_sha,
        "schedule_sha256": packet_manifest["schedule_sha256"], "children_expected": len(rows),
        "children_run": len(results), "passed": sum(item["status"] == "passed" for item in results),
        "failed": sum(item["status"] != "passed" for item in results), "retry_count": 0,
        "children": results,
        "interpretation": "Preflight contract only; no latency/allocation/adoption claim and no substitution for the frozen 60-child cohort.",
    }
    write_json_exclusive(packet / "preflight-state.json", state)
    output = {"schema": "litchi.performance.0941.real-preflight-result.v1", "status": state["status"],
              "packet": str(packet), "children": len(results), "passed": state["passed"],
              "failed": state["failed"], "real_workloads_started": True, "adoption_claim": False,
              "schedule_sha256": packet_manifest["schedule_sha256"]}
    if args.output:
        write_json_exclusive(args.output.resolve(), output)
    return output


def _synthetic_packet_binary(packet: Path, lane: str, version: str) -> tuple[Path, dict[str, Any]]:
    identity = ALLOCATOR_BINARY_IDENTITY if lane == "allocator" else NORMAL_BINARY_IDENTITY
    path = packet / "binaries" / f"{lane}-{version}"
    wrapper = """#!/usr/bin/env python3
import pathlib
import sys

route = sys.argv[sys.argv.index('--route') + 1]
diagnostic = sys.argv[sys.argv.index('--diagnostic') + 1]
output = pathlib.Path(sys.argv[sys.argv.index('--output') + 1])
source = pathlib.Path(__file__).parent.parent / 'payloads' / f'{pathlib.Path(__file__).name}-{route}-{diagnostic}.json'
output.write_bytes(source.read_bytes())
"""
    write_bytes_exclusive(path, wrapper.encode("utf-8"))
    path.chmod(0o755)
    raw = descriptor(path, root=packet)
    metadata = {
        "path": str(path), "bytes": raw["bytes"], "sha256": raw["sha256"],
        "identity": identity, "binary_identity": identity, "source_revision": SOURCE_REVISION,
        "expected_report": expected_report(identity, lane == "allocator"),
        "expected_allocation_present": lane == "allocator",
        "expected_process_present": lane == "allocator",
    }
    return path, metadata


def _synthetic_packet_manifest(packet: Path, plan_path: Path, validator_path: Path,
                               binaries: Mapping[str, Mapping[str, Mapping[str, Any]]]) -> dict[str, Any]:
    plan_target = packet / "plan.json"
    copy_custody_file(plan_path, plan_target)
    validator_target = packet / "scripts" / "report_validator.py"
    copy_custody_file(validator_path, validator_target)
    return {
        "schema": "litchi.performance.0941.capture-manifest.v1",
        "synthetic_test": True,
        "capture_origin": str(packet),
        "source_revision": SOURCE_REVISION,
        "plan": descriptor(plan_target, root=packet),
        "plan_sha256": sha256_path(plan_target),
        "packet_plan": descriptor(plan_target, root=packet),
        "packet_scripts": {"report_validator": descriptor(validator_target, root=packet)},
        "child_environment": dict(CHILD_ENVIRONMENT),
        "child_cwd": "children",
        "schedule": schedule_rows(),
        "schedule_sha256": schedule_sha256(),
        "builds": binaries,
        "source": {},
        "status": "synthetic_test_packet",
    }


def _synthetic_packet_child(helper: Any, validator: Any, fixture: Any, packet: Path,
                            row: Mapping[str, Any], binary_metadata: Mapping[str, Any],
                            index: int) -> dict[str, Any]:
    child = packet / "children" / str(row["id"])
    child.mkdir(parents=True, exist_ok=False)
    report_path = child / "report.json"
    rss_path = child / "max-rss-kib.txt"
    payload = fixture.synthetic_report(
        route=str(row["route"]), case="workload", diagnostic=bool(row["diagnostic"]),
        warmups=WARMUPS, samples=SAMPLES, binary_identity=str(binary_metadata["identity"]),
    )
    payload_path = packet / "payloads" / (
        f"{Path(str(binary_metadata['path'])).name}-{row['route']}-"
        f"{'on' if row['diagnostic'] else 'off'}.json"
    )
    if not payload_path.exists():
        write_bytes_exclusive(payload_path, json.dumps(payload, sort_keys=True).encode("utf-8"))
    command = child_command(row, binary_metadata, report_path, rss_path)
    spec = helper.ChildSpec(
        command=command, cwd=packet / "children", environment=dict(CHILD_ENVIRONMENT),
        paths=helper.ChildPaths(report=report_path, stdout=child / "stdout.log", stderr=child / "stderr.log",
                                terminal=child / "terminal.json", validation=child / "validation.json"),
        route=str(row["route"]), label=str(row["id"]), warmups=WARMUPS, samples=SAMPLES,
        expected_output_sha256=DOCX_OUTPUT_SHA256, binary=binary_metadata,
        report_contract=helper.REPORT_CONTRACT_DOCX_OWNED_WORKLOAD,
        deadline_seconds=30.0, observation_seconds=0.05, max_report_bytes=MAX_REPORT_BYTES,
    )
    outcome = helper.run_child_and_validate(spec, validator)
    terminal_path = child / "terminal.json"
    parsed = parse_rss_sidecar(rss_path)
    sidecar_desc = descriptor(rss_path, root=packet)
    terminal_desc = descriptor(terminal_path, root=packet)
    parsed.update({"path": sidecar_desc["path"], "bytes": sidecar_desc["bytes"],
                   "sha256": sidecar_desc["sha256"], "terminal_path": terminal_desc["path"],
                   "terminal_sha256": terminal_desc["sha256"]})
    rss_receipt_path = child / "rss-receipt.json"
    write_json_exclusive(rss_receipt_path, {
        "schema": "litchi.performance.0941.rss-receipt.v1", "status": "passed",
        "child_id": row["id"], "terminal_first": True, "terminal": terminal_desc,
        "sidecar": sidecar_desc, "rss": parsed,
    })
    report_path = child / "report.json"
    validation_path = child / "validation.json"
    passed = outcome["status"] == "passed"
    return {
        "index": index, "id": row["id"], "pair_id": row["pair_id"], "lane": row["lane"],
        "route": row["route"], "version": row["version"], "status": "passed" if passed else "failed",
        "terminal": terminal_desc, "report": descriptor(report_path, root=packet),
        "validation": descriptor(validation_path, root=packet),
        "rss_receipt": descriptor(rss_receipt_path, root=packet), "rss": parsed,
    }


def _copy_packet_for_negative(source: Path, destination: Path) -> Path:
    require(not destination.exists(), f"synthetic negative destination exists: {destination}")
    shutil.copytree(source, destination)
    return destination


def _expect_replay_rejection(packet: Path, label: str, mutate: Any) -> dict[str, Any]:
    case = packet.parent / f"negative-{label}"
    _copy_packet_for_negative(packet, case)
    mutate(case)
    try:
        replay_packet(case)
    except Exception as error:
        return {"label": label, "rejected": True, "error": f"{type(error).__name__}: {error}"}
    raise RunnerError(f"synthetic negative control was accepted: {label}")


def synthetic_packet(args: argparse.Namespace) -> dict[str, Any]:
    """Exercise the full 60-child packet/reducer and custody negative controls without Rust."""
    plan_path = args.plan.resolve()
    plan = validate_plan(plan_path)
    del plan
    helper_path = args.helper.resolve()
    validator_path = args.validator.resolve()
    fixture_path = args.synthetic.resolve()
    helper = load_module(helper_path, sha256_path(helper_path), "0941 child helper")
    validator = load_module(validator_path, sha256_path(validator_path), "0940 DOCX validator")
    fixture = load_module(fixture_path, sha256_path(fixture_path), "0940 synthetic fixture provider")
    root = args.scratch.resolve() if args.scratch else ARTIFACT_ROOT / "synthetic-packet" / f"run-{os.getpid()}"
    require(not root.exists(), f"synthetic packet destination exists: {root}")
    root.mkdir(parents=True)
    packet = root / "packet"
    packet.mkdir()
    (packet / "children").mkdir()
    (packet / "scripts").mkdir()
    (packet / "binaries").mkdir()
    (packet / "payloads").mkdir()
    binary_entries: dict[str, dict[str, dict[str, Any]]] = {"normal": {}, "allocator": {}}
    binary_metadata: dict[tuple[str, str], dict[str, Any]] = {}
    for lane in ("normal", "allocator"):
        for version in ("baseline", "candidate"):
            path, metadata = _synthetic_packet_binary(packet, lane, version)
            del path
            binary_entries[lane][version] = {
                "binary": descriptor(packet / "binaries" / f"{lane}-{version}", root=packet),
                "binary_identity": metadata["binary_identity"],
            }
            binary_metadata[(lane, version)] = metadata
    manifest = _synthetic_packet_manifest(packet, plan_path, validator_path, binary_entries)
    manifest_path = packet / "manifest.json"
    write_json_exclusive(manifest_path, manifest)
    manifest_sha = sha256_path(manifest_path)
    results: list[dict[str, Any]] = []
    for index, row in enumerate(schedule_rows(), start=1):
        results.append(_synthetic_packet_child(
            helper, validator, fixture, packet, row,
            binary_metadata[(str(row["binary_lane"]), str(row["version"]))], index,
        ))
    state = {
        "schema": "litchi.performance.0941.capture-state.v1", "status": "completed",
        "ended_utc": utc_now(), "manifest_sha256": manifest_sha, "schedule_sha256": schedule_sha256(),
        "children_expected": len(results), "children_run": len(results),
        "passed": sum(item["status"] == "passed" for item in results),
        "failed": sum(item["status"] != "passed" for item in results), "retry_count": 0,
        "children": results, "failure_policy": "synthetic test only; no Rust workload claim",
    }
    write_json_exclusive(packet / "capture-state.json", state)
    first_reduced = replay_packet(packet)
    moved = root / "moved-packet"
    shutil.move(str(packet), str(moved))
    moved_reduced = replay_packet(moved)

    def mutate_duplicate(case: Path) -> None:
        path = case / "capture-state.json"
        value = read_json(path)
        value["children"][-1]["id"] = value["children"][0]["id"]
        write_json_exclusive(case / "capture-state-mutated.json", value)
        path.unlink()
        (case / "capture-state-mutated.json").rename(path)

    def mutate_wrong_binary(case: Path) -> None:
        with (case / "binaries" / "normal-baseline").open("ab") as stream:
            stream.write(b"tampered")

    def mutate_environment(case: Path) -> None:
        state_value = read_json(case / "capture-state.json")
        item = state_value["children"][0]
        terminal_path = case / safe_relative(item["terminal"]["path"], "negative terminal")
        terminal = read_json(terminal_path)
        terminal["environment"]["PATH"] = "/tampered"
        terminal_path.unlink()
        write_json_exclusive(terminal_path, terminal)
        terminal_desc = descriptor(terminal_path, root=case)
        item["terminal"] = terminal_desc
        receipt_path = case / safe_relative(item["rss_receipt"]["path"], "negative RSS receipt")
        receipt = read_json(receipt_path)
        receipt["terminal"] = terminal_desc
        receipt["rss"]["terminal_path"] = terminal_desc["path"]
        receipt["rss"]["terminal_sha256"] = terminal_desc["sha256"]
        receipt_path.unlink()
        write_json_exclusive(receipt_path, receipt)
        item["rss_receipt"] = descriptor(receipt_path, root=case)
        state_value["children"][0] = item
        (case / "capture-state.json").unlink()
        write_json_exclusive(case / "capture-state.json", state_value)

    def mutate_rss_digest(case: Path) -> None:
        state_value = read_json(case / "capture-state.json")
        item = state_value["children"][0]
        receipt_path = case / safe_relative(item["rss_receipt"]["path"], "negative RSS receipt")
        receipt = read_json(receipt_path)
        sidecar_path = case / safe_relative(receipt["sidecar"]["path"], "negative RSS sidecar")
        sidecar_path.write_bytes(b"12346\n")
        fresh = descriptor(sidecar_path, root=case)
        receipt["sidecar"] = fresh
        receipt["rss"]["sha256"] = "0" * 64
        receipt_path.unlink()
        write_json_exclusive(receipt_path, receipt)
        item["rss_receipt"] = descriptor(receipt_path, root=case)
        item["rss"] = receipt["rss"]
        (case / "capture-state.json").unlink()
        write_json_exclusive(case / "capture-state.json", state_value)

    def mutate_report_digest(case: Path) -> None:
        state_value = read_json(case / "capture-state.json")
        item = state_value["children"][0]
        report_path = case / safe_relative(item["report"]["path"], "negative report")
        report_path.write_bytes(report_path.read_bytes() + b" ")

    negatives = [
        _expect_replay_rejection(moved, "wrong-binary", mutate_wrong_binary),
        _expect_replay_rejection(moved, "environment", mutate_environment),
        _expect_replay_rejection(moved, "rss-digest", mutate_rss_digest),
        _expect_replay_rejection(moved, "report-digest", mutate_report_digest),
        _expect_replay_rejection(moved, "duplicate-id", mutate_duplicate),
    ]
    output = {
        "schema": "litchi.performance.0941.synthetic-packet-test.v1",
        "status": "passed", "real_workloads_started": False,
        "rust_performance_claim": False, "children": len(results),
        "packet": str(moved), "capture_origin": str(packet), "moved_packet": str(moved),
        "original_replay_children": first_reduced["children"],
        "moved_replay_children": moved_reduced["children"],
        "negative_controls": negatives,
    }
    if args.output:
        write_json_exclusive(args.output.resolve(), output)
    return output


def midpoint(left: int | float, right: int | float) -> float:
    return (left + right) / 2.0


def nearest_rank(values: Sequence[int | float], fraction: float) -> float:
    ordered = sorted(values)
    index = max(0, min(len(ordered) - 1, int((fraction * len(ordered) + 0.999999999999) - 1)))
    return float(ordered[index])


def stats(values: Sequence[int | float]) -> dict[str, Any]:
    ordered = sorted(values)
    require(ordered, "empty statistics vector")
    p50 = midpoint(ordered[(len(ordered) - 1) // 2], ordered[len(ordered) // 2])
    return {"count": len(ordered), "min": ordered[0], "p50": p50, "p95": nearest_rank(ordered, .95),
            "p99": nearest_rank(ordered, .99), "max": ordered[-1],
            "mean": statistics.fmean(ordered), "spread": ((ordered[-1] - ordered[0]) / ordered[0]) if ordered[0] else None}


def report_vectors(report: Mapping[str, Any], key: str) -> list[int]:
    values = report.get(key)
    require(isinstance(values, list) and len(values) == SAMPLES, f"report {key} vector changed")
    require(all(type(value) is int and value >= 0 for value in values), f"report {key} values changed")
    return [int(value) for value in values]


def bootstrap_median(ratios: Sequence[float], *, seed: int = 941, draws: int = 10000) -> dict[str, Any]:
    require(len(ratios) == BLOCKS, "reader ratio count changed")
    randomizer = random.Random(seed)
    samples = [statistics.median(randomizer.choice(list(ratios)) for _ in ratios) for _ in range(draws)]
    ordered = sorted(samples)
    return {"median": float(statistics.median(ratios)), "lower_95": ordered[249], "upper_95": ordered[9749],
            "seed": seed, "draws": draws, "statistic": "median of paired baseline_p50/candidate_p50 ratios"}


def parse_packet_descriptor(packet: Path, value: Any, label: str) -> Path:
    require(isinstance(value, Mapping), f"{label} descriptor missing")
    path = packet / safe_relative(value.get("path"), label)
    actual = descriptor(path, root=packet)
    require(actual["bytes"] == value.get("bytes") and actual["sha256"] == value.get("sha256"), f"{label} changed")
    return path


def normalize_capture_path(value: Any, origin: Path, packet: Path, label: str,
                           *, packet_owned: bool = False) -> str:
    """Map a terminal's original absolute packet path into the selected packet root."""
    require(isinstance(value, str) and value, f"{label} path missing")
    path = Path(value)
    if not path.is_absolute():
        raise RunnerError(f"{label} terminal path is not absolute")
    origin = origin.resolve()
    packet = packet.resolve()
    try:
        relative = path.relative_to(origin)
    except ValueError:
        require(not packet_owned, f"{label} path escapes capture origin")
        return str(path)
    return str(packet / relative)


def normalize_terminal_descriptor(value: Any, *, origin: Path, packet: Path,
                                  label: str) -> tuple[Path, dict[str, Any]]:
    require(isinstance(value, Mapping), f"{label} terminal descriptor missing")
    normalized = dict(value)
    normalized["path"] = normalize_capture_path(value.get("path"), origin, packet, label,
                                                 packet_owned=True)
    path, actual = bound_descriptor(normalized, label)
    require(actual == {"path": normalized["path"], "bytes": value.get("bytes"),
                       "sha256": value.get("sha256")},
            f"{label} descriptor changed")
    return path, actual


def replay_packet(packet_path: Path, output: Path | None = None) -> dict[str, Any]:
    packet = packet_path.resolve()
    state = read_json(packet / "capture-state.json")
    require(state.get("schema") == "litchi.performance.0941.capture-state.v1", "capture state schema changed")
    require(state.get("status") == "completed", "cannot reduce an invalid capture")
    manifest = read_json(packet / "manifest.json")
    require(manifest.get("schema") == "litchi.performance.0941.capture-manifest.v1", "packet manifest schema changed")
    require(state.get("manifest_sha256") == sha256_path(packet / "manifest.json"), "packet manifest binding changed")
    require(state.get("schedule_sha256") == schedule_sha256(), "packet schedule hash changed")
    capture_origin_value = manifest.get("capture_origin")
    require(isinstance(capture_origin_value, str) and Path(capture_origin_value).is_absolute(),
            "packet capture origin missing")
    capture_origin = Path(capture_origin_value).resolve()
    packet_plan = manifest.get("packet_plan")
    packet_plan_path = parse_packet_descriptor(packet, packet_plan, "packet plan")
    require(manifest.get("plan_sha256") == manifest["packet_plan"]["sha256"],
            "packet plan hash binding changed")
    validate_plan(packet_plan_path)
    rows = schedule_rows()
    children = state.get("children")
    require(isinstance(children, list) and len(children) == len(rows), "packet child count changed")
    child_ids = [item.get("id") if isinstance(item, Mapping) else None for item in children]
    require(all(isinstance(child_id, str) and child_id for child_id in child_ids),
            "packet child ID is missing")
    require(len(set(child_ids)) == len(child_ids), "packet child IDs are duplicated")
    result_by_id = {item["id"]: item for item in children}
    require(set(result_by_id) == {row["id"] for row in rows}, "packet child IDs differ from schedule")
    packet_scripts = manifest.get("packet_scripts")
    require(isinstance(packet_scripts, Mapping), "packet script custody missing")
    validator_path = parse_packet_descriptor(packet, packet_scripts.get("report_validator"),
                                             "packet report validator")
    validator_desc = packet_scripts["report_validator"]
    validator = load_module(validator_path, validator_desc["sha256"], "packet DOCX validator")
    normal_pairs: dict[tuple[str, str], dict[str, Any]] = {}
    alloc_pairs: dict[tuple[str, str], dict[str, Any]] = {}
    child_metrics: list[dict[str, Any]] = []
    for row in rows:
        item = result_by_id[row["id"]]
        require(item.get("status") == "passed", f"packet child failed: {row['id']}")
        require(item.get("pair_id") == row["pair_id"] and item.get("lane") == row["lane"]
                and item.get("route") == row["route"] and item.get("version") == row["version"],
                f"packet child schedule binding changed: {row['id']}")
        terminal_path = parse_packet_descriptor(packet, item["terminal"], f"terminal {row['id']}")
        terminal = read_json(terminal_path)
        require(terminal.get("status") == "terminal" and terminal.get("exit_code") == 0
                and terminal.get("deadline_exceeded") is False and terminal.get("process_group_closed") is True,
                f"terminal status changed: {row['id']}")
        report_path = parse_packet_descriptor(packet, item["report"], f"report {row['id']}")
        report = read_json(report_path)
        binary = _binary_descriptor_for_row(manifest, row, packet=packet)
        rss_receipt_relative = safe_relative(item["rss_receipt"].get("path"), f"RSS receipt {row['id']}")
        expected_rss_path = packet / rss_receipt_relative.parent / "max-rss-kib.txt"
        require(terminal.get("environment") == manifest.get("child_environment"),
                f"child environment binding changed: {row['id']}")
        require(terminal.get("route") == row["route"] and terminal.get("label") == row["id"],
                f"child terminal route/label changed: {row['id']}")
        require(terminal.get("report_contract") == "docx-owned-workload"
                and terminal.get("expected_output_sha256") == DOCX_OUTPUT_SHA256,
                f"child terminal contract changed: {row['id']}")
        terminal_binary = terminal.get("binary")
        require(isinstance(terminal_binary, Mapping), f"child binary binding missing: {row['id']}")
        normalized_binary_path = normalize_capture_path(terminal_binary.get("path"), capture_origin,
                                                        packet, f"child binary {row['id']}")
        require(normalized_binary_path == binary["path"]
                and terminal_binary.get("identity") == binary["identity"]
                and terminal_binary.get("source_revision") == SOURCE_REVISION
                and terminal_binary.get("bytes") == binary["bytes"]
                and terminal_binary.get("sha256") == binary["sha256"],
                f"child binary binding changed: {row['id']}")
        terminal_report = terminal.get("report")
        report_terminal_path, report_terminal_actual = normalize_terminal_descriptor(
            terminal_report, origin=capture_origin, packet=packet, label=f"child report {row['id']}"
        )
        require(report_terminal_path == report_path
                and report_terminal_actual["bytes"] == item["report"].get("bytes")
                and report_terminal_actual["sha256"] == item["report"].get("sha256"),
                f"child terminal report binding changed: {row['id']}")
        for stream in ("stdout", "stderr"):
            stream_value = terminal.get(stream)
            stream_path, stream_actual = normalize_terminal_descriptor(
                stream_value, origin=capture_origin, packet=packet, label=f"child {stream} {row['id']}"
            )
            del stream_path
            del stream_actual
        child_cwd_value = manifest.get("child_cwd")
        if child_cwd_value is None:
            expected_cwd = Path(manifest["source"][row["version"]]["root"])
        else:
            expected_cwd = packet / safe_relative(child_cwd_value, "packet child cwd")
        normalized_cwd = normalize_capture_path(terminal.get("cwd"), capture_origin, packet,
                                                f"child cwd {row['id']}")
        require(normalized_cwd == str(expected_cwd), f"child cwd binding changed: {row['id']}")
        actual_argv = terminal.get("argv")
        require(isinstance(actual_argv, list) and all(isinstance(value, str) for value in actual_argv),
                f"child command missing: {row['id']}")
        normalized_argv = [
            normalize_capture_path(value, capture_origin, packet, f"child argv {row['id']}")
            if value.startswith("/") else value
            for value in actual_argv
        ]
        require(tuple(normalized_argv) == child_command(row, binary, report_path, expected_rss_path),
                f"child command binding changed: {row['id']}")
        validated = validator.validate_report(report, row["route"], warmups=WARMUPS, samples=SAMPLES,
                                              label=str(report_path), binary=binary)
        rss_receipt_path = parse_packet_descriptor(packet, item["rss_receipt"], f"RSS receipt {row['id']}")
        rss_receipt = read_json(rss_receipt_path)
        require(rss_receipt.get("status") == "passed" and rss_receipt.get("terminal") == item["terminal"],
                f"RSS receipt binding changed: {row['id']}")
        sidecar_path = parse_packet_descriptor(packet, rss_receipt.get("sidecar"),
                                               f"RSS sidecar {row['id']}")
        rss = rss_receipt.get("rss")
        require(isinstance(rss, Mapping) and type(rss.get("max_rss_kib")) is int
                and rss.get("unit") == "KiB" and rss.get("format") == "%M",
                f"RSS metric changed: {row['id']}")
        require(rss == item.get("rss"), f"RSS result binding changed: {row['id']}")
        require(rss.get("path") == rss_receipt["sidecar"].get("path")
                and rss.get("bytes") == rss_receipt["sidecar"].get("bytes")
                and rss.get("sha256") == rss_receipt["sidecar"].get("sha256"),
                f"RSS sidecar descriptor binding changed: {row['id']}")
        raw_rss = parse_rss_sidecar(sidecar_path)
        for key in ("bytes", "sha256", "unit", "format", "max_rss_kib", "max_rss_bytes"):
            require(raw_rss[key] == rss[key], f"RSS raw sidecar digest/metric changed: {row['id']}")
        require(rss.get("terminal_path") == item["terminal"]["path"]
                and rss.get("terminal_sha256") == item["terminal"]["sha256"],
                f"RSS terminal binding changed: {row['id']}")
        operation = list(validated["elapsed_ns"])
        wall = report.get("wall_elapsed_ns", {}).get("samples")
        require(isinstance(wall, list) and len(wall) == SAMPLES, f"wall vector changed: {row['id']}")
        metric = {"id": row["id"], "pair_id": row["pair_id"], "lane": row["lane"], "route": row["route"],
                  "version": row["version"], "operation": stats(operation), "wall": stats(wall),
                  "max_rss_kib": rss["max_rss_kib"], "max_rss_bytes": rss["max_rss_bytes"]}
        if row["lane"] == "normal":
            normal_pairs.setdefault((row["route"], row["pair_id"]), {})[row["version"]] = metric
        else:
            allocation = report.get("allocation")
            require(isinstance(allocation, Mapping), f"allocator report missing allocation: {row['id']}")
            opened = allocation.get("open")
            require(isinstance(opened, list) and len(opened) == SAMPLES, f"allocator open vector changed: {row['id']}")
            for sample in opened:
                require(isinstance(sample, Mapping) and sample.get("status") == "measured", f"allocator open unavailable: {row['id']}")
            fields = ("allocated_bytes", "allocation_calls", "reallocation_calls", "live_bytes_after",
                      "peak_live_bytes_after", "region_peak_live_bytes")
            metric["allocation_open"] = {field: stats([int(sample[field]) for sample in opened]) for field in fields}
            alloc_pairs.setdefault((row["route"], row["pair_id"]), {})[row["version"]] = metric
        child_metrics.append(metric)
    normal_summary: dict[str, Any] = {"pairs": [], "safety_pass": True}
    reader_ratios: list[float] = []
    review_flags: list[dict[str, Any]] = []
    for (route, pair_id), values in sorted(normal_pairs.items()):
        require(set(values) == {"baseline", "candidate"}, f"normal pair incomplete: {pair_id}")
        base, cand = values["baseline"], values["candidate"]
        ratios = {field: (float(cand[field]["p50"]) / float(base[field]["p50"]))
                  for field in ("operation", "wall")}
        for field in ("p95", "p99", "spread"):
            base_value = float(base["operation"][field])
            candidate_value = float(cand["operation"][field])
            ratio = (candidate_value / base_value) if base_value else (1.0 if candidate_value == 0 else None)
            if ratio is not None and ratio > 1.05:
                review_flags.append({"kind": "normal_tail_or_spread", "route": route, "pair_id": pair_id,
                                     "metric": field, "candidate_over_baseline": ratio})
        rss_ratio = (float(cand["max_rss_bytes"]) / float(base["max_rss_bytes"])) if base["max_rss_bytes"] else None
        if rss_ratio is not None and rss_ratio > 1.05:
            review_flags.append({"kind": "normal_max_rss", "route": route, "pair_id": pair_id,
                                 "candidate_over_baseline": rss_ratio})
        safe = ratios["operation"] <= 1.05
        normal_summary["safety_pass"] = normal_summary["safety_pass"] and safe
        if not safe:
            review_flags.append({"kind": "normal_operation_p50_safety", "route": route, "pair_id": pair_id,
                                 "candidate_over_baseline": ratios["operation"]})
        if route == "owned_reader":
            reader_ratios.append(float(base["operation"]["p50"]) / float(cand["operation"]["p50"]))
        normal_summary["pairs"].append({"route": route, "pair_id": pair_id,
                                        "baseline_operation": base["operation"], "candidate_operation": cand["operation"],
                                        "baseline_wall": base["wall"], "candidate_wall": cand["wall"],
                                        "baseline_max_rss_bytes": base["max_rss_bytes"], "candidate_max_rss_bytes": cand["max_rss_bytes"],
                                        "candidate_over_baseline_operation_p50": ratios["operation"],
                                        "baseline_over_candidate_operation_p50": 1.0 / ratios["operation"]})
    reader_bootstrap = bootstrap_median(reader_ratios)
    normal_summary["reader_primary"] = reader_bootstrap
    normal_summary["reader_primary_pass"] = reader_bootstrap["median"] >= 1.05 and reader_bootstrap["lower_95"] > 1.0
    allocation_summary: dict[str, Any] = {"pairs": [], "reader_gate_pass": True, "memory_flags": []}
    for (route, pair_id), values in sorted(alloc_pairs.items()):
        require(set(values) == {"baseline", "candidate"}, f"allocator pair incomplete: {pair_id}")
        base, cand = values["baseline"], values["candidate"]
        bopen, copen = base["allocation_open"], cand["allocation_open"]
        bytes_reduction = 1.0 - float(copen["allocated_bytes"]["p50"]) / float(bopen["allocated_bytes"]["p50"]) if bopen["allocated_bytes"]["p50"] else None
        reader_gate = (route != "owned_reader" or (
            bytes_reduction is not None and bytes_reduction >= .10
            and copen["allocation_calls"]["p50"] < bopen["allocation_calls"]["p50"]
            and copen["reallocation_calls"]["p50"] < bopen["reallocation_calls"]["p50"]))
        allocation_summary["reader_gate_pass"] = allocation_summary["reader_gate_pass"] and reader_gate
        if route == "owned_reader" and not reader_gate:
            review_flags.append({"kind": "allocator_reader_gate", "route": route, "pair_id": pair_id,
                                 "requested_bytes_reduction": bytes_reduction,
                                 "baseline_calls": bopen["allocation_calls"]["p50"], "candidate_calls": copen["allocation_calls"]["p50"],
                                 "baseline_reallocations": bopen["reallocation_calls"]["p50"], "candidate_reallocations": copen["reallocation_calls"]["p50"]})
        memory = {}
        for field in ("live_bytes_after", "peak_live_bytes_after", "region_peak_live_bytes"):
            bvalue, cvalue = bopen[field]["p50"], copen[field]["p50"]
            ratio = (float(cvalue) / float(bvalue)) if bvalue else (1.0 if cvalue == 0 else None)
            memory[field] = {"baseline": bvalue, "candidate": cvalue, "candidate_over_baseline": ratio}
            if ratio is not None and ratio > 1.05:
                flag = {"kind": "allocator_memory", "route": route, "pair_id": pair_id, "metric": field,
                        "candidate_over_baseline": ratio}
                allocation_summary["memory_flags"].append(flag)
                review_flags.append(flag)
        rss_ratio = (float(cand["max_rss_bytes"]) / float(base["max_rss_bytes"])) if base["max_rss_bytes"] else None
        if rss_ratio is not None and rss_ratio > 1.05:
            flag = {"kind": "allocator_max_rss", "route": route, "pair_id": pair_id,
                    "candidate_over_baseline": rss_ratio}
            allocation_summary["memory_flags"].append(flag)
            review_flags.append(flag)
        allocation_summary["pairs"].append({"route": route, "pair_id": pair_id,
                                             "baseline_open": bopen, "candidate_open": copen,
                                             "requested_bytes_reduction": bytes_reduction,
                                             "reader_gate_pass": reader_gate, "memory": memory,
                                             "baseline_max_rss_bytes": base["max_rss_bytes"],
                                             "candidate_max_rss_bytes": cand["max_rss_bytes"]})
    allocation_summary["memory_safe_all_routes"] = not allocation_summary["memory_flags"]
    allocation_summary["adoption_pass"] = (normal_summary["safety_pass"] and allocation_summary["reader_gate_pass"]
                                             and allocation_summary["memory_safe_all_routes"]
                                             and not review_flags)
    normal_summary["adoption_pass"] = normal_summary["safety_pass"] and normal_summary["reader_primary_pass"] and not review_flags
    result = {
        "schema": "litchi.performance.0941.reduced-evidence.v1",
        "status": "passed", "children": len(child_metrics), "normal": normal_summary,
        "allocation": allocation_summary, "review_flags": review_flags,
        "docx_latency_adoption_path": normal_summary["adoption_pass"],
        "docx_allocation_adoption_path": allocation_summary["adoption_pass"],
        "overall_adoption": "pending required XLSX 0779 controls, independent review, and root decision",
        "no_absolute_gauge_sums": True,
        "source_revision": SOURCE_REVISION,
        "schedule_sha256": schedule_sha256(rows),
    }
    if output is not None:
        require(output.resolve() != REPO_ROOT and REPO_ROOT not in output.resolve().parents,
                "reduced evidence must be external to checkout")
        write_json_exclusive(output.resolve(), result)
        result["output"] = descriptor(output.resolve())
    return result


def freeze(args: argparse.Namespace) -> dict[str, Any]:
    manifest = make_manifest(args)
    output = args.manifest.resolve()
    require(not output.exists(), f"refusing to replace manifest: {output}")
    write_json_exclusive(output, manifest)
    return {"schema": "litchi.performance.0941.freeze-result.v1", "status": "prospective_unfinished",
            "manifest": descriptor(output), "children": len(manifest["schedule"]), "capture_authorized": False}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    modes = parser.add_mutually_exclusive_group(required=True)
    modes.add_argument("--check-schedule", action="store_true")
    modes.add_argument("--check-source", action="store_true")
    modes.add_argument("--preflight", action="store_true")
    modes.add_argument("--real-preflight", action="store_true")
    modes.add_argument("--synthetic-packet", action="store_true")
    modes.add_argument("--freeze", action="store_true")
    modes.add_argument("--capture", action="store_true")
    modes.add_argument("--replay", action="store_true")
    parser.add_argument("--plan", type=Path, default=PLAN_DEFAULT)
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--packet", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--helper", type=Path, default=HELPER_DEFAULT)
    parser.add_argument("--validator", type=Path, default=VALIDATOR_DEFAULT)
    parser.add_argument("--synthetic", type=Path, default=SYNTHETIC_DEFAULT)
    parser.add_argument("--scratch", type=Path)
    for version in ("baseline", "candidate"):
        for lane in ("normal", "allocator"):
            parser.add_argument(f"--{version}-{lane}-terminal", dest=f"{version}_{lane}_terminal", type=Path)
            parser.add_argument(f"--{version}-{lane}-binary", dest=f"{version}_{lane}_binary", type=Path)
            parser.add_argument(f"--{version}-{lane}-start", dest=f"{version}_{lane}_start", type=Path)
            parser.add_argument(f"--{version}-{lane}-receipt", dest=f"{version}_{lane}_receipt", type=Path)
    args = parser.parse_args(argv)
    try:
        if args.check_schedule:
            rows = schedule_rows()
            result = {"schema": "litchi.performance.0941.schedule.v1", "status": "passed",
                      "children": len(rows), "normal_children": 48, "allocator_children": 12,
                      "warmups": WARMUPS, "samples": SAMPLES, "cpu": CPU,
                      "schedule_sha256": schedule_sha256(rows), "capture_started": False}
        elif args.check_source:
            result = check_source_contract(args.plan.resolve())
        elif args.preflight:
            result = preflight(args)
        elif args.real_preflight:
            result = real_preflight(args)
        elif args.synthetic_packet:
            result = synthetic_packet(args)
        elif args.freeze:
            require(args.manifest is not None, "freeze requires --manifest")
            required = [f"{version}_{lane}_{kind}" for version in ("baseline", "candidate")
                        for lane in ("normal", "allocator")
                        for kind in ("terminal", "binary", "start", "receipt")]
            require(all(getattr(args, name) is not None for name in required),
                    "freeze requires exact terminal, binary, start, and binary receipt paths for all builds")
            result = freeze(args)
        elif args.capture:
            require(args.manifest is not None and args.packet is not None, "capture requires --manifest and --packet")
            result = capture(args)
        else:
            require(args.packet is not None, "replay requires --packet")
            result = replay_packet(args.packet.resolve(), args.output.resolve() if args.output else None)
        print(json.dumps(result, sort_keys=True))
        return 0
    except Exception as error:
        print(json.dumps({"schema": "litchi.performance.0941.runner-error.v1", "status": "failed",
                          "error": f"{type(error).__name__}: {error}"}, sort_keys=True))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
