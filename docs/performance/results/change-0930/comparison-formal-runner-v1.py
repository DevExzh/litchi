#!/usr/bin/env python3
"""Freeze, run, replay, and reduce the prospective 0930 paired cohort.

The runner is deliberately inert until ``capture`` is given a manifest whose
root authorization flag is true and ``LITCHI_PERF_0930_CAPTURE_APPROVED=1``.
``freeze`` only verifies custody and writes an unauthorized manifest.  Raw
reports and process logs belong in an external packet; repository records keep
descriptors and scalar validation summaries.

The formal cohort is six alternating blocks, four routes, and two versions per
route (48 children).  Each child is terminally recorded before its report is
opened.  An observation timeout polls the same process; only a confirmed
process deadline terminates it, and no failed child is retried.
"""

from __future__ import annotations

import argparse
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
PLAN_DEFAULT = REPO_ROOT / "docs/performance/results/change-0930/comparison-plan-v1.json"
CORRECTION_DEFAULT = REPO_ROOT / "docs/performance/results/change-0930/comparison-formal-gate-correction-v1.json"
BASE_REVISION = "aa077c030819c3019288e350f34959666ab7204f"
CASE = "docx_source_backed_existing_section_layout_edit_save"
ROUTES = ("generic_read_at", "owned_vec", "owned_reader", "range_short")
PHASES = ("preparation_ns", "open_ns", "query_ns", "edit_ns", "commit_ns", "publication_ns")
CPU = 12
WARMUPS = 5
SAMPLES = 50
BLOCKS = 6
CHILDREN = BLOCKS * len(ROUTES) * 2
DEADLINE_SECONDS = 900.0
OBSERVE_SECONDS = 1.0
MAX_ARTIFACT_BYTES = 1 << 20
CHILD_ENV_BASE = {
    "PATH": "/usr/bin:/bin",
    "LC_ALL": "C",
    "LANG": "C",
    "TZ": "UTC",
    "RUST_BACKTRACE": "0",
}
EXPECTED_CHANGED_PATHS = (
    "crates/litchi-docx/src/section/layout.rs",
    "crates/litchi-docx/src/section/inventory.rs",
    "crates/litchi-docx/src/leniency.rs",
    "crates/litchi-docx/src/source_backed.rs",
    "crates/litchi-docx/tests/source_backed_sections.rs",
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
    spec = importlib.util.spec_from_file_location("litchi_perf_0930_comparison", path)
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
                    "id": f"formal-block-{block:02d}-{position:02d}-{route}-{version}",
                    "pair_id": f"formal-block-{block:02d}-{position:02d}-{route}",
                    "block": block,
                    "position": position,
                    "version_position": version_position,
                    "route": route,
                    "version": version,
                    "cpu": CPU,
                    "warmups": WARMUPS,
                    "samples": SAMPLES,
                })
    require(len(rows) == CHILDREN, "formal schedule child count changed")
    for route in ROUTES:
        require(sum(row["route"] == route for row in rows) == BLOCKS * 2,
                f"formal schedule route count changed: {route}")
        for version in ("baseline", "candidate"):
            require(sum(row["route"] == route and row["version"] == version for row in rows) == BLOCKS,
                    f"formal schedule version count changed: {route}/{version}")
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
    receipt = read_json(receipt_path)
    schema = str(receipt.get("schema", ""))
    require(schema.startswith("litchi.performance.0930.candidate-source-receipt")
            or schema.startswith("litchi.performance.0925"),
            f"unsupported source receipt schema: {receipt.get('schema')!r}")
    if candidate is None:
        candidate = schema.startswith("litchi.performance.0930.candidate-source-receipt")
    require(receipt.get("base_revision") == BASE_REVISION, "source base revision changed")
    source_root = Path(str(receipt.get("source_root", ""))).resolve()
    require(source_root.is_dir(), f"source root missing: {source_root}")
    listing_path = Path(str(receipt.get("full_source_listing", ""))).resolve()
    require(listing_path.is_file() and listing_path.stat().st_size < MAX_ARTIFACT_BYTES,
            "full source listing missing or too large")
    listing = descriptor(listing_path)
    require(listing["bytes"] == receipt.get("full_source_listing_bytes")
            and listing["sha256"] == receipt.get("full_source_listing_sha256"),
            "full source listing descriptor changed")
    patch: dict[str, Any] | None = None
    if candidate:
        patch_path = Path(str(receipt.get("candidate_patch", ""))).resolve()
        require(patch_path.is_file() and patch_path.stat().st_size < MAX_ARTIFACT_BYTES,
                "candidate patch missing or too large")
        patch = descriptor(patch_path)
        require(patch["bytes"] == receipt.get("candidate_patch_bytes")
                and patch["sha256"] == receipt.get("candidate_patch_sha256"),
                "candidate patch descriptor changed")
        changed = receipt.get("changed_paths")
        require(changed == list(EXPECTED_CHANGED_PATHS), "candidate changed-file allowlist changed")
        require(receipt.get("changed_path_count") == len(EXPECTED_CHANGED_PATHS),
                "candidate changed-file count changed")
    require(receipt.get("generated_targets_included") is False, "generated target entered source closure")

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
        require(type(row.get("bytes")) is int and row["bytes"] >= 0,
                f"invalid source listing size: {rel}")
        require(isinstance(row.get("sha256"), str) and len(row["sha256"]) == 64,
                f"invalid source listing hash: {rel}")
        rows.append(row)
    actual_files = {
        str(path.relative_to(source_root)).replace(os.sep, "/"): path
        for path in source_root.rglob("*")
        if path.is_file() and ".git" not in path.parts
    }
    require(set(actual_files) == seen, "source tree file set differs from full listing")
    total = 0
    for row in rows:
        path = actual_files[row["path"]]
        actual = descriptor(path)
        require(actual["bytes"] == row["bytes"] and actual["sha256"] == row["sha256"],
                f"source tree file changed: {row['path']}")
        total += actual["bytes"]
    require(len(rows) == receipt.get("source_file_count") == len(receipt.get("source_files", [])),
            "source file count changed")
    require(total == receipt.get("source_bytes"), "source byte total changed")
    if candidate:
        for path in EXPECTED_CHANGED_PATHS:
            require(path in actual_files, f"changed source file missing: {path}")
    return {
        "receipt": descriptor(receipt_path),
        "source_root": str(source_root),
        "base_revision": BASE_REVISION,
        "full_source_listing": listing,
        "candidate_patch": patch,
        "reconstruction": receipt.get("reconstruction"),
        "source_file_count": len(rows),
        "source_bytes": total,
        "changed_paths": list(EXPECTED_CHANGED_PATHS) if candidate else [],
        "candidate_source_files": ({
            path: descriptor(actual_files[path]) for path in EXPECTED_CHANGED_PATHS
        } if candidate else {}),
    }


def _command_argv(receipt: dict[str, Any]) -> list[str]:
    argv = receipt.get("command_argv")
    if isinstance(argv, list) and all(isinstance(value, str) for value in argv):
        return list(argv)
    command = receipt.get("command")
    require(isinstance(command, str), "build command is missing")
    import shlex
    return shlex.split(command)


def validate_build_receipt(receipt_path: Path, source: dict[str, Any]) -> dict[str, Any]:
    receipt = read_json(receipt_path)
    status = receipt.get("status")
    require((status in (None, "completed")) and receipt.get("exit_code") == 0,
            "build receipt does not record successful completion")
    require(receipt.get("profile") == "release", "build profile changed")
    argv = _command_argv(receipt)
    for flag in ("--offline", "--locked", "--release", "-j2"):
        require(flag in argv, f"build command lacks {flag}")
    require(any(value == "+1.95.0" or "1.95.0" in value for value in argv),
            "build command is not bound to +1.95.0")
    require(receipt.get("environment", {}) == {}, "unreviewed build environment overrides present")
    environment_scope = receipt.get("environment_scope")
    if environment_scope is None:
        require(str(receipt.get("schema", "")).startswith("litchi.performance.0925"),
                "build environment recording scope is missing")
        environment_scope = "legacy receipt: environment object is the recorded allowlist; no wholesale environment was retained"
    require(isinstance(environment_scope, str), "build environment recording scope is malformed")
    custody = receipt.get("source_custody")
    require(isinstance(custody, dict), "build source custody missing")
    require(custody.get("source_receipt") == source["receipt"],
            "build/source receipt binding differs")
    require(custody.get("full_source_listing") == source["full_source_listing"],
            "build/source listing binding differs")
    binary_value = receipt.get("binary")
    require(isinstance(binary_value, dict), "build binary descriptor missing")
    binary_path = verify_descriptor(binary_value, label="build binary")
    require(os.access(binary_path, os.X_OK), "build binary is not executable")
    terminal_value = receipt.get("terminal_receipt")
    if terminal_value is not None:
        verify_descriptor(terminal_value, label="build terminal receipt")
    log_value = receipt.get("build_log")
    if log_value is not None:
        verify_descriptor(log_value, label="build log")
    locks = receipt.get("locks")
    require(isinstance(locks, dict) and locks, "build lock identities missing")
    reconstruction = receipt.get("build_input_reconstruction")
    if reconstruction is not None:
        require(isinstance(reconstruction, dict), "build provenance bridge is malformed")
        historical_receipt = reconstruction.get("historical_source_receipt")
        historical_listing = reconstruction.get("historical_source_listing")
        historical_terminal = reconstruction.get("historical_build_terminal")
        for name, value in (("historical source receipt", historical_receipt),
                            ("historical source listing", historical_listing),
                            ("historical build terminal", historical_terminal)):
            verify_descriptor(value, label=name)
        require(reconstruction.get("reconstructed_source_root") == source["source_root"],
                "reconstructed source root bridge changed")
        require(reconstruction.get("historical_command_cwd") == receipt.get("cwd"),
                "historical build cwd bridge changed")
        require(reconstruction.get("historical_source_root") == receipt.get("cwd"),
                "historical source root bridge changed")
        require(historical_listing == source["reconstruction"]["historical_full_source_listing"],
                "historical source listing bridge differs")
        require(historical_receipt == source["reconstruction"]["historical_receipt"],
                "historical source receipt bridge differs")
        original = read_json(Path(str(historical_receipt["path"])))
        terminal = read_json(Path(str(historical_terminal["path"])))
        require(parse_utc(original["created_utc"], "historical source receipt created")
                <= parse_utc(receipt["started_utc"], "candidate build started"),
                "historical source receipt was created after build start")
        require(terminal.get("cwd") == receipt.get("cwd"),
                "historical terminal cwd differs from build cwd")
        require(parse_utc(terminal["started_utc"], "historical build started")
                == parse_utc(receipt["started_utc"], "candidate build started"),
                "historical terminal/build start differs")
    return {
        "receipt": descriptor(receipt_path),
        "binary": {"path": str(binary_path), "bytes": binary_value["bytes"],
                    "sha256": binary_value["sha256"]},
        "source_root": source["source_root"],
        "command_argv": argv,
        "profile": receipt.get("profile"),
        "environment_scope": environment_scope,
        "toolchain": receipt.get("toolchain"),
        "locks": locks,
        "build_log": log_value,
        "terminal_receipt": terminal_value,
    }


def validate_plan(plan_path: Path, correction_path: Path) -> tuple[dict[str, Any], dict[str, Any]]:
    plan = read_json(plan_path)
    correction = read_json(correction_path)
    require(plan.get("routes") == list(ROUTES), "plan route set changed")
    formal = plan.get("formal_paired_cohort")
    require(isinstance(formal, dict), "formal cohort section missing")
    require(formal.get("children") == CHILDREN
            and formal.get("blocks") == BLOCKS
            and formal.get("children_per_route_per_version") == BLOCKS
            and formal.get("measured_samples_per_route_per_version") == BLOCKS * SAMPLES,
            "formal cohort cardinality changed")
    require(formal.get("odd_block_route_order") == list(ROUTES), "odd route order changed")
    require(formal.get("even_block_route_order") == list(reversed(ROUTES)),
            "even route order changed")
    primary = correction.get("formal_gate", {}).get("primary", {})
    require(primary.get("statistic")
            == "median of paired baseline_p50/candidate_p50 ratios; ratio is baseline p50 divided by candidate p50, so >1 means candidate speedup",
            "formal ratio orientation is not corrected")
    require(primary.get("median_ratio_minimum") == 1.05
            and primary.get("bootstrap_lower_95_strictly_greater_than") == 1.0
            and primary.get("draws") == 10000 and primary.get("seed") == 930,
            "formal statistical gate changed")
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
    baseline_build = validate_build_receipt(args.baseline_build_receipt, baseline_source)
    candidate_build = validate_build_receipt(args.candidate_build_receipt, candidate_source)
    require(baseline_build["toolchain"] == candidate_build["toolchain"],
            "baseline/candidate toolchain differs")
    require(baseline_build["profile"] == candidate_build["profile"] == "release",
            "baseline/candidate release profile differs")
    rows = schedule_rows()
    return {
        "schema": "litchi.performance.0930.formal-comparison-manifest.v1",
        "created_utc": utc_now(),
        "capture_authorized": False,
        "formal_children_run": 0,
        "base_revision": BASE_REVISION,
        "claim_scope": "prospective paired normal-release cohort; no claim until all 48 children and independent replay pass",
        "failure_policy": "retain every terminal child; any failed report/control invalidates the cohort; never retry or selectively recapture",
        "source": {"baseline": baseline_source, "candidate": candidate_source},
        "build": {"baseline": baseline_build, "candidate": candidate_build},
        "custody_sources": {
            "baseline_source_receipt": baseline_source["receipt"],
            "candidate_source_receipt": candidate_source["receipt"],
            "baseline_build_receipt": baseline_build["receipt"],
            "candidate_build_receipt": candidate_build["receipt"],
            "plan": descriptor(args.plan),
            "gate_correction": descriptor(args.gate_correction),
        },
        "scripts": script_bindings(args.plan, args.gate_correction),
        "schedule": {
            "blocks": BLOCKS, "children": CHILDREN, "routes": list(ROUTES),
            "children_per_route_per_version": BLOCKS,
            "warmups": WARMUPS, "samples_per_child": SAMPLES, "cpu": CPU,
            "process_deadline_seconds": int(DEADLINE_SECONDS),
            "rows": rows, "sha256": schedule_hash(rows),
        },
        "argv_template": [
            "/usr/bin/time", "-v", "taskset", "-c", str(CPU), "<binary>",
            "--case", CASE, "--samples", str(SAMPLES), "--warmup", str(WARMUPS),
            "--json", "<report.json>",
        ],
        "route_selector": {"name": "LITCHI_DOCX_INGRESS_ROUTE", "values": list(ROUTES)},
        "child_environment": dict(CHILD_ENV_BASE),
        "corpus": {
            "archive_bytes": 2_101_427,
            "archive_sha256": "183204a832a3518bf1d4474e9b7ff347ad7d0dc95ad2e9b1dbdde405837a29e8",
            "expected_output_bytes": 2_101_446,
            "expected_output_sha256": "5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec",
        },
        "formal_gate": {
            "primary_route": "generic_read_at",
            "ratio": "baseline_p50 / candidate_p50",
            "median_ratio_minimum": 1.05,
            "bootstrap_lower_95_strictly_greater_than": 1.0,
            "bootstrap_draws": 10000,
            "bootstrap_seed": 930,
            "bootstrap_percentile": "nearest-rank percentile of paired process ratios",
            "pair_unit": "same route, block, and position; version order alternates by block",
            "secondary_flags": {
                "tail_or_pooled_p95_p99_regression": ">5%",
                "rss_regression": ">5%; setup-inclusive review flag",
                "semantic_or_counter_failure": "invalidates the complete cohort",
            },
            "source_plan": descriptor(args.plan),
            "correction": descriptor(args.gate_correction),
        },
        "status": "prospective_unfinished",
        "raw_evidence": "external packet only; terminal receipts contain descriptors and scalar status, never duplicated raw vectors",
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
                    allow_authorized: bool = False) -> dict[str, Any]:
    manifest_path = path.resolve()
    manifest = read_json(manifest_path)
    require(manifest.get("schema") == "litchi.performance.0930.formal-comparison-manifest.v1",
            "manifest schema changed")
    require(manifest.get("base_revision") == BASE_REVISION, "manifest base revision changed")
    require(manifest.get("schedule", {}).get("sha256") == schedule_hash(schedule_rows()),
            "manifest schedule hash changed")
    require(manifest.get("schedule", {}).get("rows") == schedule_rows(), "manifest schedule changed")
    require(manifest.get("schedule", {}).get("children") == CHILDREN, "manifest child count changed")
    require(manifest.get("child_environment") == CHILD_ENV_BASE,
            "child environment allowlist changed")
    if authorized:
        require(manifest.get("capture_authorized") is True,
                "capture authorization state changed")
    elif not allow_authorized:
        require(manifest.get("capture_authorized") is False,
                "unauthorized freeze manifest was changed")
    if authorized:
        require(os.environ.get("LITCHI_PERF_0930_CAPTURE_APPROVED") == "1",
                "LITCHI_PERF_0930_CAPTURE_APPROVED=1 is required")
    verify_scripts(manifest)
    validate_plan(Path(str(manifest["formal_gate"]["source_plan"]["path"])),
                  Path(str(manifest["formal_gate"]["correction"]["path"])))
    for version in ("baseline", "candidate"):
        source = manifest["source"][version]
        build = manifest["build"][version]
        # Source/build/binary are rehashed at capture start and at replay.
        source_now = validate_source_receipt(
            Path(str(source["receipt"]["path"])), candidate=(version == "candidate")
        )
        require(source_now["source_root"] == source["source_root"],
                f"{version} source root changed")
        build_now = validate_build_receipt(Path(str(build["receipt"]["path"])), source_now)
        require(build_now["binary"] == build["binary"], f"{version} binary binding changed")
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
        "schema": "litchi.performance.0930.paired-gate.v1",
        "ratio": "baseline_p50 / candidate_p50",
        "primary_route": "generic_read_at",
        "median_ratio_minimum": 1.05,
        "bootstrap_lower_95_strictly_greater_than": 1.0,
        "primary_pass": primary["median_ratio"] >= 1.05 and primary["bootstrap_lower_95"] > 1.0,
        "per_route": per_route,
        "claim": "descriptive paired gate only; adoption requires independent review and semantic controls",
    }


def child_command(row: dict[str, Any], binary: Path, report: Path) -> list[str]:
    return ["/usr/bin/time", "-v", "taskset", "-c", str(CPU), str(binary),
            "--case", CASE, "--samples", str(SAMPLES), "--warmup", str(WARMUPS),
            "--json", str(report)]


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
        "schema": "litchi.performance.0930.scalar-validation.v1",
        "status": "passed",
        "route": route,
        "samples": SAMPLES,
        "warmups": WARMUPS,
        "report": report_desc,
        "output_sha256": output,
        "phase_sample_counts": {name: len(values) for name, values in phases.items()},
    }


def run_child(row: dict[str, Any], manifest: dict[str, Any], packet: Path,
              comparison: Any, source_root: Path, binary: dict[str, Any]) -> dict[str, Any]:
    child = packet_path(packet, row["id"], "report.json").parent
    child.mkdir(parents=True, exist_ok=False)
    report_path = child / "report.json"
    stdout_path = child / "stdout.log"
    stderr_path = child / "stderr.log"
    time_path = child / "time.log"
    command = child_command(row, Path(binary["path"]), report_path)
    environment = dict(manifest["child_environment"])
    environment["LITCHI_DOCX_INGRESS_ROUTE"] = row["route"]
    env = dict(environment)
    started = utc_now()
    monotonic_start = time.monotonic()
    exit_code: int | None = None
    observation_timeouts = 0
    deadline_exceeded = False
    runner_error: str | None = None
    group_closed = False
    group_probe_count = 0
    process_group_id: int | None = None
    with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
        try:
            process = subprocess.Popen(command, cwd=str(source_root), env=env,
                                       stdout=stdout, stderr=stderr, start_new_session=True)
            process_group_id = process.pid
            while True:
                try:
                    exit_code = process.wait(timeout=OBSERVE_SECONDS)
                    break
                except subprocess.TimeoutExpired:
                    observation_timeouts += 1
                    if time.monotonic() - monotonic_start <= DEADLINE_SECONDS:
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
        "schema": "litchi.performance.0930.formal-terminal.v1",
        "status": "terminal",
        "schedule": row,
        "argv": command,
        "environment": environment,
        "cwd_basename": source_root.name,
        "started_utc": started,
        "ended_utc": ended,
        "exit_code": exit_code,
        "observation_timeout_count": observation_timeouts,
        "deadline_exceeded": deadline_exceeded,
        "runner_error": runner_error,
        "process_group_id": process_group_id,
        "process_group_closed": group_closed,
        "process_group_probe_count": group_probe_count,
        "stdout": descriptor(stdout_path, root=packet),
        "stderr": descriptor(stderr_path, root=packet),
        "time": descriptor(time_path, root=packet),
        "rss_bytes": rss_bytes(time_path),
        "report": report_desc,
        "binary": binary,
        "source_receipt_sha256": manifest["source"][row["version"]]["receipt"]["sha256"],
        "runner_sha256": manifest["scripts"]["runner"]["sha256"],
        "manifest_sha256": sha256_path(packet / "manifest.json"),
    }
    terminal_path = child / "terminal.json"
    write_new_json(terminal_path, terminal, limit=MAX_ARTIFACT_BYTES)
    valid = (exit_code == 0 and not deadline_exceeded and runner_error is None
             and group_closed
             and terminal["rss_bytes"] is not None and report_desc is not None
             and report_desc["bytes"] < MAX_ARTIFACT_BYTES)
    validation: dict[str, Any]
    if valid:
        try:
            report = read_json(report_path)
            result = comparison.validate_report(report, row["route"], warmups=WARMUPS,
                                                samples=SAMPLES, label=str(report_path),
                                                binary=binary)
            validation = scalar_validation(report, result, row["route"], report_desc)
        except Exception as error:
            validation = {"schema": "litchi.performance.0930.scalar-validation.v1",
                          "status": "failed", "error": f"{type(error).__name__}: {error}"}
            valid = False
    else:
        validation = {"schema": "litchi.performance.0930.scalar-validation.v1", "status": "failed",
                      "error": "terminal process failed, RSS missing, or report missing"}
    validation_path = child / "validation.json"
    write_new_json(validation_path, validation, limit=MAX_ARTIFACT_BYTES)
    return {
        "id": row["id"], "pair_id": row["pair_id"], "block": row["block"],
        "position": row["position"], "route": row["route"], "version": row["version"],
        "status": "passed" if valid else "failed", "exit_code": exit_code,
        "deadline_exceeded": deadline_exceeded, "rss_bytes": terminal["rss_bytes"],
        "terminal": descriptor(terminal_path, root=packet),
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
    write_new_json(packet / "manifest.json", packet_manifest, limit=MAX_ARTIFACT_BYTES)
    comparison = load_comparison(packet / "scripts" / Path(str(manifest["scripts"]["validator"]["path"])).name)
    results: list[dict[str, Any]] = []
    started = utc_now()
    for row in manifest["schedule"]["rows"]:
        version = row["version"]
        source_root = Path(manifest["source"][version]["source_root"])
        binary = manifest["build"][version]["binary"]
        results.append(run_child(row, manifest, packet, comparison, source_root, binary))
        # Keep the state scalar and durable; raw vectors remain in reports.
        write_new_json(packet / f"capture-state-{len(results):03d}.json", {
            "schema": "litchi.performance.0930.formal-state.v1",
            "status": "running", "started_utc": started,
            "children_run": len(results), "children_expected": CHILDREN,
            "children": results,
            "recapture_count": 0,
        }, limit=MAX_ARTIFACT_BYTES)
    passed = len(results) == CHILDREN and all(row["status"] == "passed" for row in results)
    state = {
        "schema": "litchi.performance.0930.formal-state.v1",
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
    manifest = read_json(manifest_path.resolve() if manifest_path else packet / "manifest.json")
    require(manifest.get("schema") == "litchi.performance.0930.formal-comparison-manifest.v1",
            "copied packet manifest schema changed")
    verify_scripts(manifest, packet)
    verify_packet_custody(manifest, packet)
    verify_descriptor(manifest["scripts"]["validator"], label="manifest validator")
    # Revalidate the external custody descriptors before reading child reports.
    verify_manifest(Path(str(manifest_path)) if manifest_path else packet / "manifest.json",
                    allow_authorized=True)
    local_validator = packet / safe_relative(
        manifest["packet_scripts"]["validator"]["path"], "packet validator"
    )
    comparison = load_comparison(local_validator)
    rows = manifest["schedule"]["rows"]
    state = read_json(packet / "capture-state.json")
    require(state.get("status") in ("completed", "invalid"), "packet is not terminal")
    children = state.get("children")
    require(isinstance(children, list) and len(children) == CHILDREN,
            "packet child count is incomplete")
    by_id = {row.get("id"): row for row in children}
    require(set(by_id) == {row["id"] for row in rows}, "packet child IDs differ from frozen schedule")
    reduced_rows: dict[tuple[str, str], list[dict[str, Any]]] = {}
    pair_values: dict[tuple[str, str], dict[str, dict[str, float | int]]] = {}
    for row in rows:
        child = by_id[row["id"]]
        require(child.get("status") == "passed", f"child is not a passed terminal: {row['id']}")
        terminal_path = packet_file(packet, child["terminal"], f"terminal {row['id']}")
        terminal = read_json(terminal_path)
        require(terminal.get("schedule") == row, f"terminal schedule changed: {row['id']}")
        require(terminal.get("exit_code") == 0 and terminal.get("deadline_exceeded") is False
                and terminal.get("runner_error") is None, f"terminal status failed: {row['id']}")
        require(terminal.get("process_group_closed") is True,
                f"process group remained alive: {row['id']}")
        expected_environment = dict(manifest["child_environment"])
        expected_environment["LITCHI_DOCX_INGRESS_ROUTE"] = row["route"]
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
        binary = manifest["build"][row["version"]]["binary"]
        require(terminal.get("binary") == binary, f"terminal binary binding changed: {row['id']}")
        report = read_json(report_path)
        validated = comparison.validate_report(report, row["route"], warmups=WARMUPS,
                                                samples=SAMPLES, label=str(report_path), binary=binary)
        validation_path = packet_file(packet, child["validation"], f"validation {row['id']}")
        validation = read_json(validation_path)
        require(validation.get("status") == "passed" and validation.get("report") == child["report"],
                f"scalar validation changed: {row['id']}")
        reduced_rows.setdefault((row["version"], row["route"]), []).append(validated)
        elapsed = list(validated["elapsed_ns"])
        pair_values.setdefault((row["route"], row["pair_id"]), {})[row["version"]] = {
            "p50": float(statistics.median(elapsed)),
            "p95": nearest_percentile([float(value) for value in elapsed], 0.95),
            "p99": nearest_percentile([float(value) for value in elapsed], 0.99),
            "rss": int(child["rss_bytes"]),
        }
    summary: dict[str, Any] = {
        "schema": "litchi.performance.0930.formal-replay.v1",
        "status": "passed",
        "packet": str(packet),
        "children": CHILDREN,
        "source": "all reports, terminal streams, and scalar validations rehashed from moved packet",
        "per_version_route": {},
    }
    for key, values in reduced_rows.items():
        version, route = key
        # The validator's reducer only consumes the raw validation vectors. It
        # is descriptive; pair gates are intentionally left to the formal
        # analysis step after replay.
        elapsed = [value for item in values for value in item["elapsed_ns"]]
        phases = {
            phase: [value for item in values for value in item["phase_ns"][phase]]
            for phase in PHASES
        }
        reduced = {
            "children": len(values),
            "samples": len(elapsed),
            "elapsed_ns": comparison._statistics(elapsed),
            "phase_statistics_ns": {
                phase: comparison._statistics(vector) for phase, vector in phases.items()
            },
        }
        summary["per_version_route"][f"{version}/{route}"] = {
            "samples": reduced["samples"],
            "children": reduced["children"],
            "elapsed_ns": reduced["elapsed_ns"],
            "phase_statistics_ns": reduced["phase_statistics_ns"],
            "rss_bytes": [by_id[row["id"]]["rss_bytes"] for row in rows
                           if row["version"] == version and row["route"] == route],
        }
    summary["paired_gate"] = paired_gate(pair_values)
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
    return {"schema": "litchi.performance.0930.formal-runner-controls.v1",
            "status": "passed", "children": len(rows), "routes": list(ROUTES),
            "children_per_route_per_version": BLOCKS, "warmups": WARMUPS,
            "samples": SAMPLES, "cpu": CPU, "deadline_seconds": DEADLINE_SECONDS,
            "schedule_sha256": schedule_hash(rows), "source_mutation": False,
            "capture_started": False}


def controls(args: argparse.Namespace) -> dict[str, Any]:
    """Exercise custody-only failure paths without starting a document child."""
    manifest = read_json(args.manifest.resolve())
    scratch = Path(tempfile.mkdtemp(prefix="formal-runner-controls-",
                                    dir="/tmp/litchi-performance-0930"))
    result: dict[str, Any] = {
        "schema": "litchi.performance.0930.formal-runner-controls.v1",
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
