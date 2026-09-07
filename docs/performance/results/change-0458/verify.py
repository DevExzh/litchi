#!/usr/bin/env python3
"""Authenticate the retained change-0458 measurement bundle.

The verifier is self contained. A normal or portable run only reads files
below this directory; --precleanup additionally checks the copied executables
and current Rust source tree that the binding captured. It never builds,
captures, or changes evidence.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import subprocess
import sys
from typing import Any, Mapping


ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3] if len(ROOT.parents) > 3 else ROOT
TASK = Path("/tmp/litchi-goal-0458")
CHANGE = 458
SCHEMA = "litchi-0458-verification-v1"
SHA256_RE = re.compile(r"[0-9a-f]{64}\Z")
RETRY_RE = re.compile(r"^(.*)-r([0-9]+)$")


class VerificationError(AssertionError):
    pass


def fail(message: str) -> None:
    raise VerificationError(message)


def load(path: Path, label: str) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{label}: invalid JSON ({error})")
    raise AssertionError("unreachable")


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: expected JSON object")
    return value


def text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label}: expected non-empty text")
    return value


def integer(value: Any, label: str, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(f"{label}: expected integer >= {minimum}")
    return value


def digest(value: Any, label: str) -> str:
    value = text(value, label).lower()
    if SHA256_RE.fullmatch(value) is None:
        fail(f"{label}: expected lowercase SHA-256")
    return value


def sha_file(path: Path) -> str:
    value = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                value.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return value.hexdigest()


def regular(path: Path, label: str) -> Path:
    if not path.is_file() or path.is_symlink():
        fail(f"{label}: missing, symlink, or non-regular file: {path}")
    return path


def safe_path(base: Path, value: Any, label: str, *, must_exist: bool = True) -> Path:
    raw = text(value, label)
    relative = Path(raw)
    if relative.is_absolute() or ".." in relative.parts:
        fail(f"{label}: path must be relative and traversal-free")
    if base.is_symlink():
        fail(f"{label}: base is a symlink")
    candidate = base / relative
    for ancestor in (base, *candidate.parents):
        if ancestor.is_symlink():
            fail(f"{label}: path has a symlinked ancestor")
    try:
        resolved = candidate.resolve(strict=must_exist)
        base_resolved = base.resolve(strict=True)
    except OSError as error:
        fail(f"{label}: cannot resolve path ({error})")
    if not resolved.is_relative_to(base_resolved):
        fail(f"{label}: path escapes its base")
    if must_exist:
        regular(candidate, label)
    elif candidate.exists() or candidate.is_symlink():
        if candidate.is_symlink() or not candidate.is_file():
            fail(f"{label}: existing path is unsafe")
    return candidate


def artifact(base: Path, value: Any, label: str) -> Path:
    row = obj(value, label)
    path = safe_path(base, row.get("path"), f"{label}.path")
    expected_bytes = integer(row.get("bytes"), f"{label}.bytes")
    expected_sha = digest(row.get("sha256"), f"{label}.sha256")
    if path.stat().st_size != expected_bytes or sha_file(path) != expected_sha:
        fail(f"{label}: artifact identity differs")
    return path


def source_record(value: Any, label: str) -> tuple[Path, dict[str, str], dict[str, Any]]:
    row = obj(value, label)
    path = safe_path(ROOT, row.get("path"), f"{label}.path")
    expected_sha = digest(row.get("sha256"), f"{label}.sha256")
    expected_files = integer(row.get("files"), f"{label}.files")
    if sha_file(path) != expected_sha:
        fail(f"{label}: source manifest hash differs")
    manifest = obj(load(path, label), label)
    if len(manifest) != expected_files:
        fail(f"{label}: source manifest file count differs")
    normalized: dict[str, str] = {}
    for name, value in manifest.items():
        rel = text(name, f"{label}.files path")
        candidate = Path(rel)
        if candidate.is_absolute() or ".." in candidate.parts:
            fail(f"{label}: source path escapes repository: {rel}")
        normalized[rel] = digest(value, f"{label}.files[{rel}]")
    return path, normalized, row


def source_equal(left: Mapping[str, Any], right: Mapping[str, Any], label: str) -> None:
    if left.get("path") != right.get("path") or left.get("sha256") != right.get("sha256") or left.get("files") != right.get("files"):
        fail(f"{label}: source identity differs")


def verify_sum_file(path: Path, base: Path) -> int:
    regular(path, path.name)
    rows: dict[str, str] = {}
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as error:
        fail(f"{path}: cannot read checksum file ({error})")
    for index, line in enumerate(lines, 1):
        fields = line.split("  ", 1)
        if len(fields) != 2 or SHA256_RE.fullmatch(fields[0]) is None:
            fail(f"{path}: malformed checksum line {index}")
        name = fields[1]
        relative = Path(name)
        if relative.is_absolute() or ".." in relative.parts or not name:
            fail(f"{path}: unsafe checksum member {name!r}")
        if name == path.name or name in rows:
            fail(f"{path}: duplicate or self checksum member {name}")
        member = safe_path(base, name, f"{path.name}.{name}")
        if sha_file(member) != fields[0]:
            fail(f"{path}: hash differs for {name}")
        rows[name] = fields[0]
    actual: set[str] = set()
    for member in base.rglob("*"):
        if member.is_symlink():
            fail(f"{base}: bundle contains symlink {member.relative_to(base)}")
        if member.is_file() and member != path:
            actual.add(member.relative_to(base).as_posix())
    if set(rows) != actual:
        missing = sorted(actual - set(rows))
        extra = sorted(set(rows) - actual)
        fail(f"{path}: checksum coverage differs (missing={missing[:3]}, extra={extra[:3]})")
    return len(rows)


def check_hash_binding(protocol: Mapping[str, Any], key: str, path: Path, *, optional: bool = False) -> None:
    expected = protocol.get(key)
    if expected is None:
        if optional:
            return
        fail(f"protocol.{key}: missing hash binding")
    if digest(expected, f"protocol.{key}") != sha_file(path):
        fail(f"protocol.{key}: hash differs for {path.name}")


def check_protocol() -> tuple[dict[str, Any], dict[str, Any]]:
    protocol_path = ROOT / "protocol.json"
    binding_path = ROOT / "binary-binding.json"
    protocol = obj(load(protocol_path, "protocol.json"), "protocol.json")
    binding = obj(load(binding_path, "binary-binding.json"), "binary-binding.json")
    if protocol.get("change") != CHANGE or not text(protocol.get("schema"), "protocol.schema").startswith("litchi-0458-protocol-"):
        fail("protocol schema/change differs")
    if binding.get("change") != CHANGE or binding.get("schema") != "litchi-0458-binary-binding-v1":
        fail("binary binding schema/change differs")
    revision = text(protocol.get("revision"), "protocol.revision")
    if binding.get("revision") != revision:
        fail("protocol and binary binding revisions differ")
    check_hash_binding(protocol, "binary_binding_sha256", ROOT / "binary-binding.json")
    check_hash_binding(protocol, "capture_sha256", ROOT / "capture.py")
    check_hash_binding(protocol, "profile_sha256", ROOT / "profile.py")
    check_hash_binding(protocol, "oracle_sha256", ROOT / "oracle.py")
    check_hash_binding(protocol, "prior_control_binding_sha256", ROOT / "prior-control-bindings.json")
    check_hash_binding(protocol, "host_sha256", ROOT / "host.json")
    check_hash_binding(protocol, "derive_sha256", ROOT / "derive.py", optional=True)
    check_hash_binding(protocol, "derive_profile_sha256", ROOT / "derive-profile.py", optional=True)
    check_hash_binding(protocol, "perf_capability_sha256", ROOT / "perf-capability.json", optional=True)

    order = protocol.get("order")
    if not isinstance(order, list) or len(order) != 24:
        fail("protocol.order must contain exactly 24 lanes")
    lanes: list[dict[str, Any]] = []
    for index, raw in enumerate(order):
        lane = obj(raw, f"protocol.order[{index}]")
        expected_keys = {"id", "repeat", "instrumentation", "shape", "scope"}
        if set(lane) != expected_keys:
            fail(f"protocol.order[{index}] lane keys differ")
        if lane["repeat"] not in {"R1", "R2"} or lane["instrumentation"] not in {"normal", "allocator"} or lane["shape"] not in {"tiny", "medium", "large"} or lane["scope"] not in {"lifecycle", "phases"}:
            fail(f"protocol.order[{index}] lane dimensions differ")
        text(lane["id"], f"protocol.order[{index}].id")
        lanes.append(lane)
    if len({lane["id"] for lane in lanes}) != 24:
        fail("protocol.order contains duplicate lane ids")
    expected_lanes = {
        (repeat, mode, shape, scope)
        for repeat in ("R1", "R2")
        for mode in ("normal", "allocator")
        for shape in ("tiny", "medium", "large")
        for scope in ("lifecycle", "phases")
    }
    observed_lanes = {(lane["repeat"], lane["instrumentation"], lane["shape"], lane["scope"]) for lane in lanes}
    if observed_lanes != expected_lanes:
        fail("protocol.order does not cover the exact 24-lane matrix")
    if protocol.get("reports") != 24 or protocol.get("retained_samples") != 720 or protocol.get("samples") != 30 or protocol.get("warmup") != 3 or protocol.get("cpu") != 2 or protocol.get("workers") != 1:
        fail("protocol matrix dimensions differ")
    if not isinstance(protocol.get("argv"), list) or not protocol["argv"] or not isinstance(protocol.get("diagnostic_argv"), list) or not protocol["diagnostic_argv"]:
        fail("protocol workload argv is missing")
    if not isinstance(protocol.get("environment"), dict):
        fail("protocol.environment must be an object")
    claims = text(protocol.get("claims"), "protocol.claims")
    if "no production" not in claims.lower() and "diagnostic" not in claims.lower():
        fail("protocol claims do not describe diagnostic scope")

    build_ref = obj(binding.get("build_receipt"), "binary-binding.build_receipt")
    build_path = artifact(ROOT, build_ref, "binary-binding.build_receipt")
    if build_ref.get("path") != "checks/build.json":
        fail("binary binding build receipt must be checks/build.json")
    build = obj(load(build_path, "checks/build.json"), "checks/build.json")
    if build.get("change") != CHANGE or build.get("status") != "pass" or build.get("exit_code") != 0 or build.get("revision") != revision or build.get("source_unchanged") is not True:
        fail("binary binding build receipt is not a successful immutable build")
    before = obj(build.get("source_before"), "build.source_before")
    after = obj(build.get("source_after"), "build.source_after")
    source_record(before, "build.source_before")
    source_record(after, "build.source_after")
    if before != after:
        fail("build source custody differs")
    source_ref = obj(binding.get("source_manifest"), "binary-binding.source_manifest")
    source_equal(source_ref, after, "binary binding source manifest")
    source_record(source_ref, "binary-binding.source_manifest")

    binaries = obj(binding.get("binaries"), "binary-binding.binaries")
    if set(binaries) != {"normal", "allocator"}:
        fail("binary binding must contain normal and allocator copies")
    for mode, raw in binaries.items():
        row = obj(raw, f"binary-binding.binaries.{mode}")
        source = text(row.get("source"), f"binary-binding.binaries.{mode}.source")
        source_path = Path(source)
        if source_path.is_absolute() or ".." in source_path.parts:
            fail(f"binary-binding.binaries.{mode}.source escapes repository")
        binary_path = Path(text(row.get("path"), f"binary-binding.binaries.{mode}.path"))
        if not binary_path.is_absolute() or not binary_path.is_relative_to(TASK):
            fail(f"binary-binding.binaries.{mode}.path is not task-owned")
        integer(row.get("bytes"), f"binary-binding.binaries.{mode}.bytes", 1)
        digest(row.get("sha256"), f"binary-binding.binaries.{mode}.sha256")
    return protocol, binding


def check_current_sources(source_ref: Mapping[str, Any]) -> None:
    _, manifest, _ = source_record(source_ref, "binary-binding.source_manifest")
    names = set(subprocess.check_output(
        ["git", "ls-files", "--cached", "--others", "--exclude-standard", "-z"],
        cwd=REPO).decode().split("\0"))
    names.add("Cargo.lock")
    current = {name for name in names if name.endswith((".rs", ".toml", ".lock")) and (REPO / name).is_file()}
    if current != set(manifest):
        fail("current Rust source manifest coverage differs")
    for relative, expected in manifest.items():
        path = REPO / relative
        regular(path, f"current source {relative}")
        if sha_file(path) != expected:
            fail(f"current source changed: {relative}")


def check_live_binding(binding: Mapping[str, Any]) -> None:
    binaries = obj(binding.get("binaries"), "binary-binding.binaries")
    for mode, raw in binaries.items():
        row = obj(raw, f"binary-binding.binaries.{mode}")
        path = Path(text(row.get("path"), f"binary-binding.binaries.{mode}.path"))
        if path.is_symlink() or not path.is_file() or not path.is_relative_to(TASK):
            fail(f"live {mode} binary copy is missing or not task-owned")
        expected_bytes = integer(row.get("bytes"), f"binary-binding.binaries.{mode}.bytes", 1)
        expected_sha = digest(row.get("sha256"), f"binary-binding.binaries.{mode}.sha256")
        if path.stat().st_size != expected_bytes or sha_file(path) != expected_sha:
            fail(f"live {mode} binary copy identity differs")
    check_current_sources(obj(binding.get("source_manifest"), "binary-binding.source_manifest"))


def check_prior_control(protocol: Mapping[str, Any]) -> None:
    path = ROOT / "prior-control-bindings.json"
    value = obj(load(path, path.name), path.name)
    if value.get("schema") != "litchi-0458-prior-control-bindings-v1" or value.get("change") != CHANGE:
        fail("prior-control-bindings schema/change differs")
    shapes = obj(value.get("shapes"), "prior-control-bindings.shapes")
    if set(shapes) != {"tiny", "medium", "large"}:
        fail("prior-control-bindings must cover tiny, medium, and large")
    for shape, raw in shapes.items():
        row = obj(raw, f"prior-control-bindings.shapes.{shape}")
        if row.get("path") != f"prior-control/{shape}.json":
            fail(f"prior-control {shape} path differs")
        artifact(ROOT, row, f"prior-control-bindings.shapes.{shape}")
        prior = obj(load(ROOT / row["path"], f"prior-control/{shape}.json"), f"prior-control/{shape}.json")
        results = prior.get("results")
        if not isinstance(results, list) or len(results) != 1:
            fail(f"prior-control/{shape}.json must contain one result")
        result = obj(results[0], f"prior-control/{shape}.json.results[0]")
        corpus = obj(result.get("corpus"), f"prior-control/{shape}.json corpus")
        source = obj(result.get("source"), f"prior-control/{shape}.json source")
        odp = obj(source.get("odp_append"), f"prior-control/{shape}.json source.odp_append")
        if corpus.get("shape") != shape or odp.get("shape") != shape:
            fail(f"prior-control/{shape}.json shape differs")
    if protocol.get("prior_control_binding_sha256") != sha_file(path):
        fail("protocol prior-control binding hash differs")


def load_oracle() -> Any:
    path = ROOT / "oracle.py"
    spec = importlib.util.spec_from_file_location("litchi0458_oracle", path)
    if spec is None or spec.loader is None:
        fail("cannot load oracle.py")
    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    except Exception as error:
        fail(f"oracle.py failed to import: {error}")
    if not callable(getattr(module, "validate", None)):
        fail("oracle.py has no callable validate")
    return module


def check_check_receipts(protocol: Mapping[str, Any]) -> dict[str, dict[str, Any]]:
    directory = ROOT / "checks"
    if not directory.is_dir():
        fail("checks directory is missing")
    receipts: dict[str, dict[str, Any]] = {}
    referenced_logs: set[str] = set()
    referenced_sources: set[str] = set()
    for path in sorted(directory.glob("*.json")):
        label = f"checks/{path.name}"
        receipt = obj(load(path, label), label)
        if receipt.get("change") != CHANGE:
            fail(f"{label}: change differs")
        if receipt.get("revision") != protocol.get("revision"):
            fail(f"{label}: revision differs")
        if digest(receipt.get("driver_sha256"), f"{label}.driver_sha256") != sha_file(ROOT / "check.py"):
            fail(f"{label}: driver hash differs from check.py")
        status = receipt.get("status")
        if status not in {"running", "pass", "failed"}:
            fail(f"{label}: unsupported status {status!r}")
        argv = receipt.get("argv")
        if not isinstance(argv, list) or not argv or not all(isinstance(item, str) and item for item in argv):
            fail(f"{label}: argv is malformed")
        cwd = text(receipt.get("cwd"), f"{label}.cwd")
        cwd_path = Path(cwd)
        if not cwd_path.is_absolute() or cwd_path != cwd_path.resolve():
            fail(f"{label}: cwd is not canonical absolute path")
        before = obj(receipt.get("source_before"), f"{label}.source_before")
        after = obj(receipt.get("source_after"), f"{label}.source_after")
        _, _, before_row = source_record(before, f"{label}.source_before")
        _, _, after_row = source_record(after, f"{label}.source_after")
        referenced_sources.add(text(before_row.get("path"), f"{label}.source_before.path"))
        referenced_sources.add(text(after_row.get("path"), f"{label}.source_after.path"))
        if status == "running":
            fail(f"{label}: check is still running")
        exit_code = receipt.get("exit_code")
        if isinstance(exit_code, bool) or not isinstance(exit_code, int):
            fail(f"{label}: completed check has no integer exit code")
        if status == "pass" and exit_code != 0:
            fail(f"{label}: pass check has nonzero exit code")
        if status == "failed" and exit_code == 0:
            fail(f"{label}: failed check has zero exit code")
        if status == "pass" and (before != after or receipt.get("source_unchanged") is not True):
            fail(f"{label}: passing source custody differs")
        log_path = artifact(ROOT, receipt.get("log"), f"{label}.log")
        referenced_logs.add(log_path.relative_to(ROOT).as_posix())
        receipts[path.stem] = receipt
    if not receipts:
        fail("checks directory contains no receipts")
    actual_logs = {path.relative_to(ROOT).as_posix() for path in directory.iterdir() if path.is_file() and path.suffix in {".log", ".gz"}}
    if actual_logs != referenced_logs:
        fail(f"checks log coverage differs (unbound={sorted(actual_logs - referenced_logs)[:3]})")
    source_directory = ROOT / "sources"
    if not source_directory.is_dir():
        fail("sources directory is missing")
    actual_sources = {path.relative_to(ROOT).as_posix() for path in source_directory.glob("*.json") if path.is_file()}
    if actual_sources != referenced_sources:
        fail(f"source manifest coverage differs (unbound={sorted(actual_sources - referenced_sources)[:3]})")
    return receipts


DEFAULT_GATES: dict[str, tuple[str, ...]] = {
    "build": ("build", "build-final", "final-build"),
    "strict": ("strict", "strict-final"),
    "tests": ("tests", "tests-final", "test", "test-final"),
    "doc": ("doc", "doc-final", "docs"),
    "format": ("format", "format-final", "rustfmt"),
    "smoke": ("smoke-r2",),
    "matrix": ("matrix", "matrix-final", "capture", "capture-final"),
    "profiles": ("profiles", "profiles-final", "profile", "profile-final"),
    "derive": ("derive", "derive-final"),
    "derive-profile": ("derive-profile", "derive-profile-final"),
    "oracle-tests": ("oracle-tests",),
    "bind": ("bind",),
}


def gate_matches(name: str, pattern: str) -> bool:
    if name == pattern:
        return True
    return bool(re.fullmatch(re.escape(pattern) + r"-r[0-9]+", name))


def latest_passing(names: list[str], receipts: Mapping[str, Mapping[str, Any]]) -> str | None:
    passing = [name for name in names if receipts[name].get("status") == "pass"]
    if not passing:
        return None

    def rank(name: str) -> tuple[int, str]:
        match = RETRY_RE.fullmatch(name)
        return (int(match.group(2)) if match else 0, name)

    return max(passing, key=rank)


def check_required_gates(protocol: Mapping[str, Any], receipts: Mapping[str, Mapping[str, Any]]) -> dict[str, str]:
    configured = protocol.get("required_gates")
    if configured is None:
        configured = {name: list(patterns) for name, patterns in DEFAULT_GATES.items()}
    elif isinstance(configured, list):
        configured = {name: [name] for name in configured}
    configured = obj(configured, "protocol.required_gates")
    selected: dict[str, str] = {}
    for group, raw in configured.items():
        if isinstance(raw, str):
            patterns = [raw]
        elif isinstance(raw, list) and all(isinstance(item, str) and item for item in raw):
            patterns = raw
        else:
            fail(f"protocol.required_gates.{group}: expected names")
        candidates = sorted(name for name in receipts if any(gate_matches(name, pattern) for pattern in patterns))
        chosen = latest_passing(candidates, receipts)
        if chosen is None:
            fail(f"required gate {group} has no passing receipt")
        selected[group] = chosen
    return selected


def path_suffix(value: Any, suffix: str, label: str) -> str:
    raw = text(value, label)
    candidate = Path(raw)
    if not candidate.is_absolute() or candidate != candidate.resolve() or not candidate.as_posix().endswith("/" + suffix):
        fail(f"{label}: expected canonical absolute path ending in {suffix}")
    return raw


def expected_workload(protocol: Mapping[str, Any], lane: Mapping[str, Any], binary: str, report: str) -> list[str]:
    try:
        return [part.format(binary=binary, report=report, **lane) for part in protocol["argv"]]
    except (KeyError, ValueError) as error:
        fail(f"protocol argv cannot be expanded: {error}")
    raise AssertionError("unreachable")


def check_run_artifacts(run_root: Path, receipt: Mapping[str, Any], label: str, expected: set[str]) -> None:
    rows = receipt.get("artifacts")
    if not isinstance(rows, list) or not rows:
        fail(f"{label}: artifacts list is missing")
    seen: set[str] = set()
    for index, raw in enumerate(rows):
        row = obj(raw, f"{label}.artifacts[{index}]")
        relative = text(row.get("path"), f"{label}.artifacts[{index}].path")
        if relative in seen:
            fail(f"{label}: duplicate artifact {relative}")
        seen.add(relative)
        path = artifact(ROOT, row, f"{label}.artifacts[{index}]")
        if not path.is_relative_to(run_root):
            fail(f"{label}: artifact escapes its run directory")
    if seen != expected:
        fail(f"{label}: artifact set differs (expected={sorted(expected)}, got={sorted(seen)})")
    for entry in run_root.iterdir():
        if entry.name != "receipt.json" and (entry.is_symlink() or not entry.is_file()):
            fail(f"{label}: unexpected non-file entry {entry.name}")
    actual = {path.relative_to(ROOT).as_posix() for path in run_root.iterdir() if path.is_file() and path.name != "receipt.json"}
    if actual != seen:
        fail(f"{label}: run directory contains unbound artifacts")


def check_runs(protocol: Mapping[str, Any], binding: Mapping[str, Any], oracle: Any) -> None:
    runs_root = ROOT / "runs"
    if not runs_root.is_dir() or runs_root.is_symlink():
        fail("runs directory is missing or unsafe")
    if any(entry.is_symlink() or not entry.is_dir() for entry in runs_root.iterdir()):
        fail("runs directory contains an unexpected non-directory entry")
    expected_directories = {lane["id"] for lane in protocol["order"]}
    actual_directories = {entry.name for entry in runs_root.iterdir() if entry.is_dir() and not entry.is_symlink()}
    if actual_directories != expected_directories:
        fail("runs directory does not contain exactly the 24 protocol lanes")
    for raw_lane in protocol["order"]:
        lane = obj(raw_lane, "protocol lane")
        lane_id = lane["id"]
        run_root = ROOT / "runs" / lane_id
        if not run_root.is_dir() or run_root.is_symlink():
            fail(f"runs/{lane_id}: run directory is missing or unsafe")
        receipt_path = run_root / "receipt.json"
        report_path = run_root / "report.json"
        receipt = obj(load(receipt_path, f"runs/{lane_id}/receipt.json"), f"runs/{lane_id}/receipt.json")
        regular(report_path, f"runs/{lane_id}/report.json")
        label = f"runs/{lane_id}"
        if receipt.get("schema") != "litchi-0458-capture-v1" or receipt.get("change") != CHANGE or receipt.get("lane") != lane:
            fail(f"{label}: receipt identity differs")
        if receipt.get("protocol_sha256") != sha_file(ROOT / "protocol.json") or receipt.get("binary_binding_sha256") != sha_file(ROOT / "binary-binding.json"):
            fail(f"{label}: protocol or binding hash differs")
        if receipt.get("environment_fixed") != protocol.get("environment") or not isinstance(receipt.get("allocator_environment"), dict):
            fail(f"{label}: environment metadata differs")
        cwd = text(receipt.get("cwd"), f"{label}.cwd")
        if not Path(cwd).is_absolute() or Path(cwd) != Path(cwd).resolve():
            fail(f"{label}: cwd is not canonical absolute")
        if receipt.get("status") != "pass" or receipt.get("exit_code") != 0:
            fail(f"{label}: capture did not pass")
        command = receipt.get("argv")
        if not isinstance(command, list) or not all(isinstance(item, str) for item in command):
            fail(f"{label}: argv is malformed")
        if len(command) < 8 or command[:2] != ["/usr/bin/time", "-v"] or command[4:6] != ["taskset", "-c"] or command[6] != str(protocol["cpu"]):
            fail(f"{label}: capture argv prefix differs")
        if command[2] != "-o":
            fail(f"{label}: capture resource output option differs")
        path_suffix(command[3], f"runs/{lane_id}/resource.log", f"{label}.argv.resource")
        binaries = obj(binding.get("binaries"), "binary-binding.binaries")
        binary = obj(binaries.get(lane["instrumentation"]), f"binary-binding.binaries.{lane['instrumentation']}")
        try:
            binaries = obj(binding.get("binaries"), "binary-binding.binaries")
            binary = obj(binaries.get(lane["instrumentation"]), f"binary-binding.binaries.{lane['instrumentation']}")
            expected_binary = text(binary.get("path"), f"{label}.binary.path")
            if command[7] != expected_binary:
                fail(f"{label}: capture binary path differs")
            workload = command[7:]
            output_index = workload.index("--output")
            report_arg = workload[output_index + 1]
        except (ValueError, IndexError):
            fail(f"{label}: capture argv has no output report")
        path_suffix(report_arg, f"runs/{lane_id}/report.json", f"{label}.argv.report")
        expected = expected_workload(protocol, lane, expected_binary, report_arg)
        if workload != expected:
            fail(f"{label}: workload argv differs from protocol")
        proof = oracle.validate(report_path, lane, protocol)
        if receipt.get("oracle") != proof:
            fail(f"{label}: retained oracle proof differs from recomputation")
        expected_artifacts = {
            f"runs/{lane_id}/report.json",
            f"runs/{lane_id}/stdout.log",
            f"runs/{lane_id}/stderr.log",
            f"runs/{lane_id}/resource.log",
        }
        check_run_artifacts(run_root, receipt, label, expected_artifacts)


def profile_command_ok(
    command: Any,
    label: str,
    kind: str,
    protocol: Mapping[str, Any],
    binary: str,
) -> None:
    if not isinstance(command, list) or not command or not all(isinstance(item, str) for item in command):
        fail(f"{label}: workload argv is malformed")
    if command[:2] != ["/usr/bin/time", "-v"] or "taskset" not in command or "--" not in command:
        fail(f"{label}: workload argv metadata differs")
    if command[2] != "-o" or command[4:6] != ["taskset", "-c"] or command[6] != str(protocol["cpu"]):
        fail(f"{label}: workload argv prefix differs")
    separator = command.index("--")
    if separator < 8:
        fail(f"{label}: workload argv separator is misplaced")
    if kind == "counters":
        if command[7] != "perf" or command[8:10] != ["stat", "-x,"] or command[10] != "-o":
            fail(f"{label}: counter profiler argv differs")
        counter_path = path_suffix(command[11], "profiling/counters/counters.csv", f"{label}.argv.counters")
        profiler = command[7:separator]
        expected_profiler = ["perf", "stat", "-x,", "-o", counter_path, "-e", "cycles,instructions,branches,branch-misses,cache-misses,page-faults,context-switches"]
    else:
        if command[7] != "perf" or command[8:10] != ["record", "--no-buildid-cache"] or command[10:12] != ["-F", "99"] or command[12:14] != ["-e", "cycles:u"] or command[14:16] != ["--call-graph", "dwarf,8192"] or command[16] != "-o":
            fail(f"{label}: sampled profiler argv differs")
        perf_data = path_suffix(command[17], "profiling/samples/perf.data", f"{label}.argv.perf_data")
        profiler = command[7:separator]
        expected_profiler = ["perf", "record", "--no-buildid-cache", "-F", "99", "-e", "cycles:u", "--call-graph", "dwarf,8192", "-o", perf_data]
    if profiler != expected_profiler:
        fail(f"{label}: profiler argv differs")
    workload = command[separator + 1:]
    try:
        output_index = workload.index("--output")
        report = workload[output_index + 1]
    except (ValueError, IndexError):
        fail(f"{label}: workload output report path is missing")
    path_suffix(report, f"profiling/{kind}/report.json", f"{label}.argv.report")
    lane = {"id": kind, "scope": "phases", "shape": "large", "instrumentation": "normal", "repeat": "diagnostic"}
    expected_workload = [part.format(binary=binary, report=report, **lane) for part in protocol["diagnostic_argv"]]
    if workload != expected_workload:
        fail(f"{label}: diagnostic workload argv differs from protocol")


def check_profiles(protocol: Mapping[str, Any], binding: Mapping[str, Any], oracle: Any) -> None:
    binaries = obj(binding.get("binaries"), "binary-binding.binaries")
    normal = obj(binaries.get("normal"), "binary-binding.binaries.normal")
    profiles_root = ROOT / "profiling"
    if not profiles_root.is_dir() or profiles_root.is_symlink():
        fail("profiling directory is missing or unsafe")
    entries = list(profiles_root.iterdir())
    if {entry.name for entry in entries} != {"counters", "samples"} or any(entry.is_symlink() or not entry.is_dir() for entry in entries):
        fail("profiling directory must contain only counters and samples")
    expected_artifacts = {
        "counters": {
            "profiling/counters/report.json",
            "profiling/counters/resource.log",
            "profiling/counters/counters.csv",
            "profiling/counters/workload.stdout",
            "profiling/counters/workload.stderr",
        },
        "samples": {
            "profiling/samples/report.json",
            "profiling/samples/resource.log",
            "profiling/samples/perf.data",
            "profiling/samples/workload.stdout",
            "profiling/samples/workload.stderr",
            "profiling/samples/top-symbols.stdout",
            "profiling/samples/top-symbols.stderr",
            "profiling/samples/perf-script.stdout",
            "profiling/samples/perf-script.stderr",
        },
    }
    for kind in ("counters", "samples"):
        run_root = ROOT / "profiling" / kind
        receipt_path = run_root / "receipt.json"
        receipt = obj(load(receipt_path, f"profiling/{kind}/receipt.json"), f"profiling/{kind}/receipt.json")
        if receipt.get("schema") != "litchi-0458-profile-v1" or receipt.get("change") != CHANGE or receipt.get("kind") != kind:
            fail(f"profiling/{kind}: receipt identity differs")
        if receipt.get("protocol_sha256") != sha_file(ROOT / "protocol.json") or receipt.get("binary") != normal:
            fail(f"profiling/{kind}: protocol or binary binding differs")
        if receipt.get("status") != "pass":
            fail(f"profiling/{kind}: profile did not pass")
        workload = obj(receipt.get("workload"), f"profiling/{kind}.workload")
        if workload.get("exit_code") != 0:
            fail(f"profiling/{kind}: workload exit code differs")
        report_path = run_root / "report.json"
        regular(report_path, f"profiling/{kind}/report.json")
        profile_command_ok(workload.get("argv"), f"profiling/{kind}", kind, protocol, text(normal.get("path"), "binary-binding.binaries.normal.path"))
        profile_protocol = dict(protocol)
        profile_protocol["samples"] = 100
        lane = {"id": kind, "scope": "phases", "shape": "large", "instrumentation": "normal", "repeat": "diagnostic"}
        proof = oracle.validate(report_path, lane, profile_protocol)
        if receipt.get("oracle") != proof:
            fail(f"profiling/{kind}: retained oracle proof differs from recomputation")
        if kind == "samples":
            for field in ("top_symbols", "script"):
                if obj(receipt.get(field), f"profiling/{kind}.{field}").get("exit_code") != 0:
                    fail(f"profiling/{kind}: {field} command did not pass")
        check_run_artifacts(run_root, receipt, f"profiling/{kind}", expected_artifacts[kind])


def recompute_json(path: Path, script: Path, label: str) -> None:
    if not path.is_file():
        fail(f"{label}: retained JSON is missing")
    environment = os.environ.copy()
    environment["PYTHONPATH"] = ""
    environment["PYTHONDONTWRITEBYTECODE"] = "1"
    result = subprocess.run([sys.executable, "-B", str(script)], cwd=ROOT, env=environment, capture_output=True, text=True)
    if result.returncode != 0:
        fail(f"{label}: recomputation failed ({result.stderr.strip() or result.stdout.strip()})")
    try:
        actual = json.loads(result.stdout)
    except json.JSONDecodeError as error:
        fail(f"{label}: recomputation did not return JSON ({error})")
    expected = load(path, label)
    if actual != expected:
        fail(f"{label}: recomputed JSON differs from retained file")


def verify(precleanup: bool = False, portable: bool = False) -> dict[str, Any]:
    if precleanup and portable:
        fail("--precleanup and --portable are mutually exclusive")
    sealed_files = verify_sum_file(ROOT / "SHA256SUMS", ROOT)
    protocol, binding = check_protocol()
    receipts = check_check_receipts(protocol)
    selected = check_required_gates(protocol, receipts)
    check_prior_control(protocol)
    oracle = load_oracle()
    check_runs(protocol, binding, oracle)
    check_profiles(protocol, binding, oracle)
    recompute_json(ROOT / "summary.json", ROOT / "derive.py", "summary.json")
    recompute_json(ROOT / "profile-summary.json", ROOT / "derive-profile.py", "profile-summary.json")
    if precleanup:
        check_live_binding(binding)
    return {
        "schema": SCHEMA,
        "change": CHANGE,
        "status": "pass",
        "precleanup": precleanup,
        "portable": portable,
        "sealed_files": sealed_files,
        "selected_gates": selected,
        "lanes": 24,
        "profiles": 2,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--precleanup", action="store_true", help="also authenticate live binaries and current source")
    parser.add_argument("--portable", action="store_true", help="mark this as a portable bundle replay")
    args = parser.parse_args(sys.argv[1:] if argv is None else argv)
    try:
        result = verify(args.precleanup, args.portable)
    except (OSError, KeyError, TypeError, ValueError, VerificationError) as error:
        result = {"schema": SCHEMA, "change": CHANGE, "status": "fail", "precleanup": args.precleanup, "portable": args.portable, "error": str(error)}
    print(json.dumps(result, ensure_ascii=False, sort_keys=True, indent=2))
    return 0 if result.get("status") == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
