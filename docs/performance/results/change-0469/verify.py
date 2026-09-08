#!/usr/bin/env python3
"""Fail-closed verifier for the 0469 XLSX compaction evidence.

The verifier authenticates the frozen protocol, role bindings, raw report
oracles, normal ABBA summary, full-matrix guard, heaptrack exports, and the
post-cleanup file seal.  It never builds or profiles.  Heaptrack elapsed time
and RSS are retained as whole-process diagnostics and are explicitly excluded
from latency comparison.
"""

from __future__ import annotations

import argparse
import copy
import datetime
import hashlib
import importlib.util
import json
import re
from pathlib import Path
import subprocess
import sys
from typing import Any, Mapping, Sequence


ROOT = Path(__file__).resolve().parent
CHANGE = 469
SCHEMA = "litchi-0469-verification-v1"
LANES = ("A1", "B1", "B2", "A2", "A-full", "B-full", "A-heap", "B-heap")
PROBE_LANES = ("A1", "B1", "B2", "A2")
FULL_LANES = ("A-full", "B-full")
HEAP_LANES = ("A-heap", "B-heap")
ROLE_FOR_LANE = {
    "A1": "control",
    "A2": "control",
    "B1": "candidate",
    "B2": "candidate",
    "A-full": "control",
    "B-full": "candidate",
    "A-heap": "control",
    "B-heap": "candidate",
}
SHA256_RE = re.compile(r"[0-9a-f]{64}\Z")
EXPECTED_TOOL = {
    "name": "litchi-perf-baseline",
    "version": "0.1.0",
    "binary": "litchi-perf-baseline",
    "profile": "release",
    "target_os": "linux",
    "target_arch": "x86_64",
    "instrumentation": "none",
}
EXPECTED_ENVIRONMENT = {
    "RUSTUP_TOOLCHAIN": "1.98.1",
    "RUSTFLAGS": "-C force-frame-pointers=yes -C force-unwind-tables=yes",
    "CARGO_PROFILE_RELEASE_DEBUG": "1",
    "DEBUGINFOD_URLS": "",
    "LC_ALL": "C",
}
EXPECTED_FIXTURES = {
    "test-data/poi/test-data/spreadsheet/54016.xls",
    "test-data/rtf/watermark.rtf",
}


class VerificationError(ValueError):
    """Raised for missing, malformed, or mismatched evidence."""


def fail(message: str) -> None:
    raise VerificationError(message)


def regular(path: Path, label: str) -> Path:
    if path.is_symlink() or not path.is_file():
        fail(f"{label}: missing, symlinked, or non-regular file")
    return path


def canonical(value: Any) -> bytes:
    try:
        return json.dumps(
            value,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
            allow_nan=False,
        ).encode("utf-8")
    except (TypeError, ValueError) as error:
        fail(f"cannot canonicalize JSON: {error}")
    raise AssertionError("unreachable")


def canonical_equal(left: Any, right: Any) -> bool:
    return canonical(left) == canonical(right)


def sha256_file(path: Path) -> tuple[str, int]:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            size = 0
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
                size += len(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest(), size


def read_json(path: Path, label: str) -> dict[str, Any]:
    regular(path, label)
    try:
        with path.open(encoding="utf-8") as stream:
            value = json.load(
                stream,
                object_pairs_hook=_reject_duplicate_keys,
                parse_constant=lambda item: (_ for _ in ()).throw(
                    ValueError(f"non-finite JSON value {item}")
                ),
            )
    except (OSError, UnicodeError, ValueError) as error:
        fail(f"{label}: invalid JSON ({error})")
    if not isinstance(value, dict):
        fail(f"{label}: expected object")
    return value


def _reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate key {key!r}")
        result[key] = value
    return result


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: expected object")
    return value


def nonempty_text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label}: expected non-empty string")
    return value


def digest(value: Any, label: str) -> str:
    value = nonempty_text(value, label).lower()
    if SHA256_RE.fullmatch(value) is None:
        fail(f"{label}: expected lowercase SHA-256")
    return value


def integer(value: Any, label: str, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(f"{label}: expected integer >= {minimum}")
    return value


def relative(value: Any, label: str) -> Path:
    raw = nonempty_text(value, label)
    path = Path(raw)
    if path.is_absolute() or path.as_posix() != raw or ".." in path.parts:
        fail(f"{label}: expected a relative traversal-free path")
    return path


def bundle_file(root: Path, value: Any, label: str) -> Path:
    path = root / relative(value, label)
    try:
        if not path.resolve(strict=True).is_relative_to(root.resolve()):
            fail(f"{label}: path escapes bundle")
    except OSError as error:
        fail(f"{label}: cannot resolve ({error})")
    return regular(path, label)


def retained_reference(root: Path, value: Any, label: str) -> Path:
    """Resolve a retained adjacent evidence reference without allowing escape."""

    raw = nonempty_text(value, label)
    path = Path(raw)
    if path.is_absolute() or path.as_posix() != raw:
        fail(f"{label}: expected a relative path")
    candidate = (root / path).resolve()
    try:
        candidate.relative_to(root.parent.resolve())
    except ValueError:
        fail(f"{label}: retained reference escapes the results directory")
    return regular(candidate, label)


def file_binding(path: Path, root: Path) -> dict[str, Any]:
    sha, size = sha256_file(path)
    try:
        name = path.resolve().relative_to(root.resolve()).as_posix()
    except ValueError:
        name = f"tools/{path.name}" if path.parent.name == "tools" else path.name
    return {"path": name, "sha256": sha, "bytes": size}


def _load_module(path: Path, name: str) -> Any:
    regular(path, str(path))
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        fail(f"cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    try:
        spec.loader.exec_module(module)
    except (OSError, SyntaxError, TypeError, ValueError) as error:
        fail(f"cannot load {path}: {error}")
    return module


def _load_analyzer(root: Path) -> tuple[Any, Path]:
    path = root / "analyze.py"
    return _load_module(path, "litchi_change0469_analyze_verify"), path


def _load_comparator(root: Path) -> tuple[Any, Path]:
    candidates = [root / "tools" / "perf_compare.py"]
    candidates.extend(parent / "tools" / "perf_compare.py" for parent in root.parents)
    for path in candidates:
        if path.is_file():
            return _load_module(path, "litchi_change0469_perf_compare_verify"), path
    fail("tools/perf_compare.py: dependency missing")


def verify_protocol(root: Path) -> dict[str, Any]:
    protocol = read_json(root / "protocol.json", "protocol.json")
    if protocol.get("schema") != "litchi-0469-protocol-v1":
        fail("protocol.json: schema differs")
    if protocol.get("cpu") != 2 or protocol.get("workers") != 1:
        fail("protocol.json: CPU or worker identity differs")
    if protocol.get("probe_order") != list(PROBE_LANES):
        fail("protocol.json: probe order differs")
    if protocol.get("probe_samples") != 100 or protocol.get("probe_warmups") != 5:
        fail("protocol.json: probe sample configuration differs")
    if protocol.get("probe_cases") != [
        "xlsx_one_cell_commit_save",
        "xlsx_one_percent_commit_save",
    ] or protocol.get("probe_shapes") != ["tiny", "medium", "dense-wide"]:
        fail("protocol.json: probe workload differs")
    if protocol.get("heap_order") != list(HEAP_LANES):
        fail("protocol.json: heap order differs")
    if protocol.get("heap_samples") != 5 or protocol.get("heap_warmups") != 1:
        fail("protocol.json: heap sample configuration differs")
    if protocol.get("heap_case") != "xlsx_one_percent_commit_save" or protocol.get("heap_shape") != "dense-wide":
        fail("protocol.json: heap workload differs")
    if protocol.get("full_guard_order") != list(FULL_LANES):
        fail("protocol.json: full guard order differs")
    if protocol.get("full_guard_samples") != 15 or protocol.get("full_guard_warmups") != 3:
        fail("protocol.json: full guard sample configuration differs")
    if "none" not in str(protocol.get("performance_claim", "")).lower():
        fail("protocol.json: performance claim is not explicitly descriptive")
    driver = regular(root / "capture.py", "capture.py")
    if protocol.get("capture_driver_sha256") != sha256_file(driver)[0]:
        fail("protocol.json: capture driver hash differs")
    return protocol


def verify_policy(root: Path, comparator: Any) -> dict[str, Any]:
    policy = read_json(root / "report-policy.json", "report-policy.json")
    try:
        checked = comparator.validate_policy(policy)
    except (AttributeError, KeyError, TypeError, ValueError) as error:
        fail(f"report-policy.json: comparator validation failed ({error})")
    if checked.get("minimum_samples") != 15 or checked.get("expected_result_count") != 201:
        fail("report-policy.json: full guard cardinality differs")
    if checked.get("require_distinct_revisions") is not True or checked.get("require_clean_worktree") is not True:
        fail("report-policy.json: guard identity requirements are weakened")
    return checked


def verify_role_binding(root: Path, role: str) -> dict[str, Any]:
    path = regular(root / f"{role}-binding.json", f"{role}-binding.json")
    binding = read_json(path, f"{role}-binding.json")
    if binding.get("schema") != "litchi-0469-role-binding-v1" or binding.get("role") != role:
        fail(f"{role}-binding.json: schema or role differs")
    nonempty_text(binding.get("revision"), f"{role}.revision")
    binary_path = Path(nonempty_text(binding.get("binary_path"), f"{role}.binary_path"))
    build_path = Path(nonempty_text(binding.get("build_path"), f"{role}.build_path"))
    if not binary_path.is_absolute() or not build_path.is_absolute():
        fail(f"{role}-binding.json: live paths must be absolute")
    manifest = bundle_file(root, binding.get("source_manifest"), f"{role}.source_manifest")
    if digest(binding.get("source_manifest_sha256"), f"{role}.source_manifest_sha256") != sha256_file(manifest)[0]:
        fail(f"{role}: source manifest hash differs")
    source_map = read_json(manifest, f"{role} source manifest")
    if len(source_map) != 6992:
        fail(f"{role} source manifest file count differs")
    for name, value in source_map.items():
        relative(name, f"{role}.sources.{name}")
        digest(value, f"{role}.sources.{name}")
    fixtures = obj(binding.get("included_fixtures"), f"{role}.included_fixtures")
    if set(fixtures) != EXPECTED_FIXTURES:
        fail(f"{role}: compile fixture set differs")
    for name, value in fixtures.items():
        digest(value, f"{role}.included_fixtures.{name}")
    integer(binding.get("bytes"), f"{role}.bytes", 1)
    digest(binding.get("binary_sha256"), f"{role}.binary_sha256")
    if "source_binding_sha256" in binding:
        source_binding = regular(root / f"{role}-source-binding.json", f"{role}-source-binding.json")
        if digest(binding["source_binding_sha256"], f"{role}.source_binding_sha256") != sha256_file(source_binding)[0]:
            fail(f"{role}: source binding hash differs")
        source_metadata = read_json(source_binding, f"{role}-source-binding.json")
        if source_metadata.get("schema") != "litchi-0469-source-binding-v1" or source_metadata.get("source_count") != 6992 or source_metadata.get("revision") != binding.get("revision") or source_metadata.get("build_path") != binding.get("build_path"):
            fail(f"{role}-source-binding.json: source identity differs")
        if source_metadata.get("changed_files") != ["crates/litchi-xlsx/src/raw/compact.rs"]:
            fail(f"{role}-source-binding.json: changed file set differs")
    for field in ("prior_binding_path", "prior_build_receipt_path"):
        if field in binding:
            retained_reference(root, binding[field], f"{role}.{field}")
    if "build_receipt_path" in binding:
        bundle_file(root, binding["build_receipt_path"], f"{role}.build_receipt_path")
    for path_field, hash_field in (
        ("prior_binding_path", "prior_binding_sha256"),
        ("prior_build_receipt_path", "prior_build_receipt_sha256"),
        ("build_receipt_path", "build_receipt_sha256"),
    ):
        if path_field in binding:
            if hash_field not in binding:
                fail(f"{role}-binding.json: {hash_field} is required with {path_field}")
            referenced = retained_reference(root, binding[path_field], f"{role}.{path_field}")
            if digest(binding[hash_field], f"{role}.{hash_field}") != sha256_file(referenced)[0]:
                fail(f"{role}-binding.json: {path_field} hash differs")
    if role == "control":
        prior_path = retained_reference(root, binding.get("prior_binding_path"), "control prior binding")
        prior = read_json(prior_path, "control prior binding")
        for field in ("revision", "binary_sha256", "bytes"):
            if prior.get(field) != binding.get(field):
                fail(f"control prior binding: {field} differs")
        if prior.get("build_receipt_sha256") != binding.get("prior_build_receipt_sha256"):
            fail("control prior binding: build receipt differs")
        prior_source_path = regular(prior_path.with_name("source-binding.json"), "control prior source binding")
        if sha256_file(prior_source_path)[0] != prior.get("source_binding_sha256"):
            fail("control prior source binding: hash differs")
        prior_source = read_json(prior_source_path, "control prior source binding")
        for field in ("revision", "source_manifest_sha256", "included_fixtures"):
            if prior_source.get(field) != binding.get(field):
                fail(f"control prior source binding: {field} differs")
        if prior_source.get("clean_tree") != binding.get("build_path"):
            fail("control prior source binding: build path differs")
    return binding


def verify_build_receipt(root: Path, binding: Mapping[str, Any], role: str) -> dict[str, Any] | None:
    field = "build_receipt_path"
    if field not in binding:
        return None
    path = bundle_file(root, binding[field], f"{role}.{field}")
    receipt = read_json(path, f"{role} build receipt")
    if receipt.get("schema") != "litchi-0469-command-v1":
        fail(f"{role} build receipt: schema differs")
    if receipt.get("exit_code") != 0:
        fail(f"{role} build receipt: command failed")
    if receipt.get("binary_sha256") not in (None, binding.get("binary_sha256")):
        fail(f"{role} build receipt: binary hash differs")
    if receipt.get("revision") not in (None, binding.get("revision")):
        fail(f"{role} build receipt: revision differs")
    if receipt.get("cwd") != binding.get("build_path"):
        fail(f"{role} build receipt: build path differs")
    argv = receipt.get("argv")
    expected = ("cargo", "build", "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--target-dir", "/home/zhuhe/code/litchi/tools/perf-baseline/target", "--bin", "litchi-perf-baseline")
    if argv != list(expected):
        fail(f"{role} build receipt: build argv differs")
    environment = obj(receipt.get("environment"), f"{role} build receipt.environment")
    if environment != {**EXPECTED_ENVIRONMENT, "CARGO_BUILD_JOBS": "4", "CARGO_INCREMENTAL": "0"}:
        fail(f"{role} build receipt: build environment differs")
    driver = regular(root / "build.py", "build.py")
    if receipt.get("driver_sha256") != sha256_file(driver)[0]:
        fail(f"{role} build receipt: build driver differs")
    started = root / "build.started.json"
    if started.is_file():
        started_receipt = read_json(started, "build.started.json")
        for field in ("schema", "argv", "cwd", "driver_sha256", "environment"):
            if started_receipt.get(field) != receipt.get(field):
                fail(f"{role} build receipts: {field} differs")
        for suffix in (".stdout", ".stderr"):
            metadata = obj(receipt.get("outputs", {}).get(suffix), f"{role}.build.outputs.{suffix}")
            output = regular(root / f"build{suffix}", f"build{suffix}")
            if digest(metadata.get("sha256"), f"{role}.build.outputs.{suffix}.sha256") != sha256_file(output)[0] or integer(metadata.get("bytes"), f"{role}.build.outputs.{suffix}.bytes") != output.stat().st_size:
                fail(f"{role} build receipt: {suffix} hash differs")
    return receipt


def verify_live_state(root: Path, bindings: Mapping[str, Mapping[str, Any]], *, live: bool) -> dict[str, Any]:
    state: dict[str, Any] = {"mode": "live" if live else "post-cleanup", "roles": {}}
    if not live:
        for role, binding in bindings.items():
            tree = Path(nonempty_text(binding.get("build_path"), f"{role}.build_path"))
            binary = Path(nonempty_text(binding.get("binary_path"), f"{role}.binary_path"))
            if tree.exists() or binary.exists() or tree.is_symlink() or binary.is_symlink():
                fail(f"post-cleanup: live {role} tree or binary remains")
            state["roles"][role] = {"tree_absent": True, "binary_absent": True}
        return state

    # The control and candidate binaries intentionally share one absolute
    # checkout path.  Inspect that tree once; whichever role's revision is
    # currently checked out gets source-file verification.  Both binary paths
    # are authenticated independently, and the older role remains bound by
    # its retained source manifest and build receipt.
    trees: dict[Path, list[tuple[str, Mapping[str, Any]]]] = {}
    for role, binding in bindings.items():
        tree = Path(nonempty_text(binding.get("build_path"), f"{role}.build_path"))
        binary = Path(nonempty_text(binding.get("binary_path"), f"{role}.binary_path"))
        trees.setdefault(tree, []).append((role, binding))
        if tree.is_symlink() or not tree.is_dir() or binary.is_symlink() or not binary.is_file():
            fail(f"live: {role} source tree or binary is absent")
        actual_sha, actual_bytes = sha256_file(binary)
        if actual_sha != binding.get("binary_sha256") or actual_bytes != binding.get("bytes"):
            fail(f"live: {role} binary identity differs")
        state["roles"][role] = {
            "binary_sha256": actual_sha,
            "binary_bytes": actual_bytes,
        }
    for tree, roles in trees.items():
        try:
            revision = subprocess.check_output(
                ["git", "rev-parse", "HEAD"], cwd=tree, text=True
            ).strip()
            dirty = subprocess.check_output(
                ["git", "status", "--porcelain"], cwd=tree, text=True
            )
        except (OSError, subprocess.CalledProcessError) as error:
            fail(f"live: cannot inspect {tree} ({error})")
        if dirty:
            fail(f"live: source tree is dirty: {tree}")
        matching = [(role, binding) for role, binding in roles if revision == binding.get("revision")]
        if not matching:
            fail(f"live: {tree} revision does not match any role binding")
        for role, binding in matching:
            source_map = read_json(
                root / relative(binding["source_manifest"], f"{role}.source_manifest"),
                f"{role} source manifest",
            )
            for name, expected in source_map.items():
                source = tree / relative(name, f"{role}.sources.{name}")
                if source.is_symlink() or not source.is_file() or sha256_file(source)[0] != expected:
                    fail(f"live: {role} source digest differs for {name}")
            for name, expected in obj(binding["included_fixtures"], f"{role}.fixtures").items():
                fixture = tree / relative(name, f"{role}.fixture.{name}")
                if fixture.is_symlink() or not fixture.is_file() or sha256_file(fixture)[0] != expected:
                    fail(f"live: {role} fixture digest differs for {name}")
            state["roles"][role].update({"tree": str(tree), "revision": revision, "source_verified": True})
    return state


def _lane_binding(binding: Mapping[str, Any], role: str, root: Path) -> str:
    expected = digest(binding.get("binary_sha256"), f"{role}.binary_sha256")
    return expected


def verify_lane_receipt(
    root: Path,
    lane: str,
    protocol: Mapping[str, Any],
    bindings: Mapping[str, Mapping[str, Any]],
) -> dict[str, Any]:
    role = ROLE_FOR_LANE[lane]
    lane_root = root / lane
    receipt = read_json(lane_root / "receipt.json", f"{lane}.receipt.json")
    started = read_json(lane_root / "started.json", f"{lane}.started.json")
    for item, label in ((started, "started"), (receipt, "receipt")):
        if item.get("schema") != "litchi-0469-capture-v1":
            fail(f"{lane}.{label}: schema differs")
        if item.get("lane") != lane or item.get("role") != role:
            fail(f"{lane}.{label}: lane or role differs")
        if item.get("revision") != bindings[role].get("revision"):
            fail(f"{lane}.{label}: revision differs")
        if item.get("binary_sha256") != bindings[role].get("binary_sha256"):
            fail(f"{lane}.{label}: binary identity differs")
        if item.get("binding_sha256") != sha256_file(root / f"{role}-binding.json")[0]:
            fail(f"{lane}.{label}: role binding identity differs")
        if item.get("driver_sha256") != protocol.get("capture_driver_sha256"):
            fail(f"{lane}.{label}: capture driver differs")
        if item.get("protocol_sha256") != sha256_file(root / "protocol.json")[0]:
            fail(f"{lane}.{label}: protocol identity differs")
    if receipt.get("exit_code") != 0 or receipt.get("clean_before") is not True or receipt.get("clean_after") is not True:
        fail(f"{lane}.receipt: capture did not attest success and cleanliness")
    if receipt.get("binary_unchanged") is not True or receipt.get("report_metadata_matches_clean_role") is not True:
        fail(f"{lane}.receipt: binary/report identity attestation missing")
    expected_samples, expected_warmups = lane_configuration(protocol, lane)
    if receipt.get("samples") != expected_samples or receipt.get("warmups") != expected_warmups:
        fail(f"{lane}.receipt: sample configuration differs")
    if started.get("argv") != receipt.get("argv") or started.get("cwd") != receipt.get("cwd"):
        fail(f"{lane}: started/final command identity differs")
    argv = receipt.get("argv")
    if not isinstance(argv, list) or any(not isinstance(value, str) for value in argv):
        fail(f"{lane}.receipt: argv is malformed")
    required_argv = ["taskset", "-c", "2", "/usr/bin/time", "-v", "-o"]
    position = 0
    for token in required_argv:
        try:
            position = argv.index(token, position) + 1
        except ValueError:
            fail(f"{lane}.receipt: argv lacks {token!r}")
    if not any(value.endswith(f"/{lane}/resource.log") for value in argv):
        fail(f"{lane}.receipt: resource log identity differs")
    if not any(value == str(bindings[role]["binary_path"]) for value in argv):
        fail(f"{lane}.receipt: binary argv identity differs")
    for token in ("--workers", "1", "--warmup", str(expected_warmups), "--samples", str(expected_samples), "--json", "--corpus-manifest"):
        if token not in argv:
            fail(f"{lane}.receipt: argv lacks {token!r}")
    if lane in HEAP_LANES:
        for token in ("heaptrack", "-o", "--"):
            if token not in argv:
                fail(f"{lane}.receipt: heaptrack argv lacks {token!r}")
    elif lane in PROBE_LANES:
        if "--case" not in argv or "--xlsx-shape" not in argv:
            fail(f"{lane}.receipt: probe selector argv is incomplete")
    else:
        if "--case" in argv or "--xlsx-shape" in argv:
            fail(f"{lane}.receipt: full guard unexpectedly narrows selectors")
    required = {"report.json", "corpus-catalog.json", "resource.log", "stdout.log", "stderr.log", "started.json"}
    for name in required:
        regular(lane_root / name, f"{lane}/{name}")
    artifacts = obj(receipt.get("artifacts"), f"{lane}.artifacts")
    for name, metadata in artifacts.items():
        path = bundle_file(root, f"{lane}/{name}", f"{lane}.artifacts.{name}")
        metadata = obj(metadata, f"{lane}.artifacts.{name}")
        if digest(metadata.get("sha256"), f"{lane}.{name}.sha256") != sha256_file(path)[0]:
            fail(f"{lane}: artifact hash differs for {name}")
        if integer(metadata.get("bytes"), f"{lane}.{name}.bytes") != path.stat().st_size:
            fail(f"{lane}: artifact size differs for {name}")
    return receipt


def lane_configuration(protocol: Mapping[str, Any], lane: str) -> tuple[int, int]:
    if lane in PROBE_LANES:
        return int(protocol["probe_samples"]), int(protocol["probe_warmups"])
    if lane in FULL_LANES:
        return int(protocol["full_guard_samples"]), int(protocol["full_guard_warmups"])
    return int(protocol["heap_samples"]), int(protocol["heap_warmups"])


def _timestamp(value: Any, label: str) -> datetime.datetime:
    raw = nonempty_text(value, label)
    try:
        parsed = datetime.datetime.fromisoformat(raw)
    except ValueError as error:
        fail(f"{label}: invalid timestamp ({error})")
    if parsed.tzinfo is None:
        fail(f"{label}: timestamp lacks timezone")
    return parsed


def verify_chronology(
    root: Path,
    receipts: Mapping[str, Mapping[str, Any]],
    build_receipt: Mapping[str, Any] | None,
    export_receipts: Mapping[str, Mapping[str, Any]],
) -> dict[str, Any]:
    """Check the retained command intervals do not overlap."""

    intervals: list[tuple[str, datetime.datetime, datetime.datetime]] = []
    for lane in LANES:
        receipt = receipts[lane]
        start = _timestamp(receipt.get("started_utc"), f"{lane}.started_utc")
        finish = _timestamp(receipt.get("finished_utc"), f"{lane}.finished_utc")
        if finish < start:
            fail(f"{lane}: capture finished before it started")
        intervals.append((lane, start, finish))
    if build_receipt is not None:
        start = _timestamp(build_receipt.get("started_utc"), "build.started_utc")
        finish = _timestamp(build_receipt.get("finished_utc"), "build.finished_utc")
        if finish < start:
            fail("build: finished before it started")
        candidate_start = min(begin for name, begin, _ in intervals if ROLE_FOR_LANE[name] == "candidate")
        if finish > candidate_start:
            fail("build: finishes after the first candidate capture starts")
        if any(start < end and finish > begin for _, begin, end in intervals):
            fail("build: overlaps a capture")
    by_name = {name: (start, finish) for name, start, finish in intervals}
    for ordered in (PROBE_LANES, FULL_LANES, HEAP_LANES):
        for previous_name, current_name in zip(ordered, ordered[1:]):
            if by_name[current_name][0] < by_name[previous_name][1]:
                fail(f"capture chronology order/overlap differs: {previous_name} and {current_name}")
    chronologically_sorted = sorted(intervals, key=lambda item: item[1])
    for previous, current in zip(chronologically_sorted, chronologically_sorted[1:]):
        if current[1] < previous[2]:
            fail(f"capture chronology overlaps: {previous[0]} and {current[0]}")
    last_capture_finish = max(finish for _, _, finish in intervals)
    exports: list[tuple[str, datetime.datetime, datetime.datetime]] = []
    for lane in HEAP_LANES:
        receipt = export_receipts[lane]
        start = _timestamp(receipt.get("started_utc"), f"{lane}.heap-export.started_utc")
        finish = _timestamp(receipt.get("finished_utc"), f"{lane}.heap-export.finished_utc")
        if finish < start:
            fail(f"{lane}.heap-export: finished before it started")
        exports.append((lane, start, finish))
        if start < last_capture_finish:
            fail(f"{lane}.heap-export: starts before the last capture finished")
    if exports[1][1] < exports[0][2]:
        fail("heap export chronology overlaps")
    return {
        "capture_order": [name for name, _, _ in chronologically_sorted],
        "capture_intervals_non_overlapping": True,
        "build_before_candidate_captures": build_receipt is not None,
        "exports_after_captures": True,
    }


def _report_rows(report: Mapping[str, Any], label: str) -> dict[tuple[str, bytes], dict[str, Any]]:
    rows = report.get("results")
    if not isinstance(rows, list):
        fail(f"{label}.results: expected list")
    indexed: dict[tuple[str, bytes], dict[str, Any]] = {}
    for index, raw in enumerate(rows):
        row = obj(raw, f"{label}.results[{index}]")
        case = nonempty_text(row.get("case"), f"{label}.results[{index}].case")
        corpus = obj(row.get("corpus"), f"{label}.results[{index}].corpus")
        key = (case, canonical(corpus))
        if key in indexed:
            fail(f"{label}: duplicate result identity")
        indexed[key] = row
    return indexed


def verify_report_metadata(
    report: Mapping[str, Any],
    label: str,
    binding: Mapping[str, Any],
    samples: int,
    warmups: int,
    *,
    cases: Sequence[str] | None = None,
    shapes: Sequence[str] | None = None,
) -> None:
    if report.get("schema_version") != 1 or report.get("tool") != EXPECTED_TOOL:
        fail(f"{label}: tool/schema identity differs")
    binary = obj(report.get("binary_identity"), f"{label}.binary_identity")
    if binary.get("binary_sha256") != binding.get("binary_sha256") or binary.get("binary_bytes") != binding.get("bytes"):
        fail(f"{label}: binary identity differs from role binding")
    environment = obj(report.get("environment"), f"{label}.environment")
    if environment.get("git_revision") != binding.get("revision") or environment.get("git_worktree_dirty") is not False:
        fail(f"{label}: source revision or cleanliness differs")
    if environment.get("cpu_affinity") != "2" or environment.get("rustflags") != EXPECTED_ENVIRONMENT["RUSTFLAGS"]:
        fail(f"{label}: CPU/build flags differ")
    configuration = obj(report.get("configuration"), f"{label}.configuration")
    if configuration.get("samples_per_case") != samples or configuration.get("warmup_iterations_per_case") != warmups:
        fail(f"{label}: sample configuration differs")
    if cases is not None and configuration.get("cases") != list(cases):
        fail(f"{label}: case set differs")
    if shapes is not None and configuration.get("xlsx_shapes") != list(shapes):
        fail(f"{label}: XLSX shape set differs")
    rows = _report_rows(report, label)
    for key, row in rows.items():
        elapsed = obj(row.get("elapsed_ns"), f"{label}.{key[0]}.elapsed_ns")
        samples_vector = elapsed.get("samples")
        if not isinstance(samples_vector, list) or len(samples_vector) != samples:
            fail(f"{label}.{key[0]}: elapsed sample cardinality differs")
        if any(isinstance(value, bool) or not isinstance(value, int) or value <= 0 for value in samples_vector):
            fail(f"{label}.{key[0]}: elapsed samples are malformed")
        if samples_vector != sorted(samples_vector):
            fail(f"{label}.{key[0]}: elapsed samples are not sorted")
        operation = row.get("operation_metrics")
        if isinstance(operation, dict) and operation.get("sample_count") != samples:
            fail(f"{label}.{key[0]}: operation sample count differs")
    catalog = obj(report.get("corpus_catalog"), f"{label}.corpus_catalog")
    if catalog.get("catalog_id") != "litchi-perf-corpus-v2" or catalog.get("manifest_version") != 2:
        fail(f"{label}.corpus_catalog: identity differs")


def verify_oracles(
    left: Mapping[str, Any],
    right: Mapping[str, Any],
    label: str,
    abba: Any,
) -> None:
    left_rows = _report_rows(left, f"{label}.left")
    right_rows = _report_rows(right, f"{label}.right")
    if set(left_rows) != set(right_rows):
        fail(f"{label}: source/result row identity differs")
    for key in sorted(left_rows, key=lambda item: (item[0], item[1])):
        before, after = left_rows[key], right_rows[key]
        for field in ("source", "sink", "output_sha256"):
            before_present = field in before
            after_present = field in after
            if before_present != after_present:
                fail(f"{label}.{key[0]}: {field} presence differs")
            if before_present and not canonical_equal(before[field], after[field]):
                fail(f"{label}.{key[0]}: {field} oracle differs")
        if "operation_metrics" in before or "operation_metrics" in after:
            if "operation_metrics" not in before or "operation_metrics" not in after:
                fail(f"{label}.{key[0]}: operation metric presence differs")
            try:
                left_identity = abba._operation_metrics_identity(
                    before, f"{label}.left.{key[0]}", 1
                )
                right_identity = abba._operation_metrics_identity(
                    after, f"{label}.right.{key[0]}", 1
                )
            except (AttributeError, KeyError, TypeError, ValueError) as error:
                fail(f"{label}.{key[0]}: operation metric validation failed ({error})")
            if left_identity != right_identity:
                fail(f"{label}.{key[0]}: source/sink operation shape differs")


def verify_heap_export(root: Path, lane: str, binding: Mapping[str, Any], protocol: Mapping[str, Any]) -> dict[str, Any]:
    lane_root = root / lane
    output = regular(lane_root / "heaptrack-print.stdout", f"{lane}/heaptrack-print.stdout")
    error_output = regular(lane_root / "heaptrack-print.stderr", f"{lane}/heaptrack-print.stderr")
    receipt = read_json(lane_root / "heaptrack-print.json", f"{lane}/heaptrack-print.json")
    if receipt.get("schema") != "litchi-0469-heap-export-v1" or receipt.get("exit_code") != 0:
        fail(f"{lane}: heap export receipt differs")
    exporter = regular(root / "export_heap.py", "export_heap.py")
    if receipt.get("driver_sha256") != sha256_file(exporter)[0]:
        fail(f"{lane}: malformed heap exporter hash")
    raw_name = nonempty_text(receipt.get("input_path"), f"{lane}.heap-export.input_path")
    raw = bundle_file(root, f"{lane}/{raw_name}", f"{lane}.heap-export.input")
    if receipt.get("input_sha256") != sha256_file(raw)[0]:
        fail(f"{lane}: heap input hash differs")
    argv = receipt.get("argv")
    expected_tokens = ["heaptrack_print", "-f", "-n", "0", "-a", "0", "-p", "0", "-T", "0", "-l", "0"]
    if not isinstance(argv, list) or any(not isinstance(value, str) for value in argv):
        fail(f"{lane}: heap exporter argv is malformed")
    position = 0
    for token in expected_tokens:
        try:
            position = argv.index(token, position) + 1
        except ValueError:
            fail(f"{lane}: heap exporter argv lacks {token!r}")
    artifacts = obj(receipt.get("artifacts"), f"{lane}.heap-export.artifacts")
    if set(artifacts) != {"heaptrack-print.stdout", "heaptrack-print.stderr"}:
        fail(f"{lane}: heap export artifact set differs")
    for name, metadata in artifacts.items():
        path = bundle_file(root, f"{lane}/{name}", f"{lane}.heap-export.{name}")
        metadata = obj(metadata, f"{lane}.heap-export.{name}")
        if digest(metadata.get("sha256"), f"{lane}.heap-export.{name}.sha256") != sha256_file(path)[0]:
            fail(f"{lane}: heap export artifact hash differs for {name}")
        if integer(metadata.get("bytes"), f"{lane}.heap-export.{name}.bytes") != path.stat().st_size:
            fail(f"{lane}: heap export artifact size differs for {name}")
    if not output.stat().st_size:
        fail(f"{lane}: empty heaptrack output")
    if not error_output.exists():
        fail(f"{lane}: heaptrack stderr is missing")
    return receipt


def verify_sha256sums(root: Path) -> int:
    sums = regular(root / "SHA256SUMS", "SHA256SUMS")
    listed: dict[str, str] = {}
    for number, line in enumerate(sums.read_text(encoding="utf-8").splitlines(), 1):
        fields = line.split("  ", 1)
        if len(fields) != 2 or SHA256_RE.fullmatch(fields[0]) is None:
            fail(f"SHA256SUMS: malformed line {number}")
        name = relative(fields[1], f"SHA256SUMS.{number}").as_posix()
        if name == "SHA256SUMS" or name in listed:
            fail(f"SHA256SUMS: duplicate or self entry {name}")
        member = bundle_file(root, name, f"SHA256SUMS.{name}")
        if sha256_file(member)[0] != fields[0]:
            fail(f"SHA256SUMS: hash differs for {name}")
        listed[name] = fields[0]
    actual = {
        path.relative_to(root).as_posix()
        for path in root.rglob("*")
        if path.is_file() and path != sums
    }
    if set(listed) != actual:
        fail("SHA256SUMS: exact file coverage differs")
    return len(listed)


def verify(*, root: Path = ROOT, live: bool = False) -> dict[str, Any]:
    root = root.resolve()
    protocol = verify_protocol(root)
    comparator, comparator_path = _load_comparator(root)
    policy = verify_policy(root, comparator)
    bindings = {role: verify_role_binding(root, role) for role in ("control", "candidate")}
    verify_build_receipt(root, bindings["control"], "control")
    build_receipt = verify_build_receipt(root, bindings["candidate"], "candidate")
    if bindings["control"].get("revision") == bindings["candidate"].get("revision"):
        fail("role bindings: control and candidate revisions must differ")
    if bindings["control"].get("build_path") != bindings["candidate"].get("build_path"):
        fail("role bindings: build paths differ")
    if bindings["control"].get("included_fixtures") != bindings["candidate"].get("included_fixtures"):
        fail("role bindings: compile fixtures differ")
    control_sources = read_json(root / relative(bindings["control"]["source_manifest"], "control.source_manifest"), "control source manifest")
    candidate_sources = read_json(root / relative(bindings["candidate"]["source_manifest"], "candidate.source_manifest"), "candidate source manifest")
    if set(control_sources) != set(candidate_sources):
        fail("role source manifests: file set differs")
    changed = [name for name in sorted(control_sources) if control_sources[name] != candidate_sources[name]]
    if changed != ["crates/litchi-xlsx/src/raw/compact.rs"]:
        fail("role source manifests: expected exactly compact.rs to differ")
    checks = read_json(root / "validation/receipt.json", "correctness checks")["checks"]
    if [check["name"] for check in checks] != ["xlsx-fmt", "xlsx-tests", "workspace-check", "xlsx-clippy", "xlsx-rustdoc", "boundaries"]:
        fail("correctness check set differs")
    for check in checks:
        log = bundle_file(root, f"validation/{check['log']}", "correctness log")
        if check["exit_code"] != 0 or sha256_file(log)[0] != check["log_sha256"]:
            fail("correctness check failed or log differs")
        start, finish = (_timestamp(check[field], check["name"]) for field in ("started_utc", "finished_utc"))
        if finish < start:
            fail("correctness check interval is reversed")
        for lane in LANES:
            capture = read_json(root / lane / "receipt.json", lane)
            if start < _timestamp(capture["finished_utc"], lane) and finish > _timestamp(capture["started_utc"], lane):
                fail("correctness check overlaps a capture")
    scope = read_json(root / "validation/scope.json", "validation source scope")
    if scope.get("changed_source_sha256") != candidate_sources["crates/litchi-xlsx/src/raw/compact.rs"]:
        fail("validation source differs from candidate")
    for lane in LANES:
        verify_lane_receipt(root, lane, protocol, bindings)
    analyzer, analyzer_path = _load_analyzer(root)
    reports = {lane: read_json(root / lane / "report.json", f"{lane}/report.json") for lane in LANES}
    verify_report_metadata(
        reports["A1"], "A1.report", bindings["control"], protocol["probe_samples"], protocol["probe_warmups"],
        cases=protocol["probe_cases"], shapes=protocol["probe_shapes"],
    )
    for lane in ("B1", "B2"):
        verify_report_metadata(reports[lane], f"{lane}.report", bindings["candidate"], protocol["probe_samples"], protocol["probe_warmups"], cases=protocol["probe_cases"], shapes=protocol["probe_shapes"])
    verify_report_metadata(reports["A2"], "A2.report", bindings["control"], protocol["probe_samples"], protocol["probe_warmups"], cases=protocol["probe_cases"], shapes=protocol["probe_shapes"])
    verify_report_metadata(reports["A-full"], "A-full.report", bindings["control"], protocol["full_guard_samples"], protocol["full_guard_warmups"])
    verify_report_metadata(reports["B-full"], "B-full.report", bindings["candidate"], protocol["full_guard_samples"], protocol["full_guard_warmups"])
    heap_policy = copy.deepcopy(policy)
    heap_policy.update({"minimum_samples": protocol["heap_samples"], "expected_result_count": 1, "expected_result_keys_sha256": None, "required_cases": [protocol["heap_case"]], "require_distinct_revisions": False})
    heap_policy["expected_configuration"] = {"samples_per_case": protocol["heap_samples"], "warmup_iterations_per_case": protocol["heap_warmups"]}
    export_receipts: dict[str, dict[str, Any]] = {}
    for lane, role in zip(HEAP_LANES, ("control", "candidate")):
        verify_report_metadata(reports[lane], f"{lane}.report", bindings[role], protocol["heap_samples"], protocol["heap_warmups"], cases=[protocol["heap_case"]], shapes=[protocol["heap_shape"]])
        try:
            comparator._validate_report_identity(dict(reports[lane]), dict(reports[lane]), heap_policy)
        except (AttributeError, KeyError, TypeError, ValueError) as error:
            fail(f"{lane}.report: comparator validation failed ({error})")
        export_receipts[lane] = verify_heap_export(root, lane, bindings[role], protocol)
    verify_oracles(reports["A-heap"], reports["B-heap"], "heap", analyzer.load_abba(root)[0])
    probe_paths = [root / lane / "report.json" for lane in PROBE_LANES]
    full_paths = [root / lane / "report.json" for lane in FULL_LANES]
    heap_paths = [root / lane / "resource.log" for lane in HEAP_LANES]
    try:
        expected_summary = analyzer.build_summary(root=root, probe_paths=probe_paths, full_paths=full_paths, policy_path=root / "report-policy.json", heap_paths=heap_paths)
    except (AttributeError, KeyError, TypeError, ValueError, OSError) as error:
        fail(f"summary recomputation failed ({error})")
    summary = read_json(root / "summary.json", "summary.json")
    if not canonical_equal(summary, expected_summary):
        fail("summary.json: analyzer recomputation differs")
    if summary.get("claim", {}).get("registered") is not False or summary.get("claim", {}).get("latency_claim") != "none":
        fail("summary.json: descriptive claim policy differs")
    abba_summary = summary.get("abba")
    if not isinstance(abba_summary, dict):
        fail("summary.json: canonical ABBA summary is missing")
    verification = obj(abba_summary.get("verification"), "summary.abba.verification")
    if verification.get("case_corpus_identity_verified") is not True or verification.get("statistics_recomputed_from_samples") is not True:
        fail("summary.abba: row or statistic oracle is not verified")
    if verification.get("sink_identity_verified") is not True:
        fail("summary.abba: sink oracle is not verified")
    operation_identity = obj(verification.get("operation_metrics_identity"), "summary.abba.verification.operation_metrics_identity")
    if operation_identity.get("verified_equal") != len(abba_summary.get("results", [])):
        fail("summary.abba: source/sink operation identity is not verified")
    normal_rss = obj(summary.get("normal_process_rss"), "summary.normal_process_rss")
    if set(normal_rss) != set((*PROBE_LANES, *FULL_LANES)):
        fail("summary.normal_process_rss: lane set differs")
    for lane, value in normal_rss.items():
        rss = obj(value, f"summary.normal_process_rss.{lane}")
        if rss.get("scope") != "whole_process_lifetime_including_setup_and_teardown" or rss.get("latency_comparison") != "descriptive_only":
            fail(f"summary.normal_process_rss.{lane}: scope semantics differ")
    full_guard = obj(summary.get("full_guard"), "summary.full_guard")
    comparison = obj(full_guard.get("comparison"), "summary.full_guard.comparison")
    if comparison.get("status") not in {"pass", "regression"}:
        fail("summary.full_guard: canonical guard rejected its inputs")
    heap = obj(summary.get("heap"), "summary.heap")
    for role in ("control", "candidate"):
        row = obj(heap.get(role), f"summary.heap.{role}")
        if obj(row.get("rss"), f"summary.heap.{role}.rss").get("latency_comparison") != "excluded":
            fail(f"summary.heap.{role}: RSS latency exclusion missing")
        if obj(row.get("heaptrack"), f"summary.heap.{role}.heaptrack").get("latency_comparison") != "excluded":
            fail(f"summary.heap.{role}: heaptrack latency exclusion missing")
    chronology = verify_chronology(root, {lane: read_json(root / lane / "receipt.json", f"{lane}.receipt.json") for lane in LANES}, build_receipt, export_receipts)
    review = _load_module(root / "review.py", "litchi_change0469_guard_review")
    expected_review = review.evaluate(root)
    if not canonical_equal(read_json(root / "review-summary.json", "targeted guard summary"), expected_review):
        fail("targeted guard summary recomputation differs")
    live_state = verify_live_state(root, bindings, live=live)
    sealed = None if live else verify_sha256sums(root)
    return {
        "schema": SCHEMA,
        "change": CHANGE,
        "status": "pass",
        "mode": "live" if live else "post-cleanup",
        "lanes": list(LANES),
        "binary_sha256": {role: bindings[role]["binary_sha256"] for role in bindings},
        "sealed_files": sealed,
        "live_state": live_state,
        "summary": {"path": "summary.json", "analyzer": file_binding(analyzer_path, root)},
        "chronology": chronology,
        "targeted_guard_strict_abba_status": expected_review["strict_abba_status"],
        "dependencies": {"perf_compare": file_binding(comparator_path, root), "perf_abba_summary": file_binding(analyzer.load_abba(root)[1], root), "report_policy": file_binding(root / "report-policy.json", root)},
    }


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--live", action="store_true", help="require the authenticated build trees and binaries")
    args = parser.parse_args(sys.argv[1:] if argv is None else argv)
    try:
        result = verify(root=args.root, live=args.live)
    except (OSError, KeyError, TypeError, ValueError, VerificationError, subprocess.SubprocessError) as error:
        result = {"schema": SCHEMA, "change": CHANGE, "status": "fail", "error": str(error)}
    print(json.dumps(result, ensure_ascii=False, sort_keys=True, indent=2))
    return 0 if result.get("status") == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
