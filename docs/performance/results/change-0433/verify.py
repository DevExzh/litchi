#!/usr/bin/env python3
"""Portable, fail-closed replay for the change-0433 ODS evidence bundle.

The replay consumes only the exported JSON, source manifests, receipts, logs,
and bound Python drivers.  It never opens a captured executable or assumes a
repository checkout.  The two after roles intentionally share an ``after``
build directory but retain separate role-prefixed capture indexes and reports.
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
import importlib.util
import json
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
from typing import Any


ROOT = Path(__file__).resolve().parent
CHANGE = 433
ROLES = ("before-buffered", "after-buffered", "after-streaming")
BUILD_DIR = {"before-buffered": "before", "after-buffered": "after", "after-streaming": "after"}
MODES = ("normal", "allocator")
HEX = set("0123456789abcdef")
EXPECTED_SHAPES = {"tiny": 64, "medium": 8_192, "large": 32_768}
EXPECTED_UNTRACKED_STATUS = [
    "?? docs/GOAL.md",
    "?? docs/performance/results/change-0433/",
]
LIFECYCLE_STAGES = ("precleanup", "aftercleanup", "final")
HARNESS_SOURCE_FILE = "tools/perf-baseline/src/lib.rs"
PRODUCTION_REVIEW_SOURCE_FILE = "crates/litchi-odf-common/src/package/model.rs"
STRICT_INITIAL_LOG = "checks/harness-strict-final.log"
STRICT_CURRENT_LOG = "checks/harness-strict-final-v2.log"
PROFILE_WORKLOAD_MARKERS = frozenset(
    f"{BUILD_DIR[role]}/profiles/{role}/{kind}/workload-verify.json"
    for role in ROLES
    for kind in ("record", "stat")
)
REPLAY_DRIVER_NAMES = (
    "replay.py",
    "capture.py",
    "check.py",
    "derive.py",
    "verify.py",
    "verify-report.py",
    "seal.py",
    "protocol.json",
    "planned-checks.json",
    "profile/capture.py",
    "profile/report.py",
)

_spec = importlib.util.spec_from_file_location("ods_report_verifier", ROOT / "verify-report.py")
if _spec is None or _spec.loader is None:
    raise RuntimeError("cannot load the bound report verifier")
report_verifier = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(report_verifier)


class Invalid(ValueError):
    """The retained 0433 bundle is invalid."""


def fail(path: str, message: str) -> None:
    raise Invalid(f"{path}: {message}")


def digest(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def safe_path(root: Path, name: str) -> Path:
    if not isinstance(name, str) or not name or Path(name).is_absolute():
        fail("path", "must be a relative non-empty path")
    path = (root / name).resolve()
    if not path.is_relative_to(root.resolve()):
        fail(name, "escapes the bundle")
    return path


def logical_bytes(root: Path, name: str) -> bytes:
    path = safe_path(root, name)
    if path.is_file():
        return path.read_bytes()
    compressed = safe_path(root, name + ".gz")
    if compressed.is_file():
        try:
            return gzip.decompress(compressed.read_bytes())
        except (OSError, EOFError, gzip.BadGzipFile) as error:
            fail(name + ".gz", f"invalid gzip stream: {error}")
    fail(name, "file is missing in raw and gzip form")
    raise AssertionError("unreachable")


def load_bytes(raw: bytes, label: str) -> Any:
    try:
        return json.loads(
            raw.decode("utf-8"),
            object_pairs_hook=report_verifier.reject_duplicate_pairs,
            parse_constant=report_verifier.reject_constant,
        )
    except (UnicodeDecodeError, json.JSONDecodeError, report_verifier.VerificationError) as error:
        fail(label, f"invalid JSON: {error}")
    raise AssertionError("unreachable")


def load(root: Path, name: str) -> Any:
    return load_bytes(logical_bytes(root, name), name)


def obj(value: Any, path: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(path, "expected an object")
    return value


def array(value: Any, path: str) -> list[Any]:
    if not isinstance(value, list):
        fail(path, "expected an array")
    return value


def text(value: Any, path: str) -> str:
    if not isinstance(value, str) or not value:
        fail(path, "expected a non-empty string")
    return value


def uint(value: Any, path: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0 or value > report_verifier.U64_MAX:
        fail(path, "expected a u64")
    return value


def sha_text(value: Any, path: str) -> str:
    value = text(value, path).lower()
    if len(value) != 64 or any(char not in HEX for char in value):
        fail(path, "expected a lowercase SHA-256 digest")
    return value


def verify_protocol(root: Path) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    protocol = obj(load(root, "protocol.json"), "protocol")
    if uint(protocol.get("change"), "protocol.change") != CHANGE:
        fail("protocol.change", f"must be {CHANGE}")
    shapes = obj(protocol.get("shapes"), "protocol.shapes")
    if shapes != EXPECTED_SHAPES:
        fail("protocol.shapes", f"must be exactly {EXPECTED_SHAPES!r}")
    if uint(protocol.get("cells_per_row"), "protocol.cells_per_row") != 4:
        fail("protocol.cells_per_row", "must be four")
    if uint(protocol.get("cpu"), "protocol.cpu") != 2:
        fail("protocol.cpu", "must be two")
    samples = uint(protocol.get("samples"), "protocol.samples")
    warmups = uint(protocol.get("warmups"), "protocol.warmups")
    workers = uint(protocol.get("workers"), "protocol.workers")
    if samples != 30 or warmups != 3 or workers != 1:
        fail("protocol", "must retain samples=30, warmups=3, workers=1")
    if uint(protocol.get("repeats"), "protocol.repeats") != 2:
        fail("protocol.repeats", "must retain two repeats")
    modes = array(protocol.get("modes"), "protocol.modes")
    if modes != list(MODES):
        fail("protocol.modes", "must be [normal, allocator]")
    roles = obj(protocol.get("roles"), "protocol.roles")
    if set(roles) != set(ROLES):
        fail("protocol.roles", "must contain exactly the three frozen roles")
    for role in ROLES:
        spec = obj(roles[role], f"protocol.roles.{role}")
        selector = text(spec.get("selector"), f"protocol.roles.{role}.selector")
        expected_selector = "ods_streaming_create" if role == "after-streaming" else "ods_buffered_create"
        if selector != expected_selector:
            fail(f"protocol.roles.{role}.selector", f"must be {expected_selector!r}")
        retention = text(spec.get("retention"), f"protocol.roles.{role}.retention")
        if retention not in {"fixed", "bounded", "nonfixed", "non_fixed", "not-fixed", "unbounded"}:
            fail(f"protocol.roles.{role}.retention", "must identify fixed or nonfixed retention")
        if role == "after-streaming" and retention not in {"fixed", "bounded"}:
            fail(f"protocol.roles.{role}.retention", "streaming role must be fixed")
        if role != "after-streaming" and retention in {"fixed", "bounded"}:
            fail(f"protocol.roles.{role}.retention", "buffered role must be nonfixed")
        window = spec.get("retained_authoring_window_bytes")
        if role == "after-streaming" and window != 4_096:
            fail(f"protocol.roles.{role}.retained_authoring_window_bytes", "must be exactly 4096")
        if role != "after-streaming" and window is not None:
            fail(f"protocol.roles.{role}.retained_authoring_window_bytes", "buffered role must not advertise a fixed window")
        oracle = obj(spec.get("oracle"), f"protocol.roles.{role}.oracle")
        expected_generator = "litchi-ods-streaming-scalar-rows-v1" if role == "after-streaming" else "litchi-ods-buffered-scalar-rows-v1"
        if oracle.get("generator") != expected_generator:
            fail(f"protocol.roles.{role}.oracle.generator", f"must be {expected_generator!r}")
        expected_prefix = "ods-streaming-scalar-rows-" if role == "after-streaming" else "ods-buffered-scalar-rows-"
        corpus_names = obj(oracle.get("corpus_names"), f"protocol.roles.{role}.oracle.corpus_names")
        if corpus_names != {shape: f"{expected_prefix}{shape}" for shape in EXPECTED_SHAPES}:
            fail(f"protocol.roles.{role}.oracle.corpus_names", "does not pin all three role-specific corpus names")
        if oracle.get("sheet_name") != "Sheet1":
            fail(f"protocol.roles.{role}.oracle.sheet_name", "must be Sheet1")
    order = array(protocol.get("order"), "protocol.order")
    if len(order) != 12:
        fail("protocol.order", "must contain 12 lanes")
    expected_r1 = [
        {"mode": mode, "shape": shape, "repeat": "r1"}
        for mode in MODES
        for shape in ("tiny", "medium", "large")
    ]
    expected = expected_r1 + [dict(lane, repeat="r2") for lane in reversed(expected_r1)]
    # The protocol owns order; validate every lane and uniqueness while not
    # imposing a timing order beyond the captured protocol's declaration.
    normalized = []
    for index, lane_value in enumerate(order):
        lane = obj(lane_value, f"protocol.order[{index}]")
        if set(lane) != {"mode", "shape", "repeat"}:
            fail(f"protocol.order[{index}]", "lane fields differ")
        if lane["mode"] not in MODES or lane["shape"] not in shapes or str(lane["repeat"]).lower() not in {"r1", "r2"}:
            fail(f"protocol.order[{index}]", "lane value is outside the frozen matrix")
        normalized.append({"mode": lane["mode"], "shape": lane["shape"], "repeat": lane["repeat"]})
    normalized_keys = [(row["mode"], row["shape"], str(row["repeat"]).lower()) for row in normalized]
    expected_keys = [(row["mode"], row["shape"], row["repeat"]) for row in expected]
    if normalized_keys != expected_keys:
        fail("protocol.order", "must run forward R1 lanes followed by reverse R2 lanes")
    return protocol, normalized


def verify_source_manifest(root: Path, manifest_value: Any, path: str, *, require_critical: bool = True) -> dict[str, str]:
    manifest = obj(manifest_value, path)
    if set(manifest) != {"path", "sha256", "files"}:
        fail(path, "fields differ")
    source_path = text(manifest["path"], f"{path}.path")
    source_hash = sha_text(manifest["sha256"], f"{path}.sha256")
    files = uint(manifest["files"], f"{path}.files")
    if files == 0:
        fail(path, "source manifest must contain at least one file")
    raw = logical_bytes(root, source_path)
    if digest(raw) != source_hash:
        fail(path, "retained source manifest hash differs")
    entries = obj(load_bytes(raw, f"{path}.payload"), f"{path}.payload")
    if len(entries) != files:
        fail(path, "source file count differs")
    for name, value in entries.items():
        text(name, f"{path}.payload.path")
        sha_text(value, f"{path}.payload.{name}")
    if require_critical:
        critical = {
            "Cargo.lock",
            "tools/perf-baseline/Cargo.lock",
            "tools/perf-baseline/Cargo.toml",
            "tools/perf-baseline/src/lib.rs",
            "tools/perf-baseline/src/operation_metrics.rs",
            "tools/perf-baseline/src/allocation_metrics.rs",
        }
        missing = sorted(critical - set(entries))
        if missing:
            fail(path, f"critical source files are missing: {missing}")
        empty_digest = hashlib.sha256(b"").hexdigest()
        empty = sorted(name for name in critical if entries[name].lower() == empty_digest)
        if empty:
            fail(path, f"critical source files have empty-content digests: {empty}")
    return entries


def verify_build(root: Path, protocol: dict[str, Any], role: str) -> dict[str, Any]:
    build_path = f"{BUILD_DIR[role]}/build.json"
    build = obj(load(root, build_path), build_path)
    required = {"revision", "source_manifest", "protocol_sha256", "verifier_sha256", "binaries"}
    if not required.issubset(build):
        fail(build_path, f"missing required fields: {sorted(required - set(build))}")
    if "role" in build and build["role"] not in {BUILD_DIR[role], role}:
        fail(f"{build_path}.role", "does not identify the build role")
    revision = text(build["revision"], f"{build_path}.revision").lower()
    if len(revision) != 40 or any(char not in HEX for char in revision):
        fail(f"{build_path}.revision", "must be a hexadecimal revision")
    if sha_text(build["protocol_sha256"], f"{build_path}.protocol_sha256") != digest((root / "protocol.json").read_bytes()):
        fail(f"{build_path}.protocol_sha256", "protocol custody mismatch")
    verifier_hash = sha_text(build["verifier_sha256"], f"{build_path}.verifier_sha256")
    if verifier_hash != digest((root / "verify-report.py").read_bytes()):
        fail(f"{build_path}.verifier_sha256", "bound verifier mismatch")
    planned_path = root / "planned-checks.json"
    if planned_path.is_file():
        planned_hash = build.get("planned_checks_sha256")
        if planned_hash is None or sha_text(planned_hash, f"{build_path}.planned_checks_sha256") != digest(planned_path.read_bytes()):
            fail(f"{build_path}.planned_checks_sha256", "does not bind planned-checks.json")
    source_entries = verify_source_manifest(root, build["source_manifest"], f"{build_path}.source_manifest")
    binaries = obj(build["binaries"], f"{build_path}.binaries")
    if set(binaries) != set(MODES):
        fail(f"{build_path}.binaries", "must contain normal and allocator binaries")
    checked_binaries = {}
    for mode in MODES:
        binary = obj(binaries[mode], f"{build_path}.binaries.{mode}")
        if set(binary) != {"path", "sha256", "bytes"}:
            fail(f"{build_path}.binaries.{mode}", "fields differ")
        checked_binaries[mode] = {
            "path": text(binary["path"], f"{build_path}.binaries.{mode}.path"),
            "sha256": sha_text(binary["sha256"], f"{build_path}.binaries.{mode}.sha256"),
            "bytes": uint(binary["bytes"], f"{build_path}.binaries.{mode}.bytes"),
        }
        if not Path(checked_binaries[mode]["path"]).is_absolute():
            fail(f"{build_path}.binaries.{mode}.path", "must be an absolute captured binary path")
        if checked_binaries[mode]["bytes"] == 0:
            fail(f"{build_path}.binaries.{mode}.bytes", "binary is empty")
    return {"role": role, "revision": revision, "source_manifest": build["source_manifest"], "source_entries": source_entries, "binaries": checked_binaries, "verifier_sha256": verifier_hash, "raw": build}


def verify_capture_state(root: Path, role: str, build: dict[str, Any]) -> None:
    """Bind the role capture to the documented dirty-worktree custody."""

    name = f"{BUILD_DIR[role]}/capture-state-{role}.json"
    state = obj(load(root, name), name)
    required = {
        "status",
        "role",
        "status_before",
        "status_after",
        "tracked_tree_clean_before_and_after",
        "report_dirty_field_expected",
        "source_manifest",
    }
    optional = {"source_manifest_before", "source_manifest_after"}
    if not required.issubset(state) or set(state) - required - optional:
        fail(name, "capture-state fields differ")
    if state["status"] != "pass" or state["role"] != role:
        fail(name, "capture-state does not identify a passing role capture")
    for key in ("status_before", "status_after"):
        statuses = array(state[key], f"{name}.{key}")
        if statuses != EXPECTED_UNTRACKED_STATUS:
            fail(f"{name}.{key}", "must contain exactly GOAL.md and the change bundle as untracked paths")
    if state["tracked_tree_clean_before_and_after"] is not True or state["report_dirty_field_expected"] is not True:
        fail(name, "tracked-tree or report dirty-worktree custody flag is false")
    if state["source_manifest"] != build["source_manifest"]:
        fail(f"{name}.source_manifest", "does not match the role build source custody")
    for key in optional:
        if key in state and state[key] != build["source_manifest"]:
            fail(f"{name}.{key}", "does not match the role build source custody")


def artifact_bytes(root: Path, name: str, metadata: Any, path: str) -> bytes:
    row = obj(metadata, path)
    if set(row) != {"sha256", "bytes"}:
        fail(path, "artifact fields differ")
    raw = logical_bytes(root, name)
    if digest(raw) != sha_text(row["sha256"], f"{path}.sha256"):
        fail(path, "artifact hash differs")
    if len(raw) != uint(row["bytes"], f"{path}.bytes"):
        fail(path, "artifact byte count differs")
    return raw


def invoke_report_verifier(root: Path, role: str, lane: dict[str, Any], logical_report: str) -> None:
    verifier = safe_path(root, "verify-report.py")
    raw = logical_bytes(root, logical_report)
    path = safe_path(root, logical_report)
    catalog_logical = logical_report[:-5] + "-catalog.json"
    catalog_raw = logical_bytes(root, catalog_logical)
    temporary_directory = None
    catalog_path = safe_path(root, catalog_logical)
    if not path.is_file() or not catalog_path.is_file():
        temporary_directory = tempfile.TemporaryDirectory(prefix="litchi-0433-report-")
        directory = Path(temporary_directory.name)
        path = directory / Path(logical_report).name
        path.write_bytes(raw)
        (directory / (path.stem + "-catalog.json")).write_bytes(catalog_raw)
    try:
        command = [sys.executable, "-B", str(verifier), "--report", str(path), "--mode", lane["mode"], "--shape", lane["shape"], "--role", role]
        result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if result.returncode != 0:
            detail = (result.stderr or result.stdout).decode("utf-8", "replace").strip()
            fail(logical_report, f"bound report verifier rejected it: {detail or result.returncode}")
    finally:
        if temporary_directory is not None:
            temporary_directory.cleanup()


def expected_name(role: str, lane: dict[str, Any]) -> str:
    return f"{role}-{lane['mode']}-{lane['shape']}-{str(lane['repeat']).lower()}"


def verify_argv(argv: Any, role: str, lane: dict[str, Any], protocol: dict[str, Any], binary: dict[str, Any], selector: str, name: str, path: str) -> None:
    values = array(argv, f"{path}.argv")
    if not all(isinstance(value, str) for value in values):
        fail(f"{path}.argv", "all arguments must be strings")
    expected_prefix = ["taskset", "-c", str(protocol["cpu"]), "/usr/bin/time", "-v", "-o"]
    if values[:6] != expected_prefix or len(values) != 22:
        fail(f"{path}.argv", "does not bind taskset and GNU time -v")
    if Path(values[6]).name != f"{name}-resource.log":
        fail(f"{path}.argv", "resource output path does not bind the lane")
    if values[7] != binary["path"]:
        fail(f"{path}.argv", "does not bind the build binary")
    tail = values[8:]
    expected = [
        "--case", selector, "--semantic-shape", lane["shape"], "--workers", str(protocol["workers"]),
        "--samples", str(protocol["samples"]), "--warmup", str(protocol["warmups"]), "--json",
    ]
    if len(tail) != 14 or tail[:11] != expected:
        fail(f"{path}.argv", "does not bind the lane configuration")
    if tail[12] != "--corpus-manifest":
        fail(f"{path}.argv", "does not request the catalog sidecar")
    if Path(tail[11]).name != Path(tail[13]).name.replace("-catalog", ""):
        fail(f"{path}.argv", "report/catalog paths do not bind the same lane")


def parse_rss(root: Path, logical: str) -> None:
    raw = logical_bytes(root, logical)
    lines = raw.decode("utf-8").splitlines()
    matches = []
    for line in lines:
        line = line.lstrip()
        prefix = "Maximum resident set size (kbytes):"
        if line.startswith(prefix):
            value = line[len(prefix):].strip()
            if not value.isdigit():
                fail(logical, "GNU time RSS value is not an integer")
            matches.append(int(value))
    if len(matches) != 1:
        fail(logical, "must contain exactly one GNU time maximum-resident-set-size line")


def semantic_key(report: dict[str, Any], path: str) -> tuple[Any, ...]:
    result = obj(array(report.get("results"), f"{path}.results")[0], f"{path}.results[0]")
    source = obj(result.get("source"), f"{path}.results[0].source")
    oracle = obj(source.get("ods_scalar_rows"), f"{path}.results[0].source.ods_scalar_rows")
    return (
        uint(oracle.get("sheet_count"), f"{path}.ods_scalar_rows.sheet_count"),
        uint(oracle.get("rows_per_sheet"), f"{path}.ods_scalar_rows.rows_per_sheet"),
        uint(oracle.get("columns_per_sheet"), f"{path}.ods_scalar_rows.columns_per_sheet"),
        sha_text(oracle.get("semantic_sha256"), f"{path}.ods_scalar_rows.semantic_sha256"),
        tuple(array(oracle.get("scalar_columns"), f"{path}.ods_scalar_rows.scalar_columns")),
    )


def verify_capture_role(root: Path, role: str, build: dict[str, Any], protocol: dict[str, Any], order: list[dict[str, Any]]) -> list[dict[str, Any]]:
    index_candidates = [f"{BUILD_DIR[role]}/capture-index-{role}.json", f"{BUILD_DIR[role]}/capture-index.json"]
    index_name = next((name for name in index_candidates if safe_path(root, name).is_file() or safe_path(root, name + ".gz").is_file()), None)
    if index_name is None:
        fail(f"{role}.capture-index", "is missing")
    index = array(load(root, index_name), index_name)
    if len(index) != len(order) or len(set(index)) != len(order):
        fail(index_name, "does not list each frozen lane exactly once")
    rows = []
    same_identity: dict[tuple[str, str, str], tuple[str, str]] = {}
    same_output: dict[tuple[str], tuple[str, str]] = {}
    spec = obj(obj(protocol["roles"], "protocol.roles")[role], f"protocol.roles.{role}")
    for position, lane in enumerate(order):
        name = expected_name(role, lane)
        expected_receipt = f"{BUILD_DIR[role]}/captures/{name}-receipt.json"
        if index[position] != f"captures/{name}-receipt.json" and index[position] != expected_receipt:
            fail(f"{index_name}[{position}]", "receipt order/path differs from the frozen role")
        receipt_name = index[position] if str(index[position]).startswith(f"{BUILD_DIR[role]}/") else expected_receipt
        receipt = obj(load(root, receipt_name), receipt_name)
        required = {"change", "role", "selector", "name", "lane", "argv", "revision", "source_manifest", "binary", "protocol_sha256", "driver_sha256", "verifier_sha256", "status", "exit_code", "artifacts"}
        if not required.issubset(receipt):
            fail(receipt_name, f"missing fields: {sorted(required - set(receipt))}")
        if uint(receipt["change"], f"{receipt_name}.change") != CHANGE or receipt["role"] != role or receipt["name"] != name or receipt["lane"] != lane:
            fail(receipt_name, "does not bind the frozen role/lane")
        if receipt["selector"] != spec["selector"] or receipt["revision"] != build["revision"] or receipt["source_manifest"] != build["source_manifest"]:
            fail(receipt_name, "source/selector identity differs from the role build")
        if receipt["status"] != "pass" or receipt["exit_code"] != 0:
            fail(receipt_name, "capture did not pass")
        binary = build["binaries"][lane["mode"]]
        if obj(receipt["binary"], f"{receipt_name}.binary") != binary:
            fail(receipt_name, "binary identity differs from the role build")
        if sha_text(receipt["protocol_sha256"], f"{receipt_name}.protocol_sha256") != digest((root / "protocol.json").read_bytes()):
            fail(receipt_name, "protocol hash differs")
        if sha_text(receipt["verifier_sha256"], f"{receipt_name}.verifier_sha256") != digest((root / "verify-report.py").read_bytes()):
            fail(receipt_name, "report verifier hash differs")
        if sha_text(receipt["driver_sha256"], f"{receipt_name}.driver_sha256") != digest((root / "capture.py").read_bytes()):
            fail(receipt_name, "capture driver hash differs")
        verify_argv(receipt["argv"], role, lane, protocol, binary, spec["selector"], name, receipt_name)
        artifacts = obj(receipt["artifacts"], f"{receipt_name}.artifacts")
        expected_artifacts = {
            f"{BUILD_DIR[role]}/captures/{name}.json",
            f"{BUILD_DIR[role]}/captures/{name}-catalog.json",
            f"{BUILD_DIR[role]}/captures/{name}.log",
            f"{BUILD_DIR[role]}/captures/{name}-resource.log",
        }
        if set(artifacts) != expected_artifacts:
            fail(receipt_name, "artifact inventory differs")
        for artifact_name, metadata in artifacts.items():
            artifact_bytes(root, artifact_name, metadata, f"{receipt_name}.artifacts.{artifact_name}")
        parse_rss(root, f"{BUILD_DIR[role]}/captures/{name}-resource.log")
        logical_report = f"{BUILD_DIR[role]}/captures/{name}.json"
        invoke_report_verifier(root, role, lane, logical_report)
        report = obj(load(root, logical_report), logical_report)
        result = obj(array(report.get("results"), f"{logical_report}.results")[0], f"{logical_report}.results[0]")
        if report.get("environment", {}).get("git_revision") != build["revision"]:
            fail(logical_report, "report revision differs from the role build")
        if report.get("environment", {}).get("git_worktree_dirty") is not True:
            fail(logical_report, "report must retain the documented dirty-worktree boolean")
        identity = obj(report.get("binary_identity"), f"{logical_report}.binary_identity")
        if identity.get("path") != binary["path"] or identity.get("binary_sha256") != binary["sha256"] or identity.get("binary_bytes") != binary["bytes"]:
            fail(logical_report, "report binary identity differs from the role build")
        key = (lane["mode"], lane["shape"], str(lane["repeat"]).lower())
        output_key = (lane["shape"],)
        semantic = semantic_key(report, logical_report)
        output = (result.get("output_sha256"), obj(result.get("corpus"), f"{logical_report}.corpus").get("archive_sha256"))
        if key in same_identity and same_identity[key] != semantic:
            fail(logical_report, "semantic oracle differs from the other role report")
        same_identity[key] = semantic
        if output_key in same_output and same_output[output_key] != output:
            fail(logical_report, "same role/mode/shape output identity changed between repeats")
        same_output[output_key] = output
        rows.append({"lane": lane, "name": name, "report": report, "semantic": semantic, "output": output})
    return rows


def verify_cross_role(rows: dict[str, list[dict[str, Any]]]) -> None:
    by_role = {
        role: {(row["lane"]["mode"], row["lane"]["shape"], str(row["lane"]["repeat"]).lower()): row for row in values}
        for role, values in rows.items()
    }
    for mode in MODES:
        for shape in ("tiny", "medium", "large"):
            for repeat in ("r1", "r2"):
                key = (mode, shape, repeat)
                values = [by_role[role][key]["semantic"] for role in ROLES]
                if any(value != values[0] for value in values[1:]):
                    fail(f"cross-role.{mode}.{shape}.{repeat}", "semantic row/cell oracle differs")


def verify_strict_receipt(
    root: Path,
    comparison: dict[str, Any],
    *,
    log_key: str,
    hash_key: str,
    bytes_key: str,
    expected_log: str,
    before_revision: str,
) -> tuple[Any, ...]:
    """Bind each reviewed strict command receipt to its comparison log.

    The first strict invocation is retained as a failed development attempt;
    the v2 invocation is the one compared with the 0429 baseline.  Keeping
    both explicit here prevents an old initial log from silently becoming the
    comparison input while preserving its diagnostic evidence.
    """

    name = f"checks/strict-debt-comparison.json"
    log_name = text(comparison.get(log_key), f"{name}.{log_key}")
    if log_name != expected_log:
        fail(f"{name}.{log_key}", f"must bind {expected_log}")
    receipt_name = str(Path(log_name).with_suffix(".json"))
    receipt = obj(load(root, receipt_name), receipt_name)
    required = {
        "change", "argv", "cwd", "revision", "driver_sha256", "source_before",
        "source_after", "source_unchanged", "status", "exit_code", "log",
    }
    missing = sorted(required - set(receipt))
    if missing:
        fail(receipt_name, f"reviewed strict receipt is missing mandatory fields: {missing}")
    if uint(receipt["change"], f"{receipt_name}.change") != CHANGE:
        fail(f"{receipt_name}.change", f"must be {CHANGE}")
    status = receipt.get("status")
    if status not in {"failed", "review"}:
        fail(receipt_name, "strict debt must remain a failed/reviewed command receipt")
    exit_code = receipt["exit_code"]
    if isinstance(exit_code, bool) or not isinstance(exit_code, int) or exit_code == 0:
        fail(f"{receipt_name}.exit_code", "strict debt receipt must retain a nonzero command result")
    argv = array(receipt["argv"], f"{receipt_name}.argv")
    if not argv or any(not isinstance(value, str) or not value for value in argv):
        fail(f"{receipt_name}.argv", "must contain non-empty command arguments")
    cwd = text(receipt["cwd"], f"{receipt_name}.cwd")
    revision = text(receipt["revision"], f"{receipt_name}.revision").lower()
    if revision != before_revision:
        fail(f"{receipt_name}.revision", "does not match the before build revision")
    driver = sha_text(receipt["driver_sha256"], f"{receipt_name}.driver_sha256")
    if receipt["source_before"] != receipt["source_after"]:
        fail(receipt_name, "strict command changed source custody")
    if receipt["source_unchanged"] is not True:
        fail(f"{receipt_name}.source_unchanged", "must be true")
    log = obj(receipt["log"], f"{receipt_name}.log")
    receipt_log = text(log.get("path"), f"{receipt_name}.log.path")
    if receipt_log != log_name:
        fail(f"{receipt_name}.log.path", "does not match the comparison log")
    raw = logical_bytes(root, log_name)
    if digest(raw) != sha_text(log.get("sha256"), f"{receipt_name}.log.sha256"):
        fail(f"{receipt_name}.log.sha256", "retained command log digest differs")
    if len(raw) != uint(log.get("bytes"), f"{receipt_name}.log.bytes"):
        fail(f"{receipt_name}.log.bytes", "retained command log byte count differs")
    if digest(raw) != sha_text(comparison.get(hash_key), f"checks/strict-debt-comparison.json.{hash_key}"):
        fail(f"checks/strict-debt-comparison.json.{hash_key}", "comparison log digest differs")
    if len(raw) != uint(comparison.get(bytes_key), f"checks/strict-debt-comparison.json.{bytes_key}"):
        fail(f"checks/strict-debt-comparison.json.{bytes_key}", "comparison log byte count differs")
    source_scope = receipt.get("source_scope", "")
    if not isinstance(source_scope, str):
        fail(f"{receipt_name}.source_scope", "must be text when present")
    return tuple(argv), cwd, driver, source_scope


def verify_strict_debt(root: Path, builds: dict[str, dict[str, Any]]) -> None:
    """Recompute strict-debt evidence and bind its source hashes to builds."""

    run_checked(
        [sys.executable, "-B", str(safe_path(root, "compare-strict.py")), "--check"],
        "strict-debt comparison replay",
    )
    name = "checks/strict-debt-comparison.json"
    comparison = obj(load(root, name), name)
    if comparison.get("status") != "pass" or uint(comparison.get("change"), f"{name}.change") != CHANGE:
        fail(name, "strict-debt comparison did not pass for change 0433")
    if comparison.get("same_message_and_source_file_multiset") is not True:
        fail(name, "strict diagnostics differ from the retained baseline")
    if comparison.get("changed_harness_src_lib_findings") != 0:
        fail(name, "strict diagnostics landed in changed harness source")
    if array(comparison.get("changed_line_findings"), f"{name}.changed_line_findings"):
        fail(name, "strict diagnostics landed in changed source lines")
    findings = array(comparison.get("findings"), f"{name}.findings")
    if uint(comparison.get("unique_findings"), f"{name}.unique_findings") != len(findings):
        fail(name, "strict finding count does not match retained findings")

    binding = obj(comparison.get("source_binding"), f"{name}.source_binding")
    if binding.get("file") != HARNESS_SOURCE_FILE:
        fail(f"{name}.source_binding.file", f"must bind {HARNESS_SOURCE_FILE}")
    baseline_revision = text(binding.get("baseline_revision"), f"{name}.source_binding.baseline_revision").lower()
    current_head_revision = text(binding.get("current_head_revision"), f"{name}.source_binding.current_head_revision").lower()
    if len(baseline_revision) != 40 or any(char not in HEX for char in baseline_revision):
        fail(f"{name}.source_binding.baseline_revision", "must be a hexadecimal revision")
    if len(current_head_revision) != 40 or any(char not in HEX for char in current_head_revision):
        fail(f"{name}.source_binding.current_head_revision", "must be a hexadecimal revision")
    if baseline_revision != builds["before-buffered"]["revision"]:
        fail(f"{name}.source_binding.baseline_revision", "does not match the before build")
    baseline_hash = sha_text(binding.get("baseline_sha256"), f"{name}.source_binding.baseline_sha256")
    current_hash = sha_text(binding.get("current_worktree_sha256"), f"{name}.source_binding.current_worktree_sha256")
    baseline_source_hash = builds["before-buffered"]["source_entries"].get(HARNESS_SOURCE_FILE)
    current_source_hash = builds["after-buffered"]["source_entries"].get(HARNESS_SOURCE_FILE)
    if baseline_source_hash is None or baseline_hash != baseline_source_hash:
        fail(f"{name}.source_binding.baseline_sha256", "does not match the before source manifest")
    if current_source_hash is None or current_hash != current_source_hash:
        fail(f"{name}.source_binding.current_worktree_sha256", "does not match the after source manifest")

    if text(comparison.get("changed_source_patch"), f"{name}.changed_source_patch") != "checks/strict-changed-source.patch":
        fail(f"{name}.changed_source_patch", "does not bind the retained changed-source patch")
    if text(comparison.get("source_binding_path"), f"{name}.source_binding_path") != "checks/strict-source-binding.json":
        fail(f"{name}.source_binding_path", "does not bind the retained source snapshot")

    for path_key, hash_key, bytes_key in (
        ("baseline_path", "baseline_original_sha256", "baseline_original_bytes"),
    ):
        artifact_name = text(comparison.get(path_key), f"{name}.{path_key}")
        raw = logical_bytes(root, artifact_name)
        if digest(raw) != sha_text(comparison.get(hash_key), f"{name}.{hash_key}"):
            fail(f"{name}.{hash_key}", "retained strict log digest differs")
        if len(raw) != uint(comparison.get(bytes_key), f"{name}.{bytes_key}"):
            fail(f"{name}.{bytes_key}", "retained strict log byte count differs")

    initial_identity = verify_strict_receipt(
        root,
        comparison,
        log_key="initial_attempt_log",
        hash_key="initial_attempt_original_sha256",
        bytes_key="initial_attempt_original_bytes",
        expected_log=STRICT_INITIAL_LOG,
        before_revision=builds["before-buffered"]["revision"],
    )
    current_identity = verify_strict_receipt(
        root,
        comparison,
        log_key="current_log",
        hash_key="current_original_sha256",
        bytes_key="current_original_bytes",
        expected_log=STRICT_CURRENT_LOG,
        before_revision=builds["before-buffered"]["revision"],
    )
    if current_identity != initial_identity:
        fail(name, "initial and v2 strict receipts do not retain an identical command identity")


def verify_production_strict_review(root: Path, builds: dict[str, dict[str, Any]]) -> None:
    """Replay and bind the explicit production pre-existing-debt review."""

    run_checked(
        [sys.executable, "-B", str(safe_path(root, "production-strict-review.py")), "--check"],
        "production strict review replay",
    )
    name = "checks/production-strict-review.json"
    review = obj(load(root, name), name)
    if review.get("status") != "pass" or uint(review.get("change"), f"{name}.change") != CHANGE:
        fail(name, "production strict review did not pass")
    if review.get("same_message_and_source_file_multiset") is not True or review.get("source_sha256_unchanged") is not True:
        fail(name, "production strict review does not prove retained pre-existing debt")
    source_file = text(review.get("source_file"), f"{name}.source_file")
    if source_file != PRODUCTION_REVIEW_SOURCE_FILE:
        fail(f"{name}.source_file", f"must bind {PRODUCTION_REVIEW_SOURCE_FILE}")
    known = obj(review.get("known_preexisting_diagnostic"), f"{name}.known_preexisting_diagnostic")
    if (
        known.get("message") != "large size difference between variants"
        or known.get("file") != PRODUCTION_REVIEW_SOURCE_FILE
        or uint(known.get("count"), f"{name}.known_preexisting_diagnostic.count") != 1
    ):
        fail(name, "known production pre-existing diagnostic differs")
    findings = array(review.get("findings"), f"{name}.findings")
    if len(findings) != 1:
        fail(f"{name}.findings", "must retain exactly one known diagnostic")
    finding = obj(findings[0], f"{name}.findings[0]")
    if (
        finding.get("message") != known["message"]
        or finding.get("file") != known["file"]
        or uint(finding.get("count"), f"{name}.findings[0].count") != 1
    ):
        fail(f"{name}.findings[0]", "does not match the known pre-existing diagnostic")

    for side, role, record_name in (
        ("baseline", "before-buffered", "checks/ods-reference-strict.json"),
        ("current", "after-buffered", "checks/production-strict-final.json"),
    ):
        binding = obj(review.get(side), f"{name}.{side}")
        if binding.get("record_path") != record_name:
            fail(f"{name}.{side}.record_path", "does not bind the retained strict command receipt")
        side_source = text(binding.get("source_sha256"), f"{name}.{side}.source_sha256")
        expected_source = builds[role]["source_entries"].get(PRODUCTION_REVIEW_SOURCE_FILE)
        if expected_source is None or side_source != expected_source:
            fail(f"{name}.{side}.source_sha256", "does not match the role build source manifest")
        for source_side in ("source_before", "source_after"):
            source = obj(binding.get(source_side), f"{name}.{side}.{source_side}")
            if source.get("source_file") != PRODUCTION_REVIEW_SOURCE_FILE or source.get("source_sha256") != expected_source:
                fail(f"{name}.{side}.{source_side}", "does not bind the role build source manifest")


def run_checked(command: list[str], label: str) -> None:
    result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if result.returncode != 0:
        detail = (result.stderr or result.stdout).decode("utf-8", "replace").strip()
        fail(label, f"bound command failed: {detail or result.returncode}")


def verify_replay_receipt(root: Path, name: str, row: dict[str, Any]) -> None:
    """Validate an external replay receipt, which intentionally has no argv."""

    required = {
        "change", "status", "command", "cwd", "portable_flag", "require_inventory",
        "stage", "exit_code", "started_utc", "finished_utc", "driver_hashes",
        "drivers_unchanged", "replayed_inventory_sha256", "inventory_unchanged",
        "inventory_rewritten", "log", "scope",
    }
    missing = sorted(required - set(row))
    if missing:
        fail(name, f"replay receipt is missing mandatory fields: {missing}")
    if uint(row["change"], f"{name}.change") != CHANGE:
        fail(f"{name}.change", f"must be {CHANGE}")
    if row["status"] not in {"pass", "failed"}:
        fail(f"{name}.status", "portable replay receipt has an invalid terminal status")
    command = array(row["command"], f"{name}.command")
    if len(command) != 7 or any(not isinstance(value, str) or not value for value in command):
        fail(f"{name}.command", "does not retain the seven-argument portable replay command")
    if command[1] != "-B" or Path(command[2]).name != "verify.py":
        fail(f"{name}.command", "does not invoke the bound verify.py driver")
    if command[3:5] != ["--portable-check", "--require-inventory"]:
        fail(f"{name}.command", "does not require portable replay with sealed inventory")
    if command[5] != "--stage" or command[6] not in LIFECYCLE_STAGES:
        fail(f"{name}.command", "does not bind a known replay lifecycle stage")
    text(row["cwd"], f"{name}.cwd")
    if row["portable_flag"] != "--portable-check" or row["require_inventory"] is not True:
        fail(name, "replay flags are not bound to portable sealed verification")
    if row["stage"] != command[6]:
        fail(f"{name}.stage", "does not match the replay command")
    text(row["started_utc"], f"{name}.started_utc")
    text(row["finished_utc"], f"{name}.finished_utc")
    exit_code = row["exit_code"]
    if isinstance(exit_code, bool) or not isinstance(exit_code, int):
        fail(f"{name}.exit_code", "must be an integer")
    for field in ("drivers_unchanged", "inventory_unchanged", "inventory_rewritten"):
        if not isinstance(row[field], bool):
            fail(f"{name}.{field}", "must be a boolean")
    if row["status"] == "pass":
        if exit_code != 0 or row["drivers_unchanged"] is not True or row["inventory_unchanged"] is not True:
            fail(name, "passing replay receipt does not prove successful custody")
        if row["inventory_rewritten"] is not False:
            fail(f"{name}.inventory_rewritten", "passing replay must not rewrite the sealed inventory")
    elif exit_code == 0 and row["drivers_unchanged"] is True and row["inventory_unchanged"] is True:
        fail(name, "failed replay receipt has no recorded failure condition")
    sha_text(row["replayed_inventory_sha256"], f"{name}.replayed_inventory_sha256")
    driver_hashes = obj(row["driver_hashes"], f"{name}.driver_hashes")
    if set(driver_hashes) != set(REPLAY_DRIVER_NAMES):
        fail(f"{name}.driver_hashes", "driver set differs from replay.py")
    for driver_name in REPLAY_DRIVER_NAMES:
        expected = sha_text(driver_hashes[driver_name], f"{name}.driver_hashes.{driver_name}")
        if row["status"] == "pass":
            driver_path = safe_path(root, driver_name)
            if digest(driver_path.read_bytes()) != expected:
                fail(f"{name}.driver_hashes.{driver_name}", "bound driver digest differs")
    log = obj(row["log"], f"{name}.log")
    log_path = text(log.get("path"), f"{name}.log.path")
    artifact_bytes(root, log_path, {"sha256": log.get("sha256"), "bytes": log.get("bytes")}, f"{name}.log")
    text(row["scope"], f"{name}.scope")


def observed_check_statuses(root: Path) -> dict[str, str]:
    """Mirror seal.py's source-before receipt inventory exactly."""

    observed: dict[str, str] = {}
    for path in sorted(root.rglob("*.json")):
        if path.name in {"compression.json", "expected-checks.json"}:
            continue
        relative = path.relative_to(root).as_posix()
        # profile/capture.py deliberately writes this verifier stdout marker
        # with a .json suffix.  It is checked byte-for-byte by profile/report.py
        # and is not a source-before command receipt.  Do not broaden this
        # exception: a malformed or misplaced marker must still fail JSON
        # parsing below.
        if relative in PROFILE_WORKLOAD_MARKERS and path.read_bytes() == b"VALID\n":
            continue
        value = load(root, relative)
        if not isinstance(value, dict) or "source_before" not in value:
            continue
        status = value.get("status")
        if status not in {"pass", "failed"}:
            fail(relative, "source-before receipt has unsupported terminal status")
        key = path.relative_to(root).with_suffix("").as_posix()
        if key in observed:
            fail("expected-checks", f"duplicate observed receipt {key}")
        observed[key] = status
    return observed


def verify_expected_checks(root: Path, *, required: bool) -> int:
    """Bind seal.py's expected-checks map to every source-before receipt."""

    path = root / "expected-checks.json"
    if not path.is_file():
        if required:
            fail("expected-checks", "sealed replay requires expected-checks.json")
        return 0
    expected = obj(load(root, "expected-checks.json"), "expected-checks")
    normalized: dict[str, str] = {}
    for key, value in expected.items():
        text(key, "expected-checks.key")
        if Path(key).name == key or not key.startswith("checks/"):
            fail("expected-checks", "keys must be checks/<receipt> paths")
        if value not in {"pass", "failed"}:
            fail(f"expected-checks.{key}", "status must be pass or failed")
        if key in normalized:
            fail("expected-checks", f"duplicate receipt {key}")
        normalized[key] = value
    observed = observed_check_statuses(root)
    if normalized != observed:
        fail("expected-checks", "expected receipt set or status differs from retained source-before receipts")
    return len(normalized)


def verify_gate_retries(root: Path, required_pass: list[Any]) -> dict[str, str]:
    """Validate optional same-command retries and return selected receipts."""

    path = root / "gate-retries.json"
    if not path.is_file():
        return {}
    retries = obj(load(root, "gate-retries.json"), "gate-retries")
    required_names = {value for value in required_pass if isinstance(value, str)}
    selected: dict[str, str] = {}
    for logical, value in retries.items():
        if logical not in required_names or Path(logical).name != logical:
            fail(f"gate-retries.{logical}", "must name a required pass gate")
        row = obj(value, f"gate-retries.{logical}")
        if set(row) != {"attempts", "selected", "reason"}:
            fail(f"gate-retries.{logical}", "fields differ")
        attempts = array(row["attempts"], f"gate-retries.{logical}.attempts")
        if len(attempts) < 2:
            fail(f"gate-retries.{logical}.attempts", "must retain the base attempt and at least one retry")
        expected_names = [logical] + [f"{logical}-v{index}" for index in range(2, len(attempts) + 1)]
        if attempts != expected_names:
            fail(f"gate-retries.{logical}.attempts", "must use base, base-v2, base-v3 naming in order")
        chosen = text(row["selected"], f"gate-retries.{logical}.selected")
        if chosen != attempts[-1]:
            fail(f"gate-retries.{logical}.selected", "must select the last retained attempt")
        text(row["reason"], f"gate-retries.{logical}.reason")

        identity: tuple[Any, ...] | None = None
        for attempt in attempts:
            receipt_name = f"checks/{attempt}.json"
            receipt = obj(load(root, receipt_name), receipt_name)
            if receipt.get("status") not in {"pass", "failed", "review"}:
                fail(receipt_name, "retry receipt is not terminal")
            argv = array(receipt.get("argv"), f"{receipt_name}.argv")
            if not argv or any(not isinstance(argument, str) or not argument for argument in argv):
                fail(f"{receipt_name}.argv", "retry command arguments are malformed")
            command_index = 0
            if Path(argv[0]).name == "env":
                command_index = 1
                assignment_count = 0
                while command_index < len(argv) and re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*=.*", argv[command_index]):
                    assignment_count += 1
                    command_index += 1
                if assignment_count == 0 or command_index >= len(argv):
                    fail(receipt_name, "leading env must contain NAME=value assignments and no flags")
            program = Path(argv[command_index]).name
            subcommand = argv[command_index + 1] if command_index + 1 < len(argv) else None
            if not (
                (program == "cargo" and subcommand in {"test", "clippy", "doc", "fmt"})
                or program == "rustfmt"
            ):
                fail(receipt_name, "retry is outside the permitted validation command set")
            current_identity = (
                tuple(argv),
                text(receipt.get("cwd"), f"{receipt_name}.cwd"),
                sha_text(receipt.get("driver_sha256"), f"{receipt_name}.driver_sha256"),
                text(receipt.get("source_scope"), f"{receipt_name}.source_scope"),
            )
            if identity is None:
                identity = current_identity
            elif current_identity != identity:
                fail(f"gate-retries.{logical}", "attempts do not retain an identical command, cwd, driver, and source scope")
        chosen_receipt = obj(load(root, f"checks/{chosen}.json"), f"checks/{chosen}.json")
        if chosen_receipt.get("status") != "pass" or chosen_receipt.get("exit_code") != 0 or chosen_receipt.get("source_unchanged") is not True:
            fail(f"gate-retries.{logical}.selected", "selected retry is not a successful unchanged-source receipt")
        selected[logical] = chosen
    return selected


def planned_passes_for_stage(required_pass: list[Any], stage: str) -> list[Any]:
    """Return only gates that must already exist at this replay lifecycle stage."""

    if stage not in LIFECYCLE_STAGES:
        fail("lifecycle.stage", f"must be one of {LIFECYCLE_STAGES!r}")
    if stage == "final":
        return required_pass
    boundary = "precleanup-portable" if stage == "precleanup" else "aftercleanup-portable"
    try:
        end = required_pass.index(boundary)
    except ValueError:
        fail("planned-checks.required_pass", f"must declare the {boundary!r} lifecycle boundary")
    return required_pass[:end]


def verify_checks(root: Path, stage: str = "final", *, require_expected: bool = False) -> int:
    checks = safe_path(root, "checks")
    if not checks.is_dir():
        fail("checks", "staged replay requires the checks directory")
    driver_hash = digest((root / "check.py").read_bytes()) if (root / "check.py").is_file() else None
    active_tag = {"precleanup": "precleanup-portable", "aftercleanup": "aftercleanup-portable"}.get(stage)
    count = 0
    for path in sorted(checks.glob("*.json")):
        name = str(path.relative_to(root))
        row = obj(load(root, name), name)
        if row.get("change") not in {None, CHANGE}:
            continue
        if active_tag is not None and name == f"checks/{active_tag}.json" and row.get("status") == "running":
            # check.py creates this receipt before launching the current
            # staged verifier and finalizes it only after this process exits.
            continue
        if any(key in row for key in ("driver_hashes", "replayed_inventory_sha256", "portable_flag")):
            verify_replay_receipt(root, name, row)
            count += 1
            continue
        # Manual custody/review rows may have no argv.  They remain inventory
        # bound, while command receipt policy applies only to actual commands.
        if "argv" not in row:
            continue
        required = {
            "change", "argv", "cwd", "revision", "driver_sha256", "environment",
            "source_before", "status", "exit_code", "finished_utc", "source_after",
            "source_unchanged", "log",
        }
        missing = sorted(required - set(row))
        if missing:
            fail(name, f"command receipt is missing mandatory fields: {missing}")
        if uint(row["change"], f"{name}.change") != CHANGE:
            fail(f"{name}.change", f"must be {CHANGE}")
        if driver_hash is None or sha_text(row["driver_sha256"], f"{name}.driver_sha256") != driver_hash:
            fail(name, "command receipt is not bound to the immutable check.py")
        argv = array(row["argv"], f"{name}.argv")
        if not argv or any(not isinstance(value, str) or not value for value in argv):
            fail(f"{name}.argv", "must contain non-empty command arguments")
        text(row["cwd"], f"{name}.cwd")
        revision = text(row["revision"], f"{name}.revision").lower()
        if len(revision) != 40 or any(char not in HEX for char in revision):
            fail(f"{name}.revision", "must be a hexadecimal revision")
        obj(row["environment"], f"{name}.environment")
        text(row["finished_utc"], f"{name}.finished_utc")
        if row.get("status") not in {"pass", "failed", "review"}:
            fail(name, "invalid command receipt status")
        exit_code = row["exit_code"]
        if isinstance(exit_code, bool) or not isinstance(exit_code, int):
            fail(f"{name}.exit_code", "must be an integer")
        if row["source_before"] != row["source_after"]:
            fail(name, "source custody changed during command")
        source_scope = row.get("source_scope", "")
        if not isinstance(source_scope, str):
            fail(f"{name}.source_scope", "must be text when present")
        require_critical = "excluding tools" not in source_scope
        verify_source_manifest(root, row["source_before"], f"{name}.source_before", require_critical=require_critical)
        verify_source_manifest(root, row["source_after"], f"{name}.source_after", require_critical=require_critical)
        if not isinstance(row["source_unchanged"], bool) or row["source_unchanged"] is not True:
            fail(f"{name}.source_unchanged", "must record unchanged source custody")
        log = obj(row["log"], f"{name}.log")
        log_path = text(log.get("path"), f"{name}.log.path")
        raw_log = artifact_bytes(root, log_path, {"sha256": log.get("sha256"), "bytes": log.get("bytes")}, f"{name}.log")
        if "test" in argv:
            matches = re.findall(rb"test result: (?:ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored;", raw_log)
            totals = [sum(int(match[index]) for match in matches) for index in range(3)]
            for index, key in enumerate(("passed_tests", "failed_tests", "ignored_tests")):
                if key not in row or row[key] != totals[index]:
                    fail(name, f"{key} does not match the retained test-result lines")
            if row["status"] == "pass" and (not matches or exit_code != 0 or row["failed_tests"] != 0 or row["passed_tests"] == 0):
                fail(name, "passing test receipt has a nonzero exit or failed/no tests")
        elif row["status"] == "pass" and exit_code != 0:
            fail(name, "passing command receipt has a nonzero exit code")
        count += 1
    verify_expected_checks(root, required=require_expected)
    planned_path = root / "planned-checks.json"
    if not planned_path.is_file():
        fail("planned-checks", "staged replay requires planned-checks.json")
    planned = obj(load(root, "planned-checks.json"), "planned-checks")
    required_pass = array(planned.get("required_pass"), "planned-checks.required_pass")
    if not required_pass:
        fail("planned-checks.required_pass", "must contain at least one terminal gate")
    required_review = array(planned.get("required_review", []), "planned-checks.required_review")
    for field, values in (("required_pass", required_pass), ("required_review", required_review)):
        if any(not isinstance(value, str) or not value or Path(value).name != value for value in values):
            fail(f"planned-checks.{field}", "receipt names must be simple non-empty filenames")
        if len(set(values)) != len(values):
            fail(f"planned-checks.{field}", "contains duplicates")
    for boundary in ("precleanup-portable", "aftercleanup-portable"):
        if boundary not in required_pass:
            fail("planned-checks.required_pass", f"must include {boundary!r}")
    retry_selected = verify_gate_retries(root, required_pass)
    for tag in planned_passes_for_stage(required_pass, stage):
        tag = text(tag, "planned-checks.required_pass")
        resolved = retry_selected.get(tag, tag)
        row = obj(load(root, f"checks/{resolved}.json"), f"checks/{resolved}.json")
        if row.get("status") != "pass" or ("exit_code" in row and row["exit_code"] != 0) or ("source_unchanged" in row and row["source_unchanged"] is not True):
            fail(f"checks/{resolved}.json", f"required check {tag} is not passing")
    for tag in required_review:
        tag = text(tag, "planned-checks.required_review")
        row = obj(load(root, f"checks/{tag}.json"), f"checks/{tag}.json")
        if row.get("status") not in {"pass", "failed", "review"}:
            fail(f"checks/{tag}.json", "required review receipt has an invalid status")
        if "source_unchanged" in row and row["source_unchanged"] is not True:
            fail(f"checks/{tag}.json", "required review receipt changed source custody")
    return count


def verify_compression(root: Path) -> int:
    path = root / "compression.json"
    if not path.is_file():
        return 0
    compression = obj(load(root, "compression.json"), "compression")
    count = 0
    for stored_name, row_value in compression.items():
        row = obj(row_value, f"compression.{stored_name}")
        if set(row) != {"original_path", "original_sha256", "original_bytes", "stored_sha256", "stored_bytes"}:
            fail(f"compression.{stored_name}", "fields differ")
        original_name = text(row["original_path"], f"compression.{stored_name}.original_path")
        if not stored_name.endswith(".gz") or stored_name[:-3] != original_name:
            fail(f"compression.{stored_name}", "original path differs")
        stored = safe_path(root, stored_name).read_bytes()
        raw = logical_bytes(root, original_name)
        if digest(stored) != sha_text(row["stored_sha256"], f"compression.{stored_name}.stored_sha256") or len(stored) != uint(row["stored_bytes"], f"compression.{stored_name}.stored_bytes"):
            fail(f"compression.{stored_name}", "stored binding differs")
        if digest(raw) != sha_text(row["original_sha256"], f"compression.{stored_name}.original_sha256") or len(raw) != uint(row["original_bytes"], f"compression.{stored_name}.original_bytes"):
            fail(f"compression.{stored_name}", "original binding differs")
        count += 1
    return count


def verify_inventory(root: Path, required: bool = False) -> int:
    inventory_path = root / "SHA256SUMS"
    if not inventory_path.is_file():
        if required:
            fail("SHA256SUMS", "sealed inventory is required")
        return 0
    rows: dict[str, str] = {}
    for line in inventory_path.read_text(encoding="utf-8").splitlines():
        if "  " not in line:
            fail("SHA256SUMS", "malformed line")
        value, name = line.split("  ", 1)
        value = value.lower()
        if name in rows or len(value) != 64 or any(char not in HEX for char in value):
            fail("SHA256SUMS", "duplicate or malformed entry")
        target = safe_path(root, name)
        if not target.is_file() or digest(target.read_bytes()) != value:
            fail(f"SHA256SUMS.{name}", "digest differs")
        rows[name] = value
    expected = {str(path.relative_to(root)) for path in root.rglob("*") if path.is_file() and path.name != "SHA256SUMS"}
    if set(rows) != expected:
        fail("SHA256SUMS", "inventory does not cover the bundle exactly")
    return len(rows)


def verify(root: Path = ROOT, *, require_inventory: bool = False, stage: str = "final") -> dict[str, Any]:
    if stage not in LIFECYCLE_STAGES:
        fail("lifecycle.stage", f"must be one of {LIFECYCLE_STAGES!r}")
    protocol, order = verify_protocol(root)
    builds = {role: verify_build(root, protocol, role) for role in ROLES}
    if builds["before-buffered"]["revision"] == builds["after-buffered"]["revision"] or builds["before-buffered"]["revision"] == builds["after-streaming"]["revision"]:
        fail("build.revision", "before and after source revisions unexpectedly match")
    if builds["before-buffered"]["binaries"]["normal"]["sha256"] == builds["after-buffered"]["binaries"]["normal"]["sha256"]:
        fail("build.binaries", "before and after normal binaries unexpectedly match")
    for role in ROLES:
        verify_capture_state(root, role, builds[role])
    captures = {role: verify_capture_role(root, role, builds[role], protocol, order) for role in ROLES}
    verify_cross_role(captures)
    verify_strict_debt(root, builds)
    verify_production_strict_review(root, builds)
    run_checked(
        [sys.executable, "-B", str(safe_path(root, "profile/report.py")), "--check"],
        "profile derivation replay",
    )
    run_checked([sys.executable, "-B", str(safe_path(root, "derive.py")), "--check"], "summary derivation replay")
    checks = verify_checks(root, stage, require_expected=require_inventory)
    compressed = verify_compression(root)
    inventory = verify_inventory(root, required=require_inventory)
    return {"status": "pass", "change": CHANGE, "stage": stage, "roles": len(ROLES), "reports": sum(len(rows) for rows in captures.values()), "samples": sum(len(rows) * 30 for rows in captures.values()), "command_receipts": checks, "compressed_artifacts": compressed, "inventory_files": inventory}


def reject_report_mutation(
    exported: Path,
    role: str,
    lane: dict[str, Any],
    logical_report: str,
    label: str,
    mutate: Any,
) -> None:
    """Run one report mutation against a copied report/catalog pair.

    Keeping the original bundle files untouched also exercises sealed bundles
    whose JSON has been losslessly compressed.  The temporary report retains
    the original basename so the verifier resolves its catalog sidecar.
    """

    original = logical_bytes(exported, logical_report)
    report = obj(load_bytes(original, logical_report), logical_report)
    mutate(report)
    catalog_logical = logical_report[:-5] + "-catalog.json"
    catalog = logical_bytes(exported, catalog_logical)
    with tempfile.TemporaryDirectory(prefix="litchi-0433-report-mutation-") as directory:
        temporary = Path(directory)
        report_path = temporary / Path(logical_report).name
        report_path.write_text(json.dumps(report, indent=2, sort_keys=True, allow_nan=False) + "\n", encoding="utf-8")
        (temporary / (report_path.stem + "-catalog.json")).write_bytes(catalog)
        command = [
            sys.executable,
            "-B",
            str(exported / "verify-report.py"),
            "--report",
            str(report_path),
            "--mode",
            lane["mode"],
            "--shape",
            lane["shape"],
            "--role",
            role,
        ]
        result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if result.returncode == 0:
            fail(label, "report verifier accepted a mutated report")


def portable_mutations(root: Path, baseline: dict[str, Any], stage: str, require_inventory: bool) -> dict[str, str]:
    """Exercise semantic, numeric, allocator, retention, role, and custody guards."""

    with tempfile.TemporaryDirectory(prefix="litchi-0433-portable-") as temporary:
        exported = Path(temporary) / "bundle"
        shutil.copytree(root, exported)
        replay = verify(exported, require_inventory=require_inventory, stage=stage)
        if replay != baseline:
            fail("portable replay", "copied bundle result differs from source bundle")
        _protocol, order = verify_protocol(exported)

        normal_role = "before-buffered"
        normal_lane = next(lane for lane in order if lane["mode"] == "normal")
        normal_name = expected_name(normal_role, normal_lane)
        normal_report = f"before/captures/{normal_name}.json"

        def mutate_elapsed(report: dict[str, Any]) -> None:
            result = obj(array(report["results"], "report.results")[0], "report.results[0]")
            elapsed = obj(result["elapsed_ns"], "report.results[0].elapsed_ns")
            elapsed["mean"] = int(elapsed["mean"]) + 1

        def mutate_output(report: dict[str, Any]) -> None:
            result = obj(array(report["results"], "report.results")[0], "report.results[0]")
            result["output_sha256"] = "0" * 64

        def mutate_semantic(report: dict[str, Any]) -> None:
            result = obj(array(report["results"], "report.results")[0], "report.results[0]")
            source = obj(result["source"], "report.results[0].source")
            oracle = obj(source["ods_scalar_rows"], "report.results[0].source.ods_scalar_rows")
            oracle["semantic_sha256"] = "0" * 64

        reject_report_mutation(exported, normal_role, normal_lane, normal_report, "portable numeric mutation", mutate_elapsed)
        reject_report_mutation(exported, normal_role, normal_lane, normal_report, "portable output mutation", mutate_output)
        reject_report_mutation(exported, normal_role, normal_lane, normal_report, "portable semantic-digest mutation", mutate_semantic)

        allocator_role = "before-buffered"
        allocator_lane = next(lane for lane in order if lane["mode"] == "allocator")
        allocator_name = expected_name(allocator_role, allocator_lane)
        allocator_report = f"before/captures/{allocator_name}.json"

        def mutate_allocator(report: dict[str, Any]) -> None:
            result = obj(array(report["results"], "report.results")[0], "report.results[0]")
            operation = obj(result["operation_metrics"], "report.results[0].operation_metrics")
            allocation = obj(operation["allocation"], "report.results[0].operation_metrics.allocation")
            after = obj(allocation["peak_live_bytes_after"], "peak_live_bytes_after")
            region = obj(allocation["region_peak_live_bytes"], "region_peak_live_bytes")
            values = array(after["values"], "peak_live_bytes_after.values")
            region_values = array(region["values"], "region_peak_live_bytes.values")
            region_values[0] = int(values[0]) + 1

        reject_report_mutation(exported, allocator_role, allocator_lane, allocator_report, "portable allocator-vector mutation", mutate_allocator)

        streaming_role = "after-streaming"
        streaming_lane = next(lane for lane in order if lane["mode"] == "normal")
        streaming_name = expected_name(streaming_role, streaming_lane)
        streaming_report = f"after/captures/{streaming_name}.json"

        def mutate_retention(report: dict[str, Any]) -> None:
            result = obj(array(report["results"], "report.results")[0], "report.results[0]")
            sink = obj(result["sink"], "report.results[0].sink")
            sink["retained_authoring_window_bytes"] = 4_095

        def mutate_role_binding(report: dict[str, Any]) -> None:
            result = obj(array(report["results"], "report.results")[0], "report.results[0]")
            source = obj(result["source"], "report.results[0].source")
            oracle = obj(source["ods_scalar_rows"], "report.results[0].source.ods_scalar_rows")
            oracle["role"] = "buffered"

        reject_report_mutation(exported, streaming_role, streaming_lane, streaming_report, "portable stream-retention mutation", mutate_retention)
        reject_report_mutation(exported, streaming_role, streaming_lane, streaming_report, "portable role-binding mutation", mutate_role_binding)

        inventory = exported / "SHA256SUMS"
        if inventory.is_file():
            inventory.unlink()
        verifier = exported / "verify-report.py"
        verifier.write_bytes(verifier.read_bytes() + b"\n# portable bound-verifier mutation\n")
        result = subprocess.run([sys.executable, "-B", str(exported / "verify.py"), "--stage", stage], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if result.returncode == 0:
            fail("portable bound-verifier mutation", "bundle accepted a changed bound verifier")
    return {
        "portable_export": "pass",
        "portable_numeric_mutation": "rejected",
        "portable_output_mutation": "rejected",
        "portable_semantic_digest_mutation": "rejected",
        "portable_allocator_vectors_mutation": "rejected",
        "portable_stream_retention_mutation": "rejected",
        "portable_role_binding_mutation": "rejected",
        "portable_bound_verifier_mutation": "rejected",
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--portable-check", action="store_true")
    parser.add_argument("--require-inventory", action="store_true")
    parser.add_argument("--stage", choices=LIFECYCLE_STAGES, default="final", help="portable evidence lifecycle gate; final requires every planned pass")
    args = parser.parse_args(argv)
    try:
        result = verify(ROOT, require_inventory=args.require_inventory, stage=args.stage)
        if args.portable_check:
            result.update(portable_mutations(ROOT, result, args.stage, args.require_inventory))
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0
    except (OSError, KeyError, TypeError, ValueError, AssertionError) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
