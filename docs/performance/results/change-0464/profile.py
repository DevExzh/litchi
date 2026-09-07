#!/usr/bin/env python3
"""Capture optional whole-process diagnostics for the retained 0464 pair.

This helper deliberately profiles the normal ``bytes`` pair-lifecycle binary
after the formal eight-lane capture.  ``perf stat`` and ``perf record`` wrap
the complete child process, so the reported counters and sampled symbols
include input loading, setup, all three warmups, all 30 retained samples (33
checked iterations), the in-process semantic/raw oracles, output writes,
teardown, and report serialization.  They do not attribute work to an
individual lifecycle operation.

The helper is optional: missing permissions, ``perf``, call-graph support, or
post-processing tools are recorded as unavailable.  It never substitutes
operation-only attribution for a whole-process result.

The measured normal binary remains bound to its original source manifest.  A
post-capture source change is admitted only through the retained compatibility
attestation, which must prove that the current source differs in exactly the
allowlisted semantic-inventory diagnostic file; the supplemental inventory
binding and build receipt are checked alongside it.
"""

from __future__ import annotations

import csv
import datetime as dt
import hashlib
import json
import math
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
from typing import Any


SCRIPT = Path(__file__).resolve()
ROOT = SCRIPT.parent
REPO = ROOT.parents[3]
OUTPUT_ROOT = ROOT / "profiling"
SUMMARY_PATH = OUTPUT_ROOT / "profile-summary.json"
PROTOCOL_PATH = ROOT / "protocol.json"
PAIR_PATH = ROOT / "pair.json"
FORMAL_BINDING_PATH = ROOT / "binding.json"
FINAL_BINDING_PATH = ROOT / "normal-final-binding.json"
COMPATIBILITY_PATH = ROOT / "source-compatibility.json"
INVENTORY_BINDING_PATH = ROOT / "inventory-binding.json"
INVENTORY_SOURCE_PATH = "tools/perf-baseline/src/bin/pptx_semantic_inventory.rs"
CPU = 2
SAMPLES = 30
WARMUP = 3
CHECKED_ITERATIONS = SAMPLES + WARMUP
REPEAT = "diagnostic"
PROVIDER = "bytes"
EVENTS = (
    "cycles",
    "instructions",
    "branches",
    "branch-misses",
    "cache-references",
    "cache-misses",
    "page-faults",
)
EVENT_STRING = ",".join(EVENTS)
WHOLE_PROCESS_SCOPE = (
    "whole-process child lifetime including input loading, setup, three "
    "warmups, 30 retained samples (33 checked iterations), semantic/raw "
    "oracle checks, output writes, teardown, and report serialization; "
    "not operation-only attribution"
)
TOP_SYMBOL_SCOPE = "whole-process sampled cycles:u; no operation-only attribution"


class ProfileError(RuntimeError):
    """The bound diagnostic inputs or retained protocol are inconsistent."""


def fail(message: str) -> None:
    raise ProfileError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def load(path: Path, label: str) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{label}: invalid JSON: {error}")


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def regular(path: Path, label: str) -> Path:
    require(path.is_file() and not path.is_symlink(), f"{label} is missing or symlinked: {path}")
    return path


def path_label(path: Path) -> str:
    path = path.resolve()
    for base in (REPO.resolve(), OUTPUT_ROOT.resolve()):
        try:
            return path.relative_to(base).as_posix()
        except ValueError:
            continue
    return str(path)


def artifact(path: Path) -> dict[str, Any]:
    regular(path, "artifact")
    return {"path": path_label(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def load_protocol() -> tuple[Path, dict[str, Any]]:
    protocol_path = regular(PROTOCOL_PATH, "0464 protocol")
    protocol = load(protocol_path, "0464 protocol")
    require(isinstance(protocol, dict), "0464 protocol must be an object")
    require(protocol.get("schema") == "litchi-0464-pptx-pair-capture-v1", "0464 protocol schema differs")
    require(protocol.get("change") == 464 and protocol.get("status") == "frozen", "0464 protocol is not frozen")
    require(protocol.get("cpu") == CPU, "0464 protocol CPU differs")
    require(protocol.get("workers") == 1, "0464 protocol must use one worker")
    require(protocol.get("samples") == SAMPLES and protocol.get("warmups") == WARMUP, "0464 protocol sample counts differ")
    matrix = protocol.get("matrix")
    require(isinstance(matrix, dict), "0464 protocol matrix is missing")
    require(matrix.get("reports_total") == 8 and matrix.get("reports_per_repeat") == 4, "0464 protocol does not declare eight formal lanes")
    require(matrix.get("serialized_lanes") is True, "0464 formal lanes are not serialized")
    orders = protocol.get("orders")
    require(isinstance(orders, dict), "0464 protocol lane orders are missing")
    seen: set[tuple[str, str]] = set()
    for repeat in ("R1", "R2"):
        lanes = orders.get(repeat)
        require(isinstance(lanes, list) and len(lanes) == 4, f"0464 protocol {repeat} lane order differs")
        for lane in lanes:
            require(isinstance(lane, dict), f"0464 protocol {repeat} lane is not an object")
            instrumentation = lane.get("instrumentation")
            provider = lane.get("provider")
            require(isinstance(instrumentation, str) and isinstance(provider, str), f"0464 protocol {repeat} lane identity is incomplete")
            seen.add((instrumentation, provider))
    require(seen == {("normal", "bytes"), ("normal", "range"), ("allocator", "bytes"), ("allocator", "range")}, "0464 formal lane coverage differs")
    return protocol_path, protocol


def digest_bytes(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def source_ref(path: Path, label: str) -> tuple[dict[str, Any], dict[str, str]]:
    require(path.resolve().is_relative_to(ROOT.resolve()), f"{label} escapes the 0464 bundle")
    regular(path, label)
    value = load(path, label)
    require(isinstance(value, dict), f"{label} must be a source manifest object")
    reference = {"path": str(path.relative_to(ROOT)), "sha256": sha(path), "files": len(value)}
    return reference, {str(name): str(value_hash) for name, value_hash in value.items()}


def current_source_manifest() -> tuple[dict[str, Any], dict[str, str], list[str]]:
    argv = ["git", "ls-files", "--cached", "--others", "--exclude-standard", "-z"]
    try:
        raw_names = subprocess.check_output(argv, cwd=REPO)
    except (OSError, subprocess.CalledProcessError) as error:
        fail(f"current source custody command failed: {error}")
    try:
        names = set(raw_names.decode().split("\0"))
    except UnicodeDecodeError as error:
        fail(f"current source custody names are not UTF-8: {error}")
    names.add("Cargo.lock")
    rows: dict[str, str] = {}
    for name in sorted(names):
        path = REPO / name
        if name.endswith((".rs", ".toml", ".lock")) and path.is_file() and not path.is_symlink():
            rows[name] = sha(path)
    require(rows, "current source custody is empty")
    canonical = (json.dumps(rows, sort_keys=True, indent=2) + "\n").encode()
    return {
        "path": None,
        "sha256": digest_bytes(canonical),
        "files": len(rows),
    }, rows, argv


def manifest_reference(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is missing")
    reference = {
        "path": value.get("path"),
        "sha256": value.get("sha256"),
        "files": value.get("files"),
    }
    require(isinstance(reference["path"], str) and reference["path"], f"{label}.path is missing")
    require(isinstance(reference["sha256"], str) and len(reference["sha256"]) == 64, f"{label}.sha256 is missing")
    require(isinstance(reference["files"], int) and reference["files"] > 0, f"{label}.files is missing")
    return reference


def first_mapping(value: dict[str, Any], keys: tuple[str, ...]) -> Any:
    for key in keys:
        if key in value:
            return value[key]
    return None


def changed_paths(value: Any, label: str) -> list[str]:
    if isinstance(value, dict):
        value = value.get("files", value.get("paths", value.get("changed")))
    require(isinstance(value, list), f"{label} must be a list")
    result: list[str] = []
    for row in value:
        if isinstance(row, str):
            path = row
        elif isinstance(row, dict):
            path = first_mapping(row, ("path", "source_path", "file"))
        else:
            path = None
        require(isinstance(path, str) and path, f"{label} contains an invalid path")
        result.append(path)
    return sorted(set(result))


def validate_source_compatibility(
    path: Path,
    measured_source: dict[str, Any],
    measured_map: dict[str, str],
) -> dict[str, Any]:
    compatibility = load(path, "source compatibility")
    require(isinstance(compatibility, dict), "source compatibility must be an object")
    schema = compatibility.get("schema")
    require(schema == "litchi-0464-source-compatibility-v1", "source compatibility schema differs")
    if "status" in compatibility:
        require(compatibility.get("status") == "pass", "source compatibility is not a passing attestation")
    baseline_value = first_mapping(compatibility, ("measured_source", "baseline_source", "old_source", "source_before", "baseline", "before"))
    current_value = first_mapping(compatibility, ("current_source", "supplemental_source", "final_source", "source_after", "current", "after"))
    baseline = manifest_reference(baseline_value, "source compatibility measured source")
    current = manifest_reference(current_value, "source compatibility current source")
    require(baseline == measured_source, "source compatibility measured source differs from normal binary binding")
    require(current["sha256"] != baseline["sha256"], "source compatibility does not attest a changed current source")
    current_manifest_path = ROOT / current["path"]
    current_manifest, current_map = source_ref(current_manifest_path, "source compatibility current manifest")
    require(current_manifest == current, "source compatibility current manifest identity differs")
    actual_current, actual_map, custody_argv = current_source_manifest()
    require(actual_current["files"] == current["files"] and actual_current["sha256"] == current["sha256"], "current source custody differs from source compatibility")
    require(actual_map == current_map, "current source manifest entries differ from source compatibility")
    allow_value = first_mapping(compatibility, ("allowlisted_changed_files", "allowed_differences", "allowlist", "allowed_changes", "changed_allowlist"))
    allowlist = changed_paths(allow_value, "source compatibility allowlist")
    require(allowlist == [INVENTORY_SOURCE_PATH], "source compatibility allowlist is broader than the diagnostic inventory source")
    changed = sorted(name for name in set(measured_map) | set(actual_map) if measured_map.get(name) != actual_map.get(name))
    require(changed == allowlist, "current source differs from measured source outside the explicit allowlist")
    changed_value = first_mapping(compatibility, ("changed_files", "actual_differences", "differences", "diff"))
    require(changed_value is not None, "source compatibility changed file list is missing")
    require(changed_paths(changed_value, "source compatibility changed files") == changed, "source compatibility changed file list differs")
    unchanged = compatibility.get("unchanged_files")
    require(unchanged == len(measured_map) - len(changed), "source compatibility unchanged file count differs")
    supplemental_artifact = compatibility.get("artifact")
    require(isinstance(supplemental_artifact, dict), "source compatibility supplemental artifact is missing")
    require(supplemental_artifact.get("source_path") == INVENTORY_SOURCE_PATH, "source compatibility artifact source path differs")
    artifact_path = ROOT / str(supplemental_artifact.get("path", ""))
    regular(artifact_path, "source compatibility supplemental artifact")
    require(supplemental_artifact.get("bytes") == artifact_path.stat().st_size and supplemental_artifact.get("sha256") == sha(artifact_path), "source compatibility supplemental artifact identity differs")
    source_bytes = (REPO / INVENTORY_SOURCE_PATH).read_bytes()
    require(supplemental_artifact.get("bytes") == len(source_bytes) and supplemental_artifact.get("sha256") == digest_bytes(source_bytes), "source compatibility artifact does not match current diagnostic source")
    return {
        "path": path,
        "schema": schema,
        "measured_source": baseline,
        "current_source": current,
        "allowlisted_changed_files": changed,
        "unchanged_files": len(measured_map) - len(changed),
        "artifact": supplemental_artifact,
        "custody_argv": custody_argv,
    }


def binding_entry(value: dict[str, Any], label: str) -> dict[str, Any]:
    entry = value.get("normal")
    if not isinstance(entry, dict) and isinstance(value.get("binaries"), dict):
        entry = value["binaries"].get("normal")
    require(isinstance(entry, dict), f"{label} normal entry is missing")
    return entry


def validate_inventory_binding(path: Path, context_source: dict[str, Any], measured_source: dict[str, Any], source_revision: str) -> dict[str, Any]:
    value = load(path, "inventory binding")
    require(isinstance(value, dict), "inventory binding must be an object")
    schema = value.get("schema")
    require(schema == "litchi-0464-supplemental-inventory-v1", "inventory binding schema differs")
    require(value.get("revision") == source_revision, "inventory binding revision differs from pair source revision")
    inventory = value.get("inventory", value.get("binary"))
    if not isinstance(inventory, dict) and isinstance(value.get("binaries"), dict):
        inventory = value["binaries"].get("inventory")
    require(isinstance(inventory, dict), "inventory binding entry is missing")
    binary_path = Path(str(inventory.get("path", "")))
    regular(binary_path, "bound inventory binary")
    require(inventory.get("bytes") == binary_path.stat().st_size and inventory.get("sha256") == sha(binary_path), "inventory binary identity differs from binding")
    require(inventory.get("source") == context_source, "inventory binding source differs from current compatibility source")
    require(value.get("measured_source") == measured_source, "inventory binding measured source differs from normal binary binding")
    build_path = ROOT / str(inventory.get("build_receipt", ""))
    regular(build_path, "inventory build receipt")
    require(inventory.get("build_receipt_sha256") == sha(build_path), "inventory build receipt hash differs from binding")
    build = load(build_path, "inventory build receipt")
    require(build.get("change") == 464 and build.get("status") == "pass" and build.get("exit_code") == 0 and build.get("source_unchanged") is True, "inventory build receipt is not a successful source-unchanged build")
    require(build.get("revision") == source_revision, "inventory build revision differs from pair source revision")
    require(build.get("source_before") == context_source and build.get("source_after") == context_source, "inventory build source differs from compatibility source")
    return {
        "path": path,
        "schema": schema,
        "binding": inventory,
        "binary_path": binary_path,
        "build_path": build_path,
    }


def load_bound_context() -> dict[str, Any]:
    protocol_path, protocol = load_protocol()
    pair_path = regular(PAIR_PATH, "pair manifest")
    protocol_binding = protocol.get("binding")
    require(isinstance(protocol_binding, dict) and isinstance(protocol_binding.get("path"), str), "0464 protocol binding is missing")
    formal_binding_path = ROOT / protocol_binding["path"]
    regular(formal_binding_path, "formal binary binding")
    formal_binding = load(formal_binding_path, "formal binary binding")
    require(isinstance(formal_binding, dict), "formal binary binding must be an object")
    measured_binding_path = FINAL_BINDING_PATH if FINAL_BINDING_PATH.is_file() and not FINAL_BINDING_PATH.is_symlink() else formal_binding_path
    measured_binding = load(measured_binding_path, "measured normal binding")
    require(isinstance(measured_binding, dict), "measured normal binding must be an object")
    pair = load(pair_path, "pair manifest")
    require(isinstance(pair, dict), "pair manifest must be an object")
    pair_sha = sha(pair_path)
    protocol_pair = protocol.get("pair_manifest")
    require(isinstance(protocol_pair, dict), "0464 protocol pair binding is missing")
    require(protocol_pair.get("path") == pair_path.name and protocol_pair.get("sha256") == pair_sha, "pair manifest differs from frozen protocol")
    require(pair.get("schema") == "pptx_pair_manifest_v1", "pair manifest schema differs")
    pair_id = pair.get("pair_id")
    source_revision = pair.get("source_revision")
    require(isinstance(pair_id, str) and pair_id, "pair manifest pair_id is missing")
    require(isinstance(source_revision, str) and source_revision, "pair manifest source_revision is missing")
    require(source_revision == protocol_pair.get("source_revision"), "pair source revision differs from frozen protocol")
    require(formal_binding.get("schema") == "litchi-0464-binaries-v1", "formal binary binding schema differs")
    require(formal_binding.get("revision") == source_revision, "formal binary binding revision differs from pair source revision")

    normal = binding_entry(measured_binding, "measured binding")
    formal_normal = binding_entry(formal_binding, "formal binding")
    for key in ("path", "sha256", "bytes", "build_receipt", "build_receipt_sha256", "source"):
        require(normal.get(key) == formal_normal.get(key), f"measured normal {key} differs from formal binding")
    binary_path = Path(str(normal.get("path", "")))
    regular(binary_path, "bound normal binary")
    expected_binary = protocol.get("binaries", {}).get("normal", {})
    require(isinstance(expected_binary, dict), "0464 protocol normal binary binding is missing")
    require(normal.get("path") == expected_binary.get("path"), "normal binary path differs from frozen protocol")
    require(normal.get("bytes") == binary_path.stat().st_size and normal.get("sha256") == sha(binary_path), "normal binary identity differs from binding")

    source = normal.get("source")
    require(isinstance(source, dict), "normal binary source binding is missing")
    source_path = ROOT / str(source.get("path", ""))
    source_manifest, source_map = source_ref(source_path, "normal binary source manifest")
    require(source_manifest == source, "normal binary source manifest differs from binding")
    build_path = ROOT / str(normal.get("build_receipt", ""))
    regular(build_path, "normal build receipt")
    require(normal.get("build_receipt_sha256") == sha(build_path), "normal build receipt hash differs from binding")
    build = load(build_path, "normal build receipt")
    require(build.get("status") == "pass" and build.get("source_unchanged") is True, "normal build receipt is not a successful source-unchanged build")
    require(build.get("revision") == source_revision, "normal build revision differs from pair source revision")
    require(build.get("source_before") == source and build.get("source_after") == source, "normal build source binding differs")

    compatibility_path = regular(COMPATIBILITY_PATH, "source compatibility")
    compatibility = validate_source_compatibility(compatibility_path, source, source_map)
    inventory_path = regular(INVENTORY_BINDING_PATH, "inventory binding")
    inventory = validate_inventory_binding(inventory_path, compatibility["current_source"], source, source_revision)
    inputs: dict[str, dict[str, Any]] = {}
    for name in ("source", "destination"):
        identity = pair.get(name)
        require(isinstance(identity, dict), f"pair {name} identity is missing")
        input_path = ROOT / str(identity.get("path", ""))
        regular(input_path, f"pair {name} input")
        require(identity.get("bytes") == input_path.stat().st_size and identity.get("sha256") == sha(input_path), f"pair {name} input identity differs")
        inputs[name] = {"path": input_path, "identity": identity}

    return {
        "protocol_path": protocol_path,
        "protocol": protocol,
        "pair_path": pair_path,
        "pair": pair,
        "pair_sha256": pair_sha,
        "formal_binding_path": formal_binding_path,
        "formal_binding": formal_binding,
        "binding_path": measured_binding_path,
        "binding": measured_binding,
        "binary_path": binary_path,
        "binary": normal,
        "build_path": build_path,
        "build": build,
        "source_path": source_path,
        "source": source,
        "source_map": source_map,
        "compatibility_path": compatibility_path,
        "compatibility": compatibility,
        "inventory_path": inventory_path,
        "inventory": inventory,
        "inputs": inputs,
        "source_revision": source_revision,
        "pair_id": pair_id,
    }


def bound_artifacts(context: dict[str, Any]) -> dict[str, dict[str, Any]]:
    inventory = context["inventory"]
    return {
        "protocol": artifact(context["protocol_path"]),
        "pair_manifest": artifact(context["pair_path"]),
        "formal_binding": artifact(context["formal_binding_path"]),
        "measured_binding": artifact(context["binding_path"]),
        "normal_binary": artifact(context["binary_path"]),
        "normal_build_receipt": artifact(context["build_path"]),
        "measured_source_manifest": artifact(context["source_path"]),
        "source_compatibility": artifact(context["compatibility_path"]),
        "current_source_manifest": artifact(ROOT / context["compatibility"]["current_source"]["path"]),
        "inventory_binding": artifact(context["inventory_path"]),
        "inventory_binary": artifact(inventory["binary_path"]),
        "inventory_build_receipt": artifact(inventory["build_path"]),
        "source_input": artifact(context["inputs"]["source"]["path"]),
        "destination_input": artifact(context["inputs"]["destination"]["path"]),
    }


def fixed_environment() -> tuple[dict[str, str], dict[str, str]]:
    environment = os.environ.copy()
    ambient = {
        key: value
        for key, value in sorted(environment.items())
        if key == "RUSTFLAGS" or key == "GLIBC_TUNABLES" or key.startswith("MALLOC_")
    }
    environment.pop("RUSTFLAGS", None)
    environment.pop("GLIBC_TUNABLES", None)
    for key in tuple(environment):
        if key.startswith("MALLOC_"):
            environment.pop(key, None)
    environment.update({
        "DEBUGINFOD_URLS": "",
        "PYTHONDONTWRITEBYTECODE": "1",
        "RUSTUP_TOOLCHAIN": "1.98.1",
    })
    return environment, ambient


def workload_argv(context: dict[str, Any], report: Path, output: Path) -> list[str]:
    return [
        str(context["binary_path"]),
        "pptx-pair-lifecycle",
        "--manifest", str(context["pair_path"]),
        "--provider", PROVIDER,
        "--samples", str(SAMPLES),
        "--warmup", str(WARMUP),
        "--repeat", REPEAT,
        "--output", str(report),
        "--output-pptx", str(output),
        "--source-revision", context["source_revision"],
    ]


def run_command(argv: list[str], stdout_path: Path, stderr_path: Path, environment: dict[str, str], cwd: Path = REPO) -> dict[str, Any]:
    record: dict[str, Any] = {
        "argv": argv,
        "cwd": str(cwd),
        "stdout": path_label(stdout_path),
        "stderr": path_label(stderr_path),
        "exit_code": None,
    }
    try:
        with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
            result = subprocess.run(argv, cwd=cwd, env=environment, stdout=stdout, stderr=stderr, check=False)
        record["exit_code"] = result.returncode
    except OSError as error:
        record["spawn_error"] = str(error)
        try:
            stderr_path.write_text(f"unavailable: {error}\n", encoding="utf-8")
        except OSError:
            pass
    return record


def resolve_report_path(value: Any) -> Path:
    require(isinstance(value, str) and value, "report output path is missing")
    path = Path(value)
    return path if path.is_absolute() else (REPO / path)


def validate_report(context: dict[str, Any], report_path: Path, output_path: Path) -> dict[str, Any]:
    regular(report_path, "lifecycle report")
    report = load(report_path, "lifecycle report")
    require(isinstance(report, dict), "lifecycle report must be an object")
    expected = {
        "schema": "pptx_pair_lifecycle_v1",
        "pair_id": context["pair_id"],
        "manifest_sha256": context["pair_sha256"],
        "provider": PROVIDER,
        "repeat": REPEAT,
        "source_revision": context["source_revision"],
        "binary_sha256": context["binary"]["sha256"],
        "binary_bytes": context["binary"]["bytes"],
        "samples": SAMPLES,
        "warmup": WARMUP,
        "checked_iteration_count": CHECKED_ITERATIONS,
        "input_source_unchanged": True,
    }
    for key, value in expected.items():
        require(report.get(key) == value, f"lifecycle report {key} differs")
    rows = report.get("samples_raw")
    require(isinstance(rows, list) and len(rows) == SAMPLES, "lifecycle report retained sample count differs")
    for index, row in enumerate(rows):
        require(isinstance(row, dict), f"lifecycle report sample {index} is not an object")
        require(isinstance(row.get("semantic_oracle"), dict) and isinstance(row.get("raw_oracle"), dict), f"lifecycle report sample {index} lacks semantic/raw oracle records")
        require(row.get("output_sha256") and row.get("output_bytes"), f"lifecycle report sample {index} lacks output identity")
    source_identity = context["inputs"]["source"]["identity"]
    destination_identity = context["inputs"]["destination"]["identity"]
    for key, identity in (("source", source_identity), ("destination", destination_identity)):
        value = report.get(key)
        require(isinstance(value, dict) and value.get("sha256") == identity["sha256"] and value.get("bytes") == identity["bytes"], f"lifecycle report {key} identity differs")
    output_artifact = report.get("output_artifact")
    require(isinstance(output_artifact, dict), "lifecycle report output artifact is missing")
    require(resolve_report_path(output_artifact.get("path")) == output_path.resolve(), "lifecycle report output path differs")
    regular(output_path, "lifecycle output")
    output_sha = sha(output_path)
    output_bytes = output_path.stat().st_size
    require(output_artifact.get("sha256") == output_sha and output_artifact.get("bytes") == output_bytes, "lifecycle output identity differs")
    for index, row in enumerate(rows):
        require(row.get("output_sha256") == output_sha and row.get("output_bytes") == output_bytes, f"lifecycle sample {index} output identity differs")
    return report


def canonical_event(value: str) -> str:
    value = value.strip().split("#", 1)[0].strip()
    value = re.sub(r":(?:u|k|h|H)$", "", value)
    return value


def numeric_counter(value: str) -> int | float | None:
    candidate = value.strip().replace(",", "")
    if not re.fullmatch(r"(?:\d+(?:\.\d*)?|\.\d+)", candidate):
        return None
    number = float(candidate)
    if not math.isfinite(number):
        return None
    return int(number) if number.is_integer() else number


def parse_counters(path: Path) -> dict[str, Any]:
    rows: list[dict[str, Any]] = []
    try:
        with path.open(newline="", encoding="utf-8", errors="replace") as stream:
            for raw in csv.reader(stream):
                if len(raw) < 3:
                    continue
                event = raw[2].strip()
                if not event:
                    continue
                rows.append({
                    "raw_value": raw[0].strip(),
                    "unit": raw[1].strip() if len(raw) > 1 else "",
                    "event": event,
                    "runtime": raw[3].strip() if len(raw) > 3 else "",
                    "running_percent": raw[4].strip() if len(raw) > 4 else "",
                    "total": numeric_counter(raw[0]),
                })
    except OSError as error:
        return {"status": "unavailable", "reason": str(error), "events": {}}
    parsed: dict[str, Any] = {}
    for expected in EVENTS:
        candidates = [row for row in rows if canonical_event(row["event"]) == expected]
        if not candidates:
            parsed[expected] = {
                "status": "unavailable",
                "reason": "event row is absent",
                "scope": "whole-process",
                "counter_scope": WHOLE_PROCESS_SCOPE,
                "normalization_denominator": CHECKED_ITERATIONS,
                "normalized_to_checked_iterations": None,
            }
            continue
        row = candidates[0]
        total = row["total"]
        value: dict[str, Any] = {
            "status": "available" if total is not None else "unavailable",
            "scope": "whole-process",
            "counter_scope": WHOLE_PROCESS_SCOPE,
            "raw_value": row["raw_value"],
            "unit": row["unit"],
            "event": row["event"],
            "runtime": row["runtime"],
            "running_percent": row["running_percent"],
            "total": total,
            "normalization_denominator": CHECKED_ITERATIONS,
            "normalized_to_checked_iterations": (total / CHECKED_ITERATIONS) if total is not None else None,
        }
        if total is None:
            value["reason"] = "perf reported a non-numeric count"
        parsed[expected] = value
    available = sum(row.get("status") == "available" for row in parsed.values())
    return {"status": "pass" if available == len(EVENTS) else ("partial" if available else "unavailable"), "rows": rows, "events": parsed}


def parse_top_symbols(path: Path) -> list[dict[str, Any]]:
    overheads: dict[str, float] = {}
    try:
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError:
        return []
    for line in lines:
        match = re.match(r"^\s*(\d+(?:\.\d+)?)%\s+(.*)$", line)
        if not match:
            continue
        percent = float(match.group(1))
        rest = match.group(2).strip()
        marker = re.search(r"\s\[[^\]]+\]\s+", rest)
        if marker:
            symbol = rest[marker.end():].strip()
        else:
            fields = rest.split(None, 1)
            symbol = fields[-1].strip() if fields else ""
        if symbol:
            overheads[symbol] = overheads.get(symbol, 0.0) + percent
    rows = sorted(overheads.items(), key=lambda item: (-item[1], item[0]))[:10]
    return [
        {"rank": index, "symbol": symbol, "overhead_percent": round(percent, 6), "scope": TOP_SYMBOL_SCOPE}
        for index, (symbol, percent) in enumerate(rows, start=1)
    ]


def probe_callgraph(directory: Path, taskset_path: str, perf_path: str, environment: dict[str, str]) -> tuple[dict[str, Any], str | None]:
    directory.mkdir(parents=True, exist_ok=True)
    stdout = directory / "callgraph-help.stdout"
    stderr = directory / "callgraph-help.stderr"
    argv = [taskset_path, "-c", str(CPU), perf_path, "record", "--call-graph", "help"]
    command = run_command(argv, stdout, stderr, environment)
    text = ""
    for path in (stdout, stderr):
        if path.is_file():
            text += path.read_text(encoding="utf-8", errors="replace") + "\n"
    lowered = text.lower()
    selected: str | None = None
    if re.search(r"\bdwarf\b", lowered):
        selected = "dwarf,8192"
    elif re.search(r"\bfp\b", lowered):
        selected = "fp"
    result: dict[str, Any] = {
        "status": "available" if selected else "unavailable",
        "argv": command["argv"],
        "exit_code": command.get("exit_code"),
        "stdout": command["stdout"],
        "stderr": command["stderr"],
        "selected": selected,
        "requested": "dwarf,8192 then fp",
    }
    if command.get("exit_code") != 0 or not selected:
        result["status"] = "unavailable"
        result["reason"] = "perf call-graph help did not advertise dwarf or fp"
        if command.get("exit_code") not in (None, 0):
            result["reason"] = f"perf call-graph help exited {command['exit_code']}"
        selected = None
        result["selected"] = None
    return result, selected


def run_profile(
    kind: str,
    context: dict[str, Any],
    taskset_path: str | None,
    perf_path: str | None,
    callgraph: dict[str, Any] | None,
    callgraph_mode: str | None,
    environment: dict[str, str],
) -> dict[str, Any]:
    directory = OUTPUT_ROOT / kind
    directory.mkdir(exist_ok=True)
    report = directory / "report.json"
    output = directory / "output.pptx"
    stdout = directory / "workload.stdout"
    stderr = directory / "workload.stderr"
    workload = workload_argv(context, report, output)
    record: dict[str, Any] = {
        "schema": "litchi-0464-profile-run-v1",
        "change": 464,
        "kind": kind,
        "status": "unavailable",
        "scope": WHOLE_PROCESS_SCOPE,
        "operation_only_attribution": False,
        "whole_process": True,
        "cpu": CPU,
        "cpu_affinity": CPU,
        "root_only": True,
        "serialized": True,
        "started_utc": now(),
        "workload_argv": workload,
        "workload_cwd": str(REPO),
        "unavailable": [],
    }
    if callgraph is not None:
        record["callgraph_probe"] = callgraph

    if taskset_path is None:
        record["unavailable"].append({"kind": kind, "reason": "taskset is unavailable; CPU 2 could not be bound"})
        record["finished_utc"] = now()
        record["artifacts"] = []
        write_json(directory / "receipt.json", record)
        return record

    if perf_path is None:
        wrapper = [taskset_path, "-c", str(CPU), *workload]
        record["instrumentation"] = "unavailable"
        record["profiler"] = {"status": "unavailable", "reason": "perf is unavailable", "argv": wrapper}
    elif kind == "counters":
        counters_path = directory / "counters.csv"
        wrapper = [taskset_path, "-c", str(CPU), perf_path, "stat", "--no-big-num", "-x,", "-o", str(counters_path), "-e", EVENT_STRING, "--", *workload]
        record["instrumentation"] = "perf_stat"
        record["profiler"] = {"status": "requested", "events": list(EVENTS), "argv": wrapper, "output": path_label(counters_path)}
    else:
        perf_data = directory / "perf.data"
        wrapper = [taskset_path, "-c", str(CPU), perf_path, "record", "--no-buildid-cache", "-F", "99", "-e", "cycles:u"]
        if callgraph_mode:
            wrapper += ["--call-graph", callgraph_mode]
        wrapper += ["-o", str(perf_data), "--", *workload]
        record["instrumentation"] = "perf_record"
        record["profiler"] = {"status": "requested", "event": "cycles:u", "argv": wrapper, "output": path_label(perf_data), "callgraph": callgraph_mode}

    command = run_command(wrapper, stdout, stderr, environment)
    record["command"] = command
    record["argv"] = wrapper
    record["exit_code"] = command.get("exit_code")
    report_value: dict[str, Any] | None = None
    report_error: str | None = None
    try:
        report_value = validate_report(context, report, output)
    except ProfileError as error:
        report_error = str(error)
        record["report_error"] = report_error
    record["report_valid"] = report_value is not None

    if kind == "counters":
        counters_path = directory / "counters.csv"
        if counters_path.is_file():
            record["counters"] = parse_counters(counters_path)
        else:
            record["counters"] = {"status": "unavailable", "reason": "perf stat did not produce counters.csv", "events": {}}
    else:
        perf_data = directory / "perf.data"
        record["perf_data"] = artifact(perf_data) if perf_data.is_file() else {"status": "unavailable", "reason": "perf record did not produce perf.data"}
        record["top_symbols"] = []
        record["postprocess"] = {}
        if perf_data.is_file() and perf_data.stat().st_size:
            top_output = directory / "top-symbols.txt"
            top_error = directory / "top-symbols.stderr"
            top_argv = [perf_path or "perf", "report", "--stdio", "--no-inline", "--no-children", "--call-graph", "none", "--percent-limit", "0", "--sort", "symbol", "-i", str(perf_data)]
            top_command = run_command(top_argv, top_output, top_error, environment)
            record["postprocess"]["perf_report"] = top_command
            if top_output.is_file():
                record["top_symbols"] = parse_top_symbols(top_output)
            script_output = directory / "perf-script.txt"
            script_error = directory / "perf-script.stderr"
            script_argv = [perf_path or "perf", "script", "--no-inline", "-i", str(perf_data)]
            record["postprocess"]["perf_script"] = run_command(script_argv, script_output, script_error, environment)
        else:
            record["postprocess"]["status"] = "unavailable"
            record["postprocess"]["reason"] = "perf.data is unavailable"

    exit_code = command.get("exit_code")
    if report_value is None:
        if exit_code is None:
            record["status"] = "unavailable"
            record["unavailable"].append({"kind": kind, "reason": "workload could not be started"})
        elif perf_path is not None and exit_code != 0:
            record["status"] = "unavailable"
            record["unavailable"].append({"kind": kind, "reason": f"profiler exited {exit_code} before a validated report was produced"})
        else:
            record["status"] = "failed"
            record["unavailable"].append({"kind": kind, "reason": "workload report or output failed validation"})
    elif perf_path is None:
        record["status"] = "unavailable"
        record["unavailable"].append({"kind": kind, "reason": "workload completed, but perf was unavailable"})
    elif exit_code != 0:
        record["status"] = "unavailable"
        record["unavailable"].append({"kind": kind, "reason": f"profiler exited {exit_code}; workload report validated"})
    elif kind == "counters":
        counter_status = record["counters"].get("status")
        record["status"] = "pass" if counter_status == "pass" else ("partial" if counter_status == "partial" else "unavailable")
        if record["status"] != "pass":
            record["unavailable"].append({"kind": kind, "reason": "one or more requested perf stat events were unavailable"})
    else:
        data = record["perf_data"]
        has_data = isinstance(data, dict) and data.get("sha256")
        report_command = record["postprocess"].get("perf_report", {})
        script_command = record["postprocess"].get("perf_script", {})
        postprocess_ok = bool(has_data and report_command.get("exit_code") == 0 and script_command.get("exit_code") == 0)
        record["status"] = "pass" if postprocess_ok else ("partial" if has_data else "unavailable")
        if not postprocess_ok:
            record["unavailable"].append({"kind": kind, "reason": "perf sample data or post-processing is unavailable"})
    if report_value is not None:
        record["report"] = artifact(report)
        record["output"] = artifact(output)
    record["finished_utc"] = now()
    record["artifacts"] = [
        artifact(path)
        for path in sorted(directory.iterdir())
        if path.is_file() and path.name != "receipt.json"
    ]
    write_json(directory / "receipt.json", record)
    return record


def profiling_artifacts() -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if not OUTPUT_ROOT.is_dir():
        return rows
    for path in sorted(OUTPUT_ROOT.rglob("*")):
        if path.is_file() and path.name != SUMMARY_PATH.name:
            rows.append(artifact(path))
    return rows


def environment_guard() -> dict[str, Any]:
    euid = getattr(os, "geteuid", lambda: None)()
    cpu_count = os.cpu_count()
    cpu_available = cpu_count is None or cpu_count > CPU
    process_affinity: list[int] | None = None
    if hasattr(os, "sched_getaffinity"):
        try:
            process_affinity = sorted(os.sched_getaffinity(0))
            cpu_available = cpu_available and CPU in process_affinity
        except OSError:
            pass
    return {
        "root_required": True,
        "effective_uid": euid,
        "root_available": euid == 0,
        "cpu": CPU,
        "logical_cpus": cpu_count,
        "process_affinity": process_affinity,
        "cpu_available": cpu_available,
        "serialized": True,
    }


def main() -> int:
    if OUTPUT_ROOT.exists():
        raise SystemExit(f"refusing to overwrite retained profiling directory: {OUTPUT_ROOT}")
    OUTPUT_ROOT.mkdir()
    summary: dict[str, Any] = {
        "schema": "litchi-0464-profile-summary-v1",
        "change": 464,
        "status": "running",
        "started_utc": now(),
        "scope": WHOLE_PROCESS_SCOPE,
        "counter_scope": "whole-process",
        "top_symbols_scope": TOP_SYMBOL_SCOPE,
        "operation_only_attribution": False,
        "cpu": CPU,
        "serialized": True,
        "provider": PROVIDER,
        "repeat": REPEAT,
        "samples": SAMPLES,
        "warmup": WARMUP,
        "checked_iterations": CHECKED_ITERATIONS,
        "events": list(EVENTS),
        "normalization": {
            "label": "normalized_to_checked_iterations",
            "denominator": CHECKED_ITERATIONS,
            "scope": "whole-process",
            "method": "whole-process counter total divided by 33 checked iterations; descriptive only",
        },
        "runner": {"path": path_label(SCRIPT), "sha256": sha(SCRIPT)},
        "environment_guard": environment_guard(),
        "unavailable": [],
        "runs": [],
        "top_symbols": [],
        "counter_totals": {},
    }
    try:
        context = load_bound_context()
        summary["bound"] = {
            "protocol": artifact(context["protocol_path"]),
            "pair_manifest": artifact(context["pair_path"]),
            "formal_binding": artifact(context["formal_binding_path"]),
            "measured_binding": artifact(context["binding_path"]),
            "binary": artifact(context["binary_path"]),
            "binary_binding": context["binary"],
            "build_receipt": artifact(context["build_path"]),
            "measured_source_manifest": artifact(context["source_path"]),
            "measured_source_binding": context["source"],
            "source_compatibility": artifact(context["compatibility_path"]),
            "source_compatibility_attestation": {
                "measured_source": context["compatibility"]["measured_source"],
                "current_source": context["compatibility"]["current_source"],
                "allowlisted_changed_files": context["compatibility"]["allowlisted_changed_files"],
                "unchanged_files": context["compatibility"]["unchanged_files"],
                "artifact": context["compatibility"]["artifact"],
                "custody_argv": context["compatibility"]["custody_argv"],
            },
            "inventory_binding": artifact(context["inventory_path"]),
            "inventory_binary": artifact(context["inventory"]["binary_path"]),
            "source_revision": context["source_revision"],
            "pair_id": context["pair_id"],
        }
        summary["bound_artifacts"] = bound_artifacts(context)
        summary["bound_artifact_hashes"] = {
            name: row["sha256"] for name, row in summary["bound_artifacts"].items()
        }
        environment, ambient = fixed_environment()
        summary["environment"] = {
            "fixed": {key: environment.get(key) for key in ("DEBUGINFOD_URLS", "PYTHONDONTWRITEBYTECODE", "RUSTUP_TOOLCHAIN")},
            "removed_ambient": ambient,
        }
        taskset_path = shutil.which("taskset")
        perf_path = shutil.which("perf")
        summary["tools"] = {"taskset": taskset_path, "perf": perf_path}
        guard = summary["environment_guard"]
        if not guard["root_available"]:
            summary["unavailable"].append({"kind": "execution", "reason": "helper is root-only"})
        elif not guard["cpu_available"]:
            summary["unavailable"].append({"kind": "execution", "reason": "CPU 2 is unavailable"})
        elif taskset_path is None:
            summary["unavailable"].append({"kind": "execution", "reason": "taskset is unavailable; CPU 2 could not be bound"})
        else:
            callgraph = None
            callgraph_mode = None
            if perf_path is not None:
                callgraph, callgraph_mode = probe_callgraph(OUTPUT_ROOT / "samples", taskset_path, perf_path, environment)
                summary["callgraph"] = callgraph
                if callgraph.get("status") != "available":
                    summary["unavailable"].append({"kind": "callgraph", "reason": callgraph.get("reason", "call-graph support is unavailable")})
            else:
                summary["unavailable"].append({"kind": "perf", "reason": "perf is unavailable; both runs will be unprofiled"})
            for kind in ("counters", "samples"):
                run = run_profile(kind, context, taskset_path, perf_path, callgraph, callgraph_mode, environment)
                summary["runs"].append(run)
                summary["unavailable"].extend(run.get("unavailable", []))
            if len(summary["runs"]) >= 2:
                counter_run = summary["runs"][0]
                summary["counter_totals"] = counter_run.get("counters", {}).get("events", {})
                sample_run = summary["runs"][1]
                summary["top_symbols"] = sample_run.get("top_symbols", [])
        statuses = [run.get("status") for run in summary["runs"]]
        if any(status == "failed" for status in statuses):
            summary["status"] = "failed"
        elif statuses and all(status == "pass" for status in statuses):
            summary["status"] = "pass"
        elif any(status in {"pass", "partial"} for status in statuses):
            summary["status"] = "partial"
        else:
            summary["status"] = "unavailable"
    except ProfileError as error:
        summary["status"] = "failed"
        summary["error"] = str(error)
    finally:
        summary["raw_artifacts"] = profiling_artifacts()
        summary["raw_artifact_hashes"] = {row["path"]: row["sha256"] for row in summary["raw_artifacts"]}
        summary["raw_artifacts_scope"] = "raw report, PPTX, perf CSV/data, command logs, post-processing text, probe output, and run receipts; profile-summary.json excluded to avoid self-reference"
        summary["finished_utc"] = now()
        write_json(SUMMARY_PATH, summary)
    return 0 if summary["status"] in {"pass", "partial", "unavailable"} else 1


if __name__ == "__main__":
    raise SystemExit(main())
