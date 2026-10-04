#!/usr/bin/env python3
"""Fail-closed offline analysis of the 0910 paired PPTX route capture.

The capture has twelve serial process pairs.  Each pair runs the existing
materialized-ReadAt route and the owning-``Vec`` route in alternating AB/BA
order.  This program only reads retained receipts and reports; it never
builds, runs the workload, reopens a package, or starts an external tool.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import random
import statistics
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Mapping, Sequence


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
SCHEMA = "litchi.performance.0910.pptx-owned-input-analysis.v1"
REPORT_SCHEMA = "pptx-owned-input-benchmark-v1"
PARENT = "468fe2931d6532c32f4c6608590816b8eaa88e31"
WARMUP = 30
SAMPLES = 20
BLOCKS = 12
BOOTSTRAP_DRAWS = 10_000
BOOTSTRAP_SEED = 910
CPU = 12
ROUTES = ("materialized", "owned")
PHASES = ("clone_ns", "ingress_ns", "snapshot_ns", "edit_commit_ns", "publish_ns")
ALL_TIMINGS = PHASES + ("total_ns",)
SOURCE_BYTES = 17_017_139
SOURCE_SHA256 = "61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757"
ENTRY_COUNT = 229
MEMBER_COUNT = 445
UNCOMPRESSED_BYTES = 17_568_429
POSITIONS = (0, 28, 57, 85, 114, 142, 171, 199)
REPLACEMENT_COUNT = 64
SINK_MAX_WRITE = 65_536
SOURCE_FILE_COUNT = 9_355


class EvidenceError(ValueError):
    """A retained result is absent, malformed, or contradictory."""


def fail(message: str) -> None:
    raise EvidenceError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path, label: str) -> Any:
    require(path.is_file() and not path.is_symlink(), f"{label}: missing regular JSON file")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{label}: invalid JSON: {error}")


def sha256(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"cannot hash missing regular file: {path}")
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def relative(path: Path, base: Path = HERE) -> str:
    try:
        return str(path.resolve().relative_to(base.resolve()))
    except ValueError as error:
        raise EvidenceError(f"path escapes packet: {path}") from error


def descriptor(path: Path, base: Path = HERE) -> dict[str, Any]:
    path = path.resolve()
    return {"path": relative(path, base), "bytes": path.stat().st_size, "sha256": sha256(path)}


def valid_sha(value: Any, label: str) -> str:
    require(isinstance(value, str) and len(value) == 64, f"{label}: expected SHA-256")
    require(all(char in "0123456789abcdef" for char in value), f"{label}: expected lowercase SHA-256")
    return value


def integer(value: Any, label: str, *, positive: bool = False) -> int:
    require(type(value) is int, f"{label}: expected integer")
    require(value > 0 if positive else value >= 0, f"{label}: expected {'positive' if positive else 'non-negative'} integer")
    return value


def finite_number(value: Any, label: str, *, positive: bool = False) -> float:
    require(isinstance(value, (int, float)) and not isinstance(value, bool), f"{label}: expected number")
    result = float(value)
    require(math.isfinite(result), f"{label}: expected finite number")
    require(result > 0 if positive else result >= 0, f"{label}: expected {'positive' if positive else 'non-negative'} number")
    return result


def resolve_descriptor(raw: Any, label: str, *, expected_name: str | None = None) -> tuple[Path, dict[str, Any]]:
    require(isinstance(raw, dict), f"{label}: descriptor is malformed")
    raw_path = raw.get("path")
    require(isinstance(raw_path, str) and raw_path, f"{label}: descriptor path is missing")
    path = Path(raw_path)
    if not path.is_absolute():
        path = ROOT / path
    path = path.resolve()
    try:
        path.relative_to(HERE.resolve())
    except ValueError as error:
        raise EvidenceError(f"{label}: descriptor escapes packet: {path}") from error
    require(path.is_file() and not path.is_symlink(), f"{label}: described file is missing")
    if expected_name is not None:
        require(path.name == expected_name, f"{label}: expected {expected_name}, got {path.name}")
    expected_bytes = integer(raw.get("bytes"), f"{label}.bytes")
    expected_sha = valid_sha(raw.get("sha256"), f"{label}.sha256")
    actual = descriptor(path)
    require(actual["bytes"] == expected_bytes, f"{label}: byte count changed")
    require(actual["sha256"] == expected_sha, f"{label}: SHA-256 changed")
    return path, actual


def descriptor_shape(raw: Any, label: str) -> dict[str, Any]:
    """Validate a descriptor whose original path may have been cleaned up."""
    require(isinstance(raw, dict), f"{label}: descriptor is malformed")
    path = raw.get("path")
    require(isinstance(path, str) and path, f"{label}: descriptor path is missing")
    require("/" not in path or not path.startswith("//"), f"{label}: descriptor path is malformed")
    return {
        "path": path,
        "bytes": integer(raw.get("bytes"), f"{label}.bytes"),
        "sha256": valid_sha(raw.get("sha256"), f"{label}.sha256"),
    }


def optional_descriptor_file(raw: Mapping[str, Any], label: str) -> None:
    """Check an original file when it remains, without requiring it after cleanup."""
    path = Path(str(raw["path"]))
    if not path.is_absolute():
        path = ROOT / path
    if not path.exists():
        return
    require(path.is_file() and not path.is_symlink(), f"{label}: retained original is not a regular file")
    require(path.stat().st_size == raw["bytes"] and sha256(path) == raw["sha256"], f"{label}: retained original changed")


def check_binary_relocation(packet: Path, build_binary_raw: Any) -> dict[str, Any]:
    relocation = read_json(packet / "binary.json", "binary")
    require(isinstance(relocation, dict), "binary: expected object")
    original = descriptor_shape(relocation.get("original"), "binary.original")
    retained_raw = descriptor_shape(relocation.get("retained"), "binary.retained")
    require(build_binary_raw == original, "binary: build receipt does not bind original descriptor")
    require(original["bytes"] == retained_raw["bytes"] and original["sha256"] == retained_raw["sha256"], "binary: original and retained identities differ")
    retained_path, retained = resolve_descriptor(retained_raw, "binary.retained")
    require(retained["bytes"] == original["bytes"] and retained["sha256"] == original["sha256"], "binary: retained bytes differ")
    optional_descriptor_file(original, "binary.original")
    return {
        "relocation": descriptor(packet / "binary.json"),
        "original": original,
        "retained": retained,
        "retained_path": relative(retained_path),
    }


def summary(values: Sequence[int], label: str) -> dict[str, Any]:
    require(values, f"{label}: no values")
    clean = [integer(value, f"{label}[{index}]") for index, value in enumerate(values)]
    ordered = sorted(clean)
    count = len(ordered)
    return {
        "samples": count,
        "min": ordered[0],
        "p50": (ordered[(count - 1) // 2] + ordered[count // 2]) // 2,
        "p95": ordered[min((95 * count + 99) // 100 - 1, count - 1)],
        "p99": ordered[min((99 * count + 99) // 100 - 1, count - 1)],
        "max": ordered[-1],
        "mean": statistics.fmean(ordered),
    }


def median(values: Sequence[float | int], label: str) -> float:
    require(values, f"{label}: no values")
    return float(statistics.median(values))


def parse_timestamp(value: Any, label: str) -> datetime:
    require(isinstance(value, str) and value, f"{label}: timestamp is missing")
    try:
        result = datetime.fromisoformat(value)
    except ValueError as error:
        raise EvidenceError(f"{label}: invalid ISO timestamp") from error
    require(result.tzinfo is not None, f"{label}: timestamp has no timezone")
    return result


def canonical(value: Any) -> str:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False)


def expected_corpus(corpus: Any, label: str) -> dict[str, Any]:
    require(isinstance(corpus, dict), f"{label}: corpus manifest is missing")
    require(corpus.get("generator") == "litchi-pptx-source-edit-media-v1", f"{label}: generator differs")
    require(corpus.get("package_format") == "PPTX/OPC/ZIP", f"{label}: package format differs")
    require(corpus.get("entry_count") == ENTRY_COUNT, f"{label}: entry count differs")
    require(corpus.get("archive_member_count") == MEMBER_COUNT, f"{label}: member count differs")
    require(corpus.get("uncompressed_payload_bytes") == UNCOMPRESSED_BYTES, f"{label}: logical bytes differ")
    require(corpus.get("shape") == "media-rich", f"{label}: corpus shape differs")
    require(corpus.get("compression") == "deflate", f"{label}: compression differs")
    valid_sha(corpus.get("archive_sha256"), f"{label}.archive_sha256")
    require(corpus["archive_sha256"] == SOURCE_SHA256, f"{label}: source digest differs")
    return corpus


def check_report_scope(report: Mapping[str, Any], label: str) -> None:
    require(report.get("schema") == REPORT_SCHEMA, f"{label}: benchmark schema differs")
    require(report.get("status") == "completed", f"{label}: benchmark did not complete")
    require(report.get("warmup_iterations") == WARMUP, f"{label}: warmup count differs")
    require(report.get("measured_samples") == SAMPLES, f"{label}: measured count differs")
    require(report.get("selected_slide_positions") == list(POSITIONS), f"{label}: selected positions differ")
    require(report.get("replacement_count") == REPLACEMENT_COUNT, f"{label}: replacement count differs")
    for field in ("timing_scope", "phase_scope", "teardown_scope", "verification_scope", "performance_claim"):
        require(isinstance(report.get(field), str) and report[field], f"{label}: {field} is missing")
    timing = " ".join(report[field].lower() for field in ("timing_scope", "phase_scope", "teardown_scope", "verification_scope"))
    for marker in ("clone", "ingress", "snapshot", "commit", "publish", "to_bytes", "drop"):
        require(marker in timing, f"{label}: timing scope omits {marker}")
    require("none" in report["performance_claim"].lower(), f"{label}: performance claim is not bounded")
    gates = report.get("gates")
    require(isinstance(gates, dict), f"{label}: gates are missing")
    for key in ("source_identity_verified", "source_shape_verified", "semantic_output_verified", "stable_output_verified", "source_immutability_verified"):
        require(gates.get(key) is True, f"{label}: gate {key} did not pass")
    require(gates.get("sink_verified") is True, f"{label}: sink gate did not pass")


def parse_process_metrics(path: Path, label: str) -> dict[str, int | float]:
    values: dict[str, int | float] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        require("=" in line, f"{label}: malformed process metric line")
        key, raw = line.split("=", 1)
        require(key and key not in values, f"{label}: duplicate process metric {key}")
        if key == "max_rss_kib" or key.endswith("faults") or key.endswith("context_switches"):
            try:
                parsed = int(raw)
            except ValueError as error:
                raise EvidenceError(f"{label}.{key}: expected integer") from error
            values[key] = integer(parsed, f"{label}.{key}", positive=key == "max_rss_kib")
        elif key in ("user_seconds", "system_seconds", "elapsed_seconds"):
            try:
                parsed = float(raw)
            except ValueError as error:
                raise EvidenceError(f"{label}.{key}: expected number") from error
            values[key] = finite_number(parsed, f"{label}.{key}")
        else:
            fail(f"{label}: unexpected process metric {key}")
    for key in ("max_rss_kib", "user_seconds", "system_seconds", "elapsed_seconds", "major_faults", "minor_faults", "voluntary_context_switches", "involuntary_context_switches"):
        require(key in values, f"{label}: missing process metric {key}")
    return values


def check_sample(sample: Any, label: str, expected_output: Mapping[str, Any]) -> dict[str, Any]:
    require(isinstance(sample, dict), f"{label}: sample is malformed")
    index = integer(sample.get("index"), f"{label}.index")
    require(0 <= index < SAMPLES, f"{label}.index: outside measured range")
    normalized: dict[str, Any] = {"index": index}
    for field in ALL_TIMINGS:
        normalized[field] = integer(sample.get(field), f"{label}.{field}", positive=True)
    require(normalized["total_ns"] >= sum(normalized[field] for field in PHASES), f"{label}: total is below phase sum")
    digest = valid_sha(sample.get("output_sha256"), f"{label}.output_sha256")
    require(digest == expected_output["sha256"], f"{label}: sample output digest differs")
    output_bytes = integer(sample.get("output_bytes"), f"{label}.output_bytes", positive=True)
    require(output_bytes == expected_output["bytes"], f"{label}: sample output size differs")
    normalized["output_sha256"] = digest
    normalized["output_bytes"] = output_bytes
    return normalized


def validate_plan(packet: Path) -> tuple[dict[str, Any], list[dict[str, Any]], str]:
    plan = read_json(packet / "capture-plan.json", "capture-plan")
    require(isinstance(plan, dict), "capture-plan: expected object")
    require(plan.get("status") == "frozen_before_capture", "capture-plan: status is not frozen_before_capture")
    require(plan.get("warmup") == WARMUP and plan.get("samples") == SAMPLES, "capture-plan: sample schedule differs")
    require(plan.get("blocks") == BLOCKS and plan.get("cpu") == CPU, "capture-plan: block or CPU binding differs")
    require(plan.get("parent") == PARENT, "capture-plan: parent binding differs")
    legs = plan.get("legs")
    require(isinstance(legs, list) and len(legs) == BLOCKS * 2, "capture-plan: expected 24 legs")
    labels: set[str] = set()
    by_block: dict[int, list[dict[str, Any]]] = {index: [] for index in range(BLOCKS)}
    for index, raw in enumerate(legs):
        require(isinstance(raw, dict), f"capture-plan.legs[{index}]: malformed")
        label = raw.get("label")
        route = raw.get("route")
        block = raw.get("block")
        position = raw.get("position")
        require(isinstance(label, str) and label and label not in labels, f"capture-plan.legs[{index}]: duplicate label")
        require(route in ROUTES, f"capture-plan.legs[{index}]: unexpected route")
        require(type(block) is int and 0 <= block < BLOCKS, f"capture-plan.legs[{index}]: unexpected block")
        require(position in (0, 1), f"capture-plan.legs[{index}]: unexpected position")
        labels.add(label)
        by_block[block].append(raw)
    for block, pair in by_block.items():
        require(len(pair) == 2, f"capture-plan block {block}: expected two legs")
        require(sorted(item["position"] for item in pair) == [0, 1], f"capture-plan block {block}: positions differ")
        require({item["route"] for item in pair} == set(ROUTES), f"capture-plan block {block}: routes differ")
        ordered = [item["route"] for item in sorted(pair, key=lambda item: item["position"])]
        expected = ["materialized", "owned"] if block % 2 == 0 else ["owned", "materialized"]
        require(ordered == expected, f"capture-plan block {block}: AB/BA order differs")
    attempt = plan.get("attempt", "v1")
    require(isinstance(attempt, str) and attempt.replace("-", "").isalnum(), "capture-plan: invalid attempt")
    return plan, legs, attempt


def check_command_receipt(
    packet: Path,
    plan: Mapping[str, Any],
    leg: Mapping[str, Any],
    attempt: str,
    binary_original: Mapping[str, Any],
    expected_source: Mapping[str, Any],
    expected_helper: Mapping[str, Any],
    expected_common: Mapping[str, Any],
) -> dict[str, Any]:
    label = leg["label"]
    folder = packet / "capture" / label
    receipt = read_json(folder / "command.json", f"{label}.command")
    require(isinstance(receipt, dict), f"{label}: command receipt is malformed")
    require(receipt.get("exit_code") == 0, f"{label}: process did not exit zero")
    require(receipt.get("leg") == leg, f"{label}: receipt leg differs from plan")
    require(receipt.get("plan") is not None, f"{label}: plan descriptor is missing")
    _, plan_descriptor = resolve_descriptor(receipt["plan"], f"{label}.plan", expected_name="capture-plan.json")
    require(plan_descriptor == descriptor(packet / "capture-plan.json"), f"{label}: plan binding differs")
    for field in ("source", "helper"):
        require(field in receipt, f"{label}: {field} descriptor is missing")
    _, source_descriptor = resolve_descriptor(receipt["source"], f"{label}.source")
    _, helper_descriptor = resolve_descriptor(receipt["helper"], f"{label}.helper")
    require(source_descriptor == expected_source, f"{label}: source manifest binding differs")
    require(helper_descriptor == expected_helper, f"{label}: helper binding differs")
    for field, expected in expected_common.items():
        require(receipt.get(field) == expected, f"{label}: common receipt field {field} differs")
    require(receipt.get("binary") == binary_original, f"{label}: binary does not bind original descriptor")
    descriptor_shape(receipt.get("binary"), f"{label}.binary")
    require(receipt["source"]["path"].endswith(f"source-{attempt}.json"), f"{label}: source attempt differs")
    require(receipt["helper"]["path"].endswith("run.py"), f"{label}: helper is not run.py")
    require(isinstance(receipt.get("argv"), list) and all(isinstance(item, str) for item in receipt["argv"]), f"{label}: argv malformed")
    argv = receipt["argv"]
    require("taskset" in argv and "-c" in argv and str(CPU) in argv, f"{label}: CPU pinning missing")
    require("/usr/bin/time" in argv, f"{label}: process timer missing")
    require(len(argv) >= 5 and argv[-4:] == [leg["route"], str(WARMUP), str(SAMPLES), str(folder / "artifacts")], f"{label}: route schedule argv differs")
    require(receipt.get("start_utc") and receipt.get("end_utc"), f"{label}: receipt timestamps missing")
    started = parse_timestamp(receipt["start_utc"], f"{label}.start_utc")
    ended = parse_timestamp(receipt["end_utc"], f"{label}.end_utc")
    require(ended >= started, f"{label}: receipt ends before it starts")
    logs = receipt.get("logs")
    require(isinstance(logs, list) and len(logs) == 2, f"{label}: expected stdout/stderr descriptors")
    for row in logs:
        resolve_descriptor(row, f"{label}.log")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, list), f"{label}: artifacts are missing")
    artifact_by_name: dict[str, tuple[Path, dict[str, Any]]] = {}
    for row in artifacts:
        path, checked = resolve_descriptor(row, f"{label}.artifact")
        require(path.parent == folder / "artifacts", f"{label}: artifact escapes artifacts directory")
        require(path.name not in artifact_by_name, f"{label}: duplicate artifact {path.name}")
        artifact_by_name[path.name] = (path, checked)
    for name in ("source.pptx", "output.pptx", "report.json"):
        require(name in artifact_by_name, f"{label}: missing {name}")
    process_raw = receipt.get("process_metrics")
    process_path, process_descriptor = resolve_descriptor(process_raw, f"{label}.process_metrics", expected_name="process.txt")
    require(process_path.parent == folder, f"{label}: process metrics path differs")
    metrics = parse_process_metrics(process_path, f"{label}.process_metrics")
    report_path = artifact_by_name["report.json"][0]
    report = read_json(report_path, f"{label}.report")
    require(isinstance(report, dict), f"{label}: report is malformed")
    require(report.get("route") == leg["route"], f"{label}: report route differs")
    check_report_scope(report, label)
    expected_corpus(report.get("corpus"), f"{label}.corpus")
    _, source_artifact = artifact_by_name["source.pptx"]
    _, output_artifact = artifact_by_name["output.pptx"]
    require(report.get("source") == {"file": "source.pptx", "bytes": source_artifact["bytes"], "sha256": source_artifact["sha256"]}, f"{label}: source artifact binding differs")
    require(report.get("output") == {"file": "output.pptx", "bytes": output_artifact["bytes"], "sha256": output_artifact["sha256"]}, f"{label}: output artifact binding differs")
    require(source_artifact["bytes"] == SOURCE_BYTES and source_artifact["sha256"] == SOURCE_SHA256, f"{label}: source archive identity differs")
    require(output_artifact["bytes"] > 0 and output_artifact["sha256"] != source_artifact["sha256"], f"{label}: output artifact is not a changed archive")
    sink_capacity = integer(report.get("sink_capacity_bytes"), f"{label}.sink_capacity_bytes", positive=True)
    sink_limit = integer(report.get("sink_max_write_bytes"), f"{label}.sink_max_write_bytes", positive=True)
    require(sink_limit == SINK_MAX_WRITE, f"{label}: sink write limit differs")
    require(sink_capacity == 2 * SOURCE_BYTES + SINK_MAX_WRITE, f"{label}: sink capacity differs")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == SAMPLES, f"{label}: measured sample cardinality differs")
    checked_samples = [check_sample(item, f"{label}.samples[{index}]", output_artifact) for index, item in enumerate(samples)]
    require([item["index"] for item in checked_samples] == list(range(SAMPLES)), f"{label}: sample indices are not ordered 0..19")
    for index, sample in enumerate(samples):
        calls = integer(sample.get("sink_write_calls"), f"{label}.samples[{index}].sink_write_calls", positive=True)
        largest = integer(sample.get("sink_largest_write"), f"{label}.samples[{index}].sink_largest_write", positive=True)
        expected_calls = (output_artifact["bytes"] + SINK_MAX_WRITE - 1) // SINK_MAX_WRITE
        expected_largest = min(output_artifact["bytes"], SINK_MAX_WRITE)
        require(calls == expected_calls, f"{label}.samples[{index}]: sink call count differs")
        require(largest == expected_largest, f"{label}.samples[{index}]: largest sink write differs")
        checked_samples[index]["sink_write_calls"] = calls
        checked_samples[index]["sink_largest_write"] = largest
    command_descriptor = descriptor(folder / "command.json")
    return {
        "label": label,
        "block": leg["block"],
        "route": leg["route"],
        "position": leg["position"],
        "folder": relative(folder),
        "command": command_descriptor,
        "bindings": {
            "source": source_descriptor,
            "helper": helper_descriptor,
            "plan": plan_descriptor,
            "binary_original": dict(binary_original),
        },
        "source": source_artifact,
        "output": output_artifact,
        "report": descriptor(report_path),
        "process_metrics": process_descriptor,
        "rss": metrics,
        "report_scope": {key: report[key] for key in ("timing_scope", "phase_scope", "teardown_scope", "verification_scope", "performance_claim")},
        "sink": {"capacity_bytes": sink_capacity, "max_write_bytes": sink_limit},
        "samples": checked_samples,
        "corpus": report["corpus"],
        "started": started,
        "ended": ended,
    }


def check_build_and_quality(packet: Path, attempt: str, plan: Mapping[str, Any]) -> dict[str, Any]:
    build = read_json(packet / f"build-{attempt}" / "command.json", f"build-{attempt}.command")
    require(isinstance(build, dict) and build.get("exit_code") == 0, f"build-{attempt}: build did not pass")
    binary_raw = descriptor_shape(build.get("binary"), f"build-{attempt}.binary")
    require(Path(binary_raw["path"]).name == "pptx_owned_benchmark", f"build-{attempt}: unexpected binary")
    relocation = check_binary_relocation(packet, binary_raw)
    source_path, build_source = resolve_descriptor(build.get("source"), f"build-{attempt}.source")
    source_manifest = read_json(source_path, f"source-{attempt}")
    require(isinstance(source_manifest, dict), f"source-{attempt}: expected object")
    require(source_manifest.get("parent") == PARENT, f"source-{attempt}: parent binding differs")
    files = source_manifest.get("files")
    require(isinstance(files, list) and len(files) == SOURCE_FILE_COUNT, f"source-{attempt}: source inventory cardinality differs")
    paths: set[str] = set()
    for index, row in enumerate(files):
        require(isinstance(row, dict), f"source-{attempt}.files[{index}]: malformed")
        path_value = row.get("path")
        require(isinstance(path_value, str) and path_value and path_value not in paths, f"source-{attempt}.files[{index}]: duplicate path")
        paths.add(path_value)
        integer(row.get("bytes"), f"source-{attempt}.files[{index}].bytes")
        valid_sha(row.get("sha256"), f"source-{attempt}.files[{index}].sha256")
    base = descriptor_shape(source_manifest.get("base"), f"source-{attempt}.base")
    source_base_actual = descriptor(packet / "source-base.json")
    require((base["bytes"], base["sha256"]) == (source_base_actual["bytes"], source_base_actual["sha256"]), f"source-{attempt}: base descriptor differs")
    resolve_descriptor(source_manifest.get("overlay_patch"), f"source-{attempt}.overlay_patch")
    expected_source = descriptor(packet / f"source-{attempt}.json")
    require(build_source == expected_source, f"build-{attempt}: source manifest binding differs")
    build_helper = resolve_descriptor(build.get("helper"), f"build-{attempt}.helper")[1]
    expected_helper = descriptor(packet / "run.py")
    require(build_helper == expected_helper, f"build-{attempt}: helper binding differs")
    common_fields = ("cwd", "environment_overrides", "removed_environment_keys", "toolchain")
    for field in common_fields:
        require(field in build, f"build-{attempt}: common receipt field {field} is missing")
    build_common = {field: build[field] for field in common_fields}
    quality: dict[str, Any] = {}
    quality_dir = packet / f"quality-{attempt}"
    require(quality_dir.is_dir(), f"quality-{attempt}: missing quality directory")
    commands = sorted(quality_dir.rglob("command.json"))
    require({path.parent.name for path in commands} == {"fmt", "clippy"}, f"quality-{attempt}: expected fmt and clippy receipts")
    for path in commands:
        item = read_json(path, str(path.relative_to(packet)))
        require(item.get("exit_code") == 0, f"{path}: quality command did not pass")
        for field in ("source", "helper"):
            checked = resolve_descriptor(item.get(field), f"{path.name}.{field}")[1]
            require(checked == (build_source if field == "source" else build_helper), f"{path}: {field} binding differs from build")
        for field, expected in build_common.items():
            require(item.get(field) == expected, f"{path}: common receipt field {field} differs from build")
        quality[path.parent.name] = descriptor(path)
    return {
        "attempt": attempt,
        "build": descriptor(packet / f"build-{attempt}" / "command.json"),
        "binary": relocation["retained"],
        "binary_path": relocation["retained_path"],
        "binary_relocation": relocation,
        "bindings": {"source": build_source, "helper": build_helper},
        "common": build_common,
        "source_manifest": {"descriptor": build_source, "files": SOURCE_FILE_COUNT},
        "quality": quality,
    }


def ratio(numerator: float, denominator: float, label: str) -> float:
    require(math.isfinite(numerator) and math.isfinite(denominator) and denominator > 0, f"{label}: invalid ratio operands")
    result = numerator / denominator
    require(math.isfinite(result) and result > 0, f"{label}: invalid ratio")
    return result


def percentile(values: Sequence[float], percent: float, label: str) -> float:
    require(values, f"{label}: no values")
    ordered = sorted(values)
    index = min(max(math.ceil(percent * len(ordered)) - 1, 0), len(ordered) - 1)
    return ordered[index]


def bootstrap_median(values: Sequence[float], label: str) -> dict[str, Any]:
    require(len(values) == BLOCKS, f"{label}: expected {BLOCKS} paired block ratios")
    generator = random.Random(BOOTSTRAP_SEED)
    draws = []
    for _ in range(BOOTSTRAP_DRAWS):
        sample = [values[generator.randrange(BLOCKS)] for _ in range(BLOCKS)]
        draws.append(float(statistics.median(sample)))
    return {
        "statistic": float(statistics.median(values)),
        "draws": BOOTSTRAP_DRAWS,
        "seed": BOOTSTRAP_SEED,
        "resampling_unit": "paired block (both routes retained together)",
        "confidence": 0.95,
        "method": "nearest-rank percentile interval over bootstrap medians",
        "lower": percentile(draws, 0.025, f"{label}.lower"),
        "upper": percentile(draws, 0.975, f"{label}.upper"),
    }


def route_summary(legs: Sequence[Mapping[str, Any]], route: str) -> dict[str, Any]:
    selected = [leg for leg in legs if leg["route"] == route]
    require(len(selected) == BLOCKS, f"{route}: expected one leg per block")
    out: dict[str, Any] = {"legs": [leg["label"] for leg in selected], "samples": BLOCKS * SAMPLES}
    for field in ALL_TIMINGS:
        values = [sample[field] for leg in selected for sample in leg["samples"]]
        out[field] = summary(values, f"{route}.{field}")
    rss = [int(leg["rss"]["max_rss_kib"]) for leg in selected]
    out["max_rss_kib"] = summary(rss, f"{route}.max_rss_kib")
    return out


def paired_analysis(legs: Sequence[Mapping[str, Any]]) -> dict[str, Any]:
    by_block: dict[int, dict[str, Mapping[str, Any]]] = {index: {} for index in range(BLOCKS)}
    for leg in legs:
        by_block[leg["block"]][leg["route"]] = leg
    rows = []
    ratios: dict[str, list[float]] = {field: [] for field in ALL_TIMINGS}
    for block in range(BLOCKS):
        pair = by_block[block]
        materialized = pair["materialized"]
        owned = pair["owned"]
        row: dict[str, Any] = {"block": block, "order": "AB" if materialized["position"] == 0 else "BA", "materialized": materialized["label"], "owned": owned["label"]}
        for field in ALL_TIMINGS:
            materialized_median = median([sample[field] for sample in materialized["samples"]], f"block {block}.materialized.{field}")
            owned_median = median([sample[field] for sample in owned["samples"]], f"block {block}.owned.{field}")
            value = ratio(materialized_median, owned_median, f"block {block}.{field}")
            row[field] = {"materialized_median": materialized_median, "owned_median": owned_median, "materialized_over_owned": value}
            ratios[field].append(value)
        row["materialized_max_rss_kib"] = materialized["rss"]["max_rss_kib"]
        row["owned_max_rss_kib"] = owned["rss"]["max_rss_kib"]
        rows.append(row)
    return {"blocks": rows, "median_ratio_bootstrap": {field: bootstrap_median(values, f"paired.{field}") for field, values in ratios.items()}}


def drift_analysis(legs: Sequence[Mapping[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for route in ROUTES:
        result[route] = {}
        for order in ("AB", "BA"):
            selected = [leg for leg in legs if leg["route"] == route and ((leg["block"] % 2 == 0) == (order == "AB"))]
            require(len(selected) == BLOCKS // 2, f"{route}/{order}: expected six legs")
            result[route][order] = {field: summary([sample[field] for leg in selected for sample in leg["samples"]], f"drift.{route}.{order}.{field}") for field in ALL_TIMINGS}
        result[route]["BA_over_AB_p50"] = {field: ratio(result[route]["BA"][field]["p50"], result[route]["AB"][field]["p50"], f"drift.{route}.{field}") for field in ALL_TIMINGS}
    return result


def analyze(packet: Path) -> dict[str, Any]:
    packet = packet.resolve()
    require(packet == HERE.resolve(), "analysis packet must be the change-0910 result directory")
    plan, plan_legs, attempt = validate_plan(packet)
    source_base = read_json(packet / "source-base.json", "source-base")
    require(isinstance(source_base, dict) and source_base.get("commit") == PARENT, "source-base: parent binding differs")
    build = check_build_and_quality(packet, attempt, plan)
    binary_original = build["binary_relocation"]["original"]
    parsed = [
        check_command_receipt(
            packet,
            plan,
            leg,
            attempt,
            binary_original,
            build["bindings"]["source"],
            build["bindings"]["helper"],
            build["common"],
        )
        for leg in plan_legs
    ]
    require(len(parsed) == BLOCKS * 2, "capture: wrong leg cardinality")
    require({leg["label"] for leg in parsed} == {leg["label"] for leg in plan_legs}, "capture: labels differ from plan")
    # Every receipt must bind the same source, binary, helper, and plan files.
    for field in ("source", "output"):
        values = [(leg[field]["bytes"], leg[field]["sha256"]) for leg in parsed]
        require(len(set(values)) == 1, f"capture: {field} differs across legs")
    corpus_values = [canonical(leg["corpus"]) for leg in parsed]
    require(len(set(corpus_values)) == 1, "capture: corpus differs across legs")
    binary_values = []
    for leg in plan_legs:
        raw_receipt = read_json(packet / "capture" / leg["label"] / "command.json", leg["label"])
        checked_binary = descriptor_shape(raw_receipt.get("binary"), leg["label"] + ".binary")
        binary_values.append((checked_binary["bytes"], checked_binary["sha256"]))
    build_binary = (build["binary"]["bytes"], build["binary"]["sha256"])
    require(len(set(binary_values)) == 1 and binary_values[0] == build_binary, "capture: binary differs from build")
    # Receipts are serial children: their process windows must not overlap and
    # must follow the explicit AB/BA plan order.
    previous_end: datetime | None = None
    for leg in parsed:
        if previous_end is not None:
            require(leg["started"] >= previous_end, f"{leg['label']}: process overlaps previous leg")
        previous_end = leg["ended"]
    source = parsed[0]["source"]
    output = parsed[0]["output"]
    require(output["sha256"] != source["sha256"], "capture: output is identical to source")
    routes = {route: route_summary(parsed, route) for route in ROUTES}
    paired = paired_analysis(parsed)
    drift = drift_analysis(parsed)
    rss = {
        "scope": "whole-process maximum RSS includes corpus generation, warmups, semantic guards, and all measured samples",
        "by_route": {route: routes[route]["max_rss_kib"] for route in ROUTES},
        "by_leg": {leg["label"]: leg["rss"]["max_rss_kib"] for leg in parsed},
    }
    return {
        "schema": SCHEMA,
        "status": "qualified_generated_fixture",
        "production_change": False,
        "performance_claim": False,
        "scope": {
            "experiment": "same-binary existing materialized route versus owning Package::from_vec control",
            "warmup_iterations_per_process": WARMUP,
            "measured_samples_per_process": SAMPLES,
            "paired_blocks": BLOCKS,
            "routes": list(ROUTES),
            "cpu": CPU,
            "timed_phases": list(ALL_TIMINGS),
            "comparison": "descriptive paired phase evidence; baseline is materialized and denominator is owned",
            "patch_replay_in_primary_timer": False,
        },
        "custody": {
            "capture_plan": descriptor(packet / "capture-plan.json"),
            "source_base": descriptor(packet / "source-base.json"),
            "build": build,
            "plan": {"status": plan["status"], "parent": plan["parent"], "blocks": BLOCKS, "warmup": WARMUP, "samples": SAMPLES, "cpu": CPU},
            "leg_count": len(parsed),
            "serial_process_windows": True,
        },
        "corpus": {"manifest": parsed[0]["corpus"], "source": source, "same_across_legs": True},
        "output": {"artifact": output, "same_across_legs": True, "per_leg": {leg["label"]: leg["output"] for leg in parsed}},
        "routes": routes,
        "paired": paired,
        "drift_by_order": drift,
        "rss": rss,
        "legs": [{key: value for key, value in leg.items() if key not in ("started", "ended", "corpus")} for leg in parsed],
        "checks": {
            "plan_status_and_schedule": True,
            "twelve_ab_ba_pairs": True,
            "twenty_samples_per_route_process": True,
            "same_binary_bound_to_build_and_capture": True,
            "source_identity_bound": True,
            "corpus_manifest_equal_and_expected": True,
            "output_identity_equal_and_semantically_verified": True,
            "reported_sample_vectors_recomputed": True,
            "phase_sum_not_above_total": True,
            "serial_nonoverlap": True,
            "paired_bootstrap_seed_910_10000_draws": True,
            "rss_includes_generation_and_oracle_scope": True,
        },
        "limitations": [
            "Pooled quantiles, paired ratios, order drift, and RSS summaries are descriptive evidence for this generated fixture and shared host.",
            "The benchmark measures existing routes; it does not implement or recommend a production provenance bridge.",
            "No allocation, instruction, compression, cold-cache, range-source, concurrency, producer, signed, encrypted, malformed, or physical-package generalization claim is made.",
            "Maximum RSS is a whole-process maximum and is not operation peak live bytes.",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True, help="new analysis JSON path")
    args = parser.parse_args()
    output = args.output.resolve()
    try:
        require(not output.exists() and not output.is_symlink(), f"refusing to overwrite analysis output: {output}")
        result = analyze(HERE)
        output.parent.mkdir(parents=True, exist_ok=True)
        with output.open("x", encoding="utf-8") as stream:
            json.dump(result, stream, indent=2, sort_keys=True)
            stream.write("\n")
    except (EvidenceError, OSError) as error:
        print(f"analyze.py: error: {error}", file=sys.stderr)
        return 2
    print("0910 paired PPTX route analysis passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
