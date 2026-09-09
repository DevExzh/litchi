#!/usr/bin/env python3
"""Validate and summarize the frozen 0484 route evidence bundle.

This is an analysis-only consumer of the route protocol and retained formal
receipts.  It deliberately does not discover or accept pilot runs, infer a
missing process, or turn two repeats into a confidence interval.  Every
accepted process remains attributable to one route/axis, workload, role, and
repeat; source and authored throughput are kept as separate metrics.

The route and axis capture handlers already run the detailed report shell.
This module rechecks receipt custody, protocol/build/gate bindings, report
identity, sample cardinality, scalar metric shape, source equality, and the
cross-role corpus identity before calculating statistics.  A changed report,
resource file, receipt, build, protocol, or preparation receipt fails closed.
"""

from __future__ import annotations

import argparse
import datetime
import hashlib
import json
import math
import os
from pathlib import Path
import re
import statistics
from typing import Any, Iterable, Mapping

import measure_routes as routes


ROOT = Path(__file__).resolve().parent
PROTOCOL_FILE = routes.ROUTE_PROTOCOL_FILE
SUMMARY_SCHEMA = "docx-replayable-tail-append-route-analysis-v1"
EXECUTION_INPUTS_FILE = "execution-inputs.json"
REVIEW_THRESHOLD_PERCENT = 5.0
FORMAL_SAMPLES = routes.FORMAL_SAMPLES
FORMAL_WARMUPS = routes.FORMAL_WARMUPS
ROLES = tuple(routes.ROLES)
REPEATS = tuple(routes.REPEATS)
SHA256 = re.compile(r"^[0-9a-f]{64}$")

PROCESS_METRICS = (
    "elapsed_ns",
    "source_throughput_bytes_per_second",
    "authored_throughput_bytes_per_second",
    "candidate_throughput_bytes_per_second",
    "process_rss_delta_bytes",
    "process_peak_rss_bytes",
    "time_max_rss_bytes",
)
ALLOCATOR_METRICS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
    "region_peak_live_bytes",
    # This is the operation-scoped peak above the live baseline at entry.
    # Keep it as a per-sample value so quantiles describe the increments,
    # rather than subtracting quantiles of two different vectors later.
    "operation_peak_increment_bytes",
)
HISTOGRAM_FIELDS = tuple(routes.base.HISTOGRAM_FIELDS)


class AnalysisError(ValueError):
    """The retained evidence is absent, malformed, or not bound."""


def fail(message: str) -> None:
    raise AnalysisError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    value: dict[str, Any] = {}
    for key, item in pairs:
        if key in value:
            raise ValueError(f"duplicate JSON key {key!r}")
        value[key] = item
    return value


def read_json(path: Path) -> Any:
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=_pairs,
            parse_constant=lambda value: (_ for _ in ()).throw(ValueError(value)),
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as error:
        raise AnalysisError(f"cannot read {path}: {error}") from error


def write_json(path: Path, value: Any) -> None:
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
        stream.write("\n")


def sha(path: Path) -> str:
    try:
        with path.open("rb") as stream:
            return hashlib.file_digest(stream, "sha256").hexdigest()
    except (OSError, ValueError) as error:
        raise AnalysisError(f"cannot hash {path}: {error}") from error


def metadata(path: Path) -> dict[str, int | str]:
    try:
        return {"bytes": path.stat().st_size, "sha256": sha(path)}
    except OSError as error:
        raise AnalysisError(f"cannot stat {path}: {error}") from error


def integer(value: Any, label: str, *, positive: bool = False) -> int:
    require(type(value) is int, f"{label}: expected an integer")
    require(value > 0 if positive else value >= 0, f"{label}: expected a {'positive' if positive else 'non-negative'} integer")
    return value


def digest(value: Any, label: str) -> str:
    require(isinstance(value, str) and SHA256.fullmatch(value) is not None, f"{label}: malformed SHA-256")
    return value


def finite_number(value: Any, label: str) -> float:
    require(isinstance(value, (int, float)) and not isinstance(value, bool), f"{label}: expected a number")
    result = float(value)
    require(math.isfinite(result), f"{label}: non-finite value")
    return result


def object_value(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: expected an object")
    return value


def validate_protocol(path: Path) -> tuple[dict[str, Any], str]:
    value = object_value(read_json(path), str(path))
    require(value.get("schema") == routes.ROUTE_SCHEMA, "route protocol schema differs")
    require(value.get("version") == routes.ROUTE_PROTOCOL_VERSION, "route protocol version differs")
    require(value.get("claim_authorized") is False, "route protocol must remain claim-disabled")
    require(value.get("performance_claim") == "none", "route protocol performance claim differs")
    formal = object_value(value.get("formal"), "protocol.formal")
    require(formal.get("samples") == FORMAL_SAMPLES, "protocol formal sample count differs")
    require(formal.get("warmups") == FORMAL_WARMUPS, "protocol formal warmup count differs")
    require(value.get("roles") == list(ROLES), "protocol role matrix differs")
    require(value.get("repeats") == list(REPEATS), "protocol repeat matrix differs")
    expected_routes = routes._run_inventory(pilot=False)
    expected_axes = routes._axis_inventory(pilot=False)
    require(value.get("formal_runs") == expected_routes, "protocol route formal inventory differs")
    require(value.get("axis_formal_runs") == expected_axes, "protocol axis formal inventory differs")
    require(value.get("expected_formal_processes") == len(expected_routes), "protocol route process count differs")
    require(value.get("expected_axis_formal_processes") == len(expected_axes), "protocol axis process count differs")
    require(len(expected_routes) == 120 and len(expected_axes) == 108, "protocol formal matrix is not the selected 120/108 inventory")
    require(isinstance(value.get("frozen_utc"), str) and value["frozen_utc"], "protocol freeze timestamp is missing")
    machine = object_value(value.get("machine"), "protocol.machine")
    require(machine.get("status") == "ready", "protocol machine inventory is not ready")
    digest(machine.get("sha256"), "protocol.machine.sha256")
    machine_path = path.parent / str(machine.get("path", ""))
    require(machine_path.is_file() and sha(machine_path) == machine["sha256"], "protocol machine inventory is missing or changed")
    try:
        routes._validate_machine_record(read_json(machine_path))
    except routes.base.MeasureError as error:
        fail(f"protocol machine inventory failed canonical validation: {error}")
    require(value.get("environment") == {key: routes.base.ENV[key] for key in routes.base.ENV_KEYS}, "protocol environment differs from current fixed environment")
    scripts = object_value(value.get("scripts"), "protocol.scripts")
    require(scripts, "protocol script bindings are missing")
    for name, bound in scripts.items():
        digest(bound, f"protocol.scripts.{name}")
    try:
        expected_scripts = routes._script_hashes()
    except routes.base.MeasureError as error:
        fail(f"cannot compute canonical route script bindings: {error}")
    require(scripts == expected_scripts, "protocol script bindings differ from current canonical route helpers")
    protocol_hash = sha(path)
    return value, protocol_hash


def _artifact(receipt_dir: Path, receipt: Mapping[str, Any], name: str) -> Path:
    artifacts = object_value(receipt.get("artifacts"), f"{receipt_dir}.artifacts")
    require(set(artifacts) == {"report.json", "resource.txt", "stdout.txt", "stderr.txt"}, f"{receipt_dir}: artifact inventory differs")
    record = object_value(artifacts.get(name), f"{receipt_dir}.artifacts.{name}")
    expected_bytes = integer(record.get("bytes"), f"{receipt_dir}.artifacts.{name}.bytes")
    expected_sha = digest(record.get("sha256"), f"{receipt_dir}.artifacts.{name}.sha256")
    path = receipt_dir / name
    require(path.is_file() and not path.is_symlink(), f"{receipt_dir}: retained {name} is missing")
    actual = metadata(path)
    require(actual == {"bytes": expected_bytes, "sha256": expected_sha}, f"{receipt_dir}: {name} artifact hash/length differs")
    return path


def _validate_gate(path: Path, attempt: str, expected_source: Mapping[str, Any]) -> str:
    value = object_value(read_json(path), str(path))
    require(value.get("schema") == "docx-stream-append-gate-v1", f"{path}: gate schema differs")
    require(value.get("attempt") == attempt and value.get("exit_code") == 0, f"{path}: gate did not pass")
    require(value.get("source_unchanged") is True, f"{path}: source custody failed")
    require(value.get("source_before") == value.get("source_after") == dict(expected_source), f"{path}: gate source differs from formal builds")
    return sha(path)


def _validate_build(
    bundle_root: Path,
    attempt: str,
    role: str,
    protocol_hash: str,
    expected_environment: Mapping[str, Any],
    expected_source: Mapping[str, Any] | None,
) -> tuple[dict[str, Any], dict[str, Any]]:
    path = bundle_root / "route-attempts" / attempt / f"build-{role}.json"
    require(path.is_file(), f"{role}: build receipt is missing")
    value = object_value(read_json(path), str(path))
    require(value.get("schema") == routes.base.BUILD_SCHEMA and value.get("version") == 1, f"{path}: build schema differs")
    require(value.get("attempt") == attempt and value.get("role") == role, f"{path}: build identity differs")
    require(value.get("version") == 1, f"{path}: build version differs")
    protocol = object_value(value.get("protocol"), f"{path}.protocol")
    require(protocol.get("path") == PROTOCOL_FILE and protocol.get("sha256") == protocol_hash, f"{path}: protocol binding differs")
    require(value.get("environment") == dict(expected_environment), f"{path}: build environment differs from frozen protocol")
    source_before = object_value(value.get("source_before"), f"{path}.source_before")
    source_after = object_value(value.get("source_after"), f"{path}.source_after")
    require(value.get("source_unchanged") is True and source_before == source_after, f"{path}: source changed during build")
    if expected_source is not None:
        require(source_after == dict(expected_source), f"{path}: source differs between formal builds")
    binary = object_value(value.get("binary"), f"{path}.binary")
    binary_path = Path(str(binary.get("path", "")))
    require(binary_path.is_file() and not binary_path.is_symlink() and os.access(binary_path, os.X_OK), f"{path}: copied binary is unavailable")
    require(metadata(binary_path) == {"bytes": integer(binary.get("bytes"), f"{path}.binary.bytes"), "sha256": digest(binary.get("sha256"), f"{path}.binary.sha256")}, f"{path}: binary changed")
    require(binary.get("executable") is True, f"{path}: binary executable binding is missing")
    original_binary = object_value(value.get("original_binary"), f"{path}.original_binary")
    require(
        integer(original_binary.get("bytes"), f"{path}.original_binary.bytes") == binary["bytes"]
        and digest(original_binary.get("sha256"), f"{path}.original_binary.sha256") == binary["sha256"]
        and original_binary.get("executable") is True
        and isinstance(original_binary.get("path"), str)
        and original_binary["path"],
        f"{path}: original/copy binary metadata differs",
    )
    command = value.get("command")
    require(isinstance(command, list) and command and all(isinstance(item, str) and item for item in command), f"{path}: build command is missing")
    gate = object_value(value.get("gate"), f"{path}.gate")
    expected_gate_path = Path("validation") / f"route-build-{role}-{attempt}.json"
    require(gate.get("path") == expected_gate_path.as_posix(), f"{path}: build gate path differs")
    gate_path = bundle_root / expected_gate_path
    require(gate_path.is_file() and sha(gate_path) == digest(gate.get("sha256"), f"{path}.gate.sha256"), f"{path}: build gate binding differs")
    gate_value = object_value(read_json(gate_path), str(gate_path))
    require(
        gate_value.get("schema") == "docx-stream-append-gate-v1"
        and gate_value.get("attempt") == attempt
        and gate_value.get("label") == f"route-build-{role}-{attempt}"
        and gate_value.get("exit_code") == 0
        and gate_value.get("source_unchanged") is True,
        f"{gate_path}: build gate failed or identity differs",
    )
    require(gate_value.get("argv") == command, f"{gate_path}: build gate command differs")
    require(gate_value.get("cwd") == str(routes.REPO), f"{gate_path}: build gate working directory differs")
    require(gate_value.get("environment") == dict(expected_environment), f"{gate_path}: build gate environment differs")
    require(gate_value.get("driver_sha256") == sha(ROOT / "gate.py"), f"{gate_path}: gate helper binding differs")
    require(gate_value.get("common_sha256") == sha(ROOT / "common.py"), f"{gate_path}: common helper binding differs")
    require(gate_value.get("source_before") == gate_value.get("source_after") == source_after, f"{gate_path}: build gate source differs")
    return value, dict(source_after)


def _validate_execution_inputs(
    bundle_root: Path,
    attempt: str,
    protocol_hash: str,
    source: Mapping[str, Any],
    prep_path: Path,
    builds: Mapping[str, Mapping[str, Any]],
) -> dict[str, Any]:
    path = bundle_root / "route-attempts" / attempt / EXECUTION_INPUTS_FILE
    require(path.is_file(), f"formal execution-input binding is missing: {path}")
    value = object_value(read_json(path), str(path))
    require(value.get("schema") == "docx-route-execution-inputs-v1", f"{path}: execution-input schema differs")
    require(value.get("attempt") == attempt, f"{path}: execution attempt differs")
    require(value.get("driver_sha256") == sha(ROOT / "bind_execution_inputs.py"), f"{path}: execution-input helper binding differs")
    protocol = object_value(value.get("protocol"), f"{path}.protocol")
    require(protocol.get("path") == PROTOCOL_FILE and protocol.get("sha256") == protocol_hash, f"{path}: protocol binding differs")
    require(protocol.get("bytes") == (bundle_root / PROTOCOL_FILE).stat().st_size, f"{path}: protocol length differs")
    machine = object_value(value.get("machine"), f"{path}.machine")
    machine_path = bundle_root / str(machine.get("path", ""))
    require(machine.get("path") == "machine.json" and machine_path.is_file(), f"{path}: machine inventory is missing")
    require(machine.get("bytes") == machine_path.stat().st_size and machine.get("sha256") == sha(machine_path), f"{path}: machine inventory binding differs")
    protocol_machine = object_value(read_json(bundle_root / "route-protocol.json").get("machine"), f"{path}.protocol.machine")
    require(machine.get("path") == protocol_machine.get("path") and machine.get("sha256") == protocol_machine.get("sha256"), f"{path}: machine differs from frozen protocol")
    require(value.get("source") == dict(source), f"{path}: source binding differs")
    require(value.get("cold_cache_claim") is False, f"{path}: cold-cache claim differs")
    source_record = object_value(value.get("source"), f"{path}.source")
    source_manifest = bundle_root / str(source_record.get("path", ""))
    require(source_manifest.is_file() and sha(source_manifest) == digest(source_record.get("sha256"), f"{path}.source.sha256"), f"{path}: source manifest changed")
    prep = object_value(value.get("preparation"), f"{path}.preparation")
    require(prep.get("path") == prep_path.relative_to(bundle_root).as_posix(), f"{path}: preparation path differs")
    require(prep.get("bytes") == prep_path.stat().st_size and sha(prep_path) == digest(prep.get("sha256"), f"{path}.preparation.sha256"), f"{path}: preparation receipt hash differs")
    execution_builds = object_value(value.get("builds"), f"{path}.builds")
    for role in ROLES:
        record = object_value(execution_builds.get(role), f"{path}.builds.{role}")
        expected_path = bundle_root / "route-attempts" / attempt / f"build-{role}.json"
        require(record.get("path") == expected_path.relative_to(bundle_root).as_posix(), f"{path}.builds.{role}.path differs")
        require(record.get("bytes") == expected_path.stat().st_size and sha(expected_path) == digest(record.get("sha256"), f"{path}.builds.{role}.sha256"), f"{path}.builds.{role}: hash differs")
        require(object_value(builds[role], f"builds.{role}") == object_value(read_json(expected_path), str(expected_path)), f"{path}.builds.{role}: receipt differs")
    filesystems = object_value(value.get("filesystems"), f"{path}.filesystems")
    filesystem_devices: dict[str, int] = {}
    expected_filesystem_paths = {
        "workspace": routes.REPO.resolve(),
        "evidence": bundle_root.resolve(),
        "scratch": routes.TEMP.resolve(),
    }
    for name in ("workspace", "evidence", "scratch"):
        item = object_value(filesystems.get(name), f"{path}.filesystems.{name}")
        filesystem_devices[name] = integer(item.get("device"), f"{path}.filesystems.{name}.device")
        recorded_path = Path(str(item.get("path", "")))
        require(recorded_path == expected_filesystem_paths[name], f"{path}.filesystems.{name}.path differs from the formal capability")
        require(recorded_path.is_dir() and recorded_path.stat().st_dev == filesystem_devices[name], f"{path}.filesystems.{name}: current path/device differs")
        mount = object_value(item.get("mount"), f"{path}.filesystems.{name}.mount")
        require(mount.get("status") == "pass" and mount.get("exit_code") == 0, f"{path}.filesystems.{name}: mount observation failed")
        mount_argv = mount.get("argv")
        require(
            isinstance(mount_argv, list)
            and len(mount_argv) >= 4
            and mount_argv[:4] == ["findmnt", "--json", "--target", str(recorded_path)],
            f"{path}.filesystems.{name}: mount target binding differs",
        )
    require(filesystem_devices["workspace"] == filesystem_devices["evidence"] == filesystem_devices["scratch"], f"{path}: workspace/evidence/scratch devices differ")
    file_inputs = value.get("file_inputs")
    expected_paths = {arm["input_file"] for arm in routes.AXIS_ARMS if arm["input_mode"] == "file"}
    require(isinstance(file_inputs, list) and {item.get("path") for item in file_inputs} == expected_paths and len(file_inputs) == len(expected_paths), f"{path}: file-input inventory differs")
    for item in file_inputs:
        record = object_value(item, f"{path}.file_inputs")
        relative = record.get("path")
        input_path = bundle_root / str(relative)
        expected_meta = {"bytes": integer(record.get("bytes"), f"{path}.file_inputs[{relative}].bytes"), "sha256": digest(record.get("sha256"), f"{path}.file_inputs[{relative}].sha256")}
        require(input_path.is_file() and not input_path.is_symlink() and metadata(input_path) == expected_meta, f"{path}: staged file input changed")
        require(record.get("absolute_path") == str(input_path.resolve()), f"{path}: staged input path binding differs")
        recorded_device = integer(record.get("device"), f"{path}.file_inputs[{relative}].device", positive=True)
        recorded_inode = integer(record.get("inode"), f"{path}.file_inputs[{relative}].inode", positive=True)
        require(recorded_device == input_path.stat().st_dev == filesystem_devices["evidence"], f"{path}: staged input filesystem differs")
        require(recorded_inode == input_path.stat().st_ino, f"{path}: staged input inode differs")
        origin = bundle_root / str(record.get("origin", ""))
        require(origin.is_file() and metadata(origin) == expected_meta, f"{path}: staged input origin differs")
    require(isinstance(value.get("replay_directory_roots"), list) and value["replay_directory_roots"] == [f"route-pilots/{attempt}", f"route-captures/{attempt}"], f"{path}: replay roots differ")
    require(value.get("replay_filesystem") == "evidence", f"{path}: replay filesystem binding differs")
    return value


def _validate_axis_preparation(
    bundle_root: Path,
    attempt: str,
    protocol_hash: str,
    source: Mapping[str, Any],
    normal_build: Mapping[str, Any],
) -> tuple[Path, dict[str, dict[str, Any]]]:
    path = bundle_root / "axis-input-origin" / attempt / "receipt.json"
    require(path.is_file(), f"axis input preparation receipt is missing: {path}")
    value = object_value(read_json(path), str(path))
    require(value.get("schema") == "docx-axis-input-preparation-v1" and value.get("status") == "pass", f"{path}: preparation did not pass")
    require(value.get("attempt") == attempt and value.get("protocol_sha256") == protocol_hash, f"{path}: preparation binding differs")
    gate_path = bundle_root / "validation" / f"axis-input-preparation-{attempt}.json"
    _validate_gate(gate_path, attempt, source)
    require(value.get("driver_sha256") == sha(ROOT / "prepare_axis_inputs.py"), f"{path}: preparation helper binding differs")
    exports = value.get("exports")
    require(isinstance(exports, list) and len(exports) == 2, f"{path}: exported source inventory differs")
    main_xml_by_source: dict[str, str] = {}
    for export in exports:
        export = object_value(export, f"{path}.exports")
        artifacts = object_value(export.get("artifacts"), f"{path}.exports.artifacts")
        for relative, expected_meta in artifacts.items():
            artifact_path = path.parent / relative
            require(artifact_path.is_file() and metadata(artifact_path) == expected_meta, f"{path}: exported artifact changed: {relative}")
        hash_files = [relative for relative in artifacts if relative.endswith("-hashes.json")]
        require(len(hash_files) == 1, f"{path}: exported source manifest is missing")
        manifest = object_value(read_json(path.parent / hash_files[0]), f"{path}.exports.source_manifest")
        source_record = object_value(manifest.get("source"), f"{path}.exports.source_manifest.source")
        candidate_record = object_value(manifest.get("candidate"), f"{path}.exports.source_manifest.candidate")
        source_hash = digest(source_record.get("sha256"), f"{path}.exports.source_manifest.source.sha256")
        main_xml_by_source[source_hash] = digest(manifest.get("source_main_xml_sha256"), f"{path}.exports.source_manifest.source_main_xml_sha256")
        for record, record_label in ((source_record, "source"), (candidate_record, "candidate")):
            record_path = path.parent / Path(hash_files[0]).parent / str(record.get("file", ""))
            expected = {"bytes": integer(record.get("bytes"), f"{path}.{record_label}.bytes"), "sha256": digest(record.get("sha256"), f"{path}.{record_label}.sha256")}
            require(record_path.is_file() and metadata(record_path) == expected, f"{path}: exported {record_label} archive differs")
    build = object_value(value.get("build"), f"{path}.build")
    require(build.get("path") == f"route-attempts/{attempt}/build-normal.json", f"{path}: preparation build path differs")
    build_path = bundle_root / str(build["path"])
    require(sha(build_path) == digest(build.get("sha256"), f"{path}.build.sha256"), f"{path}: preparation build hash differs")
    require(object_value(read_json(build_path), str(build_path)) == dict(normal_build), f"{path}: preparation build receipt differs")
    expected: dict[str, dict[str, Any]] = {}
    for arm in routes.AXIS_ARMS:
        if arm["input_mode"] == "file":
            expected[arm["input_file"]] = {"bytes": None, "sha256": None}
    staged = value.get("staged")
    require(isinstance(staged, list) and len(staged) == len(expected), f"{path}: staged input inventory differs")
    for item in staged:
        record = object_value(item, f"{path}.staged")
        staged_path = record.get("path")
        require(staged_path in expected and expected[staged_path]["bytes"] is None, f"{path}: unexpected staged input")
        bytes_value = integer(record.get("bytes"), f"{path}.staged[{staged_path}].bytes")
        hash_value = digest(record.get("sha256"), f"{path}.staged[{staged_path}].sha256")
        file_path = bundle_root / staged_path
        require(file_path.is_file() and metadata(file_path) == {"bytes": bytes_value, "sha256": hash_value}, f"{path}: staged input changed")
        origin = bundle_root / str(record.get("origin", ""))
        require(origin.is_file() and metadata(origin) == {"bytes": bytes_value, "sha256": hash_value}, f"{path}: source origin differs")
        require(hash_value in main_xml_by_source, f"{path}: staged source is absent from exported manifests")
        expected[staged_path] = {"bytes": bytes_value, "sha256": hash_value, "main_xml_sha256": main_xml_by_source[hash_value]}
    require(all(item["bytes"] is not None for item in expected.values()), f"{path}: staged input list is incomplete")
    return path, expected


def _validate_lane_gate(
    bundle_root: Path,
    attempt: str,
    label: str,
    source: Mapping[str, Any],
) -> tuple[str, dict[str, Any]]:
    path = bundle_root / "validation" / f"{label}-{attempt}.json"
    value = object_value(read_json(path), str(path))
    return _validate_gate(path, attempt, source), value


def _timestamp(value: Any, label: str) -> datetime.datetime:
    require(isinstance(value, str) and value, f"{label}: timestamp is missing")
    try:
        result = datetime.datetime.fromisoformat(value)
    except ValueError as error:
        raise AnalysisError(f"{label}: malformed timestamp") from error
    require(result.tzinfo is not None, f"{label}: timestamp must carry a timezone")
    return result


def _identity(report_case: Mapping[str, Any], label: str) -> dict[str, Any]:
    source = object_value(report_case.get("source"), f"{label}.source")
    authored = object_value(report_case.get("authored"), f"{label}.authored")
    oracle = object_value(report_case.get("oracle"), f"{label}.oracle")
    source_id = {
        "archive_bytes": integer(source.get("archive_bytes"), f"{label}.source.archive_bytes", positive=True),
        "archive_sha256": digest(source.get("archive_sha256"), f"{label}.source.archive_sha256"),
        "main_xml_bytes": integer(source.get("main_xml_bytes"), f"{label}.source.main_xml_bytes", positive=True),
        "main_xml_sha256": digest(source.get("main_xml_sha256"), f"{label}.source.main_xml_sha256"),
    }
    authored_id = {
        "encoded_xml_bytes": integer(authored.get("encoded_xml_bytes"), f"{label}.authored.encoded_xml_bytes", positive=True),
        "expected_event_sha256": digest(authored.get("expected_event_sha256"), f"{label}.authored.expected_event_sha256"),
        "expected_encoded_sha256": digest(authored.get("expected_encoded_sha256"), f"{label}.authored.expected_encoded_sha256"),
    }
    candidate_id = {
        "archive_bytes": integer(oracle.get("candidate_archive_bytes"), f"{label}.oracle.candidate_archive_bytes", positive=True),
        "archive_sha256": digest(oracle.get("candidate_archive_sha256"), f"{label}.oracle.candidate_archive_sha256"),
        "main_xml_bytes": integer(oracle.get("candidate_main_xml_bytes"), f"{label}.oracle.candidate_main_xml_bytes", positive=True),
        "main_xml_sha256": digest(oracle.get("candidate_main_xml_sha256"), f"{label}.oracle.candidate_main_xml_sha256"),
    }
    return {"source": source_id, "authored": authored_id, "candidate": candidate_id}


def _parse_resource(path: Path) -> int:
    prefix = "Maximum resident set size (kbytes):"
    values: list[int] = []
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as error:
        raise AnalysisError(f"cannot read resource receipt {path}: {error}") from error
    for line in lines:
        line = line.strip()
        if line.startswith(prefix):
            raw = line[len(prefix):].strip()
            require(raw.isdigit(), f"{path}: malformed maximum RSS")
            values.append(int(raw))
    require(len(values) == 1, f"{path}: expected exactly one maximum RSS observation")
    return values[0] * 1024


def _validate_process_sample(sample: Mapping[str, Any], label: str) -> dict[str, Any] | None:
    process = sample.get("process")
    if process is None:
        return None
    process = object_value(process, f"{label}.process")
    rss = integer(process.get("rss_bytes"), f"{label}.process.rss_bytes")
    peak = integer(process.get("peak_rss_bytes"), f"{label}.process.peak_rss_bytes")
    require(peak >= rss, f"{label}.process: peak RSS below RSS delta")
    return {"rss_bytes": rss, "peak_rss_bytes": peak}


def _validate_histogram(
    value: Any,
    label: str,
    expected_calls: int,
    *,
    observed_bytes: int | None = None,
    largest_write: int | None = None,
) -> dict[str, int]:
    """Validate and retain one of the fixed six-bin byte histograms.

    The capture validators already prove these bounds.  Rechecking the shape
    here keeps the analysis output tied to the retained report if a receipt is
    reconstructed or reviewed independently of the capture command.
    """
    histogram = object_value(value, label)
    require(set(histogram) == set(HISTOGRAM_FIELDS), f"{label}: histogram bins differ")
    counts = {field: integer(histogram.get(field), f"{label}.{field}") for field in HISTOGRAM_FIELDS}
    require(sum(counts.values()) == expected_calls, f"{label}: bins do not sum to {expected_calls}")
    lower_bounds = (0, 1, 513, 4_097, 16_385, 65_537)
    upper_bounds = (0, 512, 4_096, 16_384, 65_536, None)
    lower_total = sum(counts[field] * lower for field, lower in zip(HISTOGRAM_FIELDS, lower_bounds))
    upper_total = (
        None
        if counts[HISTOGRAM_FIELDS[-1]]
        else sum(counts[field] * upper for field, upper in zip(HISTOGRAM_FIELDS[:-1], upper_bounds[:-1]))
    )
    if observed_bytes is not None:
        integer(observed_bytes, f"{label}.observed_bytes")
        require(observed_bytes >= lower_total, f"{label}: aggregate bytes are below bucket lower bounds")
        if upper_total is not None:
            require(observed_bytes <= upper_total, f"{label}: aggregate bytes exceed bucket upper bounds")
    if largest_write is not None:
        largest = integer(largest_write, f"{label}.largest_write", positive=True)
        if observed_bytes is not None:
            require(largest <= observed_bytes, f"{label}: largest write exceeds aggregate bytes")
        if largest <= 512:
            largest_bucket = 1
        elif largest <= 4_096:
            largest_bucket = 2
        elif largest <= 16_384:
            largest_bucket = 3
        elif largest <= 65_536:
            largest_bucket = 4
        else:
            largest_bucket = 5
        highest_bucket = max((index for index, field in enumerate(HISTOGRAM_FIELDS) if counts[field]), default=0)
        require(largest_bucket == highest_bucket, f"{label}: largest write does not match highest nonzero bucket")
    return counts


def _summarize_histograms(values: list[dict[str, int]], label: str) -> dict[str, Any]:
    require(values, f"{label}: no histogram observations")
    if all(value == values[0] for value in values[1:]):
        return {"kind": "common", "bins": values[0]}
    return {"kind": "per_sample", "bins": values}


def _validate_allocation_sample(sample: Mapping[str, Any], role: str, label: str) -> dict[str, Any] | None:
    allocation = sample.get("allocation")
    if role == "normal":
        require(allocation is None, f"{label}.allocation: normal sample must omit allocator metrics")
        return None
    allocation = object_value(allocation, f"{label}.allocation")
    require(allocation.get("status") == "measured" and allocation.get("scope") == "operation_global_system_allocator", f"{label}.allocation: allocator identity differs")
    result = {field: integer(allocation.get(field), f"{label}.allocation.{field}") for field in (
        "allocation_calls", "deallocation_calls", "reallocation_calls", "failed_allocation_calls",
        "allocated_bytes", "deallocated_bytes", "live_bytes_before", "live_bytes_after",
        "peak_live_bytes_before", "peak_live_bytes_after", "region_peak_live_bytes",
    )}
    require(result["failed_allocation_calls"] == 0, f"{label}.allocation: failed allocation count is non-zero")
    require(result["live_bytes_before"] + result["allocated_bytes"] - result["deallocated_bytes"] == result["live_bytes_after"], f"{label}.allocation: conservation failed")
    require(result["peak_live_bytes_before"] <= result["peak_live_bytes_after"], f"{label}.allocation: absolute peak regressed")
    require(result["region_peak_live_bytes"] >= max(result["live_bytes_before"], result["live_bytes_after"]), f"{label}.allocation: region peak below endpoints")
    require(result["region_peak_live_bytes"] <= result["peak_live_bytes_after"], f"{label}.allocation: region peak exceeds absolute peak")
    result["operation_peak_increment_bytes"] = result["region_peak_live_bytes"] - result["live_bytes_before"]
    require(result["operation_peak_increment_bytes"] >= 0, f"{label}.allocation: operation peak increment is negative")
    return result


def _validate_replay_sample(
    sample: Mapping[str, Any],
    expected_replays: int,
    label: str,
) -> dict[str, Any] | None:
    replay = sample.get("replay")
    if expected_replays == 0:
        require(replay is None, f"{label}.replay: deterministic sample must omit replay counters")
        return None
    replay = object_value(replay, f"{label}.replay")
    for field in (
        "replay_opens", "replay_read_calls", "replay_requested_bytes",
        "replay_returned_bytes", "request_histogram", "returned_histogram",
        "file_write_calls",
    ):
        require(field in replay, f"{label}.replay.{field}: missing replay evidence")
    replay_opens = integer(replay.get("replay_opens"), f"{label}.replay.replay_opens", positive=True)
    require(replay_opens == expected_replays, f"{label}.replay.replay_opens: expected {expected_replays}")
    read_calls = integer(replay.get("replay_read_calls"), f"{label}.replay.replay_read_calls", positive=True)
    requested = integer(replay.get("replay_requested_bytes"), f"{label}.replay.replay_requested_bytes", positive=True)
    returned = integer(replay.get("replay_returned_bytes"), f"{label}.replay.replay_returned_bytes")
    require(returned <= requested, f"{label}.replay: returned bytes exceed requested bytes")
    request_histogram = _validate_histogram(
        replay.get("request_histogram"),
        f"{label}.replay.request_histogram",
        read_calls,
        observed_bytes=requested,
    )
    returned_histogram = _validate_histogram(
        replay.get("returned_histogram"),
        f"{label}.replay.returned_histogram",
        read_calls,
        observed_bytes=returned,
    )
    file_writes = replay.get("file_write_calls")
    if file_writes is not None:
        file_writes = integer(file_writes, f"{label}.replay.file_write_calls", positive=True)
    return {
        "replay_opens": replay_opens,
        "replay_read_calls": read_calls,
        "replay_requested_bytes": requested,
        "replay_returned_bytes": returned,
        "file_write_calls": file_writes,
        "request_histogram": request_histogram,
        "returned_histogram": returned_histogram,
    }


def _validate_report_shape(
    report: Mapping[str, Any],
    expected: Mapping[str, Any],
    family: str,
    role: str,
    label: str,
) -> tuple[dict[str, Any], list[dict[str, Any]], dict[str, Any]]:
    require(report.get("schema") == routes.REPORT_SCHEMA and report.get("version") == 1, f"{label}: report schema differs")
    expected_binary = {
        "binary": "litchi-perf-baseline" if role == "normal" else "litchi-perf-baseline-alloc",
        "allocator": "Rust system allocator" if role == "normal" else "CountingSystemAllocator(std::alloc::System)",
        "instrumentation": "none" if role == "normal" else "system_allocator_operation_scoped",
        "counter_revision": None if role == "normal" else "serialized_region_peak_v3",
    }
    require(report.get("binary") == expected_binary, f"{label}: report binary identity differs")
    config = object_value(report.get("config"), f"{label}.config")
    case_label = expected["case"]
    case_spec = routes.ROUTE_CASE_BY_LABEL[case_label] if family == "route" else routes._axis_case(routes.AXIS_ARM_BY_LABEL[expected["arm"]])
    require(config.get("samples") == FORMAL_SAMPLES and config.get("warmups") == FORMAL_WARMUPS, f"{label}: formal sample contract differs")
    require(config.get("provider") == (expected["route"] if family == "route" else expected["provider"]), f"{label}: provider identity differs")
    require(config.get("source_counts") == [case_spec["source_count"]], f"{label}: source workload differs")
    require(config.get("authored_counts") == [case_spec["authored_count"]], f"{label}: authored workload differs")
    require(config.get("chunk_modes") == [routes.base.REPORT_CHUNK_MODES[case_spec["chunk_mode"]]], f"{label}: chunk workload differs")
    require(config.get("text_modes") == [routes.base.REPORT_TEXT_MODES[case_spec["text_mode"]]], f"{label}: text workload differs")
    require(config.get("input_mode") == expected["input_mode"], f"{label}: input profile differs")
    require(config.get("input_storage_kind") == expected["input_backing"], f"{label}: input storage differs")
    require(config.get("input_identity_validation") == expected["input_identity_validation"], f"{label}: input identity policy differs")
    require(config.get("sink_write_bytes") == expected["sink_write_bytes"], f"{label}: sink profile differs")
    require(config.get("compression") == expected["compression"], f"{label}: compression profile differs")
    expected_authored_opens = expected.get(
        "expected_authored_opens",
        routes.ROUTE_BY_NAME[expected["route"]].expected_authored_opens
        if family == "route" else routes.base.EXPECTED_AUTHORED_OPENS,
    )
    require(config.get("expected_authored_opens") == expected_authored_opens, f"{label}: authored-open contract differs")
    expected_replays = expected.get(
        "expected_replay_opens",
        routes.ROUTE_BY_NAME[expected["route"]].expected_replay_opens
        if family == "route" else 0,
    )
    integer(expected_replays, f"{label}.expected_replay_opens")
    observed_cases = report.get("cases")
    require(isinstance(observed_cases, list) and len(observed_cases) == 1, f"{label}: report case cardinality differs")
    observed = object_value(observed_cases[0], f"{label}.cases[0]")
    for field in ("source_count", "authored_count", "chunk_mode", "text_mode"):
        expected_value = routes.base.REPORT_CHUNK_MODES[case_spec[field]] if field == "chunk_mode" else routes.base.REPORT_TEXT_MODES[case_spec[field]] if field == "text_mode" else case_spec[field]
        require(observed.get(field) == expected_value, f"{label}.cases[0].{field}: workload differs")
    require(observed.get("provider") == config.get("provider"), f"{label}.cases[0].provider: identity differs")
    identities = _identity(observed, f"{label}.cases[0]")
    samples = observed.get("samples")
    require(isinstance(samples, list) and len(samples) == FORMAL_SAMPLES, f"{label}: sample cardinality differs")
    checked: list[dict[str, Any]] = []
    for index, raw in enumerate(samples):
        sample = object_value(raw, f"{label}.cases[0].samples[{index}]")
        require(sample.get("sample") == index, f"{label}.cases[0].samples[{index}]: sample index differs")
        elapsed = integer(sample.get("elapsed_ns"), f"{label}.cases[0].samples[{index}].elapsed_ns", positive=True)
        reads = object_value(sample.get("source_reads"), f"{label}.cases[0].samples[{index}].source_reads")
        source_calls = integer(reads.get("calls"), f"{label}.cases[0].samples[{index}].source_reads.calls", positive=True)
        source_requested = integer(reads.get("requested_bytes"), f"{label}.cases[0].samples[{index}].source_reads.requested_bytes", positive=True)
        source_returned = integer(reads.get("returned_bytes"), f"{label}.cases[0].samples[{index}].source_reads.returned_bytes")
        require(source_returned > 0 and source_returned <= source_requested, f"{label}.cases[0].samples[{index}].source_reads: invalid requested/returned totals")
        source_request_histogram = _validate_histogram(
            reads.get("request_histogram"),
            f"{label}.cases[0].samples[{index}].source_reads.request_histogram",
            source_calls,
            observed_bytes=source_requested,
        )
        source_returned_histogram = _validate_histogram(
            reads.get("returned_histogram"),
            f"{label}.cases[0].samples[{index}].source_reads.returned_histogram",
            source_calls,
            observed_bytes=source_returned,
        )
        authored = object_value(sample.get("authored"), f"{label}.cases[0].samples[{index}].authored")
        authored_opens = integer(authored.get("opens"), f"{label}.cases[0].samples[{index}].authored.opens")
        require(authored_opens == expected_authored_opens, f"{label}.cases[0].samples[{index}].authored.opens: expected {expected_authored_opens}")
        integer(authored.get("events"), f"{label}.cases[0].samples[{index}].authored.events")
        integer(authored.get("text_chunks"), f"{label}.cases[0].samples[{index}].authored.text_chunks")
        integer(authored.get("text_bytes"), f"{label}.cases[0].samples[{index}].authored.text_bytes")
        sink = object_value(sample.get("sink"), f"{label}.cases[0].samples[{index}].sink")
        sink_bytes = integer(sink.get("accepted_bytes"), f"{label}.cases[0].samples[{index}].sink.accepted_bytes", positive=True)
        sink_calls = integer(sink.get("write_calls"), f"{label}.cases[0].samples[{index}].sink.write_calls", positive=True)
        sink_largest = integer(sink.get("largest_write"), f"{label}.cases[0].samples[{index}].sink.largest_write", positive=True)
        sink_histogram = _validate_histogram(
            sink.get("histogram"),
            f"{label}.cases[0].samples[{index}].sink.histogram",
            sink_calls,
            observed_bytes=sink_bytes,
            largest_write=sink_largest,
        )
        digest(sink.get("sha256"), f"{label}.cases[0].samples[{index}].sink.sha256")
        require(sink_bytes == identities["candidate"]["archive_bytes"] and sink["sha256"] == identities["candidate"]["archive_sha256"], f"{label}.cases[0].samples[{index}].sink: candidate identity differs")
        process = _validate_process_sample(sample, f"{label}.cases[0].samples[{index}]")
        allocation = _validate_allocation_sample(sample, role, f"{label}.cases[0].samples[{index}]")
        replay = _validate_replay_sample(sample, expected_replays, f"{label}.cases[0].samples[{index}]")
        checked.append({
            "elapsed_ns": elapsed,
            "source_calls": source_calls,
            "source_requested_bytes": source_requested,
            "source_returned_bytes": source_returned,
            "source_request_histogram": source_request_histogram,
            "source_returned_histogram": source_returned_histogram,
            "authored_opens": authored_opens,
            "sink_write_calls": sink_calls,
            "sink_accepted_bytes": sink_bytes,
            "sink_largest_write": sink_largest,
            "sink_histogram": sink_histogram,
            "process": process,
            "allocation": allocation,
            "replay": replay,
        })
    return identities, checked, observed


def percentile(ordered: list[float], fraction: float) -> float:
    require(ordered, "cannot calculate a percentile from an empty vector")
    if len(ordered) == 1:
        return float(ordered[0])
    position = (len(ordered) - 1) * fraction
    low = math.floor(position)
    high = math.ceil(position)
    if low == high:
        return float(ordered[low])
    ratio = position - low
    return float(ordered[low] + (ordered[high] - ordered[low]) * ratio)


def stats(values: Iterable[int | float]) -> dict[str, float | int]:
    numeric = [finite_number(value, "metric") for value in values]
    require(numeric, "cannot summarize an empty metric")
    ordered = sorted(numeric)
    return {
        "n": len(ordered),
        "min": ordered[0],
        "max": ordered[-1],
        "mean": statistics.fmean(ordered),
        "p50": percentile(ordered, 0.50),
        "p95": percentile(ordered, 0.95),
        "p99": percentile(ordered, 0.99),
    }


def _optional_metric_stats(samples: list[Mapping[str, Any]], field: str, label: str) -> dict[str, float | int] | None:
    values = [sample[field] for sample in samples if sample.get(field) is not None]
    if not values:
        return None
    require(len(values) == len(samples), f"{label}: metric is only present in some samples")
    return stats(values)


def _summarize_report(
    report: Mapping[str, Any],
    resource_path: Path,
    expected: Mapping[str, Any],
    family: str,
) -> dict[str, Any]:
    role = expected["role"]
    label = f"{family}:{expected['label']}"
    identities, samples, observed = _validate_report_shape(report, expected, family, role, label)
    resource_rss = _parse_resource(resource_path)
    source_bytes = identities["source"]["archive_bytes"]
    authored_bytes = identities["authored"]["encoded_xml_bytes"]
    candidate_bytes = identities["candidate"]["archive_bytes"]
    elapsed = [item["elapsed_ns"] for item in samples]
    metric_values: dict[str, list[int | float]] = {
        "elapsed_ns": elapsed,
        "source_throughput_bytes_per_second": [source_bytes * 1_000_000_000 / value for value in elapsed],
        "authored_throughput_bytes_per_second": [authored_bytes * 1_000_000_000 / value for value in elapsed],
        "candidate_throughput_bytes_per_second": [candidate_bytes * 1_000_000_000 / value for value in elapsed],
        # GNU time observes the whole child once, after the sample process
        # exits.  Keep it as one observation instead of repeating it across
        # the 30 in-process samples.
        "time_max_rss_bytes": [resource_rss],
    }
    process_rss = [item["process"]["rss_bytes"] for item in samples if item["process"] is not None]
    process_peak_rss = [item["process"]["peak_rss_bytes"] for item in samples if item["process"] is not None]
    if process_rss:
        metric_values["process_rss_delta_bytes"] = process_rss
        metric_values["process_peak_rss_bytes"] = process_peak_rss
    metrics = {name: stats(metric_values[name]) if name in metric_values else None for name in PROCESS_METRICS}
    allocation: dict[str, Any] | None = None
    if role == "allocator":
        for field in ALLOCATOR_METRICS:
            allocation = allocation or {}
            allocation[field] = stats(item["allocation"][field] for item in samples if item["allocation"] is not None)
    io: dict[str, Any] = {
        "source": {
            "calls": stats(item["source_calls"] for item in samples),
            "requested_bytes": stats(item["source_requested_bytes"] for item in samples),
            "returned_bytes": stats(item["source_returned_bytes"] for item in samples),
            "request_histogram": _summarize_histograms(
                [item["source_request_histogram"] for item in samples],
                f"{label}.source.request_histogram",
            ),
            "returned_histogram": _summarize_histograms(
                [item["source_returned_histogram"] for item in samples],
                f"{label}.source.returned_histogram",
            ),
        },
        "authored": {
            "opens": stats(item["authored_opens"] for item in samples),
        },
        "sink": {
            "write_calls": stats(item["sink_write_calls"] for item in samples),
            "accepted_bytes": stats(item["sink_accepted_bytes"] for item in samples),
            "largest_write": stats(item["sink_largest_write"] for item in samples),
            "histogram": _summarize_histograms(
                [item["sink_histogram"] for item in samples],
                f"{label}.sink.histogram",
            ),
        },
    }
    replay_samples = [item["replay"] for item in samples if item["replay"] is not None]
    if replay_samples:
        require(len(replay_samples) == len(samples), f"{label}.replay: replay counters are only present in some samples")
        replay_io: dict[str, Any] = {
            "replay_opens": stats(item["replay_opens"] for item in replay_samples),
            "replay_read_calls": stats(item["replay_read_calls"] for item in replay_samples),
            "replay_requested_bytes": stats(item["replay_requested_bytes"] for item in replay_samples),
            "replay_returned_bytes": stats(item["replay_returned_bytes"] for item in replay_samples),
            "request_histogram": _summarize_histograms(
                [item["request_histogram"] for item in replay_samples],
                f"{label}.replay.request_histogram",
            ),
            "returned_histogram": _summarize_histograms(
                [item["returned_histogram"] for item in replay_samples],
                f"{label}.replay.returned_histogram",
            ),
        }
        file_write_stats = _optional_metric_stats(replay_samples, "file_write_calls", f"{label}.replay.file_write_calls")
        if file_write_stats is not None:
            replay_io["file_write_calls"] = file_write_stats
        io["replay"] = replay_io
    return {
        "family": family,
        "identity": expected.get("route", expected.get("arm")),
        "route": expected.get("route"),
        "axis": expected.get("axis"),
        "value": expected.get("value"),
        "case": expected["case"],
        "role": role,
        "repeat": expected["repeat"],
        "label": expected["label"],
        "source_identity": identities["source"],
        "authored_identity": identities["authored"],
        "candidate_identity": identities["candidate"],
        "samples": FORMAL_SAMPLES,
        "metrics": metrics,
        "io": io,
        "allocator": allocation,
        "process_rss_observations": len(process_rss),
        "time_max_rss_bytes": resource_rss,
    }


def _identity_tuple(row: Mapping[str, Any], *, include_candidate: bool = True) -> tuple[Any, ...]:
    values = (row["source_identity"], row["authored_identity"])
    if include_candidate:
        values += (row["candidate_identity"],)
    return values


def _main_xml_tuple(row: Mapping[str, Any]) -> tuple[Any, ...]:
    source = row["source_identity"]
    candidate = row["candidate_identity"]
    return (
        source["main_xml_bytes"],
        source["main_xml_sha256"],
        candidate["main_xml_bytes"],
        candidate["main_xml_sha256"],
    )


def validate_identity_consistency(rows: list[Mapping[str, Any]]) -> None:
    by_key: dict[tuple[Any, ...], list[Mapping[str, Any]]] = {}
    for row in rows:
        key = (row["family"], row["identity"], row["case"], row["repeat"])
        by_key.setdefault(key, []).append(row)
    for key, group in by_key.items():
        roles = {row["role"] for row in group}
        require(roles == set(ROLES), f"{key}: normal/allocator role pair is incomplete")
        reference = group[0]
        for row in group[1:]:
            require(_identity_tuple(row) == _identity_tuple(reference), f"{key}: source/authored/candidate identity differs between roles")

    # Repeats for one route or one axis arm must retain the complete physical
    # identity.  In particular, do not silently accept a candidate archive
    # changing between repeats while source/authored metadata remains stable.
    by_arm: dict[tuple[Any, ...], list[Mapping[str, Any]]] = {}
    for row in rows:
        key = (row["family"], row["identity"], row["case"], row["role"])
        by_arm.setdefault(key, []).append(row)
    for key, group in by_arm.items():
        reference = group[0]
        for row in group[1:]:
            require(_identity_tuple(row) == _identity_tuple(reference), f"{key}: source/authored/candidate identity changes across repeats")

    # The three route providers run against the same current profile.  Their
    # source, authored, and candidate archives must therefore be byte/hash
    # identical for each workload, role, and repeat.
    by_route_workload: dict[tuple[Any, ...], list[Mapping[str, Any]]] = {}
    for row in rows:
        if row["family"] == "route":
            key = (row["case"], row["role"], row["repeat"])
            by_route_workload.setdefault(key, []).append(row)
    for key, group in by_route_workload.items():
        require({row["route"] for row in group} == set(routes.ROUTE_NAMES), f"{key}: route identity inventory is incomplete")
        reference = group[0]
        for row in group[1:]:
            require(_identity_tuple(row) == _identity_tuple(reference), f"{key}: route source/authored/candidate identity differs")

    # Each deterministic one-factor axis is compared with the matching
    # current-profile route control.  Store/Deflate may change physical ZIP
    # bytes, but their source and candidate main XML plus authored stream must
    # remain identical to that control.
    controls = {
        (row["case"], row["role"], row["repeat"]): row
        for row in rows
        if row["family"] == "route" and row.get("route") == "deterministic"
    }
    for row in rows:
        if row["family"] != "axis":
            continue
        key = (row["case"], row["role"], row["repeat"])
        require(key in controls, f"{key}: deterministic control for axis is missing")
        control = controls[key]
        compression = row.get("value") if row.get("axis") == "compression" else "current"
        if compression in ("store", "deflate"):
            require(row["authored_identity"] == control["authored_identity"], f"{key}: compression axis authored identity differs from control")
            require(_main_xml_tuple(row) == _main_xml_tuple(control), f"{key}: compression axis main XML identity differs from control")
        else:
            require(_identity_tuple(row) == _identity_tuple(control), f"{key}: axis physical identity differs from control")


def _repeat_stats(rows: list[Mapping[str, Any]]) -> list[dict[str, Any]]:
    groups: dict[tuple[Any, ...], list[Mapping[str, Any]]] = {}
    for row in rows:
        key = (row["family"], row["identity"], row["case"], row["role"])
        groups.setdefault(key, []).append(row)
    output: list[dict[str, Any]] = []
    for key, group in sorted(groups.items(), key=str):
        repeats = sorted(group, key=lambda row: row["repeat"])
        require([row["repeat"] for row in repeats] == list(REPEATS), f"{key}: repeat inventory differs")
        per_repeat = []
        for row in repeats:
            per_repeat.append({
                "repeat": row["repeat"],
                "process_count": 1,
                "samples": row["samples"],
                "metrics": row["metrics"],
                "io": row["io"],
                "allocator": row["allocator"],
            })
        uncertainty: dict[str, Any] = {
            "repeat_count": len(per_repeat),
            "method": "two-independent-process-repeat-range",
            "confidence_interval": None,
            "metrics": {},
        }
        for metric in PROCESS_METRICS:
            values = [item["metrics"][metric]["p50"] for item in per_repeat if item["metrics"].get(metric) is not None]
            if not values:
                uncertainty["metrics"][metric] = None
                continue
            low, high = min(values), max(values)
            uncertainty["metrics"][metric] = {
                "repeat_p50s": values,
                "min": low,
                "max": high,
                "range_percent": None if low == 0 else (high - low) / abs(low) * 100.0,
            }
        output.append({
            "family": key[0],
            "identity": key[1],
            "case": key[2],
            "role": key[3],
            "repeats": per_repeat,
            "uncertainty": uncertainty,
        })
    return output


def _percent_change(before: float, after: float) -> float | None:
    if before == 0:
        return 0.0 if after == 0 else None
    return (after - before) / before * 100.0


def _comparison_record(
    before: Mapping[str, Any],
    after: Mapping[str, Any],
    metric: str,
    *,
    kind: str,
    control: str,
    candidate: str,
) -> dict[str, Any]:
    before_metric = before["metrics"].get(metric)
    after_metric = after["metrics"].get(metric)
    if before_metric is None or after_metric is None:
        return {"metric": metric, "available": False, "comparison_kind": kind, "causal_speedup": False}
    before_value = float(before_metric["p50"])
    after_value = float(after_metric["p50"])
    change = _percent_change(before_value, after_value)
    return {
        "metric": metric,
        "available": True,
        "control": control,
        "candidate": candidate,
        "control_p50": before_value,
        "candidate_p50": after_value,
        "percent_change": change,
        "review_flag": change is None or abs(change) > REVIEW_THRESHOLD_PERCENT,
        "comparison_kind": kind,
        "causal_speedup": False,
        "review_threshold_percent": REVIEW_THRESHOLD_PERCENT,
    }


def route_comparisons(rows: list[Mapping[str, Any]]) -> list[dict[str, Any]]:
    by_key = {(row["case"], row["role"], row["repeat"], row["route"]): row for row in rows}
    output: list[dict[str, Any]] = []
    for case in sorted({row["case"] for row in rows}):
        for role in ROLES:
            for repeat in REPEATS:
                control = by_key[(case, role, repeat, "deterministic")]
                for candidate in ("memory_store", "file_store"):
                    row = by_key[(case, role, repeat, candidate)]
                    output.append({
                        "case": case,
                        "role": role,
                        "repeat": repeat,
                        "control_route": "deterministic",
                        "candidate_route": candidate,
                        "interpretation": "contemporaneous route comparison; no historical optimization speedup",
                        "metrics": [_comparison_record(control, row, metric, kind="contemporaneous_route", control="deterministic", candidate=candidate) for metric in PROCESS_METRICS],
                        "allocator": None if role == "normal" else {
                            metric: _comparison_record({"metrics": control["allocator"]}, {"metrics": row["allocator"]}, metric, kind="contemporaneous_route_allocator", control="deterministic", candidate=candidate)
                            for metric in ALLOCATOR_METRICS
                        },
                    })
    return output


def axis_comparisons(axis_rows: list[Mapping[str, Any]], route_rows: list[Mapping[str, Any]]) -> list[dict[str, Any]]:
    baseline = {(row["case"], row["role"], row["repeat"]): row for row in route_rows if row["route"] == "deterministic"}
    output: list[dict[str, Any]] = []
    for row in sorted(axis_rows, key=lambda item: (item["axis"], str(item["value"]), item["case"], item["role"], item["repeat"])):
        control = baseline[(row["case"], row["role"], row["repeat"])]
        output.append({
            "axis": row["axis"],
            "value": row["value"],
            "workload": row["case"],
            "role": row["role"],
            "repeat": row["repeat"],
            "control": "deterministic current profile",
            "candidate": row["identity"],
            "interpretation": "one-factor contemporaneous profile comparison; no historical optimization speedup",
            "metrics": [_comparison_record(control, row, metric, kind="one_factor_axis", control="deterministic", candidate=str(row["identity"])) for metric in PROCESS_METRICS],
            "allocator": None if row["role"] == "normal" else {
                metric: _comparison_record({"metrics": control["allocator"]}, {"metrics": row["allocator"]}, metric, kind="one_factor_axis_allocator", control="deterministic", candidate=str(row["identity"]))
                for metric in ALLOCATOR_METRICS
            },
        })
    return output


def _receipt_expected_path(bundle_root: Path, family: str, attempt: str, expected: Mapping[str, Any]) -> Path:
    directory = "route-captures" if family == "route" else "axis-captures"
    return bundle_root / directory / attempt / expected["label"] / "receipt.json"


def _canonical_report_check(
    report_path: Path,
    resource_path: Path,
    receipt_path: Path,
    expected: Mapping[str, Any],
    family: str,
    binary: Mapping[str, Any],
    argv: list[str],
    prepared_inputs: Mapping[str, Mapping[str, Any]],
) -> None:
    """Run the frozen capture acceptance contract before local extraction."""
    try:
        if family == "route":
            route = routes.ROUTE_BY_NAME[expected["route"]]
            case = routes.ROUTE_CASE_BY_LABEL[expected["case"]]
            replay_dir = receipt_path.parent / "replay" if expected["route"] == "file_store" else None
            routes.check_route_report(
                report_path,
                expected["role"],
                case,
                route,
                samples=FORMAL_SAMPLES,
                warmups=FORMAL_WARMUPS,
                binary=dict(binary),
                argv=argv,
                replay_dir=replay_dir,
            )
        else:
            arm = routes.AXIS_ARM_BY_LABEL[expected["arm"]]
            input_metadata = prepared_inputs.get(expected["input_file"]) if expected["input_mode"] == "file" else None
            routes._check_axis_report(
                report_path,
                expected["role"],
                arm,
                samples=FORMAL_SAMPLES,
                warmups=FORMAL_WARMUPS,
                binary=dict(binary),
                argv=argv,
                input_metadata=input_metadata,
            )
    except routes.base.MeasureError as error:
        fail(f"{report_path}: canonical route acceptance failed: {error}")


def _validate_cleanup_manifest(bundle_root: Path, attempt: str) -> dict[str, Any] | None:
    """Validate the optional post-capture removal record for file replay dirs."""
    path = bundle_root / f"cleanup-replay-dirs-{attempt}.json"
    if not path.is_file():
        return None
    value = object_value(read_json(path), str(path))
    require(value.get("schema") == "docx-replay-directory-cleanup-v1", f"{path}: cleanup schema differs")
    require(value.get("attempt") == attempt and value.get("status") == "pass", f"{path}: cleanup did not pass")
    entries = value.get("entries")
    require(isinstance(entries, list), f"{path}: cleanup entries are missing")
    expected_paths = {
        f"route-captures/{attempt}/{row['label']}/replay"
        for row in routes._run_inventory(pilot=False)
        if row["route"] == "file_store"
    }
    expected_paths.update(
        f"route-pilots/{attempt}/{row['label']}/replay"
        for row in routes._run_inventory(pilot=True)
        if row["route"] == "file_store"
    )
    require(len(entries) == len(expected_paths), f"{path}: cleanup entry cardinality differs")
    by_path: dict[str, dict[str, Any]] = {}
    for raw in entries:
        entry = object_value(raw, f"{path}.entries")
        require(
            set(entry) == {"path", "receipt", "device", "inode", "empty_before_removal", "removed"},
            f"{path}: cleanup entry fields differ",
        )
        relative = entry.get("path")
        require(isinstance(relative, str) and relative in expected_paths, f"{path}: unexpected cleanup path")
        require(relative not in by_path, f"{path}: duplicate cleanup path")
        receipt = object_value(entry.get("receipt"), f"{path}.entries[{relative}].receipt")
        expected_receipt = relative.removesuffix("/replay") + "/receipt.json"
        require(receipt.get("path") == expected_receipt, f"{path}: cleanup receipt path differs for {relative}")
        receipt_path = bundle_root / expected_receipt
        require(receipt_path.is_file() and not receipt_path.is_symlink(), f"{path}: cleanup receipt is missing for {relative}")
        receipt_meta = metadata(receipt_path)
        require(
            receipt_meta == {
                "bytes": integer(receipt.get("bytes"), f"{path}.entries[{relative}].receipt.bytes"),
                "sha256": digest(receipt.get("sha256"), f"{path}.entries[{relative}].receipt.sha256"),
            },
            f"{path}: cleanup receipt changed for {relative}",
        )
        device = integer(entry.get("device"), f"{path}.entries[{relative}].device", positive=True)
        integer(entry.get("inode"), f"{path}.entries[{relative}].inode", positive=True)
        require(entry.get("empty_before_removal") is True and entry.get("removed") is True, f"{path}: cleanup removal proof is incomplete for {relative}")
        retained = object_value(read_json(receipt_path), str(receipt_path))
        require(retained.get("status") == "pass" and retained.get("exit_code") == 0, f"{path}: file route receipt did not pass for {relative}")
        run = object_value(retained.get("run"), f"{receipt_path}.run")
        require(run.get("route") == "file_store", f"{path}: cleanup receipt is not a file-store run for {relative}")
        report_path = _artifact(receipt_path.parent, retained, "report.json")
        report = object_value(read_json(report_path), str(report_path))
        cases = report.get("cases")
        require(isinstance(cases, list) and len(cases) == 1, f"{path}: cleanup report case inventory differs for {relative}")
        samples = object_value(cases[0], f"{report_path}.cases[0]").get("samples")
        require(isinstance(samples, list) and samples, f"{path}: cleanup report samples are missing for {relative}")
        for index, sample in enumerate(samples):
            sample_value = object_value(sample, f"{report_path}.cases[0].samples[{index}]")
            replay = object_value(sample_value.get("replay"), f"{report_path}.cases[0].samples[{index}].replay")
            require(replay.get("file_cleanup_verified") is True, f"{path}: file cleanup flag is absent for {relative} sample {index}")
        by_path[relative] = {"entry": entry, "device": device}
    require(set(by_path) == expected_paths, f"{path}: cleanup path inventory differs")
    return {"path": path.relative_to(bundle_root).as_posix(), "sha256": sha(path), "entries": by_path}


def _validate_replay_directory(
    bundle_root: Path,
    receipt_path: Path,
    route: str,
    cleanup_manifest: Mapping[str, Any] | None,
) -> None:
    if route != "file_store":
        return
    replay_dir = receipt_path.parent / "replay"
    if replay_dir.is_symlink():
        fail(f"{receipt_path}: file replay directory is a symlink")
    if replay_dir.exists():
        require(replay_dir.is_dir(), f"{receipt_path}: file replay path is not a directory")
        require(not any(replay_dir.iterdir()), f"{receipt_path}: file replay cleanup is not proven")
        return
    require(cleanup_manifest is not None, f"{receipt_path}: removed replay directory has no cleanup manifest")
    relative = replay_dir.relative_to(bundle_root).as_posix()
    entries = object_value(cleanup_manifest.get("entries"), "cleanup_manifest.entries")
    require(relative in entries, f"{receipt_path}: removed replay directory is absent from cleanup manifest")


def _validate_receipt(
    path: Path,
    expected: Mapping[str, Any],
    family: str,
    attempt: str,
    protocol: Mapping[str, Any],
    protocol_hash: str,
    bundle_root: Path,
    source: Mapping[str, Any],
    builds: Mapping[str, Mapping[str, Any]],
    prepared_inputs: Mapping[str, Mapping[str, Any]],
    cleanup_manifest: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    require(path.is_file(), f"{expected['label']}: formal receipt is missing")
    value = object_value(read_json(path), str(path))
    expected_schema = "docx-replayable-tail-append-route-capture-v1" if family == "route" else "docx-replayable-tail-append-axis-capture-v1"
    require(value.get("schema") == expected_schema and value.get("version") == 1, f"{path}: receipt schema differs")
    require(value.get("status") == "pass" and value.get("exit_code") == 0, f"{path}: formal receipt did not pass")
    require(value.get("attempt") == attempt and value.get("missing_artifacts") == [], f"{path}: receipt completion differs")
    require(value.get("launch_error") is None and value.get("validation_error") is None, f"{path}: receipt contains an execution error")
    started_path = path.parent / "started.json"
    require(started_path.is_file(), f"{path}: started receipt is missing")
    started = object_value(read_json(started_path), str(started_path))
    require(started.get("status") == "running", f"{started_path}: started receipt status differs")
    for field in ("attempt", "run", "protocol", "machine", "build", "binary", "argv", "cwd", "environment"):
        require(started.get(field) == value.get(field), f"{path}: started/finished {field} binding differs")
    run = object_value(value.get("run"), f"{path}.run")
    for key, item in expected.items():
        # The canonical axis receipt binds the arm through axis/value/workload
        # fields and deliberately does not duplicate the internal arm label.
        if family == "axis" and key == "arm":
            continue
        require(run.get(key) == item, f"{path}.run.{key}: inventory binding differs")
    require(run.get("kind") == "formal" and run.get("label") == expected["label"], f"{path}.run: formal identity differs")
    protocol_ref = object_value(value.get("protocol"), f"{path}.protocol")
    require(protocol_ref.get("path") == PROTOCOL_FILE and protocol_ref.get("sha256") == protocol_hash, f"{path}: protocol binding differs")
    require(value.get("machine") == protocol.get("machine"), f"{path}: machine binding differs")
    require(value.get("environment") == protocol.get("environment"), f"{path}: environment binding differs")
    require(value.get("cwd") == str(routes.REPO), f"{path}: working directory differs")
    build_ref = object_value(value.get("build"), f"{path}.build")
    role = expected["role"]
    expected_build_path = Path("route-attempts") / attempt / f"build-{role}.json"
    require(build_ref.get("path") == expected_build_path.as_posix(), f"{path}: build path differs")
    build_path = bundle_root / expected_build_path
    require(build_path.is_file() and sha(build_path) == digest(build_ref.get("sha256"), f"{path}.build.sha256"), f"{path}: build receipt hash differs")
    require(builds[role].get("source_after") == dict(source), f"{path}: build source differs")
    require(value.get("binary") == builds[role].get("binary"), f"{path}: binary binding differs")
    report_path = _artifact(path.parent, value, "report.json")
    resource_path = _artifact(path.parent, value, "resource.txt")
    _artifact(path.parent, value, "stdout.txt")
    _artifact(path.parent, value, "stderr.txt")
    binary = object_value(value.get("binary"), f"{path}.binary")
    report = object_value(read_json(report_path), str(report_path))
    expected_argv = (
        routes._route_argv(binary, routes.ROUTE_CASE_BY_LABEL[expected["case"]], routes.ROUTE_BY_NAME[expected["route"]], samples=FORMAL_SAMPLES, warmups=FORMAL_WARMUPS, report=report_path, resource=resource_path, replay_dir=path.parent / "replay" if expected.get("route") == "file_store" else None)
        if family == "route" else routes._axis_argv(binary, routes._axis_case(routes.AXIS_ARM_BY_LABEL[expected["arm"]]), routes.AXIS_ARM_BY_LABEL[expected["arm"]], samples=FORMAL_SAMPLES, warmups=FORMAL_WARMUPS, report=report_path, resource=resource_path)
    )
    require(value.get("argv") == expected_argv, f"{path}: argv binding differs")
    _canonical_report_check(report_path, resource_path, path, expected, family, binary, expected_argv, prepared_inputs)
    if family == "route":
        _validate_replay_directory(bundle_root, path, expected["route"], cleanup_manifest)
    if family == "axis":
        input_file = value.get("input_file")
        if expected["input_mode"] == "file":
            expected_input = prepared_inputs[expected["input_file"]]
            input_file = object_value(input_file, f"{path}.input_file")
            input_path = bundle_root / expected["input_file"]
            require(
                input_file.get("path") == expected["input_file"]
                and input_file.get("absolute_path") == str(input_path.resolve())
                and input_file.get("bytes") == expected_input["bytes"]
                and input_file.get("sha256") == expected_input["sha256"]
                and input_file.get("identity") == "prepared_file_capability_fingerprint_must_match_report_source_archive",
                f"{path}: staged file input binding differs",
            )
        else:
            require(input_file is None, f"{path}: non-file axis has file metadata")
    identities, _, _ = _validate_report_shape(report, expected, family, role, str(report_path))
    if family == "axis" and expected["input_mode"] == "file":
        prepared = prepared_inputs[expected["input_file"]]
        require(identities["source"]["archive_bytes"] == prepared["bytes"], f"{path}: report source bytes differ from prepared input")
        require(identities["source"]["archive_sha256"] == prepared["sha256"], f"{path}: report source archive differs from prepared input")
        if prepared.get("main_xml_sha256") is not None:
            require(identities["source"]["main_xml_sha256"] == prepared["main_xml_sha256"], f"{path}: report source XML differs from prepared manifest")
    return {"receipt": value, "report": report, "report_path": report_path, "resource_path": resource_path}


def _inventory_paths(bundle_root: Path, family: str, attempt: str, expected_rows: list[dict[str, Any]]) -> None:
    directory = bundle_root / ("route-captures" if family == "route" else "axis-captures") / attempt
    require(directory.is_dir(), f"{family} formal capture directory is missing: {directory}")
    expected_labels = {row["label"] for row in expected_rows}
    actual_dirs = {item.name for item in directory.iterdir() if item.is_dir()}
    require(actual_dirs == expected_labels, f"{family} formal receipt inventory differs")
    for item in directory.iterdir():
        require(item.is_dir() and (item / "receipt.json").is_file(), f"{family} formal run directory is incomplete: {item}")


def analyze(protocol_path: Path, attempt: str) -> dict[str, Any]:
    protocol_path = protocol_path.resolve()
    protocol, protocol_hash = validate_protocol(protocol_path)
    bundle_root = protocol_path.parent
    require(re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_.-]*", attempt) is not None, "formal attempt is malformed")
    builds: dict[str, dict[str, Any]] = {}
    source: dict[str, Any] | None = None
    expected_environment = object_value(protocol.get("environment"), "protocol.environment")
    for role in ROLES:
        build, build_source = _validate_build(bundle_root, attempt, role, protocol_hash, expected_environment, source)
        builds[role] = build
        if source is None:
            source = build_source
    require(source is not None, "formal build source is missing")
    require(builds["normal"]["source_after"] == builds["allocator"]["source_after"] == source, "normal/allocator build source identities differ")
    prep_path, prepared_inputs = _validate_axis_preparation(bundle_root, attempt, protocol_hash, source, builds["normal"])
    route_gate_hash, route_gate = _validate_lane_gate(bundle_root, attempt, "route-captures", source)
    axis_gate_hash, axis_gate = _validate_lane_gate(bundle_root, attempt, "axis-captures", source)
    execution = _validate_execution_inputs(bundle_root, attempt, protocol_hash, source, prep_path, builds)
    cleanup_manifest = _validate_cleanup_manifest(bundle_root, attempt)
    execution_recorded = _timestamp(execution.get("recorded_utc"), "execution-inputs.recorded_utc")
    require(execution_recorded <= _timestamp(route_gate.get("started_utc"), "route-captures gate.started_utc"), "execution-inputs was recorded after route capture gate start")
    require(execution_recorded <= _timestamp(axis_gate.get("started_utc"), "axis-captures gate.started_utc"), "execution-inputs was recorded after axis capture gate start")
    lane_gate_hashes = {"route_captures": route_gate_hash, "axis_captures": axis_gate_hash}
    route_expected = [dict(item) for item in protocol["formal_runs"]]
    axis_expected = [dict(item) for item in protocol["axis_formal_runs"]]
    _inventory_paths(bundle_root, "route", attempt, route_expected)
    _inventory_paths(bundle_root, "axis", attempt, axis_expected)
    route_rows: list[dict[str, Any]] = []
    axis_rows: list[dict[str, Any]] = []
    for family, expected_rows, output in (("route", route_expected, route_rows), ("axis", axis_expected, axis_rows)):
        for expected in expected_rows:
            path = _receipt_expected_path(bundle_root, family, attempt, expected)
            retained = _validate_receipt(path, expected, family, attempt, protocol, protocol_hash, bundle_root, source, builds, prepared_inputs, cleanup_manifest)
            output.append(_summarize_report(retained["report"], retained["resource_path"], expected, family))
    require(len(route_rows) == 120 and len(axis_rows) == 108, "formal process cardinality differs")
    validate_identity_consistency(route_rows + axis_rows)
    route_repeat = _repeat_stats(route_rows)
    axis_repeat = _repeat_stats(axis_rows)
    return {
        "schema": SUMMARY_SCHEMA,
        "version": 1,
        "analyzer": {"path": Path(__file__).name, "sha256": sha(Path(__file__))},
        "protocol": {"path": protocol_path.name, "sha256": protocol_hash, "attempt": attempt},
        "execution_inputs": {"path": str((Path("route-attempts") / attempt / EXECUTION_INPUTS_FILE).as_posix()), "sha256": sha(bundle_root / "route-attempts" / attempt / EXECUTION_INPUTS_FILE)},
        "lane_gate_hashes": lane_gate_hashes,
        "cleanup_manifest": None if cleanup_manifest is None else {
            "path": cleanup_manifest["path"],
            "sha256": cleanup_manifest["sha256"],
            "entries": len(cleanup_manifest["entries"]),
        },
        "source_manifest": source,
        "inventory": {
            "route_processes": len(route_rows),
            "axis_processes": len(axis_rows),
            "route_samples": len(route_rows) * FORMAL_SAMPLES,
            "axis_samples": len(axis_rows) * FORMAL_SAMPLES,
            "roles": list(ROLES),
            "repeats": list(REPEATS),
            "pilots_included": False,
        },
        "route_processes": route_rows,
        "axis_processes": axis_rows,
        "route_repeat_stats": route_repeat,
        "axis_repeat_stats": axis_repeat,
        "route_comparisons": route_comparisons(route_rows),
        "axis_comparisons": axis_comparisons(axis_rows, route_rows),
        "claims": {
            "performance_claim": "none",
            "causal_speedup": False,
            "route_comparison": "contemporaneous route comparison only",
            "bounded_memory_global": False,
            "rss_scope": "whole-process GNU time maximum; not a bounded-memory proof",
            "source_and_authored_evidence": "partitioned by exact workload case; no cross-case averaging",
            "uncertainty": "two independent process repeats summarized by repeat range; no confidence interval",
            "review_threshold_percent": REVIEW_THRESHOLD_PERCENT,
        },
        "limitations": [
            "Only two process repeats are retained; repeat spread is descriptive and no strong confidence interval is reported.",
            "RSS is a whole-process maximum and does not establish a global bounded-memory claim.",
            "Source-varying and authored-varying workloads remain separate exact-case evidence.",
            "Comparisons above 5% are review flags and do not establish causal before/after speedups.",
        ],
        "execution": execution,
    }


def parser() -> argparse.ArgumentParser:
    command = argparse.ArgumentParser(description=__doc__)
    command.add_argument("--protocol", type=Path, default=ROOT / PROTOCOL_FILE)
    command.add_argument("--attempt", required=True)
    command.add_argument("--summary", type=Path, default=ROOT / "route-summary.json")
    return command


def main() -> None:
    args = parser().parse_args()
    summary = analyze(args.protocol.resolve(), args.attempt)
    write_json(args.summary.resolve(), summary)
    print(f"analyzed {summary['inventory']['route_processes']} route and {summary['inventory']['axis_processes']} axis formal processes")


if __name__ == "__main__":
    main()
