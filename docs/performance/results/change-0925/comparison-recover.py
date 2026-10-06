#!/usr/bin/env python3
"""Recover descriptive evidence from a failed 0925 terminal receipt.

This command consumes an existing packet only.  It verifies the packet-local
scripts, manifest custody, all raw child reports and GNU ``time`` logs, then
writes a compact descriptive reduction.  It never treats the result as the
original capture state and never starts or reruns a child.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import re
import shlex
import sys
from pathlib import Path
from typing import Any


MAX_ARTIFACT_BYTES = 1 << 20
RSS_RE = re.compile(r"Maximum resident set size \(kbytes\):\s*(\d+)")
EXIT_RE = re.compile(r"Exit status:\s*(\d+)")
COMMAND_RE = re.compile(r'^\s*Command being timed: "(.*)"\s*$', re.MULTILINE)


class RecoveryError(RuntimeError):
    """A raw packet or recovery contract failed closed."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RecoveryError(message)


def load_module(path: Path, name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None, f"cannot load packet script: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def load_packet_tools(packet: Path) -> tuple[Any, Any]:
    scripts = packet / "scripts"
    capture = load_module(scripts / "comparison-capture.py", "litchi_perf_0925_recovery_capture")
    comparison = capture.load_comparison(scripts / "comparison.py")
    return capture, comparison


def descriptor(path: Path, packet: Path | None = None) -> dict[str, Any]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    value: dict[str, Any] = {
        "path": str(path),
        "bytes": path.stat().st_size,
        "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
    }
    if packet is not None:
        try:
            value["packet_path"] = str(path.relative_to(packet.resolve()))
        except ValueError as error:
            raise RecoveryError(f"file is outside packet: {path}") from error
    return value


def packet_descriptor(path: Path, packet: Path, label: str) -> dict[str, Any]:
    path = path.resolve()
    packet = packet.resolve()
    try:
        relative = path.relative_to(packet)
    except ValueError as error:
        raise RecoveryError(f"{label}: file is outside packet") from error
    require(path.is_file() and not path.is_symlink(), f"{label}: file is missing")
    return {
        "path": str(relative),
        "bytes": path.stat().st_size,
        "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
    }


def load_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as error:
        raise RecoveryError(f"cannot read JSON {path}: {error}") from error
    require(isinstance(value, dict), f"expected JSON object: {path}")
    return value


def parse_time_log(path: Path, expected_report: Path, row: dict[str, Any],
                   manifest: dict[str, Any], capture: Any) -> dict[str, Any]:
    text = path.read_text(encoding="utf-8", errors="replace")
    rss_match = RSS_RE.search(text)
    exit_match = EXIT_RE.search(text)
    command_match = COMMAND_RE.search(text)
    require(rss_match is not None, f"{path}: RSS is unavailable")
    require(exit_match is not None, f"{path}: exit status is unavailable")
    require(command_match is not None, f"{path}: timed command is unavailable")
    command = command_match.group(1)
    try:
        observed_argv = shlex.split(command)
    except ValueError as error:
        raise RecoveryError(f"{path}: timed command is not shell-parseable") from error
    expected_argv = capture._child_command(row, manifest, expected_report)[2:]
    require(len(observed_argv) == len(expected_argv),
            f"{path}: timed command argument count changed")
    # The binary, case, sample, warmup, and selector prefix is immutable.
    # Only the report path may differ after a packet is moved for replay.
    require(observed_argv[:-1] == expected_argv[:-1],
            f"{path}: timed command prefix differs from frozen child binding")
    observed_report = Path(observed_argv[-1])
    require(observed_report.is_absolute() and ".." not in observed_report.parts,
            f"{path}: timed report path is not an absolute retained path")
    require(observed_report.name == expected_report.name
            and observed_report.parent.name == expected_report.parent.name,
            f"{path}: timed report child or basename differs from retained report")
    require(len(observed_report.parts) >= 3
            and observed_report.parts[-3] == "children",
            f"{path}: timed report path lacks the frozen children/<id> shape")
    return {
        "rss_bytes": int(rss_match.group(1)) * 1024,
        "exit_code": int(exit_match.group(1)),
        "command_sha256": hashlib.sha256(command.encode()).hexdigest(),
        "command_prefix_sha256": hashlib.sha256(
            json.dumps(observed_argv[:-1], separators=(",", ":")).encode()
        ).hexdigest(),
        "report_child_id": observed_report.parent.name,
        "report_basename": observed_report.name,
    }


def validate_packet_custody(packet: Path, manifest: dict[str, Any], capture: Any,
                            comparison: Any) -> None:
    require(manifest.get("schema") == "litchi.performance.0925.provider-capture-manifest.v1",
            "packet manifest schema changed")
    require(manifest.get("capture_authorized") is True,
            "packet manifest was not authorized for the retained attempt")
    require(manifest.get("formal_capture_children_run") == 0,
            "packet manifest claims completed children")
    scripts = manifest.get("scripts")
    local_scripts = manifest.get("packet_scripts")
    require(isinstance(scripts, dict) and isinstance(local_scripts, dict)
            and set(scripts) == {"comparison", "controls", "capture", "reduce"}
            and set(local_scripts) == set(scripts),
            "packet script custody set changed")
    for name, original in scripts.items():
        local = local_scripts[name]
        require(local.get("bytes") == original.get("bytes")
                and local.get("sha256") == original.get("sha256"),
                f"packet script descriptor changed: {name}")
        capture.verify_local_descriptor(packet, local, f"packet script {name}")
    capture._verify_packet_custody(manifest, packet)
    capture._validate_local_custody_semantics(manifest, packet, comparison)


def validate_failure_receipt(failure_path: Path, packet: Path, manifest: dict[str, Any],
                            authorized_manifest_path: Path, capture: Any) -> dict[str, Any]:
    failure = load_json(failure_path)
    require(failure.get("schema") == "litchi.performance.0925.capture-terminal-failure.v1",
            "capture failure receipt schema changed")
    require(failure.get("status") == "invalid"
            and failure.get("root_exit_code") == 1
            and failure.get("reports_present") == 42
            and failure.get("capture_state_present") is False,
            "capture failure receipt terminal fields changed")
    terminal = str(failure.get("terminal_output"))
    require("receipt is at least 1 MiB" in terminal and "capture-state.json" in terminal,
            "capture failure receipt does not bind the aggregate-size failure")
    authorized_manifest_path = authorized_manifest_path.resolve()
    authorized_manifest = load_json(authorized_manifest_path)
    authorized_sha256 = hashlib.sha256(authorized_manifest_path.read_bytes()).hexdigest()
    require(failure.get("authorized_manifest_sha256") == authorized_sha256,
            "capture failure receipt manifest hash differs from authorized manifest")
    for key in ("schema", "base_revision", "source", "build", "protocol", "scripts",
                "binary", "schedule", "process_deadline_seconds", "selector", "claim_scope",
                "failure_policy", "capture_authorized"):
        require(authorized_manifest.get(key) == manifest.get(key),
                f"packet manifest differs from authorized manifest: {key}")
    report_rows = failure.get("reports")
    require(isinstance(report_rows, list) and len(report_rows) == 42,
            "capture failure report inventory changed")
    by_path = {str(row.get("path")): row for row in report_rows}
    require(len(by_path) == 42, "capture failure report inventory has duplicate paths")
    return {
        "receipt": descriptor(failure_path),
        "schema": failure["schema"],
        "status": failure["status"],
        "root_exit_code": failure["root_exit_code"],
        "reports_present": failure["reports_present"],
        "capture_state_present": failure["capture_state_present"],
        "terminal_output": terminal,
        "failure_scope": failure.get("failure_scope"),
        "recovery_policy": failure.get("recovery_policy"),
        "report_inventory": report_rows,
    }


def recover(packet: Path, failure_path: Path, authorized_manifest_path: Path) -> dict[str, Any]:
    packet = packet.resolve()
    failure_path = failure_path.resolve()
    require(packet.is_dir(), f"packet is not a directory: {packet}")
    capture, comparison = load_packet_tools(packet)
    manifest_path = packet / "manifest.json"
    manifest = load_json(manifest_path)
    validate_packet_custody(packet, manifest, capture, comparison)
    require(authorized_manifest_path.is_file(),
            f"authorized manifest is missing: {authorized_manifest_path}")
    failure = validate_failure_receipt(
        failure_path, packet, manifest, authorized_manifest_path, capture
    )

    rows = manifest.get("schedule", {}).get("rows")
    expected_rows = comparison.schedule_rows()
    require(rows == expected_rows, "packet schedule differs from frozen schedule")
    children_root = packet / "children"
    child_dirs = {path.name for path in children_root.iterdir() if path.is_dir()}
    expected_ids = {row["id"] for row in rows}
    require(child_dirs == expected_ids, "raw packet child directory inventory changed")

    summaries: dict[str, list[dict[str, Any]]] = {route: [] for route in comparison.ROUTES}
    summary_by_id: dict[str, dict[str, Any]] = {}
    rss_values: dict[str, list[int]] = {route: [] for route in comparison.ROUTES}
    raw_descriptors: dict[str, list[dict[str, Any]]] = {route: [] for route in comparison.ROUTES}
    raw_by_id: dict[str, dict[str, Any]] = {}
    validation_payload_bytes = 0
    raw_report_bytes = 0
    raw_report_count = 0
    for row in rows:
        child_id = row["id"]
        child = children_root / child_id
        expected_files = {"report.json", "stdout.log", "stderr.log", "time.log"}
        actual_files = {path.name for path in child.iterdir() if path.is_file()}
        require(actual_files == expected_files, f"{child_id}: raw child file inventory changed")
        report_path = child / "report.json"
        stdout_path = child / "stdout.log"
        stderr_path = child / "stderr.log"
        time_path = child / "time.log"
        report_desc = packet_descriptor(report_path, packet, f"{child_id} report")
        failure_row = next((value for value in failure.get("report_inventory", [])
                            if value.get("path") == f"children/{child_id}/report.json"), None)
        require(isinstance(failure_row, dict)
                and report_desc["bytes"] == failure_row.get("bytes")
                and report_desc["sha256"] == failure_row.get("sha256"),
                f"{child_id}: report differs from retained failure inventory")
        report = load_json(report_path)
        summary = comparison.validate_report(
            report, row["route"], warmups=row["warmups"], samples=row["samples"],
            label=child_id, binary=manifest["binary"]
        )
        stdout_desc = packet_descriptor(stdout_path, packet, f"{child_id} stdout")
        stderr_desc = packet_descriptor(stderr_path, packet, f"{child_id} stderr")
        time_desc = packet_descriptor(time_path, packet, f"{child_id} time")
        require(stderr_path.read_bytes() == time_path.read_bytes(),
                f"{child_id}: stderr/time streams differ")
        timed = parse_time_log(time_path, report_path, row, manifest, capture)
        require(timed["exit_code"] == 0, f"{child_id}: retained process exit was nonzero")
        rss_values[row["route"]].append(timed["rss_bytes"])
        summaries[row["route"]].append(summary)
        summary_by_id[child_id] = summary
        validation_payload_bytes += len(json.dumps(summary, separators=(",", ":"), sort_keys=True).encode())
        raw_report_bytes += report_desc["bytes"]
        raw_report_count += 1
        raw_descriptor = {
            "id": child_id,
            "block": row["block"],
            "position": row["position"],
            "order": row["order"],
            "report": report_desc,
            "stdout": stdout_desc,
            "stderr": stderr_desc,
            "time": time_desc,
            "observed_exit_code": timed["exit_code"],
            "observed_rss_bytes": timed["rss_bytes"],
            "timed_command_sha256": timed["command_sha256"],
            "timed_command_prefix_sha256": timed["command_prefix_sha256"],
            "timed_report_child_id": timed["report_child_id"],
            "timed_report_basename": timed["report_basename"],
        }
        raw_descriptors[row["route"]].append(raw_descriptor)
        raw_by_id[child_id] = raw_descriptor

    reduced = comparison.reduce_reports(summaries)
    for route in comparison.ROUTES:
        require(len(summaries[route]) == 6 and len(rss_values[route]) == 6,
                f"{route}: raw child count changed")
        reduced_route = reduced["routes"][route]
        reduced_route["child_rss_bytes"] = {
            "values": rss_values[route],
            "statistics": comparison._statistics(rss_values[route]),
        }
        reduced_route["raw_child_descriptors"] = raw_descriptors[route]
    reduction_bytes = len(json.dumps(reduced, separators=(",", ":"), sort_keys=True).encode())
    require(reduction_bytes < MAX_ARTIFACT_BYTES,
            f"reconstructed aggregate is at least 1 MiB: {reduction_bytes}")

    # This is a lower-bound diagnostic for the failed original state write.  It
    # serializes the actual validated vectors and packet descriptors while
    # omitting unknown start/end/cwd fields, so it does not fabricate custody.
    state_lower_bound = {
        "schema": "litchi.performance.0925.provider-capture-state.v1",
        "children": [
            {
                "schedule": row,
                "report": raw_by_id[row["id"]]["report"],
                "stdout": raw_by_id[row["id"]]["stdout"],
                "stderr_and_time": raw_by_id[row["id"]]["stderr"],
                "time": raw_by_id[row["id"]]["time"],
                "rss_bytes": raw_by_id[row["id"]]["observed_rss_bytes"],
                "exit_code": 0,
                "validation": summary_by_id[row["id"]],
                "status": "passed",
            }
            for index, row in enumerate(rows)
        ],
        "children_run": 42,
    }
    state_lower_bound_bytes = len(
        json.dumps(state_lower_bound, separators=(",", ":"), sort_keys=True).encode()
    )

    # This is the compact shape a future runner can persist: it binds each
    # raw stream and its validated summary digest without duplicating every
    # phase/provider vector in the terminal receipt.  It is measured here
    # only as a control; the failed packet has no replacement state file.
    compact_state = {
        "schema": "litchi.performance.0925.provider-capture-state-compact.v1",
        "children": [
            {
                "id": row["id"],
                "route": row["route"],
                "block": row["block"],
                "position": row["position"],
                "order": row["order"],
                "report": raw_by_id[row["id"]]["report"],
                "stdout": raw_by_id[row["id"]]["stdout"],
                "stderr_and_time": raw_by_id[row["id"]]["stderr"],
                "time": raw_by_id[row["id"]]["time"],
                "rss_bytes": raw_by_id[row["id"]]["observed_rss_bytes"],
                "exit_code": raw_by_id[row["id"]]["observed_exit_code"],
                "validation_status": "passed",
                "validation_summary_sha256": hashlib.sha256(
                    json.dumps(summary_by_id[row["id"]], sort_keys=True,
                               separators=(",", ":")).encode()
                ).hexdigest(),
            }
            for row in rows
        ],
        "children_run": len(rows),
        "status": "completed",
    }
    compact_state_bytes = len(
        json.dumps(compact_state, separators=(",", ":"), sort_keys=True).encode()
    )
    recovery_tool = descriptor(Path(__file__))
    actual_report_aggregate = {
        "status": "passed",
        "report_count": raw_report_count,
        "raw_report_bytes": raw_report_bytes,
        "descriptive_reduction_bytes": reduction_bytes,
        "writer_limit_bytes": MAX_ARTIFACT_BYTES,
        "under_one_mib": reduction_bytes < MAX_ARTIFACT_BYTES,
        "raw_reports_external_only": True,
    }
    return {
        "schema": "litchi.performance.0925.reconstructed-raw-reduction.v1",
        "status": "invalid_terminal_reconstructed",
        "claim_scope": "descriptive raw evidence only; no completed-capture custody, timing adoption, ranking, or recapture",
        "performance_children_run": 0,
        "terminal": {
            "original_capture_state": "not_written",
            "status": "invalid",
            "reason": "all 42 child artifacts were retained, but the original runner failed writing capture-state.json because its receipt exceeded the 1 MiB limit",
            "recovery_rule": "posthoc reconstruction cannot restore original per-child start/end/cwd/validation receipt custody; no child was rerun or replaced",
        },
        "packet": descriptor(packet / "manifest.json", packet),
        "failure": failure,
        "source_build_protocol_scripts": {
            "scripts": manifest["packet_scripts"],
            "recovery_tool": recovery_tool,
            "binary": manifest["binary"],
            "source": {
                "revision": manifest["base_revision"],
                "candidate_file": manifest["source"]["candidate_file"],
                "candidate_file_sha256": manifest["source"]["candidate_file_sha256"],
                "candidate_patch_sha256": manifest["source"]["candidate_patch_sha256"],
            },
        },
        "raw_inventory": {
            "children": raw_report_count,
            "reports": raw_report_count,
            "report_bytes": raw_report_bytes,
            "packet_raw_bytes": sum(
                path.stat().st_size
                for path in (packet / "children").rglob("*")
                if path.is_file()
            ),
        },
        "state_size_diagnostic": {
            "writer_limit_bytes": MAX_ARTIFACT_BYTES,
            "validated_vector_payload_bytes": validation_payload_bytes,
            "state_shape_lower_bound_bytes": state_lower_bound_bytes,
            "lower_bound_under_writer_limit": state_lower_bound_bytes < MAX_ARTIFACT_BYTES,
            "fields_omitted_from_lower_bound": [
                "schema", "argv", "environment", "cwd", "started_utc", "ended_utc",
                "observation_timeout_count", "deadline_exceeded"
            ],
            "interpretation": "the bound includes the validated vectors and retained stream descriptors but omits runner receipt fields; it is below the writer limit, so no exact original serialized size is claimed. The retained terminal failure remains authoritative, and posthoc recovery cannot restore the omitted custody fields",
        },
        "aggregate_size_control": {
            "status": "passed",
            "serialized_reduction_bytes": reduction_bytes,
            "writer_limit_bytes": MAX_ARTIFACT_BYTES,
            "children": 42,
            "routes": len(comparison.ROUTES),
            "under_one_mib": True,
        },
        "actual_report_aggregate_size_control": actual_report_aggregate,
        "future_compact_state_control": {
            "status": "passed",
            "schema": compact_state["schema"],
            "serialized_bytes": compact_state_bytes,
            "writer_limit_bytes": MAX_ARTIFACT_BYTES,
            "under_one_mib": compact_state_bytes < MAX_ARTIFACT_BYTES,
            "not_written_to_failed_packet": True,
            "omits_full_validation_vectors": True,
            "omits_unknown_original_fields": [
                "started_utc", "ended_utc", "cwd", "observation_timeout_count"
            ],
            "purpose": "future runner design control only; no replacement capture state",
        },
        "reduction": reduced,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--packet", type=Path, required=True)
    parser.add_argument("--failure-receipt", type=Path, required=True)
    parser.add_argument("--authorized-manifest", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        output = recover(args.packet, args.failure_receipt, args.authorized_manifest)
        data = (json.dumps(output, indent=2, sort_keys=True) + "\n").encode()
        require(len(data) < MAX_ARTIFACT_BYTES, "recovery receipt is at least 1 MiB")
        output_path = args.output.resolve()
        require(not output_path.exists(), f"refusing to replace retained receipt: {output_path}")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_bytes(data)
        print(json.dumps({
            "status": output["status"],
            "output": str(output_path),
            "bytes": len(data),
            "children": output["raw_inventory"]["children"],
            "serialized_reduction_bytes": output["aggregate_size_control"]["serialized_reduction_bytes"],
            "performance_children_run": 0,
        }, sort_keys=True))
    except Exception as error:
        print(f"0925 recovery failed: {type(error).__name__}: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
