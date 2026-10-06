#!/usr/bin/env python3
"""Prospective 0931 DOCX retention-memory runner and offline reducer.

The binary route is intentionally supplied by the allocator-only diagnostic
target.  This file owns only schedule/custody, terminal receipts, checkpoint
conservation, and descriptive reduction.  It makes no latency or adoption
decision and refuses to run until an explicitly authorized manifest and the
capture environment approval are present.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import signal
import statistics
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


SCHEMA = "litchi.performance.0931.memory-manifest.v1"
REPORT_SCHEMA = "docx_section_retention_v1"
STATE_SCHEMA = "litchi.performance.0931.memory-state.v1"
TERMINAL_SCHEMA = "litchi.performance.0931.memory-terminal.v1"
VALIDATION_SCHEMA = "litchi.performance.0931.memory-validation.v1"
MAX_JSON_BYTES = 1_048_576
DEFAULT_CAP = 1_048_576
PROBE_CAP = 64 * 1024 * 1024
SECTIONS = (16, 8192)
COMMITS = (1, 8)
POLICIES = ("default", "zero", "exact", "one_under")
OWNERSHIPS = ("shared_snapshot_clones", "independent_sources")
REPETITIONS = 3
CPU = 12
DEADLINE_SECONDS = 120
APPROVAL_ENV = "LITCHI_PERF_0931_CAPTURE_APPROVED"
APPROVED_CHILD_ENV = {
    "PATH": "/usr/bin:/bin",
    "LC_ALL": "C",
    "LANG": "C",
    "RUST_BACKTRACE": "0",
    "LITCHI_PERF_0931_MEMORY_DIAGNOSTIC": "1",
}
REPO_ROOT = Path(__file__).resolve().parents[4]


class EvidenceError(RuntimeError):
    """A fail-closed custody or semantic validation error."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="microseconds")


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


def safe_relative(value: str, label: str) -> Path:
    path = Path(value)
    require(not path.is_absolute(), f"{label}: absolute path is not portable")
    require(path.parts and all(part not in ("", ".", "..") for part in path.parts),
            f"{label}: unsafe relative path {value!r}")
    return path


def read_json(path: Path, *, limit: int | None = MAX_JSON_BYTES) -> dict[str, Any]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    if limit is not None:
        require(path.stat().st_size < limit,
                f"JSON exceeds evidence bound: {path} ({path.stat().st_size} bytes)")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as error:
        raise EvidenceError(f"cannot parse JSON {path}: {error}") from error
    require(isinstance(value, dict), f"JSON is not an object: {path}")
    return value


def write_new_json(path: Path, value: Any, *, limit: int = MAX_JSON_BYTES) -> None:
    data = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    require(len(data) < limit, f"output exceeds evidence bound: {path} ({len(data)} bytes)")
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("xb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError as error:
        raise EvidenceError(f"refusing to overwrite immutable JSON: {path}") from error


def schedule_rows(source_revision: str) -> list[dict[str, Any]]:
    require(len(source_revision) == 40 and all(c in "0123456789abcdef" for c in source_revision),
            "schedule source revision must be a lowercase 40-hex Git revision")
    rows: list[dict[str, Any]] = []
    for sections in SECTIONS:
        for commits in COMMITS:
            for policy in POLICIES:
                for ownership in OWNERSHIPS:
                    for repetition in range(1, REPETITIONS + 1):
                        row_id = (f"s{sections:04d}-c{commits:02d}-{policy}-"
                                  f"{ownership}-r{repetition:02d}")
                        rows.append({
                            "id": row_id,
                            "source_revision": source_revision,
                            "section_count": sections,
                            "concurrent_commits": commits,
                            "policy": policy,
                            "ownership": ownership,
                            "repetition": repetition,
                            "cpu": CPU,
                            "deadline_seconds": DEADLINE_SECONDS,
                        })
    require(len(rows) == 96, "0931 schedule cardinality changed")
    require(len({row["id"] for row in rows}) == len(rows), "0931 schedule IDs are not unique")
    return rows


def schedule_hash(rows: list[dict[str, Any]]) -> str:
    encoded = json.dumps(rows, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def validate_child_environment(environment: Any) -> None:
    require(environment == APPROVED_CHILD_ENV,
            "child environment differs from the approved key/value allowlist")


def verify_descriptor(value: dict[str, Any], *, label: str, root: Path | None = None) -> Path:
    require(isinstance(value, dict), f"{label}: descriptor is not an object")
    raw = value.get("path")
    require(isinstance(raw, str), f"{label}: descriptor path missing")
    path = ((root / safe_relative(raw, label)).resolve()
            if root is not None else Path(raw).resolve())
    actual = descriptor(path)
    require(actual["bytes"] == value.get("bytes")
            and actual["sha256"] == value.get("sha256"),
            f"{label}: descriptor changed")
    return path


def expected_cap(policy: str, admission: dict[str, Any]) -> int:
    required = admission.get("required_retained_base_bytes")
    require(type(required) is int and required > 0, "retention admission is not positive")
    require(admission.get("probe_ceiling_bytes") == PROBE_CAP,
            "retention probe ceiling changed")
    require(admission.get("zero_cap_bytes") == 0
            and admission.get("zero_retained_base_bytes") == 0,
            "zero-cap admission changed")
    require(admission.get("exact_cap_bytes") == required
            and admission.get("exact_retained_base_bytes") == required,
            "exact-cap admission changed")
    require(admission.get("one_under_cap_bytes") == required - 1
            and admission.get("one_under_retained_base_bytes") == 0,
            "one-under-cap admission changed")
    default_retained = required if required <= DEFAULT_CAP else 0
    require(admission.get("default_cap_bytes") == DEFAULT_CAP
            and admission.get("default_retained_base_bytes") == default_retained,
            "default-cap admission changed")
    require(admission.get("boundary_checks_verified") is True,
            "retention boundary controls are not verified")
    return {"default": DEFAULT_CAP, "zero": 0, "exact": required,
            "one_under": required - 1}[policy]


def check_gauge(value: dict[str, Any], label: str) -> None:
    require(isinstance(value, dict), f"{label}: gauge is not an object")
    names = ("allocation_calls", "deallocation_calls", "reallocation_calls",
             "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
             "live_bytes", "peak_live_bytes")
    for name in names:
        require(type(value.get(name)) is int and value[name] >= 0,
                f"{label}: invalid {name}")
    require(value["allocated_bytes"] >= value["deallocated_bytes"],
            f"{label}: allocator balance underflow")
    require(value["live_bytes"] == value["allocated_bytes"] - value["deallocated_bytes"],
            f"{label}: live-byte balance changed")
    require(value["peak_live_bytes"] >= value["live_bytes"],
            f"{label}: peak is below live bytes")
    require(value.get("overflowed") is False and value.get("observer_invalid") is False,
            f"{label}: allocator observer is invalid or overflowed")
    region_peak = value.get("region_peak_live_bytes")
    require(region_peak is None or (type(region_peak) is int and region_peak >= 0),
            f"{label}: invalid region peak")


def check_phase(value: dict[str, Any], baseline: dict[str, Any], label: str) -> None:
    require(isinstance(value, dict), f"{label}: phase is not an object")
    absolute = value.get("absolute")
    check_gauge(absolute, f"{label}.absolute")
    expected_delta = absolute["live_bytes"] - baseline["live_bytes"]
    require(value.get("live_bytes_delta_from_baseline") == expected_delta,
            f"{label}: live-byte delta is inconsistent")
    region_peak = absolute.get("region_peak_live_bytes")
    if label == "commits_retained":
        require(region_peak is None, "held commit checkpoint must not claim a region peak")
    else:
        require(type(region_peak) is int and region_peak >= absolute["live_bytes"],
                f"{label}: region peak does not cover its endpoint live bytes")


def check_allocator_sequence(gauges: dict[str, Any]) -> None:
    phases = ["after_snapshot_drop", "commits_retained", "after_release_all_but_last",
              "after_release_last", "after_commit_drop"]
    sequence = [gauges["baseline"]]
    sequence.extend(gauges[name]["absolute"] for name in phases)
    monotonic = ("allocation_calls", "deallocation_calls", "reallocation_calls",
                 "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
                 "peak_live_bytes")
    for index, (previous, current) in enumerate(zip(sequence, sequence[1:])):
        for field in monotonic:
            require(current[field] >= previous[field],
                    f"allocator counter {field} is not monotonic")
        expected_live = previous["live_bytes"] + (
            current["allocated_bytes"] - previous["allocated_bytes"]
        ) - (current["deallocated_bytes"] - previous["deallocated_bytes"])
        require(current["live_bytes"] == expected_live,
                "allocator live-byte sequence is not balanced")
        region_peak = current.get("region_peak_live_bytes")
        if phases[index] != "commits_retained":
            require(type(region_peak) is int and region_peak >= previous["live_bytes"],
                    f"{phases[index]} region peak does not cover phase-begin live bytes")
    require(gauges["returned_to_baseline"] is True,
            "commit drop did not return allocator live bytes to baseline")


def validate_report(row: dict[str, Any], report: dict[str, Any]) -> dict[str, Any]:
    require(report.get("schema") == REPORT_SCHEMA, "unsupported retention report schema")
    for name in ("section_count", "concurrent_commits", "retention_policy", "ownership_mode"):
        expected = {"section_count": row["section_count"],
                    "concurrent_commits": row["concurrent_commits"],
                    "retention_policy": row["policy"],
                    "ownership_mode": row["ownership"]}[name]
        require(report.get(name) == expected, f"report {name} differs from schedule")
    require(report.get("source_revision") == row["source_revision"],
            "report source revision differs from schedule")
    source_hash = report.get("source_archive_sha256")
    require(isinstance(source_hash, str) and re.fullmatch(r"[0-9a-f]{64}", source_hash),
            "report source archive hash is malformed")
    require(type(report.get("source_archive_bytes")) is int
            and report["source_archive_bytes"] > 0,
            "report source archive size is missing")
    output_hashes = report.get("expected_output_sha256")
    output_bytes = report.get("expected_output_bytes")
    require(isinstance(output_hashes, list)
            and len(output_hashes) == row["concurrent_commits"],
            "report output hash vector has the wrong length")
    require(isinstance(output_bytes, list)
            and len(output_bytes) == row["concurrent_commits"]
            and all(type(value) is int and value >= 0 for value in output_bytes),
            "report output-size vector has the wrong shape")
    admission = report.get("retention_admission")
    require(isinstance(admission, dict), "retention admission is missing")
    cap = expected_cap(row["policy"], admission)
    require(report.get("retention_ceiling_bytes") == cap
            and report.get("default_retention_ceiling_bytes") == DEFAULT_CAP,
            "report retention cap differs from admission")
    scenarios = report.get("scenarios")
    require(isinstance(scenarios, list) and len(scenarios) == 1,
            "report must contain exactly one ownership scenario")
    scenario = scenarios[0]
    require(scenario.get("ownership_mode") == row["ownership"]
            and scenario.get("section_count") == row["section_count"]
            and scenario.get("concurrent_commits") == row["concurrent_commits"]
            and scenario.get("retention_ceiling_bytes") == cap,
            "scenario identity differs from schedule")
    each = scenario.get("retained_base_bytes_each")
    require(isinstance(each, list) and len(each) == row["concurrent_commits"]
            and all(type(value) is int and value >= 0 for value in each),
            "retained-base vector has the wrong shape")
    required = admission["required_retained_base_bytes"]
    expected_each = required if cap == required else (
        required if row["policy"] == "default" and required <= DEFAULT_CAP else 0
    )
    require(each == [expected_each] * row["concurrent_commits"],
            "retained-base vector does not match the selected cap policy")
    require(scenario.get("logical_retained_base_bytes_sum") == sum(each),
            "logical retained-byte sum is inconsistent")
    require(scenario.get("retained_commit_count") == sum(value != 0 for value in each),
            "retained commit count is inconsistent")
    pair_count = row["concurrent_commits"] * (row["concurrent_commits"] - 1) // 2
    expected_shared = pair_count if row["ownership"] == "shared_snapshot_clones" else 0
    require(scenario.get("source_inventory_pair_count") == pair_count
            and scenario.get("source_inventory_shared_pairs") == expected_shared,
            "source inventory topology differs from schedule")
    require(scenario.get("source_inventory_sharing_verified") is True
            and scenario.get("measured_commit_shape_verified") is True
            and scenario.get("measured_output_verified") is True
            and scenario.get("measured_inverse_verified") is True
            and scenario.get("preflight_output_verified") is True
            and scenario.get("preflight_inverse_verified") is True,
            "semantic or topology controls are incomplete")
    expected_validation_count = (1 if row["ownership"] == "shared_snapshot_clones"
                                 else row["concurrent_commits"])
    expected_validation_scope = (
        "one_original_lineage_commit_plus_all_preflight_commits"
        if row["ownership"] == "shared_snapshot_clones" else "every_distinct_commit"
    )
    require(scenario.get("measured_commit_validation_count") == expected_validation_count
            and scenario.get("measured_validation_scope") == expected_validation_scope,
            "measured validation scope differs from ownership topology")
    gauges = scenario.get("gauges")
    require(isinstance(gauges, dict), "lifecycle gauges are missing")
    baseline = gauges.get("baseline")
    check_gauge(baseline, "baseline")
    for name in ("after_snapshot_drop", "commits_retained",
                 "after_release_all_but_last", "after_release_last", "after_commit_drop"):
        check_phase(gauges.get(name), baseline, name)
    check_allocator_sequence(gauges)
    # Release must not create a new live allocation in an otherwise quiescent
    # journal.  The route itself also checks absolute counter conservation.
    before_last = gauges["after_release_all_but_last"]["absolute"]["live_bytes"]
    after_last = gauges["after_release_last"]["absolute"]["live_bytes"]
    after_drop = gauges["after_commit_drop"]["absolute"]["live_bytes"]
    require(after_last <= before_last and after_drop <= after_last,
            "release/drop lifecycle increased allocator live bytes")
    return {"source_archive_sha256": source_hash, "cap_bytes": cap,
            "logical_retained_base_bytes": scenario["logical_retained_base_bytes_sum"],
            "retained_commit_count": scenario["retained_commit_count"]}


def parse_rss(time_path: Path) -> int | None:
    text = time_path.read_text(encoding="utf-8", errors="replace")
    match = re.search(r"Maximum resident set size \(kbytes\):\s*(\d+)", text)
    return int(match.group(1)) * 1024 if match else None


def process_group_alive(pgid: int) -> bool:
    try:
        os.killpg(pgid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


def close_process_group(pgid: int, *, force: bool = False) -> tuple[bool, int]:
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


def command_for(row: dict[str, Any], binary: Path, report: Path) -> list[str]:
    return ["/usr/bin/time", "-v", "/usr/bin/taskset", "-c", str(row["cpu"]), str(binary),
            "docx-section-retention", "--sections", str(row["section_count"]),
            "--concurrent-commits", str(row["concurrent_commits"]),
            "--policy", row["policy"], "--ownership", row["ownership"],
            "--source-revision", row["source_revision"], "--output", str(report)]


def run_child(row: dict[str, Any], manifest: dict[str, Any], packet: Path) -> dict[str, Any]:
    child = packet / "children" / row["id"]
    child.mkdir(parents=True, exist_ok=False)
    report_path = child / "report.json"
    stdout_path = child / "stdout.log"
    stderr_path = child / "stderr.log"
    time_path = child / "time.log"
    binary = Path(str(manifest["binary"]["path"])).resolve()
    source_root = Path(str(manifest["source_root"])).resolve()
    command = command_for(row, binary, report_path)
    environment = dict(manifest["child_environment"])
    started = utc_now()
    monotonic_start = time.monotonic()
    exit_code: int | None = None
    deadline_exceeded = False
    observation_timeouts = 0
    runner_error: str | None = None
    pgid: int | None = None
    group_closed = False
    group_probes = 0
    with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
        try:
            process = subprocess.Popen(command, cwd=str(source_root), env=environment,
                                       stdout=stdout, stderr=stderr, start_new_session=True)
            pgid = process.pid
            while True:
                try:
                    exit_code = process.wait(timeout=1.0)
                    break
                except subprocess.TimeoutExpired:
                    observation_timeouts += 1
                    if time.monotonic() - monotonic_start <= row["deadline_seconds"]:
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
    if pgid is not None:
        group_closed, group_probes = close_process_group(
            pgid, force=deadline_exceeded or runner_error is not None)
    with time_path.open("xb") as stream:
        stream.write(stderr_path.read_bytes())
        stream.flush()
        os.fsync(stream.fileno())
    report_desc = descriptor(report_path, root=packet) if report_path.is_file() else None
    terminal = {
        "schema": TERMINAL_SCHEMA,
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
        "process_group_id": pgid,
        "process_group_closed": group_closed,
        "process_group_probe_count": group_probes,
        "stdout": descriptor(stdout_path, root=packet),
        "stderr": descriptor(stderr_path, root=packet),
        "time": descriptor(time_path, root=packet),
        "rss_bytes": parse_rss(time_path),
        "report": report_desc,
        "binary": manifest["binary"],
        "source_revision": row["source_revision"],
        "runner_sha256": manifest["runner_sha256"],
        "manifest_sha256": sha256_path(packet / "manifest.json"),
    }
    terminal_path = child / "terminal.json"
    write_new_json(terminal_path, terminal)
    passed = (exit_code == 0 and not deadline_exceeded and runner_error is None
              and group_closed and terminal["rss_bytes"] is not None
              and report_desc is not None and report_desc["bytes"] < MAX_JSON_BYTES)
    validation: dict[str, Any] = {"schema": VALIDATION_SCHEMA,
                                  "status": "failed", "report": report_desc}
    if passed:
        try:
            report = read_json(report_path)
            facts = validate_report(row, report)
            validation.update({"status": "passed", **facts})
        except Exception as error:
            validation["error"] = f"{type(error).__name__}: {error}"
            passed = False
    else:
        validation["error"] = "terminal process failed, RSS missing, or report missing"
    validation_path = child / "validation.json"
    write_new_json(validation_path, validation)
    return {
        "id": row["id"], "status": "passed" if passed else "failed",
        "exit_code": exit_code, "rss_bytes": terminal["rss_bytes"],
        "terminal": descriptor(terminal_path, root=packet),
        "report": report_desc,
        "validation": descriptor(validation_path, root=packet),
    }


def verify_manifest(manifest: dict[str, Any], *, authorized: bool) -> list[dict[str, Any]]:
    require(manifest.get("schema") == SCHEMA, "unsupported 0931 manifest schema")
    if authorized:
        require(manifest.get("status") == "authorized", "capture manifest is not authorized")
        require(os.environ.get(APPROVAL_ENV) == "1",
                f"{APPROVAL_ENV}=1 is required for capture")
    revision = manifest.get("source_revision")
    require(isinstance(revision, str) and re.fullmatch(r"[0-9a-f]{40}", revision),
            "manifest source revision is not a lowercase 40-hex revision")
    rows = manifest.get("schedule", {}).get("rows")
    expected = schedule_rows(revision)
    require(rows == expected, "manifest schedule differs from frozen 0931 matrix")
    require(manifest.get("schedule", {}).get("sha256") == schedule_hash(expected),
            "manifest schedule hash differs from rows")
    require(manifest.get("cpu") == CPU and manifest.get("deadline_seconds") == DEADLINE_SECONDS,
            "manifest process controls changed")
    require(manifest.get("replay_external_rehash") is True,
            "manifest must require external source/build descriptor rehash on replay")
    environment = manifest.get("child_environment")
    validate_child_environment(environment)
    binary = manifest.get("binary")
    require(isinstance(binary, dict), "manifest binary descriptor is missing")
    build_receipt = manifest.get("build_receipt")
    require(isinstance(build_receipt, dict), "manifest build receipt binding is missing")
    custody_sources = manifest.get("custody_sources")
    require(isinstance(custody_sources, dict), "manifest custody source set is missing")
    if authorized:
        verify_descriptor(binary, label="binary")
        source_root = Path(str(manifest.get("source_root", ""))).resolve()
        require(source_root.is_dir(), "manifest source root is missing")
        source_receipt = manifest.get("source_receipt")
        require(isinstance(source_receipt, dict), "manifest source receipt binding is missing")
        verify_descriptor(source_receipt, label="source receipt")
        verify_descriptor(build_receipt, label="build receipt")
        for name, value in custody_sources.items():
            verify_descriptor(value, label=f"custody source {name}")
    runner_hash = manifest.get("runner_sha256")
    require(isinstance(runner_hash, str) and re.fullmatch(r"[0-9a-f]{64}", runner_hash),
            "manifest runner hash is missing")
    if authorized:
        require(runner_hash == sha256_path(Path(__file__)),
                "manifest runner hash does not bind this script")
    return expected


def packet_path(packet: Path, value: dict[str, Any], label: str) -> Path:
    return verify_descriptor(value, label=label, root=packet)


def verify_packet_custody(packet: Path, manifest: dict[str, Any]) -> None:
    packet_scripts = manifest.get("packet_scripts")
    require(isinstance(packet_scripts, dict)
            and isinstance(packet_scripts.get("runner"), dict),
            "packet runner descriptor is missing")
    runner = packet_path(packet, packet_scripts["runner"], "packet runner")
    require(sha256_path(runner) == manifest["runner_sha256"],
            "copied packet runner does not match manifest runner hash")
    sources = manifest.get("custody_sources", {})
    copies = manifest.get("packet_custody", {})
    require(isinstance(sources, dict) and isinstance(copies, dict)
            and set(sources) == set(copies),
            "packet custody descriptors are incomplete")
    for name, original in sources.items():
        verify_descriptor(original, label=f"external custody {name}")
        local = packet_path(packet, copies[name], f"packet custody {name}")
        require(copies[name].get("bytes") == original.get("bytes")
                and copies[name].get("sha256") == original.get("sha256"),
                f"packet custody changed: {name}")
        require(local.is_file(), f"packet custody is not a regular file: {name}")


def require_packet_member(packet: Path, value: dict[str, Any], expected: Path,
                          label: str) -> Path:
    path = packet_path(packet, value, label)
    actual = path.relative_to(packet.resolve())
    require(actual == expected, f"{label}: unexpected packet member {actual}")
    return path


def check_terminal(packet: Path, manifest: dict[str, Any], row: dict[str, Any],
                   child: dict[str, Any], *, require_pass: bool) -> dict[str, Any]:
    terminal_path = packet_path(packet, child["terminal"], f"terminal {row['id']}")
    terminal = read_json(terminal_path)
    require(terminal.get("schema") == TERMINAL_SCHEMA, f"terminal schema changed: {row['id']}")
    require(terminal.get("schedule") == row, f"terminal schedule changed: {row['id']}")
    require(terminal.get("environment") == manifest["child_environment"],
            f"terminal environment changed: {row['id']}")
    require(terminal.get("source_revision") == row["source_revision"],
            f"terminal source revision changed: {row['id']}")
    require(terminal.get("cwd_basename") == Path(str(manifest["source_root"])).name,
            f"terminal source cwd changed: {row['id']}")
    require(terminal.get("binary") == manifest["binary"],
            f"terminal binary binding changed: {row['id']}")
    require(terminal.get("runner_sha256") == manifest["runner_sha256"],
            f"terminal runner binding changed: {row['id']}")
    require(terminal.get("manifest_sha256") == sha256_path(packet / "manifest.json"),
            f"terminal manifest binding changed: {row['id']}")
    report_descriptor = terminal.get("report")
    require(report_descriptor == child.get("report"),
            f"terminal/report descriptor differs: {row['id']}")
    if report_descriptor is not None:
        require(report_descriptor.get("path", "").endswith("/report.json"),
                f"terminal report path changed: {row['id']}")
    argv = terminal.get("argv")
    require(isinstance(argv, list) and len(argv) >= 13,
            f"terminal command is missing: {row['id']}")
    expected_prefix = command_for(row, Path(str(manifest["binary"]["path"])),
                                 Path("report.json"))[:-1]
    require(argv[:-1] == expected_prefix and argv[-2] == "--output"
            and Path(str(argv[-1])).name == "report.json",
            f"terminal command differs from schedule: {row['id']}")
    child_root = Path("children") / row["id"]
    for name in ("stdout", "stderr", "time"):
        require_packet_member(packet, terminal[name], child_root / f"{name}.log",
                              f"{row['id']} {name}")
    if require_pass:
        require(terminal.get("exit_code") == 0
                and terminal.get("deadline_exceeded") is False
                and terminal.get("runner_error") is None
                and terminal.get("process_group_closed") is True
                and type(terminal.get("rss_bytes")) is int,
                f"terminal process failed: {row['id']}")
    return terminal


def validate_checkpoint_chain(packet: Path, rows: list[dict[str, Any]],
                              final: dict[str, Any]) -> None:
    checkpoint_dir = packet / "checkpoints"
    checkpoints = sorted(checkpoint_dir.glob("checkpoint-*.json")) if checkpoint_dir.is_dir() else []
    require(checkpoints, "checkpoint journal is missing")
    expected_ids = [row["id"] for row in rows]
    previous: list[str] = []
    for number, path in enumerate(checkpoints, start=1):
        checkpoint = read_json(path)
        require(checkpoint.get("schema") == STATE_SCHEMA
                and checkpoint.get("status") == "running",
                f"checkpoint {path.name} is not a running state")
        completed = checkpoint.get("completed_ids")
        failed = checkpoint.get("failed_ids")
        children = checkpoint.get("children")
        require(checkpoint.get("checkpoint_index") == number,
                f"checkpoint numbering changed: {path.name}")
        require(isinstance(completed, list) and isinstance(failed, list)
                and isinstance(children, list), f"checkpoint {path.name} is incomplete")
        ids = [child.get("id") for child in children]
        require(len(ids) == len(set(ids)) and ids == expected_ids[:number],
                f"checkpoint conservation failed: {path.name}")
        expected_completed = [child["id"] for child in children if child.get("status") == "passed"]
        expected_failed = [child["id"] for child in children if child.get("status") != "passed"]
        require(completed == expected_completed and failed == expected_failed,
                f"checkpoint outcome partition changed: {path.name}")
        require(len(children) == number,
                f"checkpoint child journal changed: {path.name}")
        require(previous == expected_ids[:number - 1],
                f"checkpoint predecessor changed: {path.name}")
        previous = ids
    require(previous == expected_ids, "checkpoint journal does not cover the schedule")
    final_children = final.get("children")
    require(isinstance(final_children, list)
            and [child.get("id") for child in final_children] == expected_ids,
            "final state does not conserve checkpoint child order")
    require(final.get("children_expected") == len(rows)
            and final.get("children_run") == len(rows),
            "final state child counts are inconsistent")
    failed_count = sum(child.get("status") != "passed" for child in final_children)
    require(final.get("failed_count") == failed_count,
            "final state failed count is inconsistent")
    require(final.get("status") == ("completed" if failed_count == 0 else "invalid"),
            "final state status does not match child outcomes")


def validate_packet(packet: Path, *, require_completed: bool = False) -> dict[str, Any]:
    packet = packet.resolve()
    manifest = read_json(packet / "manifest.json")
    rows = verify_manifest(manifest, authorized=False)
    verify_descriptor(manifest["binary"], label="replayed binary")
    verify_descriptor(manifest["source_receipt"], label="replayed source receipt")
    verify_descriptor(manifest["build_receipt"], label="replayed build receipt")
    verify_packet_custody(packet, manifest)
    state = read_json(packet / "capture-state.json")
    require(state.get("schema") == STATE_SCHEMA
            and state.get("status") in ("completed", "invalid"),
            "capture state is not terminal")
    if require_completed:
        require(state.get("status") == "completed", "packet contains an invalid cohort")
    validate_checkpoint_chain(packet, rows, state)
    by_id = {child.get("id"): child for child in state["children"]}
    require(set(by_id) == {row["id"] for row in rows}, "packet child IDs differ from schedule")
    source_hashes: dict[int, str] = {}
    passed: list[tuple[dict[str, Any], dict[str, Any], dict[str, Any]]] = []
    for row in rows:
        child = by_id[row["id"]]
        require(child.get("status") in ("passed", "failed"),
                f"child status is invalid: {row['id']}")
        child_root = Path("children") / row["id"]
        require_packet_member(packet, child["terminal"], child_root / "terminal.json",
                              f"terminal {row['id']}")
        require_packet_member(packet, child["validation"], child_root / "validation.json",
                              f"validation {row['id']}")
        if child.get("report") is not None:
            require_packet_member(packet, child["report"], child_root / "report.json",
                                  f"report {row['id']}")
        terminal = check_terminal(packet, manifest, row, child,
                                  require_pass=child["status"] == "passed")
        validation_path = packet_path(packet, child["validation"], f"validation {row['id']}")
        validation = read_json(validation_path)
        require(validation.get("schema") == VALIDATION_SCHEMA,
                f"validation schema changed: {row['id']}")
        if child["status"] == "passed":
            require(validation.get("status") == "passed", f"validation failed: {row['id']}")
            report_path = packet_path(packet, child["report"], f"report {row['id']}")
            report = read_json(report_path)
            facts = validate_report(row, report)
            require(validation.get("source_archive_sha256") == facts["source_archive_sha256"],
                    f"validation source binding changed: {row['id']}")
            source_hashes.setdefault(row["section_count"], facts["source_archive_sha256"])
            require(source_hashes[row["section_count"]] == facts["source_archive_sha256"],
                    f"source corpus changed within section shape: {row['id']}")
            passed.append((row, terminal, report))
        else:
            require(validation.get("status") == "failed",
                    f"failed child has a passing validation: {row['id']}")
    return {"manifest": manifest, "rows": rows, "state": state, "passed": passed,
            "source_hashes": source_hashes}


def delta(value: dict[str, Any], baseline: dict[str, Any], field: str) -> int:
    require(type(value.get(field)) is int and type(baseline.get(field)) is int,
            f"missing allocator field {field}")
    return value[field] - baseline[field]


def phase_metrics(gauges: dict[str, Any], name: str) -> dict[str, Any]:
    baseline = gauges["baseline"]
    absolute = baseline if name == "baseline" else gauges[name]["absolute"]
    result = {
        "allocated_bytes_delta": delta(absolute, baseline, "allocated_bytes"),
        "deallocated_bytes_delta": delta(absolute, baseline, "deallocated_bytes"),
        "live_bytes_delta": delta(absolute, baseline, "live_bytes"),
        "peak_live_bytes_delta": delta(absolute, baseline, "peak_live_bytes"),
        "region_peak_live_bytes": absolute.get("region_peak_live_bytes"),
    }
    region_peak = absolute.get("region_peak_live_bytes")
    if region_peak is not None:
        predecessor = {
            "after_snapshot_drop": baseline,
            "commits_retained": None,
            "after_release_all_but_last": gauges["commits_retained"]["absolute"],
            "after_release_last": gauges["after_release_all_but_last"]["absolute"],
            "after_commit_drop": gauges["after_release_last"]["absolute"],
        }.get(name)
        require(predecessor is not None, f"{name}: region peak predecessor is missing")
        result["region_peak_live_bytes_delta_from_begin_live"] = (
            region_peak - predecessor["live_bytes"]
        )
    return result


def reduce_packet(packet: Path, output: Path, *, allow_invalid: bool = False) -> dict[str, Any]:
    checked = validate_packet(packet, require_completed=not allow_invalid)
    cells: dict[tuple[Any, ...], list[dict[str, Any]]] = {}
    for row, terminal, report in checked["passed"]:
        scenario = report["scenarios"][0]
        gauges = scenario["gauges"]
        baseline = gauges["baseline"]
        retained = gauges["commits_retained"]["absolute"]
        dropped = gauges["after_commit_drop"]["absolute"]
        sample = {
            "id": row["id"], "section_count": row["section_count"],
            "concurrent_commits": row["concurrent_commits"],
            "policy": row["policy"], "ownership": row["ownership"],
            "repetition": row["repetition"], "rss_bytes": terminal["rss_bytes"],
            "source_archive_sha256": report["source_archive_sha256"],
            "logical_retained_base_bytes": scenario["logical_retained_base_bytes_sum"],
            "retained_commit_count": scenario["retained_commit_count"],
            "allocated_bytes_delta": delta(retained, baseline, "allocated_bytes"),
            "deallocated_bytes_delta": delta(retained, baseline, "deallocated_bytes"),
            "live_bytes_delta": delta(retained, baseline, "live_bytes"),
            "peak_live_bytes_delta": delta(retained, baseline, "peak_live_bytes"),
            "release_live_bytes_delta": delta(dropped, baseline, "live_bytes"),
            "phases": {name: phase_metrics(gauges, name) for name in (
                "after_snapshot_drop", "commits_retained",
                "after_release_all_but_last", "after_release_last",
                "after_commit_drop")},
        }
        key = (row["section_count"], row["concurrent_commits"], row["policy"],
               row["ownership"])
        cells.setdefault(key, []).append(sample)
    reduced: list[dict[str, Any]] = []
    for key in sorted(cells):
        samples = cells[key]
        reduced.append({
            "section_count": key[0], "concurrent_commits": key[1],
            "policy": key[2], "ownership": key[3],
            "replications": len(samples), "samples": samples,
            "median": {field: statistics.median(item[field] for item in samples)
                       for field in ("rss_bytes", "logical_retained_base_bytes",
                                     "retained_commit_count", "allocated_bytes_delta",
                                     "deallocated_bytes_delta", "live_bytes_delta",
                                     "peak_live_bytes_delta", "release_live_bytes_delta")},
            "phase_median": {
                phase: {
                    field: (statistics.median(values) if values else None)
                    for field in ("allocated_bytes_delta", "deallocated_bytes_delta",
                                  "live_bytes_delta", "peak_live_bytes_delta",
                                  "region_peak_live_bytes",
                                  "region_peak_live_bytes_delta_from_begin_live")
                    for values in [[item["phases"][phase][field]
                                    for item in samples
                                    if item["phases"][phase].get(field) is not None]]
                }
                for phase in ("after_snapshot_drop", "commits_retained",
                              "after_release_all_but_last", "after_release_last",
                              "after_commit_drop")
            },
        })
    result = {
        "schema": "litchi.performance.0931.memory-reduction.v1",
        "claim": "descriptive allocator/live-memory diagnostic; no latency or adoption claim",
        "packet_status": checked["state"]["status"],
        "passed_children": len(checked["passed"]),
        "scheduled_children": len(checked["rows"]),
        "source_archive_sha256_by_sections": {str(k): v
                                               for k, v in checked["source_hashes"].items()},
        "cells": reduced,
    }
    write_new_json(output, result)
    return result


def synthetic_report(row: dict[str, Any]) -> dict[str, Any]:
    required = 512 * row["section_count"]
    retained = required if row["policy"] == "exact" else 0
    if row["policy"] == "default" and required <= DEFAULT_CAP:
        retained = required
    if row["policy"] == "one_under":
        retained = 0
    n = row["concurrent_commits"]
    each = [retained] * n
    def gauge(allocated: int, deallocated: int, peak: int) -> dict[str, Any]:
        return {"allocation_calls": 1, "deallocation_calls": 1,
                "reallocation_calls": 0, "failed_allocation_calls": 0,
                "allocated_bytes": allocated, "deallocated_bytes": deallocated,
                "live_bytes": allocated - deallocated, "peak_live_bytes": peak,
                "region_peak_live_bytes": peak, "overflowed": False,
                "observer_invalid": False}
    base = gauge(1000, 900, 100)
    live = 100 + retained * n
    endpoint = gauge(1000 + retained * n, 900, max(100, live))
    final = gauge(1000 + retained * n, 900 + retained * n, max(100, live))
    phase = lambda value, region_peak=True: {
        "absolute": ({**value, "region_peak_live_bytes": value["peak_live_bytes"]}
                      if region_peak else {**value, "region_peak_live_bytes": None}),
        "live_bytes_delta_from_baseline": value["live_bytes"] - 100,
    }
    pair_count = n * (n - 1) // 2
    admission = {"probe_ceiling_bytes": PROBE_CAP, "required_retained_base_bytes": required,
                 "zero_cap_bytes": 0, "zero_retained_base_bytes": 0,
                 "one_under_cap_bytes": required - 1, "one_under_retained_base_bytes": 0,
                 "exact_cap_bytes": required, "exact_retained_base_bytes": required,
                 "default_cap_bytes": DEFAULT_CAP,
                 "default_retained_base_bytes": required if required <= DEFAULT_CAP else 0,
                 "boundary_checks_verified": True}
    return {"schema": REPORT_SCHEMA, "section_count": row["section_count"],
            "concurrent_commits": n, "retention_policy": row["policy"],
            "ownership_mode": row["ownership"], "source_revision": row["source_revision"],
            "retention_ceiling_bytes": expected_cap(row["policy"], admission),
            "default_retention_ceiling_bytes": DEFAULT_CAP, "source_archive_bytes": 1,
            "source_archive_sha256": "0" * 64,
            "expected_output_sha256": ["1" * 64] * n, "expected_output_bytes": [1] * n,
            "retention_admission": admission,
            "scenarios": [{"ownership_mode": row["ownership"],
                           "section_count": row["section_count"],
                           "concurrent_commits": n,
                           "retention_ceiling_bytes": expected_cap(row["policy"], admission),
                           "retained_base_bytes_each": each,
                           "retained_commit_count": sum(x != 0 for x in each),
                           "logical_retained_base_bytes_sum": sum(each),
                           "source_inventory_shared_pairs":
                               pair_count if row["ownership"] == "shared_snapshot_clones" else 0,
                           "source_inventory_pair_count": pair_count,
                           "source_inventory_sharing_verified": True,
                           "measured_commit_shape_verified": True,
                           "measured_commit_validation_count":
                               1 if row["ownership"] == "shared_snapshot_clones" else n,
                           "measured_validation_scope":
                               ("one_original_lineage_commit_plus_all_preflight_commits"
                                if row["ownership"] == "shared_snapshot_clones"
                                else "every_distinct_commit"),
                           "measured_output_verified": True,
                           "measured_inverse_verified": True,
                           "preflight_output_verified": True,
                           "preflight_inverse_verified": True,
                           "gauges": {"baseline": base,
                                      "after_snapshot_drop": phase(endpoint),
                                      "commits_retained": phase(endpoint, False),
                                      "after_release_all_but_last": phase(endpoint),
                                      "after_release_last": phase(final),
                                      "after_commit_drop": phase(final),
                                      "returned_to_baseline": True}}]}


def self_test() -> dict[str, Any]:
    revision = "a" * 40
    rows = schedule_rows(revision)
    report = synthetic_report(rows[0])
    validate_report(rows[0], report)
    bad_topology = json.loads(json.dumps(report))
    bad_topology["scenarios"][0]["source_inventory_shared_pairs"] = 1
    rejected_topology = False
    try:
        validate_report(rows[0], bad_topology)
    except EvidenceError:
        rejected_topology = True
    require(rejected_topology, "synthetic topology negative control did not reject")
    bad_cap = json.loads(json.dumps(report))
    bad_cap["retention_admission"]["exact_retained_base_bytes"] = 0
    rejected_cap = False
    try:
        validate_report(rows[0], bad_cap)
    except EvidenceError:
        rejected_cap = True
    require(rejected_cap, "synthetic cap negative control did not reject")
    bad_flags = json.loads(json.dumps(report))
    bad_flags["scenarios"][0]["gauges"]["baseline"]["overflowed"] = True
    rejected_flags = False
    try:
        validate_report(rows[0], bad_flags)
    except EvidenceError:
        rejected_flags = True
    require(rejected_flags, "synthetic observer-flag negative control did not reject")
    bad_peak = json.loads(json.dumps(report))
    bad_peak["scenarios"][0]["gauges"]["after_release_last"]["absolute"][
        "region_peak_live_bytes"] = 0
    rejected_peak = False
    try:
        validate_report(rows[0], bad_peak)
    except EvidenceError:
        rejected_peak = True
    require(rejected_peak, "synthetic region-peak negative control did not reject")
    rejected_path = False
    try:
        safe_relative("../escape.json", "synthetic path")
    except EvidenceError:
        rejected_path = True
    require(rejected_path, "synthetic path negative control did not reject")
    rejected_environment = False
    try:
        validate_child_environment({"PATH": "/usr/bin:/bin"})
    except EvidenceError:
        rejected_environment = True
    require(rejected_environment, "synthetic environment negative control did not reject")
    valid_children = [{"id": row["id"], "status": "passed"} for row in rows]
    final = {"children": valid_children, "children_expected": len(rows),
             "children_run": len(rows), "failed_count": 0, "status": "completed"}
    with tempfile_directory() as packet:
        checkpoint_dir = packet / "checkpoints"
        checkpoint_dir.mkdir()
        for index, child in enumerate(valid_children, start=1):
            write_new_json(checkpoint_dir / f"checkpoint-{index:03d}.json", {
                "schema": STATE_SCHEMA, "status": "running", "checkpoint_index": index,
                "completed_ids": [item["id"] for item in valid_children[:index]],
                "failed_ids": [], "children": valid_children[:index],
            })
        validate_checkpoint_chain(packet, rows, final)
    return {"schema": "litchi.performance.0931.memory-self-test.v1",
            "schedule_children": len(rows), "topology_negative": rejected_topology,
            "cap_negative": rejected_cap, "observer_flags_negative": rejected_flags,
            "region_peak_negative": rejected_peak, "path_negative": rejected_path,
            "environment_negative": rejected_environment,
            "checkpoint_conservation": True}


class tempfile_directory:
    """Small context manager without retaining synthetic evidence."""
    def __enter__(self) -> Path:
        import tempfile
        self._context = tempfile.TemporaryDirectory(prefix="litchi-0931-")
        return Path(self._context.name)

    def __exit__(self, exc_type: Any, exc: Any, traceback: Any) -> None:
        self._context.cleanup()


def capture(args: argparse.Namespace) -> dict[str, Any]:
    manifest_path = args.manifest.resolve()
    manifest = read_json(manifest_path)
    rows = verify_manifest(manifest, authorized=True)
    packet = args.packet.resolve()
    require(not packet.exists(), f"refusing to replace packet: {packet}")
    require(REPO_ROOT not in packet.parents and packet != REPO_ROOT,
            "capture packet must remain outside the checkout")
    packet.mkdir(parents=True)
    (packet / "children").mkdir()
    (packet / "checkpoints").mkdir()
    script_copy = packet / "scripts" / Path(__file__).name
    script_copy.parent.mkdir()
    with script_copy.open("xb") as stream:
        stream.write(Path(__file__).read_bytes())
    packet_manifest = dict(manifest)
    packet_manifest["packet_scripts"] = {"runner": descriptor(script_copy, root=packet)}
    packet_custody: dict[str, dict[str, Any]] = {}
    custody_sources = manifest.get("custody_sources", {})
    if custody_sources:
        custody_dir = packet / "custody"
        custody_dir.mkdir()
        for name, value in custody_sources.items():
            source = verify_descriptor(value, label=f"custody source {name}")
            safe_name = Path(name)
            require(len(safe_name.parts) == 1 and safe_name.name == name,
                    f"unsafe custody name: {name!r}")
            target = custody_dir / name
            with target.open("xb") as stream:
                stream.write(source.read_bytes())
                stream.flush()
                os.fsync(stream.fileno())
            copied = descriptor(target, root=packet)
            require(copied["bytes"] == value["bytes"]
                    and copied["sha256"] == value["sha256"],
                    f"custody copy changed: {name}")
            packet_custody[name] = copied
    packet_manifest["packet_custody"] = packet_custody
    write_new_json(packet / "manifest.json", packet_manifest)
    children: list[dict[str, Any]] = []
    started = utc_now()
    for index, row in enumerate(rows, start=1):
        children.append(run_child(row, packet_manifest, packet))
        passed = [child["id"] for child in children if child["status"] == "passed"]
        failed = [child["id"] for child in children if child["status"] != "passed"]
        write_new_json(packet / "checkpoints" / f"checkpoint-{index:03d}.json", {
            "schema": STATE_SCHEMA, "status": "running", "checkpoint_index": index,
            "started_utc": started, "completed_ids": passed, "failed_ids": failed,
            "children": children,
        })
    failed_count = sum(child["status"] != "passed" for child in children)
    state = {"schema": STATE_SCHEMA,
             "status": "completed" if failed_count == 0 else "invalid",
             "started_utc": started, "ended_utc": utc_now(),
             "children_run": len(children), "children_expected": len(rows),
             "failed_count": failed_count, "children": children,
             "failure_policy": "retain all terminal children; never retry or recapture"}
    write_new_json(packet / "capture-state.json", state)
    return {"status": state["status"], "packet": str(packet),
            "children_run": len(children), "failed": failed_count}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    capture_parser = sub.add_parser("capture")
    capture_parser.add_argument("--manifest", type=Path, required=True)
    capture_parser.add_argument("--packet", type=Path, required=True)
    validate_parser = sub.add_parser("validate")
    validate_parser.add_argument("--packet", type=Path, required=True)
    validate_parser.add_argument("--require-completed", action="store_true")
    reduce_parser = sub.add_parser("reduce")
    reduce_parser.add_argument("--packet", type=Path, required=True)
    reduce_parser.add_argument("--output", type=Path, required=True)
    reduce_parser.add_argument("--allow-invalid", action="store_true")
    sub.add_parser("self-test")
    args = parser.parse_args()
    try:
        if args.command == "capture":
            result = capture(args)
        elif args.command == "validate":
            checked = validate_packet(args.packet, require_completed=args.require_completed)
            result = {"schema": "litchi.performance.0931.memory-validation-result.v1",
                      "status": checked["state"]["status"],
                      "passed_children": len(checked["passed"]),
                      "scheduled_children": len(checked["rows"])}
        elif args.command == "reduce":
            result = reduce_packet(args.packet, args.output, allow_invalid=args.allow_invalid)
        else:
            result = self_test()
        print(json.dumps(result, sort_keys=True))
        return 0
    except (EvidenceError, OSError, ValueError) as error:
        print(f"0931 validation error: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
