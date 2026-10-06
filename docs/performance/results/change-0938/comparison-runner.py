#!/usr/bin/env python3
"""Run and replay the 0938 identical-binary variability control.

This runner has one normal lane and two metadata labels, A and B.  The labels
select the same retained executable, source revision, arguments, environment,
and corpus; they only identify the paired observation in the packet.  The
runner never treats either label as a candidate and never starts a child until
an authorized manifest and the explicit approval environment are present.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import importlib.util
import json
import math
import os
import random
import re
import shutil
import signal
import statistics
import subprocess
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


REPO_ROOT = Path("/home/zhuhe/code/litchi")
COMPARISON_PATH = REPO_ROOT / "docs/performance/results/change-0937/comparison.py"
COMPARISON_SHA256 = "b70a2b38d098e3abd2b567d1bea70f40e9d27e62acaca8f80ff633dd2c596712"
PLAN_DEFAULT = REPO_ROOT / "docs/performance/results/change-0938/comparison-plan-v1.json"
GATE_DEFAULT = REPO_ROOT / "docs/performance/results/change-0938/comparison-gate-v1.json"
BASE_REVISION = "3ae856d76424f2f9b58d26dcd20fd074b8e97541"
CURRENT_HEAD = "3e4ea43db4a5c2bd2d2fd8e16ba18d9c887b7809"
SOURCE_RECEIPT = Path("/home/zhuhe/litchi-performance-0937/receipts/baseline-source-final-v1.json")
SOURCE_EQUIVALENCE = Path(
    "/home/zhuhe/litchi-performance-0937/retained-baseline-runtime-equivalence-v1.json"
)
BUILD_RECEIPT = Path(
    "/home/zhuhe/litchi-performance-0937/receipts/baseline-normal-build-bound-v1.json"
)
BINARY = Path(
    "/home/zhuhe/litchi-performance-0933/target-candidate-normal/release/litchi-perf-baseline"
)
BINARY_SHA256 = "8a81df641151e5ed06b8f700c5d4cf068d0fee9c9e131fa147bf707a799c4c66"
BINARY_BYTES = 69_934_624
SOURCE_ROOT = Path("/home/zhuhe/litchi-performance-0937/baseline-source")
CASE = "docx_source_backed_existing_section_layout_edit_save"
ROUTES = ("generic_read_at", "owned_vec", "owned_reader", "range_short")
LABELS = ("A", "B")
PHASES = ("preparation_ns", "open_ns", "query_ns", "edit_ns", "commit_ns", "publication_ns")
CPU = 12
WARMUPS = 5
SAMPLES = 50
BLOCKS = 12
CHILDREN = len(ROUTES) * len(LABELS) * BLOCKS
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
TIME_FIELDS = {
    "rss_bytes": (r"Maximum resident set size \(kbytes\):\s*(\d+)", "kbytes"),
    "major_page_faults": (r"Major \(requiring I/O\) page faults:\s*(\d+)", "int"),
    "minor_page_faults": (r"Minor \(reclaiming a frame\) page faults:\s*(\d+)", "int"),
    "user_cpu_seconds": (r"User time \(seconds\):\s*([0-9.]+)", "float"),
    "system_cpu_seconds": (r"System time \(seconds\):\s*([0-9.]+)", "float"),
    "voluntary_context_switches": (r"Voluntary context switches:\s*(\d+)", "int"),
    "involuntary_context_switches": (r"Involuntary context switches:\s*(\d+)", "int"),
}


class EvidenceError(RuntimeError):
    """A custody, schedule, report, or replay contract failed closed."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


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
    value: dict[str, Any] = {
        "path": str(path if root is None else path.relative_to(root.resolve())),
        "bytes": path.stat().st_size,
        "sha256": sha256_path(path),
    }
    return value


def verify_descriptor(value: dict[str, Any], *, label: str, root: Path | None = None) -> Path:
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label}: path missing")
    path = Path(raw)
    if root is not None:
        require(not path.is_absolute(), f"{label}: packet path must be relative")
        path = (root.resolve() / path).resolve()
        require(path.is_relative_to(root.resolve()), f"{label}: path escapes packet")
    else:
        path = path.resolve()
    actual = descriptor(path)
    require(actual["bytes"] == value.get("bytes") and actual["sha256"] == value.get("sha256"),
            f"{label}: descriptor changed")
    return path


def read_json(path: Path, *, max_bytes: int | None = MAX_ARTIFACT_BYTES) -> dict[str, Any]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    if max_bytes is not None:
        require(path.stat().st_size < max_bytes, f"JSON is at least 1 MiB: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as error:
        raise EvidenceError(f"cannot parse JSON {path}: {error}") from error
    require(isinstance(value, dict), f"JSON is not an object: {path}")
    return value


def write_new_json(path: Path, value: dict[str, Any], *, limit: int = MAX_ARTIFACT_BYTES) -> None:
    data = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode()
    require(len(data) < limit, f"artifact is at least 1 MiB: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("xb") as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())


def load_comparison(path: Path = COMPARISON_PATH) -> Any:
    require(sha256_path(path) == COMPARISON_SHA256, "0937 validator dependency hash changed")
    spec = importlib.util.spec_from_file_location("comparison_0937_dependency", path)
    require(spec is not None and spec.loader is not None, "cannot load comparison dependency")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def schedule_rows() -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for block in range(1, BLOCKS + 1):
        routes = list(ROUTES if block % 2 else tuple(reversed(ROUTES)))
        labels = list(LABELS if block % 2 else tuple(reversed(LABELS)))
        order = "forward" if block % 2 else "reverse"
        label_order = "A_then_B" if block % 2 else "B_then_A"
        for position, route in enumerate(routes, start=1):
            pair_id = f"block-{block:02d}-{route}"
            for label in labels:
                rows.append({
                    "id": f"block-{block:02d}-{position:02d}-{route}-{label}",
                    "pair_id": pair_id,
                    "block": block,
                    "position": position,
                    "route": route,
                    "label": label,
                    "route_order": order,
                    "label_order": label_order,
                    "cpu": CPU,
                    "warmups": WARMUPS,
                    "samples": SAMPLES,
                    "deadline_seconds": int(DEADLINE_SECONDS),
                    "failure_policy": "retain_child_and_invalidate_cohort; never_rerun_or_recapture",
                })
    require(len(rows) == CHILDREN, "0938 schedule cardinality changed")
    require(len({row["id"] for row in rows}) == CHILDREN, "0938 schedule has duplicate IDs")
    return rows


def schedule_sha256(rows: list[dict[str, Any]] | None = None) -> str:
    payload = json.dumps(rows if rows is not None else schedule_rows(), sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(payload.encode()).hexdigest()


def child_environment(row: dict[str, Any]) -> dict[str, str]:
    environment = dict(CHILD_ENV_BASE)
    environment["LITCHI_DOCX_INGRESS_ROUTE"] = row["route"]
    return environment


def child_command(row: dict[str, Any], report: Path) -> list[str]:
    return [
        "/usr/bin/time", "-v", "taskset", "-c", str(CPU), str(BINARY),
        "--case", CASE, "--samples", str(SAMPLES), "--warmup", str(WARMUPS),
        "--json", str(report),
    ]


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


def parse_time_metrics(path: Path) -> dict[str, int | float]:
    text = path.read_text(encoding="utf-8", errors="replace")
    result: dict[str, int | float] = {}
    for field, (pattern, kind) in TIME_FIELDS.items():
        match = re.search(pattern, text)
        require(match is not None, f"time log missing {field}")
        raw = match.group(1)
        if kind == "kbytes":
            result[field] = int(raw) * 1024
        elif kind == "int":
            result[field] = int(raw)
        else:
            result[field] = float(raw)
    return result


def source_and_binary_binding() -> dict[str, Any]:
    source = read_json(SOURCE_RECEIPT)
    require(source.get("source_revision") == BASE_REVISION, "source revision changed")
    require(source.get("source_root") == str(SOURCE_ROOT), "source root changed")
    require(source.get("source_file_count") == 10779, "source file count changed")
    require(source.get("source_bytes") == 259962058, "source byte total changed")
    source_desc = descriptor(SOURCE_RECEIPT)
    equivalence = read_json(SOURCE_EQUIVALENCE)
    require(equivalence.get("status") == "verified_runtime_equivalence_with_test_overlay",
            "runtime equivalence status changed")
    equivalence_desc = descriptor(SOURCE_EQUIVALENCE)
    build = read_json(BUILD_RECEIPT)
    require(build.get("status") == "passed" and build.get("source_revision") == BASE_REVISION,
            "retained build binding changed")
    binary = descriptor(BINARY)
    require(binary["sha256"] == BINARY_SHA256 and binary["bytes"] == BINARY_BYTES,
            "retained normal binary identity changed")
    require(os.access(BINARY, os.X_OK), "retained normal binary is not executable")
    return {
        "source": source_desc,
        "source_revision": BASE_REVISION,
        "source_root": str(SOURCE_ROOT),
        "source_equivalence": equivalence_desc,
        "build_receipt": descriptor(BUILD_RECEIPT),
        "binary": binary,
        "historical_build_source_revision": build.get("source_revision"),
        "current_head": CURRENT_HEAD,
        "current_head_scope": "HEAD changes from e3b5acdd are docs/performance evidence only; no runtime source changes.",
    }


def validate_plan(plan_path: Path, gate_path: Path) -> tuple[dict[str, Any], dict[str, Any]]:
    plan = read_json(plan_path)
    gate = read_json(gate_path)
    require(plan.get("schema") == "litchi.performance.0938.variability-plan.v1", "plan schema changed")
    require(gate.get("schema") == "litchi.performance.0938.variability-gate.v1", "gate schema changed")
    require(plan.get("children") == CHILDREN and plan.get("blocks") == BLOCKS,
            "plan cardinality changed")
    require(plan.get("routes") == list(ROUTES) and plan.get("labels") == list(LABELS),
            "plan routes or labels changed")
    require(plan.get("warmups") == WARMUPS and plan.get("samples") == SAMPLES
            and plan.get("cpu") == CPU, "plan workload changed")
    require(plan.get("schedule_sha256") == schedule_sha256(), "plan schedule hash changed")
    require(gate.get("seed") == 938 and gate.get("bootstrap_draws") == 10000,
            "variability bootstrap changed")
    require(gate.get("no_adoption_gate") is True and gate.get("no_causal_probability_claim") is True,
            "variability scope changed")
    return plan, gate


def script_bindings(plan_path: Path, gate_path: Path) -> dict[str, dict[str, Any]]:
    return {
        "runner": descriptor(Path(__file__).resolve()),
        "validator_dependency": descriptor(COMPARISON_PATH),
        "plan": descriptor(plan_path.resolve()),
        "gate": descriptor(gate_path.resolve()),
    }


def make_manifest(args: argparse.Namespace) -> dict[str, Any]:
    plan, gate = validate_plan(args.plan, args.gate)
    binding = source_and_binary_binding()
    rows = schedule_rows()
    return {
        "schema": "litchi.performance.0938.variability-manifest.v1",
        "created_utc": utc_now(),
        "capture_authorized": False,
        "formal_children_run": 0,
        "claim_scope": "Identical-binary A/B variability control only; no candidate, optimization, adoption, causal, or probability claim.",
        "failure_policy": "retain every terminal child; no retries, selective recapture, or replacement after any failure",
        "source_binding": binding,
        "binary": binding["binary"],
        "scripts": script_bindings(args.plan, args.gate),
        "plan": descriptor(args.plan.resolve()),
        "gate": descriptor(args.gate.resolve()),
        "schedule": {"rows": rows, "children": CHILDREN, "blocks": BLOCKS, "routes": list(ROUTES),
                     "labels": list(LABELS), "warmups": WARMUPS, "samples": SAMPLES, "cpu": CPU,
                     "sha256": schedule_sha256(rows)},
        "child_environment": dict(CHILD_ENV_BASE),
        "route_selector": "LITCHI_DOCX_INGRESS_ROUTE",
        "argv_template": ["/usr/bin/time", "-v", "taskset", "-c", str(CPU), str(BINARY),
                          "--case", CASE, "--samples", str(SAMPLES), "--warmup", str(WARMUPS),
                          "--json", "<report.json>"],
        "stats": {
            "pair_unit": "same block and route, labels A/B; label order alternates by block",
            "route_order_even": list(reversed(ROUTES)),
            "route_order_odd": list(ROUTES),
            "p50": "child_p50_ns.p50, median of twelve child p50 values per route and label",
            "p95_p99": "nearest-rank quantiles of all 600 raw elapsed samples per route and label",
            "spread": "max minus min of all 600 raw elapsed samples per route and label",
            "rss_and_process_metrics": "median of twelve child terminal metrics per route and label",
            "directional_counts": "count paired blocks with B/A >1.05 and A/B >1.05, descriptive only",
            "seed": 938, "bootstrap_draws": 10000,
        },
        "status": "prospective_unauthorized",
        "no_source_copy_or_build": True,
        "iwork": "excluded",
    }


def verify_scripts(manifest: dict[str, Any], packet: Path | None = None) -> None:
    scripts = manifest.get("scripts")
    require(isinstance(scripts, dict), "manifest scripts missing")
    for name, value in scripts.items():
        require(isinstance(value, dict), f"script descriptor malformed: {name}")
        if packet is None:
            verify_descriptor(value, label=f"manifest script {name}")
        else:
            local = packet / "scripts" / Path(str(value["path"])).name
            actual = descriptor(local, root=packet)
            require(actual["bytes"] == value["bytes"] and actual["sha256"] == value["sha256"],
                    f"packet script differs from manifest: {name}")


def verify_manifest(path: Path, *, authorized: bool = False, packet: Path | None = None) -> dict[str, Any]:
    manifest_path = path.resolve()
    manifest = read_json(manifest_path)
    require(manifest.get("schema") == "litchi.performance.0938.variability-manifest.v1",
            "manifest schema changed")
    require(manifest.get("schedule", {}).get("rows") == schedule_rows(), "schedule rows changed")
    require(manifest.get("schedule", {}).get("sha256") == schedule_sha256(), "schedule hash changed")
    require(manifest.get("schedule", {}).get("children") == CHILDREN, "child count changed")
    require(manifest.get("binary") == descriptor(BINARY), "binary binding changed")
    require(manifest.get("source_binding") == source_and_binary_binding(), "source binding changed")
    validate_plan(manifest_path.parent / Path(str(manifest["plan"]["path"]))
                  if not Path(str(manifest["plan"]["path"])).is_absolute()
                  else Path(str(manifest["plan"]["path"])),
                  manifest_path.parent / Path(str(manifest["gate"]["path"]))
                  if not Path(str(manifest["gate"]["path"])).is_absolute()
                  else Path(str(manifest["gate"]["path"])))
    require(manifest.get("child_environment") == CHILD_ENV_BASE, "child environment changed")
    verify_scripts(manifest, packet)
    if authorized:
        require(manifest.get("capture_authorized") is True, "capture authorization missing")
        require(os.environ.get("LITCHI_PERF_0938_CAPTURE_APPROVED") == "1",
                "LITCHI_PERF_0938_CAPTURE_APPROVED=1 is required")
    else:
        require(manifest.get("capture_authorized") is False, "unauthorized manifest changed")
    if packet is not None:
        packet_manifest = read_json(packet / "manifest.json")
        require(packet_manifest.get("schema") == manifest.get("schema"), "packet manifest schema changed")
        require(packet_manifest.get("schedule", {}).get("sha256") == manifest["schedule"]["sha256"],
                "packet manifest schedule changed")
    return manifest


def parse_time_and_check(path: Path, expected: dict[str, Any] | None = None) -> dict[str, int | float]:
    metrics = parse_time_metrics(path)
    if expected is not None:
        require(metrics == expected, "time metrics differ from terminal record")
    return metrics


def packet_path(packet: Path, child_id: str, name: str) -> Path:
    require(child_id and all(ch.isalnum() or ch in "-_" for ch in child_id),
            f"unsafe child ID: {child_id}")
    return packet / "children" / child_id / name


def run_child(row: dict[str, Any], packet: Path) -> dict[str, Any]:
    child = packet / "children" / row["id"]
    child.mkdir(parents=True, exist_ok=False)
    report = child / "report.json"
    stdout = child / "stdout.log"
    stderr = child / "stderr.log"
    time_log = child / "time.log"
    command = child_command(row, report)
    environment = child_environment(row)
    started = utc_now()
    monotonic_start = time.monotonic()
    exit_code: int | None = None
    observation_timeouts = 0
    deadline_exceeded = False
    runner_error: str | None = None
    pgid: int | None = None
    with stdout.open("xb") as out, stderr.open("xb") as err:
        try:
            process = subprocess.Popen(command, cwd=str(SOURCE_ROOT), env=environment,
                                       stdout=out, stderr=err, start_new_session=True)
            pgid = process.pid
            while True:
                try:
                    exit_code = process.wait(timeout=OBSERVE_SECONDS)
                    break
                except subprocess.TimeoutExpired:
                    observation_timeouts += 1
                    if time.monotonic() - monotonic_start <= float(row["deadline_seconds"]):
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
    group_closed, probes = (False, 0)
    if pgid is not None:
        group_closed, probes = close_process_group(pgid, force=deadline_exceeded or runner_error is not None)
    time_log.write_bytes(stderr.read_bytes())
    metrics: dict[str, int | float] | None = None
    try:
        metrics = parse_time_metrics(time_log)
    except Exception:
        metrics = None
    report_desc = descriptor(report, root=packet) if report.is_file() else None
    terminal = {
        "schema": "litchi.performance.0938.formal-terminal.v1",
        "status": "terminal", "schedule": row, "label": row["label"], "argv": command,
        "environment": environment, "cwd": str(SOURCE_ROOT), "started_utc": started,
        "ended_utc": ended, "exit_code": exit_code, "observation_timeout_count": observation_timeouts,
        "deadline_seconds": float(row["deadline_seconds"]), "deadline_exceeded": deadline_exceeded,
        "runner_error": runner_error, "process_group_id": pgid, "process_group_closed": group_closed,
        "process_group_probe_count": probes, "stdout": descriptor(stdout, root=packet),
        "stderr": descriptor(stderr, root=packet), "time": descriptor(time_log, root=packet),
        "time_metrics": metrics, "report": report_desc, "binary": descriptor(BINARY),
        "source_revision": BASE_REVISION, "manifest_sha256": sha256_path(packet / "manifest.json"),
        "runner_sha256": sha256_path(Path(__file__).resolve()),
    }
    write_new_json(child / "terminal.json", terminal)
    valid = (exit_code == 0 and not deadline_exceeded and runner_error is None and group_closed
             and metrics is not None and report_desc is not None and report_desc["bytes"] < MAX_ARTIFACT_BYTES)
    validation: dict[str, Any]
    if valid:
        try:
            comparison = load_comparison()
            report_value = read_json(report)
            result = comparison.validate_report(report_value, row["route"], warmups=WARMUPS,
                                                samples=SAMPLES, label=str(report), binary=descriptor(BINARY))
            validation = {"schema": "litchi.performance.0938.validation.v1", "status": "passed",
                          "label": row["label"], "route": row["route"], "pair_id": row["pair_id"],
                          "report": report_desc, "output_sha256": result["output_sha256"],
                          "sample_count": len(result["elapsed_ns"])}
        except Exception as error:
            validation = {"schema": "litchi.performance.0938.validation.v1", "status": "failed",
                          "label": row["label"], "route": row["route"],
                          "error": f"{type(error).__name__}: {error}", "report": report_desc}
            valid = False
    else:
        validation = {"schema": "litchi.performance.0938.validation.v1", "status": "failed",
                      "label": row["label"], "route": row["route"],
                      "error": "terminal/process/time/report validation failed", "report": report_desc}
    write_new_json(child / "validation.json", validation)
    return {
        "id": row["id"], "pair_id": row["pair_id"], "block": row["block"], "position": row["position"],
        "route": row["route"], "label": row["label"], "status": "passed" if valid else "failed",
        "exit_code": exit_code, "rss_bytes": None if metrics is None else metrics["rss_bytes"],
        "terminal": descriptor(child / "terminal.json", root=packet),
        "report": report_desc, "validation": descriptor(child / "validation.json", root=packet),
    }


def capture(args: argparse.Namespace) -> dict[str, Any]:
    manifest = verify_manifest(args.manifest, authorized=True)
    packet = args.packet.resolve()
    require(not packet.exists(), f"refusing existing packet: {packet}")
    packet.mkdir(parents=True)
    (packet / "children").mkdir()
    (packet / "scripts").mkdir()
    for name, value in manifest["scripts"].items():
        source = Path(str(value["path"]))
        target = packet / "scripts" / source.name
        with target.open("xb") as stream:
            stream.write(source.read_bytes())
        require(descriptor(target, root=packet)["sha256"] == value["sha256"],
                f"packet script changed: {name}")
    packet_manifest = copy.deepcopy(manifest)
    packet_manifest["packet_scripts"] = {
        name: descriptor(packet / "scripts" / Path(str(value["path"])).name, root=packet)
        for name, value in manifest["scripts"].items()
    }
    write_new_json(packet / "manifest.json", packet_manifest)
    rows = schedule_rows()
    state_rows: list[dict[str, Any]] = []
    for index, row in enumerate(rows, start=1):
        result = run_child(row, packet)
        state_rows.append(result)
        write_new_json(packet / f"capture-state-{index:03d}.json", {
            "schema": "litchi.performance.0938.capture-state.v1", "status": "running",
            "completed": index, "planned": CHILDREN, "failed": sum(x["status"] == "failed" for x in state_rows),
            "rows": state_rows,
        })
    failed = sum(value["status"] == "failed" for value in state_rows)
    final = {"schema": "litchi.performance.0938.capture-state.v1", "status": "completed" if not failed else "failed",
             "completed": len(state_rows), "planned": CHILDREN, "failed": failed, "rows": state_rows,
             "retries": 0, "recaptures": 0}
    write_new_json(packet / "capture-state.json", final)
    return {"status": final["status"], "packet": str(packet), "children_run": len(state_rows),
            "passed": len(state_rows) - failed, "failed": failed, "retries": 0, "recaptures": 0}


def stats(values: list[float | int]) -> dict[str, float | int]:
    require(values, "empty metric vector")
    ordered = sorted(values)
    def nearest(fraction: float) -> float:
        index = min(max(math.ceil(fraction * len(ordered)) - 1, 0), len(ordered) - 1)
        return float(ordered[index])
    return {"min": min(values), "p50": float(statistics.median(values)), "p95": nearest(.95),
            "p99": nearest(.99), "max": max(values), "mean": statistics.fmean(values),
            "spread": max(values) - min(values)}


def bootstrap_median(values: list[float]) -> dict[str, float | int]:
    require(values, "empty bootstrap vector")
    rng = random.Random(938)
    draws = [statistics.median(rng.choice(values) for _ in values) for _ in range(10000)]
    ordered = sorted(draws)
    lower = ordered[max(math.ceil(0.025 * len(ordered)) - 1, 0)]
    upper = ordered[max(math.ceil(0.975 * len(ordered)) - 1, 0)]
    return {
        "seed": 938,
        "draws": 10000,
        "statistic": "median of twelve paired ratios",
        "median": float(statistics.median(values)),
        "lower_95_nearest_rank": float(lower),
        "upper_95_nearest_rank": float(upper),
    }


def replay(args: argparse.Namespace) -> dict[str, Any]:
    packet = args.packet.resolve()
    manifest = verify_manifest(args.manifest, authorized=True, packet=packet)
    state = read_json(packet / "capture-state.json")
    require(state.get("status") == "completed" and state.get("completed") == CHILDREN,
            "capture state is not a completed full cohort")
    require(state.get("failed") == 0 and state.get("retries") == 0 and state.get("recaptures") == 0,
            "capture state contains failure/retry/recapture")
    rows = schedule_rows(); expected = {row["id"]: row for row in rows}
    children = state.get("rows")
    require(isinstance(children, list) and len(children) == CHILDREN, "capture state row count changed")
    require({value.get("id") for value in children} == set(expected), "capture state IDs changed")
    comparison = load_comparison()
    validated: list[dict[str, Any]] = []
    for child in children:
        row = expected[child["id"]]
        require(child.get("status") == "passed", f"child failed: {row['id']}")
        terminal_path = verify_descriptor(child["terminal"], label=f"{row['id']} terminal", root=packet)
        terminal = read_json(terminal_path)
        require(terminal.get("schedule") == row and terminal.get("label") == row["label"],
                f"{row['id']}: terminal schedule/label changed")
        require(terminal.get("exit_code") == 0 and terminal.get("deadline_exceeded") is False
                and terminal.get("runner_error") is None and terminal.get("process_group_closed") is True,
                f"{row['id']}: terminal status changed")
        time_path = verify_descriptor(terminal["time"], label=f"{row['id']} time", root=packet)
        metrics = parse_time_and_check(time_path, terminal.get("time_metrics"))
        require(metrics["rss_bytes"] == terminal.get("time_metrics", {}).get("rss_bytes"),
                f"{row['id']}: RSS changed")
        report_path = verify_descriptor(child["report"], label=f"{row['id']} report", root=packet)
        report = read_json(report_path)
        result = comparison.validate_report(report, row["route"], warmups=WARMUPS, samples=SAMPLES,
                                            label=str(report_path), binary=descriptor(BINARY))
        validation_path = verify_descriptor(child["validation"], label=f"{row['id']} validation", root=packet)
        validation = read_json(validation_path)
        require(validation.get("status") == "passed" and validation.get("label") == row["label"],
                f"{row['id']}: validation changed")
        validated.append({"row": row, "result": result, "time": metrics})
    require(len(validated) == CHILDREN, "validated child count changed")
    summary = reduce_validated(validated)
    summary.update({"schema": "litchi.performance.0938.variability-replay.v1", "status": "passed",
                    "children": CHILDREN, "packet": str(packet),
                    "source": "packet manifest, capture state, child terminal/time/report/validation descriptors, and parsed reports rehashed from packet; full stdout/stderr and terminal-field audit is reserved for the independent checker"})
    return summary


def reduce_validated(validated: list[dict[str, Any]]) -> dict[str, Any]:
    by_route_label: dict[tuple[str, str], list[dict[str, Any]]] = {}
    by_pair: dict[tuple[str, str], dict[str, dict[str, Any]]] = {}
    for item in validated:
        row = item["row"]
        by_route_label.setdefault((row["route"], row["label"]), []).append(item)
        by_pair.setdefault((row["pair_id"], row["route"]), {})[row["label"]] = item
    routes: dict[str, Any] = {}
    for route in ROUTES:
        labels: dict[str, Any] = {}
        for label in LABELS:
            items = by_route_label[(route, label)]
            raw = [value for item in items for value in item["result"]["elapsed_ns"]]
            child_p50 = [float(statistics.median(item["result"]["elapsed_ns"])) for item in items]
            labels[label] = {
                "children": len(items), "samples": len(raw), "elapsed_ns": stats(raw),
                "child_p50_ns": stats(child_p50),
                "phase_statistics_ns": {phase: stats([value for item in items for value in item["result"]["phase_ns"][phase]]) for phase in PHASES},
                "rss_bytes": stats([item["time"]["rss_bytes"] for item in items]),
                "major_page_faults": stats([item["time"]["major_page_faults"] for item in items]),
                "minor_page_faults": stats([item["time"]["minor_page_faults"] for item in items]),
                "user_cpu_seconds": stats([item["time"]["user_cpu_seconds"] for item in items]),
                "system_cpu_seconds": stats([item["time"]["system_cpu_seconds"] for item in items]),
                "voluntary_context_switches": stats([item["time"]["voluntary_context_switches"] for item in items]),
                "involuntary_context_switches": stats([item["time"]["involuntary_context_switches"] for item in items]),
            }
        pairs: list[dict[str, Any]] = []
        for (pair_id, pair_route), values in sorted(by_pair.items()):
            if pair_route != route:
                continue
            require(set(values) == set(LABELS), f"{pair_id}: label pair incomplete")
            a = float(statistics.median(values["A"]["result"]["elapsed_ns"]))
            b = float(statistics.median(values["B"]["result"]["elapsed_ns"]))
            pairs.append({"pair_id": pair_id, "a_p50_ns": a, "b_p50_ns": b,
                          "b_over_a": b / a, "a_over_b": a / b,
                          "a_rss_bytes": values["A"]["time"]["rss_bytes"],
                          "b_rss_bytes": values["B"]["time"]["rss_bytes"],
                          "rss_b_over_a": values["B"]["time"]["rss_bytes"] / values["A"]["time"]["rss_bytes"]})
        routes[route] = {
            "labels": labels, "pairs": pairs,
            "directional_counts": {
                "b_over_a_gt_1_05": sum(pair["b_over_a"] > 1.05 for pair in pairs),
                "a_over_b_gt_1_05": sum(pair["a_over_b"] > 1.05 for pair in pairs),
                "pairs": len(pairs),
            },
            "paired_p50_bootstrap": {
                "b_over_a": bootstrap_median([pair["b_over_a"] for pair in pairs]),
                "a_over_b": bootstrap_median([pair["a_over_b"] for pair in pairs]),
            },
            "rss_ratio": {
                "b_over_a": stats([pair["rss_b_over_a"] for pair in pairs]),
                "a_over_b": stats([1.0 / pair["rss_b_over_a"] for pair in pairs]),
            },
            "claim": "descriptive identical-binary variability control; no causal probability or adoption claim",
        }
    return {"routes": routes, "seed": 938, "bootstrap_draws": 10000}


def controls(args: argparse.Namespace) -> dict[str, Any]:
    manifest = read_json(args.manifest.resolve())
    scratch = Path(tempfile.mkdtemp(prefix="variability-controls-", dir="/tmp/litchi-performance-0938"))
    result: dict[str, Any] = {"schema": "litchi.performance.0938.controls.v1", "status": "failed",
                              "capture_started": False, "document_children_started": 0,
                              "runner": descriptor(Path(__file__).resolve()), "tests": {}}
    try:
        rows = schedule_rows()
        require(len(rows) == CHILDREN and len({row["id"] for row in rows}) == CHILDREN, "schedule positive failed")
        result["tests"]["schedule_positive"] = {"passed": True, "children": len(rows), "sha256": schedule_sha256(rows)}
        require(rows[0]["label"] == "A" and rows[1]["label"] == "B"
                and rows[2]["route"] == "owned_vec" and rows[2]["label"] == "A"
                and rows[4]["route"] == "owned_reader" and rows[4]["label"] == "A",
                "schedule alternation positive failed")
        result["tests"]["route_and_label_order"] = {"passed": True}
        comparison = load_comparison()
        fake = {"path": str(scratch / "synthetic-binary"), "bytes": 123456, "sha256": "a" * 64}
        (scratch / "synthetic-binary").write_bytes(b"x")
        report = comparison.synthetic_report("generic_read_at", binary=fake)
        comparison.validate_report(report, "generic_read_at", binary=fake)
        result["tests"]["synthetic_positive"] = {"passed": True}
        bad = copy.deepcopy(report); bad["binary_identity"]["binary_sha256"] = "b" * 64
        rejected = False
        try:
            comparison.validate_report(bad, "generic_read_at", binary=fake)
        except Exception:
            rejected = True
        require(rejected, "wrong binary negative accepted")
        bad_phase = copy.deepcopy(report); bad_phase["results"][0]["source"]["docx_section_layout"]["query_ns"][0] += 1
        rejected = False
        try:
            comparison.validate_report(bad_phase, "generic_read_at", binary=fake)
        except Exception:
            rejected = True
        require(rejected, "phase-sum negative accepted")
        result["tests"]["report_negatives"] = {"passed": True, "wrong_binary_rejected": True, "phase_sum_rejected": True}
        a = child_command(rows[0], scratch / "a.json"); b = child_command(rows[1], scratch / "b.json")
        a[-1] = "<report.json>"; b[-1] = "<report.json>"
        require(a == b, "A/B command differs beyond report path")
        require(child_environment(rows[0]) == {**CHILD_ENV_BASE, "LITCHI_DOCX_INGRESS_ROUTE": rows[0]["route"]},
                "clean child environment changed")
        result["tests"]["identical_binary_command_env"] = {"passed": True, "label_difference": "metadata only"}
        result["status"] = "passed"
        return result
    finally:
        shutil.rmtree(scratch, ignore_errors=True)
        write_new_json(args.output.resolve(), result)


def freeze(args: argparse.Namespace) -> dict[str, Any]:
    plan, gate = validate_plan(args.plan, args.gate)
    binding = source_and_binary_binding()
    manifest = make_manifest(args)
    output = args.manifest.resolve()
    require(not output.exists(), f"refusing to replace manifest: {output}")
    write_new_json(output, manifest)
    return {"status": manifest["status"], "manifest": str(output), "bytes": output.stat().st_size,
            "sha256": sha256_path(output), "children": CHILDREN, "capture_authorized": False,
            "source_revision": binding["source_revision"]}


def reduce_command(args: argparse.Namespace) -> dict[str, Any]:
    summary = replay(args)
    output = args.output.resolve()
    require(not output.is_relative_to(REPO_ROOT), "reduced evidence must be external")
    write_new_json(output, summary)
    summary["output"] = descriptor(output)
    return summary


def parser() -> argparse.ArgumentParser:
    value = argparse.ArgumentParser(description=__doc__)
    sub = value.add_subparsers(dest="command", required=True)
    check = sub.add_parser("check-schedule")
    check.set_defaults(function=lambda args: {"status": "passed", "children": CHILDREN,
                                               "schedule_sha256": schedule_sha256(), "capture_started": False})
    controls_parser = sub.add_parser("controls")
    controls_parser.add_argument("--manifest", type=Path, required=True)
    controls_parser.add_argument("--output", type=Path, required=True)
    controls_parser.set_defaults(function=controls)
    freeze_parser = sub.add_parser("freeze")
    freeze_parser.add_argument("--manifest", type=Path, required=True)
    freeze_parser.add_argument("--plan", type=Path, default=PLAN_DEFAULT)
    freeze_parser.add_argument("--gate", type=Path, default=GATE_DEFAULT)
    freeze_parser.set_defaults(function=freeze)
    capture_parser = sub.add_parser("capture")
    capture_parser.add_argument("--manifest", type=Path, required=True)
    capture_parser.add_argument("--packet", type=Path, required=True)
    capture_parser.set_defaults(function=capture)
    replay_parser = sub.add_parser("replay")
    replay_parser.add_argument("--manifest", type=Path, required=True)
    replay_parser.add_argument("--packet", type=Path, required=True)
    replay_parser.set_defaults(function=replay)
    reduce_parser = sub.add_parser("reduce")
    reduce_parser.add_argument("--manifest", type=Path, required=True)
    reduce_parser.add_argument("--packet", type=Path, required=True)
    reduce_parser.add_argument("--output", type=Path, required=True)
    reduce_parser.set_defaults(function=reduce_command)
    return value


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    print(json.dumps(args.function(args), indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"0938 variability runner failed: {type(error).__name__}: {error}", file=sys.stderr)
        raise
