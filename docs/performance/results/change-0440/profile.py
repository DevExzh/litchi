#!/usr/bin/env python3
"""Capture an optional large normal whole-process profile for one 0440 role."""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent
EVENTS = ("cycles:u", "instructions:u", "branches:u", "branch-misses:u", "L1-dcache-load-misses:u")


class ProfileError(ValueError):
    pass


def fail(message: str) -> None:
    raise ProfileError(message)


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def load(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{path}: invalid JSON: {error}")
    raise AssertionError("unreachable")


def write(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def record(path: Path) -> dict[str, Any]:
    return {"path": str(path.relative_to(ROOT)), "bytes": path.stat().st_size, "sha256": sha(path)}


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: expected object")
    return value


def import_capture():
    spec = importlib.util.spec_from_file_location("change0440_capture", ROOT / "capture.py")
    if spec is None or spec.loader is None:
        fail("cannot import capture.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


capture = import_capture()


def run_profile(role: str, kind: str, repo: Path, protocol_path: Path, custody_path: Path) -> int:
    if role not in {"before", "after"} or kind not in {"stat", "record"}:
        fail("role/kind must be before|after and stat|record")
    protocol = obj(load(protocol_path), str(protocol_path))
    if protocol.get("change") != 440 or protocol.get("status") != "frozen" or protocol.get("cpu") != 2 or protocol.get("workers") != 1 or protocol.get("samples") != 30 or protocol.get("warmups") != 3:
        fail("protocol profile bindings differ")
    capture.oracle_files(protocol)
    profile_spec = obj(protocol.get("profiles"), "protocol.profiles")
    if profile_spec.get("events") != list(EVENTS) or profile_spec.get("record_frequency_hz") != 999 or profile_spec.get("call_graph") != "fp,127":
        fail("protocol perf event/frequency/call-graph binding differs")
    role_spec = obj(obj(protocol.get("roles"), "protocol.roles").get(role), f"protocol.roles.{role}")
    build_dir, build = capture.build_for(role, protocol, repo, protocol_path)
    custody_spec = importlib.util.spec_from_file_location("change0440_profile_custody", custody_path)
    if custody_spec is None or custody_spec.loader is None:
        fail(f"cannot load custody driver {custody_path}")
    custody = importlib.util.module_from_spec(custody_spec)
    custody_spec.loader.exec_module(custody)
    source_before = custody.sources()
    if source_before != build.get("source_manifest"):
        fail("profile source manifest differs from build")
    status_before = capture.status_outside_bundle(repo)
    directory = ROOT / "profiles" / role / kind / "formal"
    if directory.exists() and any(directory.iterdir()):
        fail(f"profile directory already contains evidence: {directory}")
    directory.mkdir(parents=True, exist_ok=True)
    report = directory / "report.json"
    catalog = directory / "report-catalog.json"
    resource = directory / "resource.log"
    workload = directory / "workload.log"
    oracle_log = directory / "oracle.log"
    stat_output = directory / "perf-stat.txt"
    data = directory / "perf.data"
    perf_script = directory / "perf-script.txt"
    perf_report = directory / "perf-report.txt"
    receipt = directory / "receipt.json"
    binary_desc = obj(build["binaries"]["normal"], f"{build_dir}/build.json.normal")
    binary = capture.resolve_repo(binary_desc["path"], repo, "profile binary")
    base = capture.workload_argv(protocol, binary, role_spec["selector"], "large", report, catalog)
    if kind == "stat":
        profiler = ["perf", "stat", "--no-big-num", "-x,", "-e", ",".join(EVENTS), "-o", str(stat_output), "--", *base]
    else:
        profiler = ["perf", "record", "--no-buildid-cache", "-o", str(data), "-F", "999", "-e", "cycles:u", "--call-graph", "fp,127", "--", *base]
    argv = ["taskset", "-c", "2", "/usr/bin/time", "-v", "-o", str(resource), *profiler]
    row: dict[str, Any] = {
        "schema": "litchi-0440-profile-receipt-v1", "change": 440, "role": role, "kind": kind,
        "scope": "whole fresh executable including setup, corpus generation, warmups, samples, hashing, and checks; Python oracle afterward",
        "selector": role_spec["selector"], "source_field": role_spec["source_field"], "shape": "large", "argv": argv,
        "cwd": str(repo), "revision": build["revision"], "source_manifest": build["source_manifest"], "binary": binary_desc,
        "protocol_sha256": sha(protocol_path), "driver_sha256": sha(Path(__file__)), "capture_driver_sha256": sha(ROOT / "capture.py"),
        "source_before": source_before, "stat_events": list(EVENTS) if kind == "stat" else [],
        "record_event": "cycles:u" if kind == "record" else None, "record_frequency_hz": 999 if kind == "record" else None,
        "call_graph": "fp,127" if kind == "record" else None, "started_utc": now(), "status": "running",
    }
    write(receipt, row)
    env = os.environ.copy()
    env.update({"RUSTUP_TOOLCHAIN": "1.98.1", "DEBUGINFOD_URLS": "", "PYTHONDONTWRITEBYTECODE": "1"})
    try:
        with workload.open("xb") as output:
            result = subprocess.run(argv, cwd=repo, env=env, stdout=output, stderr=subprocess.STDOUT)
        row["exit_code"] = result.returncode
        if result.returncode != 0:
            raise RuntimeError(f"profile workload exited {result.returncode}")
        capture.verify_report_identity(report, binary_desc, build["revision"])
        oracle_argv = capture.oracle_command(protocol, report, "normal", "large", role)
        oracle_result = subprocess.run(oracle_argv, cwd=repo, env=env, capture_output=True, text=True)
        oracle_log.write_text("argv=" + json.dumps(oracle_argv) + "\nstdout=" + oracle_result.stdout + "stderr=" + oracle_result.stderr, encoding="utf-8")
        row["oracle_argv"] = oracle_argv
        row["oracle_exit_code"] = oracle_result.returncode
        if oracle_result.returncode != 0 or oracle_result.stdout.strip() != "VALID":
            raise RuntimeError("profile oracle rejected report")
        if kind == "record":
            for command, output in ((["perf", "report", "--stdio", "--no-children", "--percent-limit", "0", "-i", str(data)], perf_report), (["perf", "script", "-i", str(data)], perf_script)):
                with output.open("xb") as stream:
                    subprocess.run(command, cwd=repo, env=env, stdout=stream, stderr=subprocess.STDOUT, check=True)
        row["status"] = "pass"
    except Exception as error:
        row["status"], row["error"] = "failed", repr(error)
    finally:
        row["finished_utc"] = now()
        row["source_after"] = custody.sources()
        row["status_before"] = status_before
        row["status_after"] = capture.status_outside_bundle(repo)
        row["source_unchanged"] = row["source_before"] == row["source_after"]
        row["outside_bundle_status_unchanged"] = row["status_before"] == row["status_after"]
        if not row["source_unchanged"] or not row["outside_bundle_status_unchanged"]:
            row["status"], row["error"] = "failed", "source custody or outside-bundle worktree status changed"
        row["artifacts"] = {key: record(path) for key, path in {"report": report, "catalog": catalog, "resource": resource, "workload": workload, "oracle": oracle_log, "perf_stat": stat_output, "perf_data": data, "perf_script": perf_script, "perf_report": perf_report}.items() if path.is_file()}
        write(receipt, row)
    return 0 if row["status"] == "pass" else 1


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--role", choices=("before", "after"), required=True)
    parser.add_argument("--kind", choices=("stat", "record"), required=True)
    parser.add_argument("--repo-root", type=Path, required=True)
    parser.add_argument("--protocol", type=Path, default=ROOT / "protocol.json")
    parser.add_argument("--custody-driver", type=Path, required=True)
    args = parser.parse_args()
    try:
        return run_profile(args.role, args.kind, args.repo_root.resolve(), args.protocol.resolve(), args.custody_driver.resolve())
    except (OSError, KeyError, TypeError, ValueError, subprocess.CalledProcessError, ProfileError) as error:
        print(f"PROFILE INVALID: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
