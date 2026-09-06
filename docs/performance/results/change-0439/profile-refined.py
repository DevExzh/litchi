#!/usr/bin/env python3
"""Capture one whole-process profile for the 0439 append baseline.

This is deliberately separate from the 12-report timing matrix. It runs the
same large normal selector once under ``perf stat`` or ``perf record`` after
the matrix has completed. The report includes process setup, corpus creation
inside the workload, warmups, samples, output hashing, and the retained
workload's own checks. The copied Python oracle is run afterward and is not
part of the profiled command.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import subprocess
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent
EVENTS = (
    "cycles:u",
    "instructions:u",
    "branches:u",
    "branch-misses:u",
    "L1-dcache-load-misses:u",
)
ATTEMPT_RE = re.compile(r"^[a-z0-9][a-z0-9_-]{0,63}$")
HEX64 = re.compile(r"^[0-9a-fA-F]{64}$")


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
    return {
        "path": str(path.relative_to(ROOT)),
        "bytes": path.stat().st_size,
        "sha256": sha(path),
    }


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: expected an object")
    return value


def import_capture():
    path = ROOT / "capture.py"
    spec = importlib.util.spec_from_file_location("change0439_capture_for_profile", path)
    if spec is None or spec.loader is None:
        fail(f"cannot load capture helper: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


capture = import_capture()


def resolve_repo(value: Any, repo: Path, label: str) -> Path:
    path = Path(value)
    if not path.is_absolute():
        path = repo / path
    return path.resolve()


def checked_build(protocol: dict[str, Any], repo: Path, protocol_path: Path) -> tuple[Path, dict[str, Any]]:
    roles = obj(protocol.get("roles"), "protocol.roles")
    role = obj(roles.get("after"), "protocol.roles.after")
    build_directory = role.get("build_directory")
    if not isinstance(build_directory, str) or not build_directory:
        fail("protocol.roles.after.build_directory: required")
    build_path = ROOT / build_directory / "build.json"
    build = obj(load(build_path), str(build_path))
    if build.get("change") != 439 or build.get("role") not in {None, "after", build_directory}:
        fail(f"{build_path}: wrong change or role")
    if build.get("protocol_sha256") != sha(protocol_path):
        fail(f"{build_path}: protocol hash is stale")
    oracle = protocol["oracle"]
    for path_key, hash_key in (("verifier_path", "verifier_sha256"), ("protocol_path", "protocol_sha256")):
        oracle_path = (ROOT / oracle[path_key]).resolve()
        if not oracle_path.is_relative_to(ROOT.resolve()) or sha(oracle_path) != oracle[hash_key]:
            fail("frozen oracle identity differs")
    binaries = obj(build.get("binaries"), f"{build_path}.binaries")
    if set(binaries) != {"normal", "allocator"}:
        fail(f"{build_path}: normal and allocator binaries are required")
    for mode, identity in binaries.items():
        identity = obj(identity, f"{build_path}.binaries.{mode}")
        value = identity.get("path")
        if not isinstance(value, str) or not value:
            fail(f"{build_path}.binaries.{mode}.path: required")
        binary = resolve_repo(value, repo, f"{build_path}.binaries.{mode}.path")
        size = identity.get("bytes")
        expected = identity.get("sha256")
        if isinstance(size, bool) or not isinstance(size, int) or size <= 0:
            fail(f"{build_path}.binaries.{mode}.bytes: expected positive integer")
        if not isinstance(expected, str) or HEX64.fullmatch(expected) is None:
            fail(f"{build_path}.binaries.{mode}.sha256: expected SHA-256")
        if not binary.is_file() or binary.stat().st_size != size or sha(binary) != expected.lower():
            fail(f"{build_path}.binaries.{mode}: binary identity is stale")
    return build_path.parent, build


def source_status(repo: Path) -> list[str]:
    return capture.status_outside_bundle(repo)


def run_profile(kind: str, attempt: str, repo: Path, protocol_path: Path, custody_path: Path) -> int:
    if kind not in {"stat", "record"}:
        fail("profile kind must be stat or record")
    protocol = obj(load(protocol_path), str(protocol_path))
    if protocol.get("change") != 439:
        fail("protocol.change must be 439")
    if protocol.get("cpu") != 2 or protocol.get("workers") != 1:
        fail("protocol must bind CPU 2 and one worker")
    if protocol.get("samples") != 30 or protocol.get("warmups") != 3:
        fail("profile workload must use samples=30 and warmups=3")
    shapes = protocol.get("shapes")
    if shapes != {"tiny": 64, "medium": 4_096, "large": 8_192}:
        fail("protocol.shapes must bind tiny=64, medium=4096, large=8192")
    build_dir, build = checked_build(protocol, repo, protocol_path)
    custody_spec = importlib.util.spec_from_file_location("change0439_profile_custody", custody_path)
    if custody_spec is None or custody_spec.loader is None:
        fail(f"cannot load custody driver: {custody_path}")
    custody = importlib.util.module_from_spec(custody_spec)
    custody_spec.loader.exec_module(custody)
    source_before = custody.sources()
    if source_before != build.get("source_manifest"):
        fail("ambient source manifest differs from after build")
    status_before = source_status(repo)

    profile_dir = ROOT / "profiles" / "after" / kind / attempt
    if profile_dir.exists() and any(profile_dir.iterdir()):
        fail(f"profile directory already contains evidence: {profile_dir}")
    profile_dir.mkdir(parents=True, exist_ok=True)
    report = profile_dir / "report.json"
    catalog = profile_dir / "report-catalog.json"
    resource = profile_dir / "resource.log"
    workload_log = profile_dir / "workload.log"
    oracle_log = profile_dir / "oracle.log"
    stat_output = profile_dir / "perf-stat.txt"
    data = profile_dir / "perf.data"
    script_output = profile_dir / "perf-script.txt"
    perf_report = profile_dir / "perf-report.txt"
    receipt_path = profile_dir / "receipt.json"
    role_spec = obj(obj(protocol["roles"], "protocol.roles").get("after"), "protocol.roles.after")
    selector = role_spec.get("selector")
    if selector != "odp_existing_append_lifecycle":
        fail("protocol.roles.after.selector must be odp_existing_append_lifecycle")
    binary_desc = obj(build["binaries"].get("normal"), "after build normal binary")
    binary_path = resolve_repo(binary_desc["path"], repo, "profile binary")
    base = capture.workload_argv(protocol, binary_path, selector, "large", "normal", report, catalog)
    if kind == "stat":
        profiler = [
            "perf", "stat", "--no-big-num", "-x,", "-e", ",".join(EVENTS),
            "-o", str(stat_output), "--", *base,
        ]
    else:
        profiler = [
            "perf", "record", "--no-buildid-cache", "-o", str(data),
            "-F", "999", "-e", "cycles:u", "--call-graph", "fp,127", "--", *base,
        ]
    argv = ["taskset", "-c", "2", "/usr/bin/time", "-v", "-o", str(resource), *profiler]
    row: dict[str, Any] = {
        "schema": "litchi-0439-profile-receipt-v1",
        "change": 439,
        "role": "after",
        "kind": kind,
        "attempt": attempt,
        "scope": (
            "whole fresh process: setup, corpus generation, warmups, timed samples, "
            "output hashing, and workload checks; copied Python oracle runs afterward "
            "outside the profiled command and GNU-time scope"
        ),
        "selector": selector,
        "source_field": role_spec.get("source_field"),
        "shape": "large",
        "argv": argv,
        "cwd": str(repo),
        "revision": build.get("revision"),
        "source_manifest": build.get("source_manifest"),
        "binary": binary_desc,
        "build_directory": build_dir.name,
        "protocol_path": str(protocol_path.relative_to(ROOT)),
        "protocol_sha256": sha(protocol_path),
        "driver_sha256": sha(Path(__file__)),
        "capture_helper_sha256": sha(ROOT / "capture.py"),
        "custody_driver": str(custody_path),
        "source_before": source_before,
        "status_before": status_before,
        "stat_events": list(EVENTS) if kind == "stat" else [],
        "record_event": "cycles:u" if kind == "record" else None,
        "record_frequency_hz": 999 if kind == "record" else None,
        "call_graph": "fp,127" if kind == "record" else None,
        "started_utc": now(),
        "status": "running",
    }
    write(receipt_path, row)
    env = os.environ.copy()
    env.update({"RUSTUP_TOOLCHAIN": "1.98.1", "DEBUGINFOD_URLS": "", "PYTHONDONTWRITEBYTECODE": "1"})
    try:
        with workload_log.open("xb") as output:
            result = subprocess.run(argv, cwd=repo, env=env, stdout=output, stderr=subprocess.STDOUT)
        row["exit_code"] = result.returncode
        if result.returncode != 0:
            raise RuntimeError(f"profiler workload exited {result.returncode}")
        capture.verify_report_identity(report, binary_desc, build["revision"])
        oracle_argv = capture.oracle_command(protocol, report, "normal", "large", "after")
        oracle_result = subprocess.run(oracle_argv, cwd=repo, env=env, capture_output=True, text=True)
        oracle_log.write_text(
            "argv=" + json.dumps(oracle_argv) + "\nstdout=" + oracle_result.stdout + "stderr=" + oracle_result.stderr,
            encoding="utf-8",
        )
        row["oracle_exit_code"] = oracle_result.returncode
        row["oracle_stdout_sha256"] = sha(oracle_log)
        if oracle_result.returncode != 0 or oracle_result.stdout.strip() != "VALID":
            raise RuntimeError("copied ODP oracle rejected profile report")
        if kind == "record":
            for command, output in (
                (["perf", "report", "--stdio", "--no-children", "--percent-limit", "0", "-i", str(data)], perf_report),
                (["perf", "script", "-i", str(data)], script_output),
            ):
                with output.open("xb") as stream:
                    subprocess.run(command, cwd=repo, env=env, stdout=stream, stderr=subprocess.STDOUT, check=True)
        row["status"] = "pass"
    except Exception as error:
        row["status"] = "failed"
        row["error"] = repr(error)
    finally:
        row["source_after"] = custody.sources()
        row["source_unchanged"] = row["source_before"] == row["source_after"]
        row["status_after"] = source_status(repo)
        row["outside_bundle_status_unchanged"] = status_before == row["status_after"]
        if not row["source_unchanged"] or not row["outside_bundle_status_unchanged"]:
            row["status"] = "failed"
            row["error"] = "source custody or repository status changed during profile"
        row["finished_utc"] = now()
        row["artifacts"] = {
            key: record(path)
            for key, path in {
                "report": report,
                "catalog": catalog,
                "resource": resource,
                "workload_log": workload_log,
                "oracle_log": oracle_log,
                "perf_stat": stat_output,
                "perf_data": data,
                "perf_script": script_output,
                "perf_report": perf_report,
            }.items()
            if path.is_file()
        }
        write(receipt_path, row)
    print(json.dumps({"status": row["status"], "role": "after", "kind": kind}))
    return 0 if row["status"] == "pass" else 1


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--kind", choices=("stat", "record"), required=True)
    parser.add_argument("--attempt", default="formal")
    parser.add_argument("--repo-root", type=Path, required=True)
    parser.add_argument("--protocol", type=Path, default=ROOT / "protocol.json")
    parser.add_argument("--custody-driver", type=Path, required=True)
    args = parser.parse_args()
    if ATTEMPT_RE.fullmatch(args.attempt) is None:
        parser.error("--attempt must match [a-z0-9][a-z0-9_-]{0,63}")
    try:
        return run_profile(
            args.kind,
            args.attempt,
            args.repo_root.resolve(),
            args.protocol.resolve(),
            args.custody_driver.resolve(),
        )
    except (OSError, KeyError, TypeError, ValueError, subprocess.CalledProcessError, ProfileError) as error:
        print(f"PROFILE INVALID: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
