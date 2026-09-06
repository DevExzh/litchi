#!/usr/bin/env python3
"""Capture one required 0436 large-normal whole-process profile.

The four invocations are independent fresh processes: before/after role by
``--role`` and perf-stat/perf-record by ``--kind``.  They are diagnostics only;
the formal latency matrix remains in ``capture.py``.
"""

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
REPO = ROOT.parents[3]
ROLES = {"before-streaming": "before", "after-streaming": "after"}
KINDS = {"stat", "record"}
EVENTS = ("cycles:u", "instructions:u", "branches:u", "branch-misses:u", "L1-dcache-load-misses:u")


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def load(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def write(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def resolve(value: str) -> Path:
    path = Path(value)
    return path if path.is_absolute() else REPO / path


def record(path: Path) -> dict[str, Any]:
    return {"path": str(path.relative_to(ROOT)), "bytes": path.stat().st_size, "sha256": sha(path)}


def custody_sources() -> dict[str, Any]:
    spec = importlib.util.spec_from_file_location("change0436_profile_custody", ROOT / "check.py")
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load frozen custody driver")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.sources()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--role", choices=tuple(ROLES), required=True)
    parser.add_argument("--kind", choices=tuple(sorted(KINDS)), required=True)
    args = parser.parse_args()
    role_dir = ROLES[args.role]
    protocol = load(ROOT / "protocol.json")
    build_path = ROOT / role_dir / "build.json"
    build = load(build_path)
    ambient_build = load(ROOT / "after" / "build.json")
    ambient_source = ambient_build["source_manifest"]
    source_before = custody_sources()
    if source_before != ambient_source:
        print("PROFILE INVALID: ambient source differs from after build", file=sys.stderr)
        return 2
    # Keep the evidence directory keyed by the build copy (before/after),
    # while the receipt retains the semantic role name used by the protocol.
    profile_dir = ROOT / "profiles" / role_dir / args.kind
    if profile_dir.exists() and any(profile_dir.iterdir()):
        print(f"PROFILE INVALID: evidence already exists at {profile_dir}", file=sys.stderr)
        return 2
    profile_dir.mkdir(parents=True, exist_ok=True)
    report = profile_dir / "report.json"
    catalog = profile_dir / "report-catalog.json"
    resource = profile_dir / "resource.log"
    workload_log = profile_dir / "workload.log"
    stat_output = profile_dir / "perf-stat.txt"
    data = profile_dir / "perf.data"
    script_output = profile_dir / "perf-script.txt"
    perf_report = profile_dir / "perf-report.txt"
    receipt = profile_dir / "receipt.json"
    binary = build["binaries"]["normal"]
    binary_path = resolve(binary["path"])
    if not binary_path.is_file() or sha(binary_path) != binary["sha256"] or binary_path.stat().st_size != binary["bytes"]:
        print("PROFILE INVALID: normal binary identity is unavailable or stale", file=sys.stderr)
        return 2
    selector = protocol["roles"][args.role]["selector"]
    base = [
        str(binary_path), "--case", selector, "--semantic-shape", "large",
        "--workers", str(protocol["workers"]), "--samples", str(protocol["samples"]),
        "--warmup", str(protocol["warmups"]), "--json", str(report), "--corpus-manifest", str(catalog),
    ]
    if args.kind == "stat":
        profiler = ["perf", "stat", "--no-big-num", "-x,", "-e", ",".join(EVENTS), "-o", str(stat_output), "--"] + base
    else:
        profiler = ["perf", "record", "--no-buildid-cache", "-o", str(data), "-F", "999", "-e", "cycles:u", "--call-graph", "fp,127", "--"] + base
    argv = ["taskset", "-c", str(protocol["cpu"]), "/usr/bin/time", "-v", "-o", str(resource)] + profiler
    row: dict[str, Any] = {
        "schema": "litchi-0436-profile-receipt-v1", "change": 436, "role": args.role, "kind": args.kind,
        "attempt": "formal", "preparatory": False,
        "scope": "whole fresh process including setup, corpus generation, warmups, samples, output hashing, and oracle",
        "selector": selector, "shape": "large", "argv": argv, "cwd": str(REPO),
        "revision": build["revision"], "source_manifest": build["source_manifest"], "binary": binary,
        "ambient_build_directory": "after", "ambient_source_manifest": ambient_source, "source_before": source_before,
        "protocol_path": "protocol.json", "protocol_sha256": sha(ROOT / "protocol.json"),
        "oracle_verifier_path": "oracle/verify-report.py", "oracle_verifier_sha256": sha(ROOT / "oracle" / "verify-report.py"),
        "oracle_role": "after-streaming", "driver_sha256": sha(Path(__file__)),
        "stat_events": list(EVENTS) if args.kind == "stat" else [],
        "record_event": "cycles:u" if args.kind == "record" else None,
        "record_frequency_hz": 999 if args.kind == "record" else None,
        "call_graph": "fp,127" if args.kind == "record" else None,
        "started_utc": now(), "status": "running",
    }
    write(receipt, row)
    env = os.environ.copy()
    env.update({"RUSTUP_TOOLCHAIN": "1.98.1", "DEBUGINFOD_URLS": "", "PYTHONDONTWRITEBYTECODE": "1"})
    try:
        with workload_log.open("xb") as stream:
            result = subprocess.run(argv, cwd=REPO, env=env, stdout=stream, stderr=subprocess.STDOUT)
        row["exit_code"] = result.returncode
        if result.returncode != 0:
            raise RuntimeError(f"profiler workload exited {result.returncode}")
        oracle_argv = [sys.executable, "-B", str(ROOT / "oracle" / "verify-report.py"), "--report", str(report), "--mode", "normal", "--shape", "large", "--role", "after-streaming"]
        oracle_result = subprocess.run(oracle_argv, cwd=REPO, env=env, capture_output=True, text=True)
        (profile_dir / "oracle.log").write_text("stdout=" + oracle_result.stdout + "stderr=" + oracle_result.stderr, encoding="utf-8")
        row["oracle_exit_code"] = oracle_result.returncode
        if oracle_result.returncode != 0 or oracle_result.stdout.strip() != "VALID":
            raise RuntimeError("copied oracle rejected profile workload report")
        if args.kind == "record":
            for command, output in ((["perf", "report", "--stdio", "--no-children", "--percent-limit", "0", "-i", str(data)], perf_report), (["perf", "script", "-i", str(data)], script_output)):
                with output.open("xb") as stream:
                    subprocess.run(command, cwd=REPO, env=env, stdout=stream, stderr=subprocess.STDOUT, check=True)
        row["status"] = "pass"
    except Exception as error:
        row["status"] = "failed"
        row["error"] = repr(error)
    finally:
        row["source_after"] = custody_sources()
        row["source_unchanged"] = row["source_before"] == row["source_after"]
        if not row["source_unchanged"]:
            row["status"] = "failed"
            row["error"] = "source custody changed during profile"
        row["finished_utc"] = now()
        row["artifacts"] = {
            key: record(path) for key, path in {
                "report": report, "catalog": catalog, "resource": resource, "workload_log": workload_log,
                "perf_stat": stat_output, "perf_data": data, "perf_script": script_output,
                "perf_report": perf_report, "oracle_log": profile_dir / "oracle.log",
            }.items() if path.is_file()
        }
        write(receipt, row)
    print(json.dumps({"status": row["status"], "role": args.role, "kind": args.kind}))
    return 0 if row["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
