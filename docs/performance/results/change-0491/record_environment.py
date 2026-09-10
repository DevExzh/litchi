#!/usr/bin/env python3
"""Record current host and explicit scratch-filesystem evidence before capture."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import platform
import shutil
import subprocess

from support import ENV, ENV_KEYS, REPO, now, write


def command(argv: list[str]) -> dict:
    started = now()
    executable = shutil.which(argv[0])
    if executable is None:
        return {"argv": argv, "status": "unavailable", "started_utc": started}
    try:
        result = subprocess.run(argv, cwd=REPO, env=ENV, capture_output=True, timeout=30)
    except (OSError, subprocess.TimeoutExpired) as error:
        return {"argv": argv, "status": "failed", "error": str(error), "started_utc": started, "finished_utc": now()}
    return {
        "argv": argv, "executable": executable,
        "status": "pass" if result.returncode == 0 else "failed",
        "exit_code": result.returncode,
        "started_utc": started, "finished_utc": now(),
        "stdout": result.stdout.decode("utf-8", errors="replace"),
        "stderr": result.stderr.decode("utf-8", errors="replace"),
        "stdout_sha256": hashlib.sha256(result.stdout).hexdigest(),
        "stderr_sha256": hashlib.sha256(result.stderr).hexdigest(),
    }


def optional_text(path: str) -> dict:
    try:
        value = Path(path).read_text()
        return {"path": path, "status": "observed", "text": value}
    except OSError as error:
        return {"path": path, "status": "unavailable", "error": str(error)}


def collect(scratch: Path, cpu: int) -> dict:
    scratch = scratch.resolve(strict=True)
    if not scratch.is_dir():
        raise ValueError("scratch capability must be an existing directory")
    affinity = sorted(os.sched_getaffinity(0)) if hasattr(os, "sched_getaffinity") else None
    if affinity is not None and cpu not in affinity:
        raise ValueError(f"selected CPU {cpu} is outside coordinator affinity {affinity}")
    stat = os.statvfs(scratch)
    tools = {
        "cpu": ["lscpu", "--json"],
        "rustc": ["rustc", "-Vv"],
        "cargo": ["cargo", "-V"],
        "perf": ["perf", "--version"],
        "strace": ["strace", "-V"],
        "heaptrack": ["heaptrack", "--version"],
        "time": ["/usr/bin/time", "--version"],
        "libreoffice": ["libreoffice", "--version"],
        "scratch_mount": ["findmnt", "--json", "--target", str(scratch), "--output", "TARGET,SOURCE,FSTYPE,OPTIONS"],
        "workspace_mount": ["findmnt", "--json", "--target", str(REPO), "--output", "TARGET,SOURCE,FSTYPE,OPTIONS"],
    }
    receipts = {name: command(argv) for name, argv in tools.items()}
    for required in ("cpu", "rustc", "cargo", "time", "scratch_mount"):
        if receipts[required]["status"] != "pass":
            raise ValueError(f"required environment observation unavailable: {required}")
    return {
        "schema": "docx-stream-route-machine-v1",
        "recorded_utc": now(),
        "driver_sha256": hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
        "platform": platform.platform(),
        "machine": platform.machine(),
        "python": platform.python_version(),
        "logical_cpu_count": os.cpu_count(),
        "coordinator_affinity": affinity,
        "selected_cpu": cpu,
        "environment": {key: ENV[key] for key in ENV_KEYS},
        "os_release": optional_text("/etc/os-release"),
        "meminfo": optional_text("/proc/meminfo"),
        "cgroup_cpu": optional_text("/sys/fs/cgroup/cpu.max"),
        "cgroup_memory": optional_text("/sys/fs/cgroup/memory.max"),
        "scratch": {
            "path": str(scratch),
            "device": scratch.stat().st_dev,
            "block_size": stat.f_bsize,
            "fragment_size": stat.f_frsize,
            "total_bytes": stat.f_blocks * stat.f_frsize,
            "available_bytes_at_recording": stat.f_bavail * stat.f_frsize,
        },
        "cache_policy": {
            "warm": "fresh child plus untimed priming before measured child",
            "cold_requested": "advisory DONTNEED; no verified eviction claim",
            "cold_verified": "private aligned copy; sync, DONTNEED, zero clean residency proof and positive timed process read_bytes",
            "provider_matrix": "owned or recently staged file; explicitly warm with logical simulated ranges",
            "host_wide_cache_drops": False,
        },
        "commands": receipts,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--scratch-dir", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--cpu", type=int, default=2)
    args = parser.parse_args()
    if args.output.exists():
        parser.error(f"refusing to replace {args.output}")
    value = collect(args.scratch_dir, args.cpu)
    write(args.output, value)
    print(json.dumps({"output": str(args.output), "schema": value["schema"], "cpu": args.cpu}))


if __name__ == "__main__":
    main()
