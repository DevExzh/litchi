#!/usr/bin/env python3
"""Record the real launch PID/affinity, then exec the exact baseline binary."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path


SCHEMA = "litchi.performance.0970.launch.v2"


def utc() -> str:
    return datetime.now(timezone.utc).isoformat()


def helper_sha256() -> str:
    digest = hashlib.sha256()
    with Path(__file__).resolve().open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def cgroup_lines() -> list[str]:
    return Path("/proc/self/cgroup").read_text(encoding="utf-8").splitlines()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--metadata", type=Path, required=True)
    parser.add_argument("--requested-affinity", required=True)
    parser.add_argument("target", nargs=argparse.REMAINDER)
    args = parser.parse_args()
    try:
        requested = json.loads(args.requested_affinity)
        if (
            not isinstance(requested, list)
            or not requested
            or any(type(item) is not int or item < 0 for item in requested)
            or sorted(set(requested)) != requested
        ):
            raise ValueError("requested affinity must be a sorted unique CPU list")
        if args.target and args.target[0] == "--":
            args.target = args.target[1:]
        if not args.target:
            raise ValueError("exec target is missing")
        target = str(Path(args.target[0]).resolve())
        argv = [target, *args.target[1:]]
        if not os.path.isabs(target) or not os.access(target, os.X_OK):
            raise ValueError(f"target is not executable: {target}")
        actual = sorted(os.sched_getaffinity(0))
        if actual != requested:
            raise ValueError(f"requested/actual affinity mismatch: {requested} != {actual}")
        metadata = args.metadata.resolve()
        packet_id = metadata.name[:-len(".launch.json")] if metadata.name.endswith(".launch.json") else metadata.stem
        payload = {
            "schema": SCHEMA,
            "packet_id": packet_id,
            "helper_path": str(Path(__file__).resolve()),
            "helper_sha256": helper_sha256(),
            "pid": os.getpid(),
            "exec_pid": os.getpid(),
            "requested_affinity": requested,
            "actual_affinity": actual,
            "cgroup": cgroup_lines(),
            "created_utc": utc(),
            "exec_ready_utc": None,
            "exec_path": target,
            "exec_argv": argv,
            "execve": True,
        }
        payload["exec_ready_utc"] = utc()
        encoded = (json.dumps(payload, indent=2, sort_keys=True, allow_nan=False) + "\n").encode("utf-8")
        metadata.parent.mkdir(parents=True, exist_ok=True)
        fd = os.open(metadata, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o644)
        try:
            with os.fdopen(fd, "wb") as handle:
                handle.write(encoded)
                handle.flush()
                os.fsync(handle.fileno())
        except Exception:
            try:
                metadata.unlink(missing_ok=True)
            except OSError:
                pass
            raise
        os.execve(target, argv, os.environ.copy())
    except (OSError, ValueError, TypeError, json.JSONDecodeError) as exc:
        print(f"launch_baseline_v2 rejected launch: {exc}", file=sys.stderr)
        return 2
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
