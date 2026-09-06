#!/usr/bin/env python3
"""Capture one frozen 0433 role matrix serially on protocol CPU 2."""

from __future__ import annotations

import datetime
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys


ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
ROLES = {"before-buffered": "before", "after-buffered": "after", "after-streaming": "after"}


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def load(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def role_spec(protocol: dict, role: str) -> dict:
    roles = protocol.get("roles")
    if not isinstance(roles, dict) or role not in roles or not isinstance(roles[role], dict):
        raise ValueError(f"protocol.roles.{role} is missing")
    spec = roles[role]
    selector = spec.get("selector")
    if not isinstance(selector, str) or not selector:
        raise ValueError(f"protocol.roles.{role}.selector is missing")
    return spec


def main() -> int:
    import argparse

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--role", choices=tuple(sorted(ROLES)), required=True)
    args = parser.parse_args()
    role = args.role
    protocol = load(ROOT / "protocol.json")
    spec = role_spec(protocol, role)
    build_root = ROOT / ROLES[role]
    build = load(build_root / "build.json")
    if build.get("change") != 433:
        raise ValueError("build.json change does not match 0433")

    custody_spec = importlib.util.spec_from_file_location("change0433_custody", ROOT / "check.py")
    if custody_spec is None or custody_spec.loader is None:
        raise RuntimeError("cannot load immutable custody driver")
    custody = importlib.util.module_from_spec(custody_spec)
    custody_spec.loader.exec_module(custody)
    if custody.sources() != build["source_manifest"]:
        raise ValueError("build source manifest is stale")
    cpu = int(protocol["cpu"])
    if cpu not in os.sched_getaffinity(0):
        raise ValueError("protocol CPU is not available to the capture process")
    if sha(ROOT / "protocol.json") != build["protocol_sha256"]:
        raise ValueError("protocol hash is stale in build.json")
    if sha(ROOT / "verify-report.py") != build["verifier_sha256"]:
        raise ValueError("verifier hash is stale in build.json")
    subprocess.run(["git", "diff", "--exit-code", "HEAD", "--"], cwd=REPO, check=True, stdout=subprocess.DEVNULL)
    status_before = subprocess.check_output(["git", "status", "--short"], cwd=REPO, text=True).splitlines()

    directory = build_root / "captures"
    directory.mkdir(parents=True, exist_ok=True)
    index_path = build_root / f"capture-index-{role}.json"
    state_path = build_root / f"capture-state-{role}.json"
    if index_path.exists() or state_path.exists():
        raise ValueError(f"capture for {role} already exists")
    order = protocol.get("order")
    if not isinstance(order, list) or not order:
        raise ValueError("protocol.order must contain the frozen lanes")
    index: list[str] = []
    for lane in order:
        if not isinstance(lane, dict):
            raise ValueError("protocol.order contains a non-object lane")
        mode = lane["mode"]
        shape = lane["shape"]
        repeat = lane["repeat"]
        name = f"{role}-{mode}-{shape}-{str(repeat).lower()}"
        binary = build["binaries"][mode]
        binary_path = Path(binary["path"])
        if sha(binary_path) != binary["sha256"] or binary_path.stat().st_size != binary["bytes"]:
            raise ValueError(f"binary identity is stale for {mode}")
        report = directory / f"{name}.json"
        catalog = directory / f"{name}-catalog.json"
        log = directory / f"{name}.log"
        resource = directory / f"{name}-resource.log"
        receipt = directory / f"{name}-receipt.json"
        argv = [
            "taskset", "-c", str(cpu), "/usr/bin/time", "-v", "-o", str(resource),
            binary["path"], "--case", spec["selector"], "--semantic-shape", shape,
            "--workers", str(protocol["workers"]), "--samples", str(protocol["samples"]),
            "--warmup", str(protocol["warmups"]), "--json", str(report),
            "--corpus-manifest", str(catalog),
        ]
        row = {
            "change": 433,
            "role": role,
            "selector": spec["selector"],
            "name": name,
            "lane": lane,
            "argv": argv,
            "revision": build["revision"],
            "source_manifest": build["source_manifest"],
            "binary": binary,
            "protocol_sha256": sha(ROOT / "protocol.json"),
            "driver_sha256": sha(Path(__file__)),
            "verifier_sha256": sha(ROOT / "verify-report.py"),
            "started_utc": now(),
            "status": "running",
        }
        receipt.write_text(json.dumps(row, indent=2) + "\n", encoding="utf-8")
        print(json.dumps({"status": "running", "role": role, "lane": name}), flush=True)
        try:
            with log.open("xb") as stream:
                result = subprocess.run(argv, cwd=REPO, stdout=stream, stderr=subprocess.STDOUT)
            row["exit_code"] = result.returncode
            if result.returncode != 0:
                raise RuntimeError(f"producer exited {result.returncode}")
            subprocess.run(
                [sys.executable, "-B", str(ROOT / "verify-report.py"), "--report", str(report),
                 "--mode", mode, "--shape", shape, "--role", role],
                check=True,
            )
            row["status"] = "pass"
        finally:
            if row["status"] == "running":
                row["status"] = "failed"
            row["finished_utc"] = now()
            row["artifacts"] = {
                str(path.relative_to(ROOT)): {"sha256": sha(path), "bytes": path.stat().st_size}
                for path in (report, catalog, log, resource) if path.is_file()
            }
            receipt.write_text(json.dumps(row, indent=2) + "\n", encoding="utf-8")
        index.append(str(receipt.relative_to(build_root)))

    if len(index) != len(order):
        raise ValueError("capture did not complete every frozen lane")
    if custody.sources() != build["source_manifest"]:
        raise ValueError("source manifest changed during capture")
    subprocess.run(["git", "diff", "--exit-code", "HEAD", "--"], cwd=REPO, check=True, stdout=subprocess.DEVNULL)
    status_after = subprocess.check_output(["git", "status", "--short"], cwd=REPO, text=True).splitlines()
    if status_after != status_before:
        raise ValueError("tracked or out-of-bundle files changed during capture")
    index_path.write_text(json.dumps(index, indent=2) + "\n", encoding="utf-8")
    state_path.write_text(json.dumps({
        "status": "pass",
        "role": role,
        "status_before": status_before,
        "status_after": status_after,
        "tracked_tree_clean_before_and_after": True,
        "report_dirty_field_expected": True,
        "source_manifest": build["source_manifest"],
    }, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"status": "pass", "role": role, "reports": len(index)}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
