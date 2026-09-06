#!/usr/bin/env python3
"""Capture one required 0435 large-normal whole-process profile.

The six formal invocations are independent fresh processes: each of the
three ODT roles by ``--role`` and perf-stat/perf-record by ``--kind``.  They
are diagnostics only; the formal latency matrix remains in ``capture.py``.
An independent before-build diagnostic can use ``--preparatory`` and an
attempt tag before the final descriptors and frozen protocol exist.
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
REPO = ROOT.parents[3]
ROLES = {
    "before-buffered": "before",
    "after-buffered": "after",
    "after-streaming": "after",
}
KINDS = {"stat", "record"}
EVENTS = ("cycles:u", "instructions:u", "branches:u", "branch-misses:u", "L1-dcache-load-misses:u")
ATTEMPT_RE = re.compile(r"^[a-z0-9][a-z0-9_-]{0,63}$")


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
    spec = importlib.util.spec_from_file_location("change0435_profile_custody", ROOT / "check.py")
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load frozen custody driver")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.sources()


def oracle_role(protocol: dict[str, Any], role: str) -> str:
    """Resolve the role-specific oracle role without a hardcoded candidate."""

    spec = protocol["roles"][role]
    if isinstance(spec.get("oracle_role"), str):
        return spec["oracle_role"]
    role_oracle = spec.get("oracle")
    if isinstance(role_oracle, dict) and isinstance(role_oracle.get("role"), str):
        return role_oracle["role"]
    mapped = protocol.get("oracle_roles")
    if isinstance(mapped, dict) and isinstance(mapped.get(role), str):
        return mapped[role]
    global_oracle = protocol.get("oracle")
    if isinstance(global_oracle, dict):
        mapped = global_oracle.get("roles")
        if isinstance(mapped, dict) and isinstance(mapped.get(role), str):
            return mapped[role]
    return role


def preparatory_inputs(protocol_path: Path, protocol: dict[str, Any]) -> dict[str, Any]:
    """Load the ambient before build without requiring final descriptors.

    A preparatory profile deliberately binds the copied before executable to
    the passing build receipt and draft protocol.  The formal descriptors are
    created later, so consulting them here would make a before profile depend
    on the candidate/frozen-input phase it is intended to precede.
    """

    if protocol.get("status") != "draft":
        raise ValueError("protocol-draft.json must retain status=draft for preparatory profiling")
    build_path = ROOT / "checks" / "before-build.json"
    build = load(build_path)
    if build.get("change") != 435 or build.get("status") != "pass" or build.get("exit_code") != 0:
        raise ValueError("checks/before-build.json is not a passing 0435 build receipt")
    if build.get("source_unchanged") is not True:
        raise ValueError("checks/before-build.json does not prove source custody")
    source_before = build.get("source_before")
    source_after = build.get("source_after")
    if not isinstance(source_before, dict) or source_before != source_after:
        raise ValueError("checks/before-build.json has no stable source manifest")
    revision = build.get("revision")
    if not isinstance(revision, str) or not revision:
        raise ValueError("checks/before-build.json has no source revision")

    copies_path = ROOT / "before" / "binary-copies.json"
    copies = load(copies_path)
    if not isinstance(copies, dict) or set(copies) != {"normal", "allocator"}:
        raise ValueError("before/binary-copies.json must retain normal and allocator binaries")
    for mode in ("normal", "allocator"):
        identity = copies[mode]
        if not isinstance(identity, dict):
            raise ValueError(f"before/binary-copies.json {mode} identity is malformed")
        binary_path_value = identity.get("path")
        binary_bytes = identity.get("bytes")
        binary_sha = identity.get("sha256")
        if not isinstance(binary_path_value, str) or not binary_path_value:
            raise ValueError(f"before/binary-copies.json {mode} path is malformed")
        if isinstance(binary_bytes, bool) or not isinstance(binary_bytes, int) or binary_bytes <= 0:
            raise ValueError(f"before/binary-copies.json {mode} byte count is malformed")
        if not isinstance(binary_sha, str) or re.fullmatch(r"[0-9a-fA-F]{64}", binary_sha) is None:
            raise ValueError(f"before/binary-copies.json {mode} hash is malformed")
        binary_path = resolve(binary_path_value)
        if not binary_path.is_file() or binary_path.stat().st_size != binary_bytes or sha(binary_path) != binary_sha:
            raise ValueError(f"before/{mode} copied binary is stale or unavailable")

    oracle_path = ROOT / "verify-report.py"
    if not oracle_path.is_file():
        raise ValueError("root verify-report.py is unavailable for preparatory verification")
    return {
        "protocol_path": protocol_path,
        "protocol": protocol,
        "build_path": build_path,
        "build": build,
        "copies_path": copies_path,
        "copies": copies,
        "ambient_source": source_before,
        "source_manifest": source_before,
        "binary": copies["normal"],
        "revision": revision,
        "oracle_path": oracle_path,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--role", choices=tuple(ROLES), required=True)
    parser.add_argument("--kind", choices=tuple(sorted(KINDS)), required=True)
    parser.add_argument("--preparatory", action="store_true",
                        help="profile the ambient before build using protocol-draft.json")
    parser.add_argument("--attempt", help="unique preparatory attempt tag retained under profiles/")
    args = parser.parse_args()
    if args.attempt is not None and ATTEMPT_RE.fullmatch(args.attempt) is None:
        parser.error("--attempt must match [a-z0-9][a-z0-9_-]{0,63}")
    if args.preparatory:
        if args.role != "before-buffered":
            parser.error("--preparatory is only valid for --role before-buffered")
        attempt = args.attempt or "initial"
        protocol_path = ROOT / "protocol-draft.json"
        protocol = load(protocol_path)
        context = preparatory_inputs(protocol_path, protocol)
        build = context["build"]
        build_path = context["build_path"]
        copies_path = context["copies_path"]
        ambient_dir = "before"
        ambient_source = context["ambient_source"]
        source_manifest = context["source_manifest"]
        binary = context["binary"]
        revision = context["revision"]
        oracle_path = context["oracle_path"]
    else:
        if args.attempt is not None:
            parser.error("--attempt requires --preparatory")
        attempt = "formal"
        role_dir = ROLES[args.role]
        protocol_path = ROOT / "protocol.json"
        protocol = load(protocol_path)
        build_path = ROOT / role_dir / "build.json"
        build = load(build_path)
        # Matched final profiles all run from the candidate checkout. The
        # historical executable keeps its own source/build identity; the
        # pre-implementation profile is retained separately above.
        ambient_dir = "after"
        ambient_build = load(ROOT / ambient_dir / "build.json")
        ambient_source = ambient_build["source_manifest"]
        source_manifest = build["source_manifest"]
        binary = build["binaries"]["normal"]
        revision = build["revision"]
        copies_path = None
        oracle_path = ROOT / "oracle" / "verify-report.py"
    source_before = custody_sources()
    if source_before != ambient_source:
        print(f"PROFILE INVALID: ambient source differs from {args.role} build", file=sys.stderr)
        return 2
    profile_root = "preparatory-before-buffered" if args.preparatory else args.role
    profile_dir = ROOT / "profiles" / profile_root
    if args.preparatory:
        profile_dir /= attempt
    profile_dir /= args.kind
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
        "schema": "litchi-0435-profile-receipt-v1", "change": 435, "role": args.role, "kind": args.kind,
        "scope": "whole fresh process including setup, corpus generation, warmups, samples, output hashing, and oracle",
        "selector": selector, "shape": "large", "argv": argv, "cwd": str(REPO),
        "revision": revision, "source_manifest": source_manifest, "binary": binary,
        "ambient_source_manifest": ambient_source, "source_before": source_before,
        "protocol_path": str(protocol_path.relative_to(ROOT)), "protocol_sha256": sha(protocol_path),
        "driver_sha256": sha(Path(__file__)), "oracle_verifier_path": str(oracle_path.relative_to(ROOT)),
        "oracle_verifier_sha256": sha(oracle_path), "attempt": attempt,
        "preparatory": args.preparatory,
        "ambient_role": args.role,
        "ambient_build_directory": ambient_dir,
        "oracle_role": oracle_role(protocol, args.role),
        "stat_events": list(EVENTS) if args.kind == "stat" else [],
        "record_event": "cycles:u" if args.kind == "record" else None,
        "record_frequency_hz": 999 if args.kind == "record" else None,
        "call_graph": "fp,127" if args.kind == "record" else None,
        "started_utc": now(), "status": "running",
    }
    if args.preparatory:
        row.update({
            "build_receipt": str(build_path.relative_to(ROOT)),
            "build_receipt_sha256": sha(build_path),
            "binary_copies_receipt": str(copies_path.relative_to(ROOT)),
            "binary_copies_receipt_sha256": sha(copies_path),
            "preparatory_protocol_status": protocol.get("status"),
        })
    write(receipt, row)
    env = os.environ.copy()
    env.update({"RUSTUP_TOOLCHAIN": "1.98.1", "DEBUGINFOD_URLS": "", "PYTHONDONTWRITEBYTECODE": "1"})
    try:
        with workload_log.open("xb") as stream:
            result = subprocess.run(argv, cwd=REPO, env=env, stdout=stream, stderr=subprocess.STDOUT)
        row["exit_code"] = result.returncode
        if result.returncode != 0:
            raise RuntimeError(f"profiler workload exited {result.returncode}")
        oracle_argv = [sys.executable, "-B", str(oracle_path), "--report", str(report), "--mode", "normal", "--shape", "large", "--role", row["oracle_role"]]
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
