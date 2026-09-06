#!/usr/bin/env python3
"""Capture one serialized 0434 ABBA phase.

The formal matrix is split into four invocations so that each fresh process
has one unambiguous phase tag: A1 (before, forward), B1 (after, forward), B2
(after, reverse), and A2 (before, reverse).  A failed attempt is left in its
tagged directory and can be retried with a different ``--attempt`` value.
This driver does not build binaries and never edits source files.
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
PHASES = {
    "A1": ("before-streaming", 0, 6),
    "B1": ("after-streaming", 0, 6),
    "B2": ("after-streaming", 6, 12),
    "A2": ("before-streaming", 6, 12),
}
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


def file_record(path: Path) -> dict[str, Any]:
    return {"path": str(path.relative_to(ROOT)), "bytes": path.stat().st_size, "sha256": sha(path)}


def resolve_repo_path(value: str) -> Path:
    path = Path(value)
    return path if path.is_absolute() else REPO / path


def custody_module():
    spec = importlib.util.spec_from_file_location("change0434_custody", ROOT / "check.py")
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load the frozen source-custody driver")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def status_outside_bundle() -> list[str]:
    raw = subprocess.check_output(
        ["git", "status", "--short", "--untracked-files=all"], cwd=REPO, text=True
    )
    bundle = ROOT.relative_to(REPO).as_posix() + "/"
    return [line for line in raw.splitlines() if not line[3:].startswith(bundle)]


def build_for(role: str, protocol: dict[str, Any]) -> tuple[Path, dict[str, Any]]:
    spec = protocol["roles"][role]
    build_dir = ROOT / spec["build_directory"]
    build_path = build_dir / "build.json"
    build = load(build_path)
    if build.get("change") != 434 or build.get("role") not in {None, spec["build_directory"], role}:
        raise ValueError(f"{build_path}: wrong change or role")
    if build.get("protocol_sha256") != sha(ROOT / "protocol.json"):
        raise ValueError(f"{build_path}: protocol hash is stale")
    outer_verifier_sha = build.get("outer_verifier_sha256", build.get("verifier_sha256"))
    # A schema-compatible descriptor may retain the copied oracle under the
    # historical verifier_sha256 key and add an explicit outer binding.  The
    # preferred 0434 descriptor uses verifier_sha256 for the outer verifier.
    if outer_verifier_sha != sha(ROOT / "verify.py") and build.get("verifier_sha256") != sha(ROOT / "oracle" / "verify-report.py"):
        raise ValueError(f"{build_path}: outer verifier hash is stale")
    if build.get("capture_driver_sha256") != sha(Path(__file__)):
        raise ValueError(f"{build_path}: capture driver hash is stale")
    source = build.get("source_manifest")
    if not isinstance(source, dict) or source.get("path", "").startswith("/"):
        raise ValueError(f"{build_path}: malformed source manifest")
    if not (ROOT / source["path"]).is_file():
        raise ValueError(f"{build_path}: source manifest is not retained")
    oracle = protocol["oracle"]
    if build.get("oracle_protocol_sha256", oracle["sha256"]) != oracle["sha256"]:
        raise ValueError(f"{build_path}: oracle protocol hash is stale")
    if build.get("oracle_verifier_sha256", oracle["verifier_sha256"]) != oracle["verifier_sha256"]:
        raise ValueError(f"{build_path}: oracle verifier hash is stale")
    binaries = build.get("binaries")
    if not isinstance(binaries, dict) or set(binaries) != {"normal", "allocator"}:
        raise ValueError(f"{build_path}: normal/allocator binary identities are required")
    for mode in ("normal", "allocator"):
        identity = binaries[mode]
        path = resolve_repo_path(identity["path"])
        if not path.is_file():
            raise ValueError(f"{build_path}: {mode} binary is missing: {path}")
        if path.stat().st_size != identity["bytes"] or sha(path) != identity["sha256"]:
            raise ValueError(f"{build_path}: {mode} binary identity is stale")
    return build_dir, build


def lane_artifacts(directory: Path, name: str) -> dict[str, Path]:
    return {
        "report": directory / f"{name}.json",
        "catalog": directory / f"{name}-catalog.json",
        "workload_log": directory / f"{name}.log",
        "resource_log": directory / f"{name}-resource.log",
        "oracle_log": directory / f"{name}-oracle.log",
        "receipt": directory / f"{name}-receipt.json",
    }


def expected_lanes(protocol: dict[str, Any], phase: str) -> list[dict[str, Any]]:
    order = protocol.get("order")
    if not isinstance(order, list) or len(order) != 12:
        raise ValueError("protocol.order must contain exactly 12 frozen lanes")
    _, first, last = PHASES[phase]
    lanes = order[first:last]
    if len(lanes) != 6:
        raise ValueError("each ABBA phase must contain six lanes")
    return lanes


def run_phase(phase: str, attempt: str) -> int:
    protocol = load(ROOT / "protocol.json")
    role, first, last = PHASES[phase]
    lanes = expected_lanes(protocol, phase)
    build_root, build = build_for(role, protocol)
    # Both role binaries are executed from the same ambient candidate
    # checkout.  The before binary's build manifest intentionally describes a
    # different historical source revision, so ambient custody is bound to the
    # after build for every ABBA phase and remains separate from the executable
    # role binding in each receipt.
    _, ambient_build = build_for("after-streaming", protocol)
    custody = custody_module()
    source_before = custody.sources()
    if source_before != ambient_build["source_manifest"]:
        raise ValueError("ambient source manifest differs from the after build")
    status_before = status_outside_bundle()
    directory = ROOT / "runs" / phase / attempt
    if directory.exists() and any(directory.iterdir()):
        raise ValueError(f"capture directory already contains evidence: {directory}")
    directory.mkdir(parents=True, exist_ok=True)
    outer_protocol_sha = sha(ROOT / "protocol.json")
    oracle = protocol["oracle"]
    driver_sha = sha(Path(__file__))
    outer_verifier_sha = sha(ROOT / "verify.py")
    spec = protocol["roles"][role]
    cpu = int(protocol["cpu"])
    if cpu not in os.sched_getaffinity(0):
        raise ValueError(f"protocol CPU {cpu} is unavailable")
    index: list[str] = []
    failed = False
    for lane in lanes:
        mode = lane["mode"]
        shape = lane["shape"]
        repeat = lane["repeat"]
        name = f"{phase}-{mode}-{shape}-{str(repeat).lower()}"
        artifacts = lane_artifacts(directory, name)
        if any(path.exists() for path in artifacts.values()):
            raise ValueError(f"lane already has an artifact: {name}")
        binary = build["binaries"][mode]
        binary_path = resolve_repo_path(binary["path"])
        argv = [
            "taskset", "-c", str(cpu), "/usr/bin/time", "-v", "-o", str(artifacts["resource_log"]),
            str(binary_path), "--case", spec["selector"], "--semantic-shape", shape,
            "--workers", str(protocol["workers"]), "--samples", str(protocol["samples"]),
            "--warmup", str(protocol["warmups"]), "--json", str(artifacts["report"]),
            "--corpus-manifest", str(artifacts["catalog"]),
        ]
        row: dict[str, Any] = {
            "schema": "litchi-0434-capture-receipt-v1",
            "change": 434,
            "phase": phase,
            "attempt": attempt,
            "role": role,
            "build_directory": build_root.name,
            "selector": spec["selector"],
            "name": name,
            "lane": lane,
            "argv": argv,
            "cwd": str(REPO),
            "revision": build["revision"],
            "source_manifest": build["source_manifest"],
            "binary": binary,
            "protocol_sha256": outer_protocol_sha,
            "oracle_protocol_sha256": oracle["sha256"],
            "oracle_verifier_sha256": oracle["verifier_sha256"],
            "driver_sha256": driver_sha,
            "verifier_sha256": outer_verifier_sha,
            "source_before": source_before,
            "ambient_source_manifest": ambient_build["source_manifest"],
            "status_before": status_before,
            "started_utc": now(),
            "status": "running",
        }
        write(artifacts["receipt"], row)
        print(json.dumps({"status": "running", "phase": phase, "role": role, "lane": name}), flush=True)
        try:
            env = os.environ.copy()
            env.update({"RUSTUP_TOOLCHAIN": "1.98.1", "PYTHONDONTWRITEBYTECODE": "1"})
            with artifacts["workload_log"].open("xb") as output:
                result = subprocess.run(argv, cwd=REPO, env=env, stdout=output, stderr=subprocess.STDOUT)
            row["exit_code"] = result.returncode
            if result.returncode != 0:
                raise RuntimeError(f"producer exited {result.returncode}")
            oracle_argv = [
                sys.executable, "-B", str(ROOT / "oracle" / "verify-report.py"),
                "--report", str(artifacts["report"]), "--mode", mode,
                "--shape", shape, "--role", oracle["role"],
            ]
            oracle_result = subprocess.run(oracle_argv, cwd=REPO, capture_output=True, text=True)
            artifacts["oracle_log"].write_text(
                "argv=" + json.dumps(oracle_argv) + "\nstdout=" + oracle_result.stdout
                + "stderr=" + oracle_result.stderr,
                encoding="utf-8",
            )
            row["oracle_exit_code"] = oracle_result.returncode
            row["oracle_stdout_sha256"] = sha(artifacts["oracle_log"])
            if oracle_result.returncode != 0 or oracle_result.stdout.strip() != "VALID":
                raise RuntimeError("copied 0433 oracle rejected the report")
            row["status"] = "pass"
        except Exception as error:  # retain a complete, uniquely tagged failure receipt
            row["error"] = repr(error)
            row["status"] = "failed"
            failed = True
        finally:
            row["finished_utc"] = now()
            row["source_after"] = custody.sources()
            row["status_after"] = status_outside_bundle()
            row["source_unchanged"] = row["source_before"] == row["source_after"]
            row["outside_bundle_status_unchanged"] = row["status_before"] == row["status_after"]
            row["artifacts"] = {
                key: file_record(path) for key, path in artifacts.items() if key != "receipt" and path.is_file()
            }
            write(artifacts["receipt"], row)
        index.append(str(artifacts["receipt"].relative_to(ROOT)))
        if failed:
            break
    state = {
        "schema": "litchi-0434-capture-state-v1",
        "change": 434,
        "phase": phase,
        "attempt": attempt,
        "role": role,
        "expected_lanes": len(lanes),
        "completed_lanes": len(index),
        "status": "failed" if failed else "pass",
        "source_before": source_before,
        "ambient_source_manifest": ambient_build["source_manifest"],
        "source_after": custody.sources(),
        "status_before": status_before,
        "status_after": status_outside_bundle(),
        "source_manifest": build["source_manifest"],
        "build_source_manifest": build["source_manifest"],
        "ambient_source_manifest": ambient_build["source_manifest"],
        "protocol_sha256": outer_protocol_sha,
        "oracle_protocol_sha256": oracle["sha256"],
        "oracle_verifier_sha256": oracle["verifier_sha256"],
        "driver_sha256": driver_sha,
        "index": index,
        "finished_utc": now(),
    }
    write(directory / "capture-index.json", index)
    write(directory / "capture-state.json", state)
    if failed or len(index) != len(lanes):
        return 1
    if custody.sources() != source_before or status_outside_bundle() != status_before:
        raise RuntimeError("source custody or outside-bundle status changed during capture")
    print(json.dumps({"status": "pass", "phase": phase, "role": role, "reports": len(index)}))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--phase", choices=tuple(PHASES), required=True)
    parser.add_argument("--attempt", default="formal", help="unique attempt tag retained under runs/<phase>")
    args = parser.parse_args()
    if ATTEMPT_RE.fullmatch(args.attempt) is None:
        parser.error("--attempt must match [a-z0-9][a-z0-9_-]{0,63}")
    try:
        return run_phase(args.phase, args.attempt)
    except (OSError, KeyError, TypeError, ValueError, subprocess.CalledProcessError) as error:
        print(f"CAPTURE INVALID: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
