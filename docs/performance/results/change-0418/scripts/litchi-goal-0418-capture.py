#!/usr/bin/env python3
"""Run the serialized 0418 ABBA normal and allocator captures.

The protocol supplies the four selectors, their normal sample counts, and the
ABBA role order.  Every selector is run as ``A1, B1, B2, A2`` for both the
normal and allocator binaries; the protocol marks allocator latency as
excluded, and later validation decides which rows expose operation-local
allocation metrics.  Before any formal run, samples=1/warmup=0 preflights
cover every selector, every ABBA leg, and both binary phases.

Each command gets a fresh process, CPU 2, a matching clean role worktree,
``/usr/bin/time -v``, and separate report/catalog/timing/stdout/stderr files.
The manifest is atomically updated after every command so a failure leaves a
complete command journal without being mistaken for a successful capture.
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

from _common import (
    ABBA_LEGS,
    ABBA_ROLES,
    CHANGE_ROOT,
    BatchError,
    append_failure,
    assert_source_matches,
    host_identity,
    load_build_identity,
    load_json,
    load_protocol,
    relative_path,
    require_new,
    require_output_files,
    role_binary,
    selector_stem,
    sha256_file,
    utc_now,
    validate_binary_descriptor,
    write_json,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=CHANGE_ROOT)
    parser.add_argument("--protocol", type=Path)
    parser.add_argument("--build-identity", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument(
        "--mode",
        choices=("all", "preflight", "normal", "allocator"),
        default="all",
        help="run preflight and formal lanes (default), or one explicit lane",
    )
    parser.add_argument(
        "--resume",
        action="store_true",
        help="reuse only existing successful run records with intact artifacts",
    )
    return parser.parse_args()


def _run_key(preflight: bool, phase: str, leg: str, selector: str) -> str:
    prefix = "preflight" if preflight else phase
    # Normal and allocator preflights intentionally exercise the same
    # selectors, so phase must remain part of their identity.
    return f"{prefix}:{phase}:{leg}:{selector}"


def _jobs_for_phase(protocol: dict[str, Any], phase: str) -> list[dict[str, Any]]:
    return protocol["jobs"] if phase == "normal" else protocol["allocator_jobs"]


def _plan(protocol: dict[str, Any], mode: str) -> list[dict[str, Any]]:
    phases: list[str]
    if mode == "all":
        phases = ["normal", "allocator"]
        include_preflight = True
    elif mode == "preflight":
        phases = ["normal", "allocator"]
        include_preflight = True
    else:
        phases = [mode]
        include_preflight = False
    entries: list[dict[str, Any]] = []
    if include_preflight:
        for phase in phases:
            for job in _jobs_for_phase(protocol, phase):
                for leg, role in zip(ABBA_LEGS, ABBA_ROLES):
                    entries.append({
                        "preflight": True,
                        "phase": phase,
                        "leg": leg,
                        "role": role,
                        "selector": job["selector"],
                        "samples": 1,
                        "warmups": 0,
                    })
    if mode != "preflight":
        for phase in phases:
            for job in _jobs_for_phase(protocol, phase):
                for leg, role in zip(ABBA_LEGS, ABBA_ROLES):
                    if phase == "normal":
                        samples, warmups = job["samples"], job["warmups"]
                    else:
                        samples = protocol["allocator"]["samples"]
                        warmups = protocol["allocator"]["warmups"]
                    entries.append({
                        "preflight": False,
                        "phase": phase,
                        "leg": leg,
                        "role": role,
                        "selector": job["selector"],
                        "samples": samples,
                        "warmups": warmups,
                    })
    return entries


def _initial_manifest(
    root: Path, protocol: dict[str, Any], build: dict[str, Any],
    build_path: Path,
) -> dict[str, Any]:
    expected_all = _plan(protocol, "all")
    expected = [
        {
            "key": _run_key(item["preflight"], item["phase"], item["leg"], item["selector"]),
            **item,
        }
        for item in expected_all
    ]
    return {
        "schema_version": 1,
        "change": 418,
        "status": "starting",
        "protocol": {
            "path": relative_path(protocol["path"], root),
            "sha256": protocol["sha256"],
        },
        "build_identity": {
            "path": relative_path(build_path, root),
            "sha256": sha256_file(build_path),
        },
        "source_roles": {
            role: build["roles"][role]["source"] for role in ("control", "candidate")
        },
        "binaries": {
            role: build["roles"][role]["binaries"]
            for role in ("control", "candidate")
        },
        "execution": {
            "cpu": protocol["cpu"],
            "workers": protocol["workers"],
            "abba_legs": list(ABBA_LEGS),
            "abba_roles": list(ABBA_ROLES),
            "common_flags": protocol["common_flags"],
            "common_flags_source": protocol["common_flags_source"],
            "filesystem_cache": "warm",
            "fresh_process_per_selector_per_leg": True,
            "allocator_latency_comparison": "excluded by protocol",
        },
        "host": host_identity(include_perf=False),
        "environment_overrides": {"RUSTUP_TOOLCHAIN": "1.98.1"},
        "expected": expected,
        "runs": [],
    }


def _check_existing_manifest(
    manifest: dict[str, Any], protocol: dict[str, Any], build: dict[str, Any],
    build_path: Path, *, root: Path,
) -> None:
    if manifest.get("change") != 418:
        raise BatchError("existing capture manifest is not change 0418")
    if manifest.get("protocol", {}).get("sha256") != protocol["sha256"]:
        raise BatchError("existing capture protocol hash differs")
    if manifest.get("build_identity", {}).get("sha256") != sha256_file(build_path):
        raise BatchError("existing capture build identity hash differs")
    for run in manifest.get("runs", []):
        if not isinstance(run, dict):
            raise BatchError("existing capture contains a malformed run")
        if run.get("exit_code") != 0:
            raise BatchError("refusing to resume over a previously failed run")
        artifact_paths: dict[str, Path] = {}
        for key in ("report", "catalog", "time_v", "stdout", "stderr"):
            value = run.get(key)
            if not isinstance(value, str):
                raise BatchError(f"existing run {run.get('key')!r} is missing {key}")
            path = Path(value)
            artifact_paths[key] = path if path.is_absolute() else root / path
            if not artifact_paths[key].is_file():
                raise BatchError(f"existing run artifact is missing: {artifact_paths[key]}")
        for key in ("report", "catalog", "time_v"):
            if artifact_paths[key].stat().st_size == 0:
                raise BatchError(f"existing run artifact is empty: {artifact_paths[key]}")
        recorded_hashes = run.get("artifact_sha256")
        if not isinstance(recorded_hashes, dict):
            raise BatchError(f"existing run {run.get('key')!r} has no artifact hashes")
        for key, path in artifact_paths.items():
            if recorded_hashes.get(key) != sha256_file(path):
                raise BatchError(f"existing run artifact hash changed: {path}")


def _require_preflight_complete(
    manifest: dict[str, Any], protocol: dict[str, Any]
) -> None:
    expected = {
        _run_key(item["preflight"], item["phase"], item["leg"], item["selector"])
        for item in _plan(protocol, "preflight")
    }
    completed = {
        run["key"] for run in manifest.get("runs", [])
        if isinstance(run, dict) and run.get("exit_code") == 0
    }
    missing = sorted(expected - completed)
    if missing:
        raise BatchError(
            "formal capture requires successful preflight for every selector/leg/phase; "
            f"missing {len(missing)} run(s), first={missing[0]!r}"
        )


def _artifact_paths(root: Path, item: dict[str, Any]) -> dict[str, Path]:
    selector = selector_stem(item["selector"])
    leg = item["leg"].lower()
    if item["preflight"]:
        stem = f"{leg}-{item['phase']}-{selector}"
        directory = root / "preflight"
    else:
        stem = f"{leg}-{selector}"
        directory = root / item["phase"]
    return {
        "report": directory / f"{stem}.json",
        "catalog": directory / f"{stem}.catalog.json",
        "time_v": directory / f"{stem}.time.txt",
        "stdout": directory / f"{stem}.stdout.txt",
        "stderr": directory / f"{stem}.stderr.txt",
    }


def _run_one(
    root: Path, protocol: dict[str, Any], build: dict[str, Any],
    item: dict[str, Any], manifest: dict[str, Any], manifest_path: Path,
) -> dict[str, Any]:
    role = item["role"]
    phase = item["phase"]
    descriptor = role_binary(build, role, phase)
    actual_binary = validate_binary_descriptor(descriptor, label=f"{role}/{phase}")
    worktree = Path(build["roles"][role]["source"]["worktree"]).resolve()
    source_before = assert_source_matches(build, role)
    paths = _artifact_paths(root, item)
    paths["report"].parent.mkdir(parents=True, exist_ok=True)
    require_new(list(paths.values()))

    args = [
        actual_binary["path"],
        "--case",
        item["selector"],
        *protocol["common_flags"],
        "--samples",
        str(item["samples"]),
        "--warmup",
        str(item["warmups"]),
        "--json",
        str(paths["report"].resolve()),
        "--corpus-manifest",
        str(paths["catalog"].resolve()),
    ]
    argv = [
        "taskset",
        "-c",
        str(protocol["cpu"]),
        "/usr/bin/time",
        "-v",
        "-o",
        str(paths["time_v"].resolve()),
        *args,
    ]
    env = os.environ.copy()
    env["RUSTUP_TOOLCHAIN"] = "1.98.1"
    started = utc_now()
    monotonic_started = time.monotonic()
    with paths["stdout"].open("wb") as stdout, paths["stderr"].open("wb") as stderr:
        process = subprocess.run(
            argv,
            cwd=worktree,
            env=env,
            stdout=stdout,
            stderr=stderr,
            check=False,
        )
    finished = utc_now()
    # The descriptors and source binding are checked after every child too;
    # an accidental source or copied-binary mutation cannot be hidden by a
    # later successful leg.
    source_after = assert_source_matches(build, role)
    actual_after = validate_binary_descriptor(descriptor, label=f"{role}/{phase}")
    if process.returncode == 0:
        try:
            require_output_files([paths["report"], paths["catalog"], paths["time_v"]])
        except BatchError as exc:
            append_failure(
                root,
                {
                    "stage": "capture-artifacts",
                    "utc": finished,
                    "run_key": _run_key(
                        item["preflight"], phase, item["leg"], item["selector"]
                    ),
                    "argv": argv,
                    "cwd": str(worktree),
                    "message": str(exc),
                },
            )
            raise
    record: dict[str, Any] = {
        "key": _run_key(item["preflight"], phase, item["leg"], item["selector"]),
        **item,
        "argv": argv,
        "cwd": str(worktree),
        "source_revision": source_before["revision"],
        "source_files": source_before["source_files"],
        "binary_sha256": actual_after["sha256"],
        "binary_bytes": actual_after["bytes"],
        "environment_overrides": {"RUSTUP_TOOLCHAIN": "1.98.1"},
        "started_utc": started,
        "finished_utc": finished,
        "elapsed_seconds": time.monotonic() - monotonic_started,
        "exit_code": process.returncode,
        "report": relative_path(paths["report"], root),
        "catalog": relative_path(paths["catalog"], root),
        "time_v": relative_path(paths["time_v"], root),
        "stdout": relative_path(paths["stdout"], root),
        "stderr": relative_path(paths["stderr"], root),
        "artifact_sha256": {
            key: sha256_file(path) for key, path in paths.items() if path.is_file()
        },
        "source_after": source_after,
    }
    manifest["runs"].append(record)
    write_json(manifest_path, manifest)
    if process.returncode != 0:
        append_failure(
            root,
            {
                "stage": "capture",
                "utc": finished,
                "run_key": record["key"],
                "argv": argv,
                "cwd": str(worktree),
                "exit_code": process.returncode,
                "stderr": record["stderr"],
            },
        )
        raise BatchError(f"capture failed for {record['key']} with exit {process.returncode}")
    return record


def main() -> int:
    args = parse_args()
    root = args.root.resolve()
    protocol_path = (args.protocol or (root / "protocol.json")).resolve()
    build_path = (args.build_identity or (root / "build-identity.json")).resolve()
    manifest_path = (args.output or (root / "capture.json")).resolve()
    try:
        protocol = load_protocol(root, protocol_path)
        build = load_build_identity(build_path)
        if build.get("protocol", {}).get("sha256") != protocol["sha256"]:
            raise BatchError("build identity is bound to a different 0418 protocol")
        for role in ("control", "candidate"):
            assert_source_matches(build, role)

        if manifest_path.exists():
            if not args.resume:
                raise BatchError(f"refusing to overwrite existing capture manifest: {manifest_path}")
            manifest = load_json(manifest_path)
            if not isinstance(manifest, dict):
                raise BatchError("existing capture manifest is not an object")
            _check_existing_manifest(manifest, protocol, build, build_path, root=root)
        else:
            manifest = _initial_manifest(root, protocol, build, build_path)
            write_json(manifest_path, manifest)

        if args.mode in ("normal", "allocator"):
            _require_preflight_complete(manifest, protocol)

        planned = _plan(protocol, args.mode)
        completed = {
            run["key"] for run in manifest.get("runs", [])
            if isinstance(run, dict) and run.get("exit_code") == 0
        }
        expected_keys = {
            _run_key(item["preflight"], item["phase"], item["leg"], item["selector"])
            for item in planned
        }
        if expected_keys.intersection(completed) and not args.resume:
            raise BatchError("capture manifest already contains requested run keys")

        for item in planned:
            key = _run_key(item["preflight"], item["phase"], item["leg"], item["selector"])
            if key in completed:
                continue
            _run_one(root, protocol, build, item, manifest, manifest_path)
            completed.add(key)
            print(f"completed {key}", flush=True)

        full_expected = {
            entry["key"] for entry in manifest.get("expected", [])
            if isinstance(entry, dict) and isinstance(entry.get("key"), str)
        }
        if full_expected and full_expected <= completed:
            manifest["status"] = "complete"
        elif args.mode == "preflight":
            manifest["status"] = "preflight-complete"
        else:
            manifest["status"] = f"{args.mode}-complete"
        manifest["completed_utc"] = utc_now()
        manifest["completed_run_keys"] = sorted(completed)
        write_json(manifest_path, manifest)
        print(f"0418 capture complete: {manifest_path}")
        return 0
    except (BatchError, OSError, subprocess.SubprocessError) as exc:
        append_failure(
            root,
            {"stage": "capture-driver", "utc": utc_now(), "message": str(exc)},
        )
        print(f"0418 capture failed: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
