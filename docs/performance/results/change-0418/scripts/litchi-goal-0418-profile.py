#!/usr/bin/env python3
"""Capture bounded 0418 CPU call graphs and optional paired PMU counters.

The default ``record`` mode profiles the protocol's primary lifecycle selector
once for each clean role worktree with 20 retained samples and 3 warmups,
``cycles:u`` at 999 Hz, and ``fp,127`` call chains.  It then emits standard
``perf report`` and ``perf script`` text from the retained ``perf.data``.
``--mode stat`` runs the protocol's paired whole-command PMU event set, and
``--mode all`` performs both lanes serially.  These are whole-process
diagnostics: corpus generation, setup, warmups, timed harness children,
verification, and report writing are included; no phase-local or speedup
claim is inferred.
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
    require_int,
    require_output_files,
    role_binary,
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
    parser.add_argument("--capture", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument(
        "--mode", choices=("record", "stat", "all"), default="record",
        help="CPU record (default), PMU stat, or both serially",
    )
    parser.add_argument(
        "--selector", action="append", dest="selectors",
        help="profile this selector; repeat for more (default: protocol primary)",
    )
    return parser.parse_args()


def _successful_capture_keys(capture: dict[str, Any]) -> set[str]:
    result: set[str] = set()
    for run in capture.get("runs", []):
        if isinstance(run, dict) and run.get("exit_code") == 0:
            key = run.get("key")
            if isinstance(key, str):
                result.add(key)
    return result


def _check_capture(
    capture_path: Path, capture: dict[str, Any], protocol: dict[str, Any],
    build_path: Path,
) -> None:
    if capture.get("change") != 418:
        raise BatchError("profile requires a change-0418 capture manifest")
    if capture.get("protocol", {}).get("sha256") != protocol["sha256"]:
        raise BatchError("capture and protocol hashes differ")
    if capture.get("build_identity", {}).get("sha256") != sha256_file(build_path):
        raise BatchError("capture and build identity hashes differ")
    expected = {
        entry.get("key") for entry in capture.get("expected", [])
        if isinstance(entry, dict) and isinstance(entry.get("key"), str)
    }
    successful = _successful_capture_keys(capture)
    missing = sorted(expected - successful)
    if missing:
        raise BatchError(
            f"profile requires all successful capture legs; missing {len(missing)}, "
            f"first={missing[0]!r}"
        )
    if capture.get("status") != "complete":
        raise BatchError("profile requires capture status=complete")


def _profile_paths(root: Path, role: str, selector: str) -> dict[str, Path]:
    stem = f"{role}-{selector}"
    directory = root / "profile"
    return {
        "data": directory / f"{stem}.perf.data",
        "report": directory / f"{stem}.perf-report.txt",
        "script": directory / f"{stem}.perf-script.txt",
        "time_v": directory / f"{stem}.time.txt",
        "workload_report": directory / f"{stem}.json",
        "workload_catalog": directory / f"{stem}.catalog.json",
        "stdout": directory / f"{stem}.stdout.txt",
        "stderr": directory / f"{stem}.stderr.txt",
        "postprocess_stderr": directory / f"{stem}.postprocess.stderr.txt",
    }


def _run_record(
    root: Path, protocol: dict[str, Any], build: dict[str, Any],
    role: str, selector: str, record: dict[str, Any], manifest_path: Path,
    manifest: dict[str, Any],
) -> None:
    descriptor = role_binary(build, role, "normal")
    binary = validate_binary_descriptor(descriptor, label=f"{role}/normal")
    source_before = assert_source_matches(build, role)
    worktree = Path(build["roles"][role]["source"]["worktree"]).resolve()
    paths = _profile_paths(root, role, selector)
    require_new(list(paths.values()))
    profile = protocol["profile"]
    args = [
        binary["path"], "--case", selector, *protocol["common_flags"],
        "--samples", str(profile["samples"]), "--warmup", str(profile["warmups"]),
        "--json", str(paths["workload_report"].resolve()),
        "--corpus-manifest", str(paths["workload_catalog"].resolve()),
    ]
    paths["data"].parent.mkdir(parents=True, exist_ok=True)
    argv = [
        "taskset", "-c", str(protocol["cpu"]), "/usr/bin/time", "-v", "-o",
        str(paths["time_v"].resolve()), "perf", "record", "--no-buildid-cache",
        "-e", profile["event"], "-F", str(profile["frequency"]),
        "--call-graph", profile["call_graph"], "-o", str(paths["data"].resolve()),
        "--", *args,
    ]
    env = os.environ.copy()
    env.update({"RUSTUP_TOOLCHAIN": "1.98.1", "DEBUGINFOD_URLS": ""})
    started = utc_now()
    monotonic_started = time.monotonic()
    with paths["stdout"].open("wb") as stdout, paths["stderr"].open("wb") as stderr:
        process = subprocess.run(
            argv, cwd=worktree, env=env, stdout=stdout, stderr=stderr, check=False
        )
    finished = utc_now()
    source_after = assert_source_matches(build, role)
    actual_binary = validate_binary_descriptor(descriptor, label=f"{role}/normal")
    if process.returncode != 0:
        record.update({
            "stage": "perf-record",
            "exit_code": process.returncode,
            "argv": argv,
            "cwd": str(worktree),
            "started_utc": started,
            "finished_utc": finished,
            "time_seconds": time.monotonic() - monotonic_started,
            "stderr": relative_path(paths["stderr"], root),
        })
        manifest["records"].append(record)
        write_json(manifest_path, manifest)
        append_failure(
            root,
            {"stage": "profile-record", "utc": finished, "argv": argv,
             "cwd": str(worktree), "exit_code": process.returncode,
             "stderr": relative_path(paths["stderr"], root)},
        )
        raise BatchError(f"perf record failed for {role}/{selector}")
    require_output_files([paths["data"], paths["time_v"]])

    postprocess = [
        (
            "perf-report",
            ["perf", "report", "--stdio", "--no-children", "--percent-limit", "0",
             "-i", str(paths["data"].resolve())],
            paths["report"],
        ),
        (
            "perf-script",
            ["perf", "script", "-i", str(paths["data"].resolve())],
            paths["script"],
        ),
    ]
    for stage, post_argv, output in postprocess:
        with output.open("wb") as stream, paths["postprocess_stderr"].open("ab") as error:
            post = subprocess.run(
                post_argv, cwd=worktree, env=env, stdout=stream, stderr=error, check=False
            )
        if post.returncode != 0:
            append_failure(
                root,
                {"stage": stage, "utc": utc_now(), "argv": post_argv,
                 "cwd": str(worktree), "exit_code": post.returncode,
                 "stderr": relative_path(paths["postprocess_stderr"], root)},
            )
            raise BatchError(f"{stage} failed for {role}/{selector}")
        require_output_files([output])

    record.update({
        "stage": "complete",
        "role": role,
        "selector": selector,
        "phase": "cpu-record",
        "samples": protocol["profile"]["samples"],
        "warmups": protocol["profile"]["warmups"],
        "argv": argv,
        "cwd": str(worktree),
        "source": source_before,
        "source_after": source_after,
        "binary_sha256": actual_binary["sha256"],
        "binary_bytes": actual_binary["bytes"],
        "environment_overrides": {"RUSTUP_TOOLCHAIN": "1.98.1", "DEBUGINFOD_URLS": ""},
        "started_utc": started,
        "finished_utc": finished,
        "elapsed_seconds": time.monotonic() - monotonic_started,
        "exit_code": 0,
        "artifacts": {
            key: {
                "path": relative_path(path, root),
                "bytes": path.stat().st_size,
                "sha256": sha256_file(path),
            }
            for key, path in paths.items() if path.is_file()
        },
        "scope": protocol["profile"]["scope"],
    })
    manifest["records"].append(record)
    write_json(manifest_path, manifest)


def _pmu_config(protocol: dict[str, Any]) -> dict[str, Any]:
    pmu = protocol["raw"].get("pmu")
    if not isinstance(pmu, dict):
        raise BatchError("--mode stat/all requires protocol.pmu")
    selector = pmu.get("selector")
    if not isinstance(selector, str):
        raise BatchError("protocol.pmu.selector must be a string")
    events = pmu.get("events")
    if not isinstance(events, list) or not events or not all(isinstance(v, str) for v in events):
        raise BatchError("protocol.pmu.events must be a non-empty string list")
    order = pmu.get("order")
    if order != ["control", "candidate"]:
        raise BatchError("protocol.pmu.order must be [control, candidate]")
    return {
        "selector": selector,
        "samples": require_int(pmu.get("samples"), "pmu.samples", minimum=1),
        "warmups": require_int(pmu.get("warmups"), "pmu.warmups", minimum=0),
        "events": events,
        "order": order,
        "scope": pmu.get("scope"),
    }


def _run_stat(
    root: Path, protocol: dict[str, Any], build: dict[str, Any],
    role: str, pmu: dict[str, Any], record: dict[str, Any], manifest_path: Path,
    manifest: dict[str, Any],
) -> None:
    descriptor = role_binary(build, role, "normal")
    binary = validate_binary_descriptor(descriptor, label=f"{role}/normal")
    source_before = assert_source_matches(build, role)
    worktree = Path(build["roles"][role]["source"]["worktree"]).resolve()
    stem = f"pmu-{role}-{pmu['selector']}"
    directory = root / "profile"
    paths = {
        "stat": directory / f"{stem}.stat.csv",
        "time_v": directory / f"{stem}.time.txt",
        "report": directory / f"{stem}.json",
        "catalog": directory / f"{stem}.catalog.json",
        "stdout": directory / f"{stem}.stdout.txt",
        "stderr": directory / f"{stem}.stderr.txt",
    }
    directory.mkdir(parents=True, exist_ok=True)
    require_new(list(paths.values()))
    args = [
        binary["path"], "--case", pmu["selector"], *protocol["common_flags"],
        "--samples", str(pmu["samples"]), "--warmup", str(pmu["warmups"]),
        "--json", str(paths["report"].resolve()),
        "--corpus-manifest", str(paths["catalog"].resolve()),
    ]
    argv = [
        "taskset", "-c", str(protocol["cpu"]), "/usr/bin/time", "-v", "-o",
        str(paths["time_v"].resolve()), "perf", "stat", "--no-big-num", "-x,",
        "-e", ",".join(pmu["events"]), "-o", str(paths["stat"].resolve()), "--", *args,
    ]
    env = os.environ.copy()
    env.update({"RUSTUP_TOOLCHAIN": "1.98.1", "DEBUGINFOD_URLS": ""})
    started = utc_now()
    monotonic_started = time.monotonic()
    with paths["stdout"].open("wb") as stdout, paths["stderr"].open("wb") as stderr:
        process = subprocess.run(
            argv, cwd=worktree, env=env, stdout=stdout, stderr=stderr, check=False
        )
    finished = utc_now()
    source_after = assert_source_matches(build, role)
    if process.returncode != 0:
        append_failure(
            root,
            {"stage": "profile-stat", "utc": finished, "argv": argv,
             "cwd": str(worktree), "exit_code": process.returncode,
             "stderr": relative_path(paths["stderr"], root)},
        )
        raise BatchError(f"perf stat failed for {role}/{pmu['selector']}")
    require_output_files([paths["stat"], paths["time_v"], paths["report"], paths["catalog"]])
    for log in (paths["stdout"], paths["stderr"]):
        if not log.is_file():
            raise BatchError(f"PMU command log is missing: {log}")
    manifest["records"].append({
        "stage": "complete",
        "role": role,
        "selector": pmu["selector"],
        "phase": "pmu-stat",
        "samples": pmu["samples"],
        "warmups": pmu["warmups"],
        "events": pmu["events"],
        "argv": argv,
        "cwd": str(worktree),
        "source": source_before,
        "source_after": source_after,
        "binary_sha256": binary["sha256"],
        "binary_bytes": binary["bytes"],
        "environment_overrides": {"RUSTUP_TOOLCHAIN": "1.98.1", "DEBUGINFOD_URLS": ""},
        "started_utc": started,
        "finished_utc": finished,
        "elapsed_seconds": time.monotonic() - monotonic_started,
        "exit_code": 0,
        "artifacts": {
            key: {
                "path": relative_path(path, root),
                "bytes": path.stat().st_size,
                "sha256": sha256_file(path),
            }
            for key, path in paths.items()
        },
        "scope": pmu["scope"],
    })
    write_json(manifest_path, manifest)


def main() -> int:
    args = parse_args()
    root = args.root.resolve()
    protocol_path = (args.protocol or (root / "protocol.json")).resolve()
    build_path = (args.build_identity or (root / "build-identity.json")).resolve()
    capture_path = (args.capture or (root / "capture.json")).resolve()
    manifest_path = (args.output or (root / "profile.json")).resolve()
    try:
        protocol = load_protocol(root, protocol_path)
        build = load_build_identity(build_path)
        capture = load_json(capture_path)
        if not isinstance(capture, dict):
            raise BatchError("capture manifest must be an object")
        _check_capture(capture_path, capture, protocol, build_path)
        selectors = args.selectors or protocol["profile_selectors"]
        for selector in selectors:
            if selector not in protocol["job_by_selector"]:
                raise BatchError(f"profile selector is absent from protocol jobs: {selector}")
        if manifest_path.exists():
            raise BatchError(f"refusing to overwrite existing profile manifest: {manifest_path}")
        manifest: dict[str, Any] = {
            "schema_version": 1,
            "change": 418,
            "status": "starting",
            "protocol": {"path": relative_path(protocol["path"], root), "sha256": protocol["sha256"]},
            "build_identity": {"path": relative_path(build_path, root), "sha256": sha256_file(build_path)},
            "capture": {"path": relative_path(capture_path, root), "sha256": sha256_file(capture_path)},
            "source_roles": {role: build["roles"][role]["source"] for role in ("control", "candidate")},
            "binaries": {role: build["roles"][role]["binaries"] for role in ("control", "candidate")},
            "host": host_identity(include_perf=True),
            "execution": {
                "cpu": protocol["cpu"], "workers": protocol["workers"],
                "selectors": selectors, "profile": protocol["profile"],
                "environment_overrides": {"RUSTUP_TOOLCHAIN": "1.98.1", "DEBUGINFOD_URLS": ""},
            },
            "records": [],
        }
        write_json(manifest_path, manifest)
        if args.mode in ("record", "all"):
            for selector in selectors:
                for role in ("control", "candidate"):
                    _run_record(
                        root, protocol, build, role, selector,
                        {"role": role, "selector": selector}, manifest_path, manifest,
                    )
                    print(f"completed CPU profile {role}/{selector}", flush=True)
        if args.mode in ("stat", "all"):
            pmu = _pmu_config(protocol)
            if pmu["selector"] not in protocol["job_by_selector"]:
                raise BatchError("protocol PMU selector is absent from protocol jobs")
            for role in pmu["order"]:
                _run_stat(root, protocol, build, role, pmu, {}, manifest_path, manifest)
                print(f"completed PMU profile {role}/{pmu['selector']}", flush=True)
        manifest["status"] = "complete"
        manifest["completed_utc"] = utc_now()
        write_json(manifest_path, manifest)
        print(f"0418 profile complete: {manifest_path}")
        return 0
    except (BatchError, OSError, subprocess.SubprocessError) as exc:
        append_failure(root, {"stage": "profile-driver", "utc": utc_now(), "message": str(exc)})
        print(f"0418 profile failed: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
