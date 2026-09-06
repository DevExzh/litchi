#!/usr/bin/env python3
"""Capture one whole-process 0433 ODS CPU profile.

This diagnostic lane is deliberately separate from the formal latency capture.
It profiles only the normal release binary, the frozen ``large`` (32768-row)
shape, one worker, three warmups, and thirty samples.  The process scope is
the complete harness process: setup/corpus generation, warmups, measured
samples, output hashing, and the harness oracle/report.  ``perf`` report
post-processing and the external report verifier are outside that scope.

Run these six commands serially, after the before/after normal binaries have
been built and their build.json files are present::

    python3 docs/performance/results/change-0433/profile/capture.py --role before-buffered --kind stat
    python3 docs/performance/results/change-0433/profile/capture.py --role before-buffered --kind record
    python3 docs/performance/results/change-0433/profile/capture.py --role after-buffered --kind stat
    python3 docs/performance/results/change-0433/profile/capture.py --role after-buffered --kind record
    python3 docs/performance/results/change-0433/profile/capture.py --role after-streaming --kind stat
    python3 docs/performance/results/change-0433/profile/capture.py --role after-streaming --kind record

The raw ``perf.data`` and symbolized ``perf report`` text are retained for
record jobs before the temporary binaries are removed.  The script does not
run Cargo or modify the source tree.
"""

from __future__ import annotations

import argparse
import datetime
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import platform
import subprocess
import sys
from typing import Any


PROFILE_ROOT = Path(__file__).resolve().parent
BUNDLE_ROOT = PROFILE_ROOT.parent
REPO_ROOT = BUNDLE_ROOT.parents[3]
CHANGE = 433
CPU = 2
WORKERS = 1
SHAPE = "large"
NORMAL = "normal"
KINDS = ("stat", "record")
ROLES = {
    "before-buffered": ("before", "ods_buffered_create"),
    "after-buffered": ("after", "ods_buffered_create"),
    "after-streaming": ("after", "ods_streaming_create"),
}
STAT_EVENTS = (
    "cycles",
    "instructions",
    "branches",
    "branch-misses",
    "L1-dcache-load-misses",
)
RECORD_EVENT = "cycles:u"
RECORD_FREQUENCY = 999
CALL_GRAPH = "fp,127"
ENVIRONMENT_OVERRIDES = {"DEBUGINFOD_URLS": "", "RUSTUP_TOOLCHAIN": "1.98.1"}
PROFILE_SCOPE = (
    "whole process including setup/corpus generation, warmups, measured samples, "
    "output hashing, and the harness oracle/report; excludes perf postprocessing "
    "and the external report verifier"
)


def now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def sha_bytes(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def sha_path(path: Path) -> str:
    return sha_bytes(path.read_bytes())


def duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key: {key}")
        result[key] = value
    return result


def load(path: Path) -> Any:
    return json.loads(
        path.read_text(encoding="utf-8"),
        object_pairs_hook=duplicate_keys,
        parse_constant=lambda value: (_ for _ in ()).throw(
            ValueError(f"non-finite JSON constant: {value}")
        ),
    )


def write_json(path: Path, value: Any) -> None:
    path.write_text(
        json.dumps(value, indent=2, sort_keys=True, allow_nan=False) + "\n",
        encoding="utf-8",
    )


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ValueError(f"{label} must be an object")
    return value


def text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        raise ValueError(f"{label} must be nonempty text")
    return value


def uint(value: Any, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise ValueError(f"{label} must be an unsigned integer")
    return value


def bundle_path(name: str) -> Path:
    path = Path(name)
    if path.is_absolute():
        raise ValueError(f"absolute bundle path: {name}")
    resolved = (BUNDLE_ROOT / path).resolve()
    if not resolved.is_relative_to(BUNDLE_ROOT.resolve()):
        raise ValueError(f"bundle path escapes root: {name}")
    return resolved


def verify_source_manifest(manifest: dict[str, Any]) -> None:
    path_name = text(manifest.get("path"), "source_manifest.path")
    source_path = bundle_path(path_name)
    raw = source_path.read_bytes()
    assert sha_bytes(raw) == text(manifest.get("sha256"), "source_manifest.sha256")
    entries = obj(json.loads(raw.decode("utf-8"), object_pairs_hook=duplicate_keys), "source manifest payload")
    assert len(entries) == uint(manifest.get("files"), "source_manifest.files")


def verify_protocol() -> dict[str, Any]:
    protocol = obj(load(BUNDLE_ROOT / "protocol.json"), "protocol")
    assert protocol.get("change") == CHANGE
    assert protocol.get("cpu") == CPU
    assert protocol.get("workers") == WORKERS
    assert protocol.get("samples") == 30
    assert protocol.get("warmups") == 3
    assert obj(protocol.get("shapes"), "protocol.shapes").get(SHAPE) == 32_768
    scope = text(protocol.get("profile_scope"), "protocol.profile_scope")
    for marker in ("setup", "oracle", "samples", "hash"):
        assert marker in scope.lower(), f"protocol.profile_scope omits {marker}"
    return protocol


def role_spec(protocol: dict[str, Any], role: str) -> tuple[str, str, dict[str, Any]]:
    if role not in ROLES:
        raise ValueError(f"unknown role: {role}")
    build_dir, expected_selector = ROLES[role]
    roles = obj(protocol.get("roles"), "protocol.roles")
    spec = obj(roles.get(role), f"protocol.roles.{role}")
    selector = text(spec.get("selector"), f"protocol.roles.{role}.selector")
    assert selector == expected_selector
    return build_dir, selector, spec


def load_build(protocol: dict[str, Any], role: str) -> tuple[dict[str, Any], dict[str, Any]]:
    build_dir, _selector, _spec = role_spec(protocol, role)
    build = obj(load(BUNDLE_ROOT / build_dir / "build.json"), f"{build_dir}/build.json")
    assert build.get("change") == CHANGE
    source_manifest = obj(build.get("source_manifest"), f"{build_dir}.source_manifest")
    verify_source_manifest(source_manifest)
    protocol_hash = text(build.get("protocol_sha256"), f"{build_dir}.protocol_sha256")
    assert protocol_hash == sha_path(BUNDLE_ROOT / "protocol.json")
    verifier_hash = text(build.get("verifier_sha256"), f"{build_dir}.verifier_sha256")
    assert verifier_hash == sha_path(BUNDLE_ROOT / "verify-report.py")
    binaries = obj(build.get("binaries"), f"{build_dir}.binaries")
    binary = obj(binaries.get(NORMAL), f"{build_dir}.binaries.normal")
    assert set(binary) == {"path", "sha256", "bytes"}
    expected_path = Path("/tmp/litchi-goal-0433-binaries") / build_dir / NORMAL
    assert Path(text(binary["path"], "normal binary path")) == expected_path
    binary_path = Path(binary["path"])
    assert binary_path.is_file()
    assert sha_path(binary_path) == text(binary["sha256"], "normal binary sha256")
    assert binary_path.stat().st_size == uint(binary["bytes"], "normal binary bytes")
    return build, binary


def source_custody(expected: dict[str, Any]) -> list[str]:
    """Recompute the Rust source manifest and ensure no tracked file is dirty."""
    module_spec = importlib.util.spec_from_file_location("change0433_custody", BUNDLE_ROOT / "check.py")
    if module_spec is None or module_spec.loader is None:
        raise RuntimeError("cannot load the bound source-custody driver")
    module = importlib.util.module_from_spec(module_spec)
    module_spec.loader.exec_module(module)
    if module.sources() != expected:
        raise ValueError("current Rust/TOML/lock source manifest differs from the build")
    status = subprocess.check_output(
        ["git", "status", "--porcelain"], cwd=REPO_ROOT, text=True
    ).splitlines()
    tracked = [line for line in status if not line.startswith("?? ")]
    if tracked:
        raise ValueError(f"tracked worktree changes are present: {tracked}")
    return status


def tool_version(command: list[str]) -> str:
    result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, check=True)
    line = result.stdout.strip().splitlines()
    if not line:
        raise ValueError(f"tool returned no version: {command[0]}")
    return line[0]


def artifact_record(path: Path) -> dict[str, Any]:
    return {"sha256": sha_path(path), "bytes": path.stat().st_size}


def relative(path: Path) -> str:
    return str(path.relative_to(BUNDLE_ROOT))


def expected_artifacts(run_dir: Path, kind: str) -> list[Path]:
    common = [
        run_dir / "workload.json",
        run_dir / "workload-catalog.json",
        run_dir / "workload-verify.json",
        run_dir / "workload-verify.stderr.log",
        run_dir / "process.stdout.log",
        run_dir / "process.stderr.log",
        run_dir / "resource.log",
    ]
    if kind == "stat":
        return common + [run_dir / "perf-stat.csv"]
    return common + [run_dir / "perf.data", run_dir / "perf-report.txt", run_dir / "perf-report.stderr.log"]


def append_index(receipt_path: Path) -> None:
    index_path = BUNDLE_ROOT / "profile-index.json"
    index = load(index_path) if index_path.exists() else []
    if not isinstance(index, list):
        raise ValueError("profile-index.json must be an array")
    name = relative(receipt_path)
    if name in index:
        raise ValueError(f"profile receipt already indexed: {name}")
    index.append(name)
    write_json(index_path, index)


def run_profile(role: str, kind: str) -> int:
    if kind not in KINDS:
        raise ValueError(f"unknown profile kind: {kind}")
    protocol = verify_protocol()
    build_dir, selector, _spec = role_spec(protocol, role)
    build, binary = load_build(protocol, role)
    status_before = source_custody(build["source_manifest"])
    samples = uint(protocol["samples"], "protocol.samples")
    warmups = uint(protocol["warmups"], "protocol.warmups")
    # Keep raw profiles below the provider's sealed before/after trees.  seal.py
    # losslessly compresses .data and .log files only in these two locations.
    run_dir = BUNDLE_ROOT / build_dir / "profiles" / role / kind
    if run_dir.exists():
        raise ValueError(f"refusing to overwrite profile directory: {run_dir}")
    run_dir.mkdir(parents=True)

    workload_report = run_dir / "workload.json"
    corpus_catalog = run_dir / "workload-catalog.json"
    resource = run_dir / "resource.log"
    process_stdout = run_dir / "process.stdout.log"
    process_stderr = run_dir / "process.stderr.log"
    verify_stdout = run_dir / "workload-verify.json"
    verify_stderr = run_dir / "workload-verify.stderr.log"
    stat_path = run_dir / "perf-stat.csv"
    data_path = run_dir / "perf.data"
    symbolized_path = run_dir / "perf-report.txt"
    symbolized_stderr = run_dir / "perf-report.stderr.log"
    binary_path = Path(binary["path"])
    workload_argv = [
        str(binary_path),
        "--case", selector,
        "--semantic-shape", SHAPE,
        "--workers", str(WORKERS),
        "--samples", str(samples),
        "--warmup", str(warmups),
        "--json", str(workload_report),
        "--corpus-manifest", str(corpus_catalog),
    ]
    if kind == "stat":
        profiler_argv = [
            "perf", "stat", "--no-big-num", "-x,", "-e", ",".join(STAT_EVENTS),
            "-o", str(stat_path), "--", *workload_argv,
        ]
    else:
        profiler_argv = [
            "perf", "record", "--no-buildid-cache", "-e", RECORD_EVENT,
            "-F", str(RECORD_FREQUENCY), "--call-graph", CALL_GRAPH,
            "-o", str(data_path), "--", *workload_argv,
        ]
    profile_argv = [
        "taskset", "-c", str(CPU), "/usr/bin/time", "-v", "-o", str(resource),
        *profiler_argv,
    ]
    report_argv = None
    if kind == "record":
        report_argv = [
            "perf", "report", "--stdio", "--no-children", "--percent-limit", "0",
            "-i", str(data_path),
        ]

    env = os.environ.copy()
    env.update(ENVIRONMENT_OVERRIDES)
    receipt_path = run_dir / "receipt.json"
    receipt: dict[str, Any] = {
        "schema": "ods_process_profile_v1",
        "change": CHANGE,
        "role": role,
        "build_directory": build_dir,
        "kind": kind,
        "selector": selector,
        "mode": NORMAL,
        "shape": SHAPE,
        "workers": WORKERS,
        "samples": samples,
        "warmups": warmups,
        "revision": text(build.get("revision"), f"{build_dir}.revision"),
        "source_manifest": build["source_manifest"],
        "binary": binary,
        "protocol_sha256": sha_path(BUNDLE_ROOT / "protocol.json"),
        "profile_driver_sha256": sha_path(Path(__file__)),
        "verifier_sha256": sha_path(BUNDLE_ROOT / "verify-report.py"),
        "scope": PROFILE_SCOPE,
        "protocol_profile_scope": text(protocol["profile_scope"], "protocol.profile_scope"),
        "stat_events": list(STAT_EVENTS) if kind == "stat" else [],
        "record_event": RECORD_EVENT if kind == "record" else None,
        "record_frequency_hz": RECORD_FREQUENCY if kind == "record" else None,
        "call_graph": CALL_GRAPH if kind == "record" else None,
        "environment_overrides": ENVIRONMENT_OVERRIDES,
        "source_status_before": status_before,
        "tool_versions": {
            "perf": tool_version(["perf", "--version"]),
            "time": tool_version(["/usr/bin/time", "--version"]),
            "python": platform.python_version(),
        },
        "workload_argv": workload_argv,
        "profile_argv": profile_argv,
        "report_argv": report_argv,
        "started_utc": now(),
        "status": "running",
    }
    write_json(receipt_path, receipt)
    process = None
    failure: str | None = None
    try:
        with process_stdout.open("xb") as stdout, process_stderr.open("xb") as stderr:
            process = subprocess.run(
                profile_argv,
                cwd=REPO_ROOT,
                env=env,
                stdout=stdout,
                stderr=stderr,
            )
        receipt["profile_exit_code"] = process.returncode
        if process.returncode != 0:
            raise RuntimeError(f"profile command exited {process.returncode}")
        verify_argv = [
            sys.executable, "-B", str(BUNDLE_ROOT / "verify-report.py"),
            "--report", str(workload_report), "--mode", NORMAL,
            "--shape", SHAPE, "--role", role,
        ]
        receipt["verifier_argv"] = verify_argv
        with verify_stdout.open("xb") as stdout, verify_stderr.open("xb") as stderr:
            verifier = subprocess.run(
                verify_argv,
                cwd=REPO_ROOT,
                env=env,
                stdout=stdout,
                stderr=stderr,
            )
        receipt["verifier_exit_code"] = verifier.returncode
        if verifier.returncode != 0:
            raise RuntimeError(f"workload report verifier exited {verifier.returncode}")
        if kind == "record":
            assert data_path.is_file() and data_path.stat().st_size > 0
            assert report_argv is not None
            with symbolized_path.open("xb") as stdout, symbolized_stderr.open("xb") as stderr:
                report_process = subprocess.run(
                    report_argv,
                    cwd=REPO_ROOT,
                    env=env,
                    stdout=stdout,
                    stderr=stderr,
                )
            receipt["report_exit_code"] = report_process.returncode
            if report_process.returncode != 0:
                raise RuntimeError(f"perf report exited {report_process.returncode}")
            if not symbolized_path.read_text(encoding="utf-8", errors="replace").strip():
                raise RuntimeError("perf report produced an empty symbolized report")
        else:
            assert stat_path.is_file() and stat_path.stat().st_size > 0
        receipt["status"] = "pass"
    except Exception as error:
        failure = str(error)
        receipt["status"] = "failed"
        receipt["error"] = failure
    finally:
        try:
            status_after = source_custody(build["source_manifest"])
            receipt["source_status_after"] = status_after
            receipt["tracked_tree_clean_before_and_after"] = True
        except Exception as error:
            receipt["source_status_after"] = []
            receipt["tracked_tree_clean_before_and_after"] = False
            receipt["status"] = "failed"
            receipt["error"] = f"{receipt.get('error', '')}; source custody: {error}".lstrip("; ")
            failure = receipt["error"]
        receipt["finished_utc"] = now()
        receipt["artifacts"] = {
            relative(path): artifact_record(path)
            for path in expected_artifacts(run_dir, kind)
            if path.is_file()
        }
        write_json(receipt_path, receipt)
        append_index(receipt_path)
    if failure is not None:
        raise RuntimeError(failure)
    print(json.dumps({"status": "pass", "role": role, "kind": kind, "receipt": relative(receipt_path)}))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--role", choices=tuple(ROLES), required=True)
    parser.add_argument("--kind", choices=KINDS, required=True)
    args = parser.parse_args()
    return run_profile(args.role, args.kind)


if __name__ == "__main__":
    raise SystemExit(main())
