#!/usr/bin/env python3
"""External, per-fixture retained-memory protocol for signed native routes.

The probe remains uninstrumented.  Each process includes fixture open and
setup, and is profiled independently at 1, 8 and 128 calls.  Heaptrack raw
captures and reports must live under ``--raw-root`` outside this checkout;
only the small, hash-bound summary is written to ``--summary``.
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import signal
import subprocess
import sys
import time
from typing import Any

HERE = Path(__file__).resolve().parent
TIMEOUT_S = 300.0
OUTPUT_LIMIT = 4 * 1024 * 1024
RSS_LIMIT = 20 * 1024 * 1024 * 1024
CALL_COUNTS = (1, 8, 128)

VALIDATE_SPEC = importlib.util.spec_from_file_location(
    "validate_signed_native", HERE / "validate_signed_native.py"
)
assert VALIDATE_SPEC and VALIDATE_SPEC.loader
VALIDATE = importlib.util.module_from_spec(VALIDATE_SPEC)
VALIDATE_SPEC.loader.exec_module(VALIDATE)
PROFILE_SPEC = importlib.util.spec_from_file_location(
    "perf_resource_profile", HERE.parents[3] / "tools/perf_resource_profile.py"
)
assert PROFILE_SPEC and PROFILE_SPEC.loader
PROFILE = importlib.util.module_from_spec(PROFILE_SPEC)
PROFILE_SPEC.loader.exec_module(PROFILE)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def artifact(path: Path) -> dict[str, object]:
    path = path.resolve()
    if not path.is_file():
        raise ValueError(f"artifact is not a file: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def canonical(value: object) -> bytes:
    return (
        json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        + "\n"
    ).encode()


def tool_version(tool: str) -> str:
    result = subprocess.run([tool, "--version"], check=True, text=True, stdout=subprocess.PIPE)
    return result.stdout.strip().splitlines()[0]


def outside_git(path: Path) -> None:
    path = path.resolve()
    try:
        root = Path(
            subprocess.run(
                ["git", "rev-parse", "--show-toplevel"],
                cwd=HERE,
                check=True,
                text=True,
                stdout=subprocess.PIPE,
            ).stdout.strip()
        ).resolve()
    except (OSError, subprocess.CalledProcessError) as error:
        raise ValueError(f"could not resolve repository root: {error}") from error
    if path == root or root in path.parents:
        raise ValueError(f"raw memory artifacts must be outside the checkout: {path}")


def run_profile(
    *,
    binary: Path,
    endpoint: dict[str, Any],
    fixture_name: str,
    route: str,
    calls: int,
    cpu: int,
    raw: Path,
    label: str,
) -> dict[str, Any]:
    work = raw / label
    work.mkdir(parents=True, exist_ok=False)
    stdout_path = work / "probe.stdout"
    stderr_path = work / "probe.stderr"
    prefix = work / "heaptrack-profile"
    command = [
        "/usr/bin/heaptrack",
        "--record-only",
        "-o",
        str(prefix),
        "taskset",
        "-c",
        str(cpu),
        str(binary),
        "--mode",
        "memory",
        "--route",
        route,
        "--calls",
        str(calls),
        "--blocks",
        "1",
    ]
    environment = dict(os.environ)
    environment["LITCHI_SIGNED_FIXTURE_ROOT"] = endpoint["fixture_root"]
    environment["LITCHI_SIGNED_ENDPOINT"] = "after"
    environment["LITCHI_SIGNED_FIXTURE"] = fixture_name
    environment["LITCHI_SIGNED_TARGET_NS"] = str(VALIDATE.TARGET_NS)
    started = time.time_ns()
    stdout = stdout_path.open("wb")
    stderr = stderr_path.open("wb")
    process = subprocess.Popen(
        command,
        stdout=stdout,
        stderr=stderr,
        env=environment,
        start_new_session=True,
    )
    timed_out = False
    exit_code: int | None = None
    max_rss = 0

    def rss_bytes(pid: int) -> int:
        try:
            for line in Path(f"/proc/{pid}/status").read_text().splitlines():
                if line.startswith("VmRSS:"):
                    return int(line.split()[1]) * 1024
        except (FileNotFoundError, PermissionError, ValueError, IndexError):
            return 0
        return 0

    try:
        deadline = time.monotonic() + TIMEOUT_S
        while process.poll() is None:
            max_rss = max(max_rss, rss_bytes(process.pid))
            if stdout_path.stat().st_size > OUTPUT_LIMIT or stderr_path.stat().st_size > OUTPUT_LIMIT:
                os.killpg(process.pid, signal.SIGKILL)
                raise RuntimeError(f"{label}: profiler child exceeded output cap")
            if max_rss > RSS_LIMIT:
                os.killpg(process.pid, signal.SIGKILL)
                raise RuntimeError(f"{label}: profiler child exceeded RSS cap")
            if time.monotonic() >= deadline:
                timed_out = True
                os.killpg(process.pid, signal.SIGKILL)
                break
            time.sleep(0.1)
        exit_code = process.wait(timeout=5)
    finally:
        if process.poll() is None:
            os.killpg(process.pid, signal.SIGKILL)
            process.wait(timeout=5)
        stdout.close()
        stderr.close()
    finished = time.time_ns()
    captures = sorted(path for path in work.glob("heaptrack-profile*") if path.is_file())
    if timed_out or exit_code != 0 or len(captures) != 1:
        raise RuntimeError(f"{label}: heaptrack run failed (exit={exit_code}, captures={len(captures)})")
    capture = captures[0]
    printed = work / "heaptrack-print.txt"
    histogram = work / "heaptrack-histogram.tsv"
    print_stderr = work / "heaptrack-print.stderr"
    with printed.open("wb") as printed_stream, print_stderr.open("wb") as print_stderr_stream:
        subprocess.run(
            ["/usr/bin/heaptrack_print", "-f", str(capture), "-H", str(histogram)],
            check=True,
            stdout=printed_stream,
            stderr=print_stderr_stream,
        )
    parsed = PROFILE.parse_heaptrack_print(printed, retained=False)
    parsed["allocated_bytes"] = PROFILE.parse_heaptrack_histogram(histogram)
    if parsed.get("status") != "ok":
        raise ValueError(f"{label}: heaptrack_print did not produce process totals")
    return {
        "label": label,
        "fixture": fixture_name,
        "route": route,
        "calls": calls,
        "scope": "whole_process_including_fixture_open_and_setup; profiler overhead included",
        "command": command,
        "started_ns": started,
        "finished_ns": finished,
        "rss_max_bytes": max_rss,
        "rss_limit_bytes": RSS_LIMIT,
        "binary": artifact(binary),
        "capture": artifact(capture),
        "print": artifact(printed),
        "histogram": artifact(histogram),
        "stdout": artifact(stdout_path),
        "stderr": artifact(stderr_path),
        "parsed": parsed,
    }


def run(args: argparse.Namespace) -> dict[str, Any]:
    summary = args.summary.resolve()
    raw_root = args.raw_root.resolve()
    outside_git(raw_root)
    outside_git(summary.parent)
    if summary.exists():
        raise ValueError(f"refusing to overwrite summary: {summary}")
    before = VALIDATE.validate_endpoint_manifest(args.before_manifest.resolve(), "before")
    after = VALIDATE.validate_endpoint_manifest(args.after_manifest.resolve(), "after")
    pair = VALIDATE.validate_pair(before, after)
    binary = args.binary.resolve()
    if VALIDATE.artifact(binary)["sha256"] != after["binary"]["sha256"]:
        raise ValueError("memory binary does not match the final endpoint manifest")
    fixture_value = VALIDATE.load_json(Path(after["fixtures"]["path"]), "after fixtures")
    VALIDATE.load_fixture_manifest(fixture_value, "after fixtures")
    routes = tuple(args.routes) if args.routes else tuple(sorted(VALIDATE.FINAL_ONLY))
    if set(routes) - VALIDATE.FINAL_ONLY or len(set(routes)) != len(routes):
        raise ValueError("memory routes must be unique final-only signed routes")
    if args.cpu not in os.sched_getaffinity(0):
        raise ValueError(f"CPU {args.cpu} is unavailable")
    raw_root.mkdir(parents=True, exist_ok=False)
    rows = []
    for fixture in fixture_value["fixtures"]:
        fixture_name = fixture["name"]
        for route in routes:
            for calls in CALL_COUNTS:
                label = f"{fixture_name.replace('/', '_')}-{route.replace('/', '_')}-{calls}"
                rows.append(
                    run_profile(
                        binary=binary,
                        endpoint=after,
                        fixture_name=fixture_name,
                        route=route,
                        calls=calls,
                        cpu=args.cpu,
                        raw=raw_root,
                        label=label,
                    )
                )
    result = {
        "schema": "litchi_signed_native_memory.v1",
        "source": pair,
        "endpoint": {
            "manifest": str(args.after_manifest.resolve()),
            "source_commit": after["source_commit"],
            "binary": after["binary"],
            "fixture_manifest": after["fixtures"],
        },
        "tooling": {
            "heaptrack": tool_version("/usr/bin/heaptrack"),
            "heaptrack_print": tool_version("/usr/bin/heaptrack_print"),
            "parser": artifact(HERE.parents[3] / "tools/perf_resource_profile.py"),
        },
        "cpu": args.cpu,
        "call_counts": list(CALL_COUNTS),
        "scope": "whole_process_including_fixture_open_and_setup; profiler overhead included",
        "raw_root": str(raw_root),
        "rows": rows,
    }
    VALIDATE.validate_memory_summary(
        result,
        after,
        {fixture["name"] for fixture in fixture_value["fixtures"]},
        routes,
    )
    summary.parent.mkdir(parents=True, exist_ok=True)
    summary.write_bytes(canonical(result))
    return {"summary": str(summary), "rows": len(rows), "raw_root": str(raw_root)}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--before-manifest", required=True, type=Path)
    parser.add_argument("--after-manifest", required=True, type=Path)
    parser.add_argument("--binary", required=True, type=Path)
    parser.add_argument("--raw-root", required=True, type=Path)
    parser.add_argument("--summary", required=True, type=Path)
    parser.add_argument("--cpu", required=True, type=int)
    parser.add_argument("--routes", nargs="*")
    args = parser.parse_args()
    print(json.dumps(run(args), indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"signed memory profiling failed: {error}", file=sys.stderr)
        raise
