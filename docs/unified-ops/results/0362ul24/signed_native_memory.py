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
W8_SPEC = importlib.util.spec_from_file_location(
    "w8_run_guard", HERE.parent / "0362h" / "run.py"
)
assert W8_SPEC and W8_SPEC.loader
W8 = importlib.util.module_from_spec(W8_SPEC)
W8_SPEC.loader.exec_module(W8)


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


def process_group_members(pgid: int) -> set[int]:
    members: set[int] = set()
    for entry in Path("/proc").iterdir():
        if not entry.name.isdigit():
            continue
        try:
            stat = (entry / "stat").read_text()
            close = stat.rfind(")")
            fields = stat[close + 2 :].split()
            if int(fields[2]) == pgid:
                members.add(int(entry.name))
        except (FileNotFoundError, PermissionError, ValueError, IndexError):
            continue
    return members


def process_rss(pid: int) -> int:
    try:
        for line in Path(f"/proc/{pid}/status").read_text().splitlines():
            if line.startswith("VmRSS:"):
                return int(line.split()[1]) * 1024
    except (FileNotFoundError, PermissionError, ValueError, IndexError):
        return 0
    return 0


def aggregate_group_rss(pgid: int) -> int:
    return sum(process_rss(pid) for pid in process_group_members(pgid))


def stop_group(process: subprocess.Popen[bytes], pgid: int) -> None:
    if process_group_members(pgid):
        try:
            os.killpg(pgid, signal.SIGKILL)
        except ProcessLookupError:
            pass
    try:
        process.wait(timeout=5)
    except subprocess.TimeoutExpired:
        process.kill()
        process.wait(timeout=5)
    deadline = time.monotonic() + 5
    while time.monotonic() < deadline and process_group_members(pgid):
        time.sleep(0.1)


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
    environment["LITCHI_SIGNED_FIXTURE_MANIFEST"] = endpoint["fixtures"]["path"]
    environment["LITCHI_SIGNED_FIXTURE_MANIFEST_SHA256"] = endpoint["fixtures"]["sha256"]
    before_host = W8.guard(raw, f"{label}.before")
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
    # ``start_new_session`` makes the child PID the process-group ID.  Keep
    # that stable even when the root exits before the first polling turn.
    process_group = process.pid
    during_hosts: list[str] = []
    guard_error: str | None = None

    try:
        deadline = time.monotonic() + TIMEOUT_S
        checks = 0
        while process.poll() is None:
            max_rss = max(max_rss, aggregate_group_rss(process_group))
            if stdout_path.stat().st_size > OUTPUT_LIMIT or stderr_path.stat().st_size > OUTPUT_LIMIT:
                guard_error = f"{label}: profiler child exceeded output cap"
            if time.monotonic() >= deadline:
                timed_out = True
                guard_error = f"{label}: profiler child exceeded {TIMEOUT_S:.1f}s timeout"
            if max_rss > RSS_LIMIT and guard_error is None:
                guard_error = f"{label}: profiler child exceeded RSS cap"
            checks += 1
            during_label = f"{label}.during-{checks:05}"
            try:
                during_hosts.append(W8.guard(raw, during_label, {process.pid}))
            except RuntimeError as error:
                if guard_error is None:
                    guard_error = str(error)
                during_path = raw / f"{during_label}.host.json"
                if during_path.is_file():
                    during_hosts.append(str(during_path.resolve()))
            if guard_error is not None:
                stop_group(process, process_group)
                break
            time.sleep(0.1)
        exit_code = process.wait(timeout=5)
    finally:
        stop_group(process, process_group)
        stdout.close()
        stderr.close()
    finished = time.time_ns()
    if (
        stdout_path.stat().st_size > OUTPUT_LIMIT or stderr_path.stat().st_size > OUTPUT_LIMIT
    ) and guard_error is None:
        guard_error = f"{label}: profiler child exceeded output cap"
    after_host: str | None = None
    after_error: str | None = None
    try:
        after_host = W8.guard(raw, f"{label}.after")
    except RuntimeError as error:
        after_error = str(error)
        if guard_error is None:
            guard_error = after_error
    failed_after = {
        "exit_code": exit_code,
        "guard_error": guard_error,
        "after_guard_error": after_error,
        "timed_out": timed_out,
        "process_group_empty": not process_group_members(process_group),
        "stdout_bytes": stdout_path.stat().st_size,
        "stderr_bytes": stderr_path.stat().st_size,
    }
    captures = sorted(path for path in work.glob("heaptrack-profile*") if path.is_file())
    if guard_error or after_error or timed_out or exit_code != 0 or len(captures) != 1:
        raise RuntimeError(
            f"{label}: heaptrack run failed (exit={exit_code}, captures={len(captures)}, error={guard_error or after_error})"
        )
    capture = captures[0]
    try:
        probe_value = json.loads(stdout_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ValueError(f"{label}: memory probe output is invalid") from error
    if (
        not isinstance(probe_value, dict)
        or probe_value.get("schema") != VALIDATE.SCHEMA
        or probe_value.get("route") != route
        or probe_value.get("mode") != "memory"
        or probe_value.get("calls") != calls
        or probe_value.get("fixture_count") != 1
        or probe_value.get("operations") != calls
        or probe_value.get("scope") != "whole_process_including_fixture_open_and_setup"
    ):
        raise ValueError(f"{label}: memory probe counter/scope proof is invalid")
    fixture_rows = probe_value.get("fixtures")
    if (
        not isinstance(fixture_rows, list)
        or len(fixture_rows) != 1
        or not isinstance(fixture_rows[0], dict)
        or fixture_rows[0].get("name") != fixture_name
    ):
        raise ValueError(f"{label}: memory probe fixture identity is invalid")
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
    host_paths = [before_host, *during_hosts, after_host] if after_host else [before_host, *during_hosts]
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
        "aggregate_rss_max_bytes": max_rss,
        "host_snapshots": [VALIDATE.artifact(Path(path)) for path in host_paths],
        "during_snapshots": len(during_hosts),
        "process_group": process_group,
        "failed_after": failed_after,
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
        "status": "complete",
        "source": pair,
        "endpoint": {
            "manifest": str(args.after_manifest.resolve()),
            "manifest_artifact": VALIDATE.artifact(args.after_manifest.resolve()),
            "source_commit": after["source_commit"],
            "binary": after["binary"],
            "fixture_manifest": after["fixtures"],
            "build_receipt": after["build_receipt"],
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
