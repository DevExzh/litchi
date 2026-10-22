#!/usr/bin/env python3
"""Fail-closed, calibrated ABBA runner for the W8 File probe.

The probe performs leaf/facade ABBA internally. This runner performs the
outer before/after ABBA and keeps the four legs of every block in separate
files. Every process snapshot uses /proc cmdlines (rather than the truncated
``comm`` column) and is taken before, after, and every 100 ms while a child
tree is alive.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import signal
import subprocess
import sys
import time
from typing import Any

MIN_SAMPLE_NS = 150_000_000
TARGET_NS = 500_000_000
HEADROOM = 1.25
MAX_CALLS = 2_000_000
SHA256 = re.compile(r"^[0-9a-f]{64}$")
SERVICE_EXEMPT_EXE = "/usr/local/bin/peculiar_shell"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def artifact(path: Path) -> dict[str, object]:
    path = path.resolve()
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def save(path: Path, value: object) -> None:
    path.write_text(json.dumps(value, indent=2) + "\n")


def command_line(pid: int) -> tuple[str, str, str]:
    """Return comm, full cmdline, and executable identity."""
    try:
        comm = Path(f"/proc/{pid}/comm").read_text().strip()
        raw = Path(f"/proc/{pid}/cmdline").read_bytes()
        executable = os.readlink(f"/proc/{pid}/exe")
    except (FileNotFoundError, PermissionError):
        return "", "", ""
    return comm, raw.replace(b"\0", b" ").decode(errors="replace").strip(), executable


def process_rows() -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for entry in Path("/proc").iterdir():
        if not entry.name.isdigit():
            continue
        pid = int(entry.name)
        try:
            stat = (entry / "stat").read_text()
            close = stat.rfind(")")
            if close < 0:
                continue
            # After the comm field, field 3 is state and field 4 is ppid.
            parent = int(stat[close + 2 :].split()[1])
        except (FileNotFoundError, PermissionError, ValueError, IndexError):
            continue
        comm, args, executable = command_line(pid)
        rows.append({"pid": pid, "ppid": parent, "comm": comm, "args": args, "exe": executable})
    return rows


def descendants(rows: list[dict[str, object]], roots: set[int]) -> set[int]:
    children: dict[int, list[int]] = {}
    for row in rows:
        children.setdefault(int(row["ppid"]), []).append(int(row["pid"]))
    result = set(roots)
    pending = list(roots)
    while pending:
        parent = pending.pop()
        for child in children.get(parent, []):
            if child not in result:
                result.add(child)
                pending.append(child)
    return result


FORBIDDEN_NAMES = re.compile(
    r"^(?:cargo|rustc|rustdoc|rustfmt|clippy-driver|heaptrack|perf|"
    r"valgrind|callgrind|clang\+\+|ld|lld|ld\.lld|script_program)$"
)
FORBIDDEN_COMMANDS = re.compile(
    r"(?:peculiar(?:[_-]|$)|test-adblock|scoped_patch_probe|"
    r"file_w8_(?:probe|qualified_probe)|run-abba|callgrind|heaptrack|"
    r"(?:pytest|nextest|criterion|hyperfine|strace|dtrace|gdb|lldb|massif)|"
    r"(?:cargo|rustc|rustdoc|rustfmt|clippy)\s+(?:test|bench|run|build|doc|fmt)|"
    r"(?:^|[ /])(?:cargo|rustc|rustdoc|rustfmt|clippy-driver|perf|"
    r"heaptrack|valgrind|callgrind)(?:\s|$))",
    re.IGNORECASE,
)


def forbidden(row: dict[str, object]) -> bool:
    executable = str(row.get("exe", ""))
    args = str(row.get("args", ""))
    comm = str(row.get("comm", ""))
    # This exact installed service is long-lived and unrelated to builds. A
    # build-tree copy or wrapper remains a blocker.
    if executable == SERVICE_EXEMPT_EXE and Path(executable).name == "peculiar_shell":
        return False
    return (
        FORBIDDEN_NAMES.fullmatch(comm) is not None
        or FORBIDDEN_NAMES.fullmatch(Path(executable).name) is not None
        or FORBIDDEN_COMMANDS.search(args) is not None
    )


def guard(out: Path, label: str, allowed_roots: set[int] | None = None) -> str:
    rows = process_rows()
    allowed = descendants(rows, {os.getpid()})
    allowed |= descendants(rows, allowed_roots or set())
    blockers = [
        row
        for row in rows
        if int(row["pid"]) not in allowed
        and (
            forbidden(row)
        )
    ]
    save(
        out / f"{label}.host.json",
        {"time_ns": time.time_ns(), "allowed": sorted(allowed), "blockers": blockers, "processes": rows},
    )
    if blockers:
        raise RuntimeError(f"host overlap at {label}: {blockers}")
    return str((out / f"{label}.host.json").resolve())


def run_probe(
    out: Path,
    label: str,
    binary: Path,
    route: str,
    mode: str,
    calls: int | None,
    blocks: int,
    cpu: int,
) -> dict:
    stdout = out / f"{label}.stdout"
    stderr = out / f"{label}.stderr"
    command = [
        "taskset", "-c", str(cpu), str(binary), "--mode", mode,
        "--route", route, "--blocks", str(blocks),
    ]
    env = os.environ.copy()
    env["LITCHI_FILE_TARGET_NS"] = str(TARGET_NS)
    env["LITCHI_FILE_HEADROOM"] = str(HEADROOM)
    if calls is not None:
        env["LITCHI_FILE_CALLS"] = str(calls)
    before_host = guard(out, f"{label}.before")
    started = time.time_ns()
    with stdout.open("w") as output, stderr.open("w") as errors:
        child = subprocess.Popen(
            command,
            stdout=output,
            stderr=errors,
            env=env,
            start_new_session=True,
        )
        overlap: str | None = None
        during_hosts: list[str] = []
        checks = 0
        while True:
            try:
                status = child.wait(timeout=0.1)
                break
            except subprocess.TimeoutExpired:
                checks += 1
                if overlap is None:
                    try:
                        during_hosts.append(guard(out, f"{label}.during-{checks:05}", {child.pid}))
                    except RuntimeError as error:
                        overlap = str(error)
                        try:
                            os.killpg(child.pid, signal.SIGTERM)
                        except ProcessLookupError:
                            pass
        finished = time.time_ns()
    after_host: str | None = None
    after_error: str | None = None
    try:
        after_host = guard(out, f"{label}.after")
    except RuntimeError as error:
        after_error = str(error)
    host_paths = [before_host, *during_hosts, after_host] if after_host else [before_host, *during_hosts]
    host_artifacts = [artifact(Path(path)) for path in host_paths]
    receipt = {
        "schema": "litchi_file_probe_receipt.v2",
        "label": label,
        "binary": artifact(binary),
        "route": route,
        "mode": mode,
        "blocks": blocks,
        "calls": calls,
        "command": command,
        "environment": {
            "LITCHI_FILE_CALLS": env.get("LITCHI_FILE_CALLS"),
            "target_ns": TARGET_NS,
            "headroom": HEADROOM,
        },
        "started_ns": started,
        "finished_ns": finished,
        "exit_code": status,
        "overlap": overlap,
        "guard_error": after_error,
        "host_snapshots": host_artifacts,
        "stdout": artifact(stdout),
        "stderr": artifact(stderr),
    }
    save(out / f"{label}.receipt.json", receipt)
    if after_error:
        raise RuntimeError(after_error)
    if overlap:
        raise RuntimeError(overlap)
    if status != 0:
        raise RuntimeError(f"probe {label} exited {status}; see {stderr}")
    try:
        value = json.loads(stdout.read_text())
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise RuntimeError(f"invalid probe output {label}: {error}") from error
    if not isinstance(value, dict):
        raise RuntimeError(f"probe {label} did not return an object")
    value["_receipt"] = str((out / f"{label}.receipt.json").resolve())
    value["_host_snapshots"] = host_artifacts
    return value


def routes(binary: Path, out: Path, cpu: int, suffix: str) -> tuple[list[str], dict]:
    row = run_probe(out, f"route-list-{suffix}", binary, "*", "list", None, 1, cpu)
    values = row.get("routes")
    if not isinstance(values, list) or not values or not all(isinstance(item, str) for item in values):
        raise RuntimeError("probe route list is empty or malformed")
    return values, row


def endpoint_manifest(path: Path, binary: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text())
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise RuntimeError(f"{label} endpoint manifest is invalid: {error}") from error
    if not isinstance(value, dict) or value.get("schema") != "litchi_file_endpoint.v1":
        raise RuntimeError(f"{label} endpoint manifest has an invalid schema")

    def checked_artifact(name: str) -> dict[str, Any]:
        row = value.get(name)
        if (
            not isinstance(row, dict)
            or not isinstance(row.get("path"), str)
            or not isinstance(row.get("sha256"), str)
            or not SHA256.fullmatch(row["sha256"])
        ):
            raise RuntimeError(f"{label} endpoint manifest lacks a valid {name} artifact")
        actual = artifact(Path(row["path"]))
        if actual["sha256"] != row["sha256"]:
            raise RuntimeError(f"{label} {name} artifact hash does not match its file")
        return actual

    binary_row = checked_artifact("binary")
    actual_binary = artifact(binary)
    if binary_row["sha256"] != actual_binary["sha256"]:
        raise RuntimeError(f"{label} endpoint binary hash does not match the selected executable")
    source_commit = value.get("source_commit")
    dependency = value.get("dependency_projection_sha256")
    if not isinstance(source_commit, str) or not source_commit:
        raise RuntimeError(f"{label} endpoint manifest lacks source_commit")
    if not isinstance(dependency, str) or not SHA256.fullmatch(dependency):
        raise RuntimeError(f"{label} endpoint manifest lacks dependency_projection_sha256")
    return {
        "manifest": artifact(path),
        "source_commit": source_commit,
        "dependency_projection_sha256": dependency,
        "binary": binary_row,
        "source": checked_artifact("source"),
        "lock": checked_artifact("lock"),
        "probe": checked_artifact("probe"),
        "fixture": checked_artifact("fixture"),
    }


def shared_identity(before: dict[str, Any], after: dict[str, Any]) -> dict[str, str]:
    fields = ("dependency_projection_sha256", "lock", "probe", "fixture")
    for field in fields:
        if field == "dependency_projection_sha256":
            left, right = before[field], after[field]
        else:
            left, right = before[field]["sha256"], after[field]["sha256"]
        if left != right:
            raise RuntimeError(f"endpoint {field} identities differ: {left} != {right}")
    return {
        "dependency_projection_sha256": before["dependency_projection_sha256"],
        "lock_sha256": before["lock"]["sha256"],
        "probe_sha256": before["probe"]["sha256"],
        "fixture_sha256": before["fixture"]["sha256"],
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("before", type=Path)
    parser.add_argument("after", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("rounds", type=int, nargs="?", default=5)
    parser.add_argument("cpu", type=int, nargs="?", default=29)
    parser.add_argument("--before-manifest", required=True, type=Path)
    parser.add_argument("--after-manifest", required=True, type=Path)
    args = parser.parse_args()
    if args.rounds < 5:
        parser.error("at least five ABBA blocks are required")
    if args.cpu not in os.sched_getaffinity(0):
        parser.error(f"CPU {args.cpu} is unavailable")
    before, after, out = (p.resolve() for p in (args.before, args.after, args.output))
    if before == after:
        parser.error("before and after must be distinct protected binaries")
    if not before.is_file() or not os.access(before, os.X_OK) or not after.is_file() or not os.access(after, os.X_OK):
        parser.error("before and after must be executable protected binaries")
    if out.exists() and any(out.iterdir()):
        parser.error(f"refusing to overwrite nonempty output directory {out}")
    out.mkdir(parents=True, exist_ok=True)
    before_endpoint = endpoint_manifest(args.before_manifest.resolve(), before, "before")
    after_endpoint = endpoint_manifest(args.after_manifest.resolve(), after, "after")
    shared = shared_identity(before_endpoint, after_endpoint)
    guard(out, "initial")
    before_routes, before_route_receipt = routes(before, out, args.cpu, "before")
    after_routes, after_route_receipt = routes(after, out, args.cpu, "after")
    if before_routes != after_routes:
        raise RuntimeError("before/after route lists differ")
    manifest = {
        "schema": "litchi_file_qualification.v3",
        "before": artifact(before),
        "after": artifact(after),
        "routes": before_routes,
        "route_lists": {"before": before_route_receipt, "after": after_route_receipt},
        "rounds": args.rounds,
        "cpu": args.cpu,
        "target_ns": TARGET_NS,
        "minimum_sample_ns": MIN_SAMPLE_NS,
        "headroom": HEADROOM,
        "guard": "full /proc comm+cmdline+exe snapshots before/after and every 100ms; only this runner and the child subtree are exempt",
        "provenance": {
            "before": before_endpoint,
            "after": after_endpoint,
            "shared": shared,
        },
        "status": "running",
        "calibration": [],
        "abba": [],
    }
    save(out / "manifest.json", manifest)
    for index, route in enumerate(before_routes):
        safe = re.sub(r"[^A-Za-z0-9_.-]", "_", route)
        calibrations = []
        for side, binary in (("before", before), ("after", after)):
            calibration = run_probe(out, f"{safe}-{side}-calibration", binary, route, "calibrate", None, 1, args.cpu)
            if calibration.get("route") != route or not isinstance(calibration.get("common_calls"), int):
                raise RuntimeError(f"malformed calibration for {route}/{side}")
            common = int(calibration["common_calls"])
            if common <= 0 or common > MAX_CALLS or calibration.get("minimum_elapsed_ns", 0) < MIN_SAMPLE_NS:
                raise RuntimeError(f"calibration cannot meet target for {route}/{side}: {calibration}")
            calibrations.append(calibration)
        common_calls = max(int(item["common_calls"]) for item in calibrations)
        if common_calls > MAX_CALLS:
            raise RuntimeError(f"shared calibration exceeds ceiling for {route}: {common_calls}")
        manifest["calibration"].append({"route": route, "rows": calibrations, "common_calls": common_calls})
        save(out / "manifest.json", manifest)
        blocks = []
        route_row = {"route": route, "calls": common_calls, "blocks": blocks}
        manifest["abba"].append(route_row)
        save(out / "manifest.json", manifest)
        for block_index in range(args.rounds):
            legs = []
            for leg_index, side in enumerate(("before", "after", "after", "before")):
                binary = before if side == "before" else after
                value = run_probe(
                    out,
                    f"{safe}-{index:02}-block-{block_index:02}-leg-{leg_index}-{side}",
                    binary,
                    route,
                    "measure",
                    common_calls,
                    1,
                    args.cpu,
                )
                if value.get("route") != route or value.get("calls") != common_calls:
                    raise RuntimeError(f"shared call count mismatch for {route}/{side}")
                legs.append({"side": side, "value": value})
            blocks.append({"block": block_index, "legs": legs})
            save(out / "manifest.json", manifest)
        save(out / "manifest.json", manifest)
    guard(out, "finished")
    manifest["status"] = "complete; analyze.py review pending"
    save(out / "manifest.json", manifest)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"0362h qualification failed: {error}", file=sys.stderr)
        raise
