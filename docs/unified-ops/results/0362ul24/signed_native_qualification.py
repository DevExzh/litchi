#!/usr/bin/env python3
"""Guarded signed-native preflight and ABBA runner.

The Rust probe is the only code that calls the format APIs.  This process
guard records full /proc snapshots, pins the child to one CPU, validates every
correctness receipt before timing, and keeps final-only capability evidence
separate from a before/after ratio.
"""
from __future__ import annotations

import argparse
import importlib.util
import json
import os
from pathlib import Path
import re
import signal
import statistics
import subprocess
import sys
import time
from typing import Any

HERE = Path(__file__).resolve().parent
VALIDATOR_SPEC = importlib.util.spec_from_file_location(
    "validate_signed_native", HERE / "validate_signed_native.py"
)
assert VALIDATOR_SPEC and VALIDATOR_SPEC.loader
VALIDATE = importlib.util.module_from_spec(VALIDATOR_SPEC)
VALIDATOR_SPEC.loader.exec_module(VALIDATE)

# Reuse the reviewed W8 /proc implementation.  It is imported, never edited,
# so the signed runner inherits its full cmdline/executable overlap guard.
W8_SPEC = importlib.util.spec_from_file_location(
    "w8_run_guard", HERE.parent / "0362h" / "run.py"
)
assert W8_SPEC and W8_SPEC.loader
W8 = importlib.util.module_from_spec(W8_SPEC)
W8_SPEC.loader.exec_module(W8)

SAFE_ROUTE = re.compile(r"^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$")
CHILD_TIMEOUT_S = VALIDATE.CHILD_TIMEOUT_S
CHILD_OUTPUT_LIMIT = VALIDATE.CHILD_OUTPUT_LIMIT
CHILD_RSS_LIMIT = VALIDATE.CHILD_RSS_LIMIT


def save(path: Path, value: object) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def load(path: Path, label: str) -> dict[str, Any]:
    return VALIDATE.load_json(path, label)


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


def aggregate_tree_rss(process_tree: VALIDATE.OwnedProcessTree) -> int:
    return sum(process_rss(pid) for pid in process_tree.pids())


def wait_for_group_exit(pgid: int, timeout: float = 5.0) -> None:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline and process_group_members(pgid):
        time.sleep(0.1)


def child(
    output: Path,
    label: str,
    binary: Path,
    endpoint: dict[str, Any],
    route: str,
    mode: str,
    cpu: int,
    calls: int | None = None,
) -> dict[str, Any]:
    if route != "*" and not SAFE_ROUTE.fullmatch(route):
        raise ValueError(f"invalid route name: {route!r}")
    if route == "*" and mode != "list":
        raise ValueError("wildcard route is reserved for list mode")
    stdout_path = output / f"{label}.stdout.json"
    stderr_path = output / f"{label}.stderr"
    command = [
        "taskset",
        "-c",
        str(cpu),
        str(binary),
        "--mode",
        mode,
        "--route",
        route,
        "--blocks",
        "1",
    ]
    if calls is not None:
        command.extend(("--calls", str(calls)))
    environment = dict(os.environ)
    environment["LITCHI_SIGNED_FIXTURE_ROOT"] = endpoint["fixture_root"]
    environment["LITCHI_SIGNED_FIXTURE_MANIFEST"] = endpoint["fixtures"]["path"]
    environment["LITCHI_SIGNED_FIXTURE_MANIFEST_SHA256"] = endpoint["fixtures"]["sha256"]
    environment["LITCHI_SIGNED_ENDPOINT"] = (
        "before" if label.startswith("before") else "after"
    )
    environment["LITCHI_SIGNED_TARGET_NS"] = str(VALIDATE.TARGET_NS)
    environment["LITCHI_SIGNED_MAX_CALLS"] = str(VALIDATE.MAX_CALLS)

    before_host = W8.guard(output, f"{label}.before")
    started_ns = time.time_ns()
    during_hosts: list[str] = []
    guard_error: str | None = None
    failed_after: dict[str, Any] | None = None
    timed_out = False
    rss_max = 0
    process_group: int | None = None
    process: subprocess.Popen[bytes] | None = None
    process_tree: VALIDATE.OwnedProcessTree | None = None
    exit_code: int | None = None
    pending_error: BaseException | None = None

    def stop_group() -> None:
        if process is None or process_group is None:
            return
        try:
            os.killpg(process_group, signal.SIGKILL)
        except ProcessLookupError:
            pass
        try:
            process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            process.kill()
            process.wait(timeout=5)
        if process_tree is not None:
            process_tree.kill()
            process_tree.reap()
            process_tree.release()
        wait_for_group_exit(process_group)

    try:
        process_tree = VALIDATE.OwnedProcessTree()
        with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
            process = subprocess.Popen(
                command,
                stdout=stdout,
                stderr=stderr,
                env=environment,
                start_new_session=True,
            )
            # ``start_new_session`` makes the child PID the process-group ID;
            # using it directly also remains safe if the root exits before
            # the first polling turn.
            process_group = process.pid
            process_tree.attach(process.pid)
            deadline = time.monotonic() + CHILD_TIMEOUT_S
            checks = 0
            while process.poll() is None:
                if time.monotonic() >= deadline:
                    timed_out = True
                    guard_error = f"child exceeded {CHILD_TIMEOUT_S:.1f}s timeout"
                rss_max = max(rss_max, aggregate_tree_rss(process_tree))
                if rss_max > CHILD_RSS_LIMIT and guard_error is None:
                    guard_error = f"child exceeded {CHILD_RSS_LIMIT} byte RSS cap"
                if (
                    (stdout_path.exists() and stdout_path.stat().st_size > CHILD_OUTPUT_LIMIT)
                    or (stderr_path.exists() and stderr_path.stat().st_size > CHILD_OUTPUT_LIMIT)
                ) and guard_error is None:
                    guard_error = f"child exceeded {CHILD_OUTPUT_LIMIT} byte output cap"
                checks += 1
                during_label = f"{label}.during-{checks:05}"
                try:
                    during_hosts.append(W8.guard(output, during_label, {process.pid}))
                except RuntimeError as error:
                    if guard_error is None:
                        guard_error = str(error)
                    during_path = output / f"{during_label}.host.json"
                    if during_path.is_file() and not during_path.is_symlink():
                        during_hosts.append(str(during_path.resolve()))
                if guard_error is not None:
                    stop_group()
                    break
                time.sleep(0.1)
            if process is not None:
                exit_code = process.wait(timeout=5)
    except BaseException as error:
        if guard_error is None:
            guard_error = str(error)
        pending_error = error
        stop_group()
        if process is not None:
            exit_code = process.wait(timeout=5)
    finally:
        if process_tree is not None and process is not None and process.poll() is not None:
            try:
                if process_tree.descendant_pids() and guard_error is None:
                    guard_error = "child process tree outlived its root process"
            except RuntimeError as error:
                if guard_error is None:
                    guard_error = str(error)
        if process_group is not None:
            stop_group()
        elif process_tree is not None:
            process_tree.release()
        finished_ns = time.time_ns()
        failed_after = {
            "time_ns": finished_ns,
            "exit_code": process.returncode if process is not None else None,
            "guard_error": guard_error,
            "timed_out": timed_out,
            "rss_max_bytes": rss_max,
            "aggregate_rss_max_bytes": rss_max,
            "rss_limit_bytes": CHILD_RSS_LIMIT,
            "output_limit_bytes": CHILD_OUTPUT_LIMIT,
            "process_group_empty": process_tree is None or process_tree.empty(),
            "stdout_bytes": stdout_path.stat().st_size if stdout_path.exists() else 0,
            "stderr_bytes": stderr_path.stat().st_size if stderr_path.exists() else 0,
        }
    if (
        (stdout_path.exists() and stdout_path.stat().st_size > CHILD_OUTPUT_LIMIT)
        or (stderr_path.exists() and stderr_path.stat().st_size > CHILD_OUTPUT_LIMIT)
    ) and guard_error is None:
        guard_error = f"child exceeded {CHILD_OUTPUT_LIMIT} byte output cap"
    after_host: str | None = None
    after_error: str | None = None
    try:
        after_host = W8.guard(output, f"{label}.after")
    except RuntimeError as error:
        after_error = str(error)
        if guard_error is None:
            guard_error = after_error
    exit_code = process.returncode if process is not None else exit_code
    if failed_after is not None:
        failed_after.update(
            {
                "guard_error": guard_error,
                "after_guard_error": after_error,
                "stdout_bytes": stdout_path.stat().st_size if stdout_path.exists() else 0,
                "stderr_bytes": stderr_path.stat().st_size if stderr_path.exists() else 0,
                "process_group_empty": process_tree is None or process_tree.empty(),
                "aggregate_rss_max_bytes": rss_max,
            }
        )
    host_paths = [before_host, *during_hosts, after_host] if after_host else [before_host, *during_hosts]
    receipt = {
        "schema": "litchi_signed_native_child.v1",
        "label": label,
        "command": command,
        "binary": VALIDATE.artifact(binary),
        "route": route,
        "mode": mode,
        "calls": calls,
        "fixture_root": endpoint["fixture_root"],
        "fixture_manifest": endpoint["fixtures"],
        "started_ns": started_ns,
        "finished_ns": finished_ns,
        "exit_code": exit_code,
        "host_snapshots": [VALIDATE.artifact(Path(path)) for path in host_paths],
        "during_snapshots": len(during_hosts),
        "process_group": process_group,
        "rss_max_bytes": rss_max,
        "aggregate_rss_max_bytes": rss_max,
        "timeout_seconds": CHILD_TIMEOUT_S,
        "output_limit_bytes": CHILD_OUTPUT_LIMIT,
        "rss_limit_bytes": CHILD_RSS_LIMIT,
        "timed_out": timed_out,
        "guard_error": guard_error,
        "failed_after": failed_after,
        "after_guard_error": after_error,
        "stdout": VALIDATE.artifact(stdout_path),
        "stderr": VALIDATE.artifact(stderr_path),
    }
    receipt_path = output / f"{label}.receipt.json"
    receipt_path.write_bytes(VALIDATE.canonical(receipt))
    if pending_error is not None:
        raise pending_error
    if guard_error or after_error:
        raise RuntimeError(guard_error or after_error or f"{label}: child guard failed")
    if exit_code != 0:
        raise RuntimeError(f"{label} exited {exit_code}; see {stderr_path}")
    try:
        value = json.loads(stdout_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise RuntimeError(f"{label} returned invalid JSON: {error}") from error
    if not isinstance(value, dict):
        raise RuntimeError(f"{label} returned a non-object")
    value["_receipt"] = str(receipt_path.resolve())
    value["_receipt_artifact"] = VALIDATE.artifact(receipt_path)
    return value


def route_list(output: Path, binary: Path, endpoint: dict[str, Any], cpu: int, label: str) -> list[str]:
    value = child(output, label, binary, endpoint, "*", "list", cpu)
    routes = value.get("routes")
    if routes != list(VALIDATE.ROUTES):
        raise ValueError(f"{label}: route list is not the reviewed signed matrix")
    return routes


def preflight(
    output: Path,
    binary: Path,
    endpoint: dict[str, Any],
    endpoint_label: str,
    cpu: int,
) -> dict[str, dict[str, Any]]:
    rows: dict[str, dict[str, Any]] = {}
    for index, route in enumerate(VALIDATE.ROUTES):
        value = child(
            output,
            f"{endpoint_label}.validate-{index:02}",
            binary,
            endpoint,
            route,
            "validate",
            cpu,
        )
        capability = value.get("capability")
        if capability == "not_applicable":
            if endpoint_label != "before" or route not in VALIDATE.BASELINE_NOT_APPLICABLE:
                raise ValueError(f"{endpoint_label}/{route}: invalid not_applicable disposition")
            VALIDATE.validate_correctness_receipt(value, route, endpoint="before")
        elif capability == "capability_missing":
            if endpoint_label == "after" or route not in VALIDATE.FINAL_ONLY:
                raise ValueError(f"{endpoint_label}/{route}: required capability missing")
            VALIDATE.validate_correctness_receipt(value, route, endpoint="before")
        else:
            VALIDATE.validate_correctness_receipt(
                value,
                route,
                endpoint="final" if endpoint_label == "after" else "before",
            )
        rows[route] = value
    return rows


def calibrate(
    output: Path,
    binary: Path,
    endpoint: dict[str, Any],
    route: str,
    cpu: int,
    label: str,
) -> tuple[int, dict[str, Any]]:
    value = child(output, label, binary, endpoint, route, "calibrate", cpu)
    calls = VALIDATE.validate_calibration(value, route)
    return calls, value


def measure_block(
    output: Path,
    binary: Path,
    endpoint: dict[str, Any],
    route: str,
    calls: int,
    cpu: int,
    label: str,
    sides: tuple[str, ...],
    expected_route: str | None = None,
) -> dict[str, Any]:
    legs = []
    for index, side in enumerate(sides):
        value = child(
            output,
            f"{label}-leg-{index}-{side}",
            binary,
            endpoint,
            route,
            "measure",
            cpu,
            calls,
        )
        VALIDATE.validate_measure(value, route, calls, expected_route=expected_route)
        legs.append({"side": side, "value": value})
    return {"legs": legs}


def median_per_call(blocks: list[dict[str, Any]], side: str) -> float:
    values = [
        leg["value"]["elapsed_ns"] / leg["value"]["calls"]
        for block in blocks
        for leg in block["legs"]
        if leg["side"] == side
    ]
    if not values:
        raise ValueError(f"no {side} measurements")
    return float(statistics.median(values))


def paired_route(route: str) -> str:
    if route.startswith("signed_leaf/"):
        return "signed_facade/" + route.split("/", 1)[1]
    if route.startswith("signed_facade/"):
        return "signed_leaf/" + route.split("/", 1)[1]
    raise ValueError(f"invalid signed route {route}")


def final_only_row(
    output: Path,
    after: dict[str, Any],
    after_binary: Path,
    route: str,
    rounds: int,
    cpu: int,
) -> dict[str, Any]:
    leaf_route = route if route.startswith("signed_leaf/") else paired_route(route)
    facade_route = route if route.startswith("signed_facade/") else paired_route(route)
    leaf_calls, leaf_calibration = calibrate(
        output,
        after_binary,
        after,
        leaf_route,
        cpu,
        f"{route.replace('/', '_')}-leaf-calibration",
    )
    facade_calls, facade_calibration = calibrate(
        output,
        after_binary,
        after,
        facade_route,
        cpu,
        f"{route.replace('/', '_')}-facade-calibration",
    )
    calls = max(leaf_calls, facade_calls)
    blocks: list[dict[str, Any]] = []
    for block_index in range(rounds):
        legs = []
        for leg_index, side in enumerate(("leaf", "facade", "facade", "leaf")):
            side_route = leaf_route if side == "leaf" else facade_route
            legs.append(
                measure_block(
                    output,
                    after_binary,
                    after,
                    side_route,
                    calls,
                    cpu,
                    f"{route.replace('/', '_')}-leaf-facade-{block_index:02}-{leg_index}-{side}",
                    (side,),
                    expected_route=side_route,
                )["legs"][0]
            )
        blocks.append({"legs": legs})
    leaf_ns = median_per_call(blocks, "leaf")
    facade_ns = median_per_call(blocks, "facade")
    return {
        "route": route,
        "paired_route": paired_route(route),
        "baseline": "leaf_facade",
        "comparison": "final_leaf_vs_facade",
        "calls": calls,
        "calibration": {"leaf": leaf_calibration, "facade": facade_calibration},
        "blocks": blocks,
        "leaf_ns_per_call": leaf_ns,
        "facade_ns_per_call": facade_ns,
        "facade_leaf_ratio": facade_ns / leaf_ns,
    }


def run(
    before_binary: Path,
    after_binary: Path,
    before_manifest: Path,
    after_manifest: Path,
    output: Path,
    rounds: int,
    cpu: int,
) -> dict[str, Any]:
    if rounds < VALIDATE.MIN_ABBA_ROUNDS:
        raise ValueError("at least five ABBA rounds are required")
    if cpu not in os.sched_getaffinity(0):
        raise ValueError(f"CPU {cpu} is unavailable")
    VALIDATE.reject_symlink_path(output, "qualification output")
    if output.exists() and any(output.iterdir()):
        raise ValueError(f"refusing to overwrite nonempty output {output}")
    repository = Path(__file__).resolve().parents[4]
    if output.resolve() == repository or repository in output.resolve().parents:
        raise ValueError("qualification output must be outside the checkout")
    output.mkdir(parents=True, exist_ok=True)
    VALIDATE.reject_symlink_path(before_manifest, "before endpoint manifest")
    VALIDATE.reject_symlink_path(after_manifest, "after endpoint manifest")
    before = VALIDATE.validate_endpoint_manifest(before_manifest, "before")
    after = VALIDATE.validate_endpoint_manifest(after_manifest, "after")
    pair = VALIDATE.validate_pair(before, after)
    VALIDATE.reject_symlink_path(before_binary, "before endpoint binary")
    VALIDATE.reject_symlink_path(after_binary, "after endpoint binary")
    before_binary = before_binary.resolve()
    after_binary = after_binary.resolve()
    if before_binary == after_binary:
        raise ValueError("before and after binaries must be distinct")
    if VALIDATE.artifact(before_binary)["sha256"] != before["binary"]["sha256"]:
        raise ValueError("before binary does not match its endpoint manifest")
    if VALIDATE.artifact(after_binary)["sha256"] != after["binary"]["sha256"]:
        raise ValueError("after binary does not match its endpoint manifest")
    W8.guard(output, "initial")
    route_list(output, before_binary, before, cpu, "before-list")
    route_list(output, after_binary, after, cpu, "after-list")
    before_rows = preflight(output, before_binary, before, "before", cpu)
    after_rows = preflight(output, after_binary, after, "after", cpu)

    manifest: dict[str, Any] = {
        "schema": "litchi_signed_native_qualification.v1",
        "rounds": rounds,
        "cpu": cpu,
        "target_ns": VALIDATE.TARGET_NS,
        "minimum_sample_ns": VALIDATE.MIN_SAMPLE_NS,
        "guard": "reviewed W8 full /proc guard before/during/after every child",
        "output": str(output.resolve()),
        "provenance": {
            "pair": pair,
            "before_manifest": VALIDATE.artifact(before_manifest),
            "after_manifest": VALIDATE.artifact(after_manifest),
            "before_binary": before["binary"],
            "after_binary": after["binary"],
            "before_source_commit": before["source_commit"],
            "after_source_commit": after["source_commit"],
            "cpu": cpu,
        },
        "status": "running",
        "routes": [],
    }
    save(output / "manifest.json", manifest)
    for index, route in enumerate(VALIDATE.TIMED):
        if route in VALIDATE.FINAL_ONLY:
            row = final_only_row(output, after, after_binary, route, rounds, cpu)
        else:
            before_calls, before_calibration = calibrate(
                output,
                before_binary,
                before,
                route,
                cpu,
                f"{route.replace('/', '_')}-before-calibration",
            )
            after_calls, after_calibration = calibrate(
                output,
                after_binary,
                after,
                route,
                cpu,
                f"{route.replace('/', '_')}-after-calibration",
            )
            calls = max(before_calls, after_calls)
            blocks = []
            for block_index in range(rounds):
                block_legs = []
                for leg_index, side in enumerate(("before", "after", "after", "before")):
                    block_legs.append(
                        measure_block(
                            output,
                            before_binary if side == "before" else after_binary,
                            before if side == "before" else after,
                            route,
                            calls,
                            cpu,
                            f"{route.replace('/', '_')}-abba-{block_index:02}-{leg_index}-{side}",
                            (side,),
                        )["legs"][0]
                    )
                blocks.append({"legs": block_legs})
            row = {
                "route": route,
                "baseline": "matched",
                "calls": calls,
                "calibration": {
                    "before": before_calibration,
                    "after": after_calibration,
                },
                "blocks": blocks,
                "before_ns_per_call": median_per_call(blocks, "before"),
                "after_ns_per_call": median_per_call(blocks, "after"),
            }
            row["after_before_ratio"] = row["after_ns_per_call"] / row["before_ns_per_call"]
        manifest["routes"].append(row)
        save(output / "manifest.json", manifest)
    W8.guard(output, "finished")
    manifest["status"] = "latency_complete"
    save(output / "manifest.json", manifest)
    VALIDATE.validate_abba(manifest)
    return manifest


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
    manifest = run(
        args.before,
        args.after,
        args.before_manifest,
        args.after_manifest,
        args.output,
        args.rounds,
        args.cpu,
    )
    print(
        json.dumps(
            {
                "status": manifest["status"],
                "routes": len(manifest["routes"]),
                "output": str(Path(args.output).resolve()),
            }
        )
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"0362ul24 signed qualification failed: {error}", file=sys.stderr)
        raise
