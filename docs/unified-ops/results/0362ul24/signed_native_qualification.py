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
from pathlib import Path
import re
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


def save(path: Path, value: object) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def load(path: Path, label: str) -> dict[str, Any]:
    return VALIDATE.load_json(path, label)


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
    if not SAFE_ROUTE.fullmatch(route):
        raise ValueError(f"invalid route name: {route!r}")
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
    environment = dict(__import__("os").environ)
    environment["LITCHI_SIGNED_FIXTURE_ROOT"] = endpoint["fixture_root"]
    environment["LITCHI_SIGNED_ENDPOINT"] = (
        "before" if label.startswith("before") else "after"
    )
    environment["LITCHI_SIGNED_TARGET_NS"] = str(VALIDATE.TARGET_NS)
    environment["LITCHI_SIGNED_MAX_CALLS"] = str(VALIDATE.MAX_CALLS)

    before_host = W8.guard(output, f"{label}.before")
    started_ns = time.time_ns()
    with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
        process = subprocess.Popen(
            command,
            stdout=stdout,
            stderr=stderr,
            env=environment,
            start_new_session=True,
        )
        while process.poll() is None:
            time.sleep(0.1)
            W8.guard(output, f"{label}.during-{time.time_ns()}", {process.pid})
        exit_code = process.returncode
    finished_ns = time.time_ns()
    after_host = W8.guard(output, f"{label}.after")
    receipt = {
        "schema": "litchi_signed_native_child.v1",
        "label": label,
        "command": command,
        "binary": VALIDATE.artifact(binary),
        "route": route,
        "mode": mode,
        "calls": calls,
        "started_ns": started_ns,
        "finished_ns": finished_ns,
        "exit_code": exit_code,
        "host_snapshots": [before_host, after_host],
        "stdout": VALIDATE.artifact(stdout_path),
        "stderr": VALIDATE.artifact(stderr_path),
    }
    save(output / f"{label}.receipt.json", receipt)
    if exit_code != 0:
        raise RuntimeError(f"{label} exited {exit_code}; see {stderr_path}")
    try:
        value = json.loads(stdout_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise RuntimeError(f"{label} returned invalid JSON: {error}") from error
    if not isinstance(value, dict):
        raise RuntimeError(f"{label} returned a non-object")
    value["_receipt"] = str((output / f"{label}.receipt.json").resolve())
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
        if capability == "capability_missing":
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
        VALIDATE.validate_measure(value, route, calls)
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
    if cpu not in __import__("os").sched_getaffinity(0):
        raise ValueError(f"CPU {cpu} is unavailable")
    if output.exists() and any(output.iterdir()):
        raise ValueError(f"refusing to overwrite nonempty output {output}")
    output.mkdir(parents=True, exist_ok=True)
    before = VALIDATE.validate_endpoint_manifest(before_manifest.resolve(), "before")
    after = VALIDATE.validate_endpoint_manifest(after_manifest.resolve(), "after")
    pair = VALIDATE.validate_pair(before, after)
    if before_binary.resolve() == after_binary.resolve():
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
        "provenance": pair,
        "status": "running",
        "routes": [],
    }
    save(output / "manifest.json", manifest)
    for index, route in enumerate(VALIDATE.TIMED):
        baseline_missing = before_rows[route].get("capability") == "capability_missing"
        if baseline_missing:
            calls, calibration = calibrate(
                output,
                after_binary,
                after,
                route,
                cpu,
                f"{route.replace('/', '_')}-after-calibration",
            )
            blocks = [
                measure_block(
                    output,
                    after_binary,
                    after,
                    route,
                    calls,
                    cpu,
                    f"{route.replace('/', '_')}-final-{block_index:02}",
                    ("after", "after", "after", "after"),
                )
                for block_index in range(rounds)
            ]
            row = {
                "route": route,
                "baseline": "capability_missing",
                "calls": calls,
                "calibration": {"after": calibration},
                "blocks": blocks,
                "final_after_ns_per_call": median_per_call(blocks, "after"),
            }
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
    manifest["status"] = "complete"
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
        args.before.resolve(),
        args.after.resolve(),
        args.before_manifest.resolve(),
        args.after_manifest.resolve(),
        args.output.resolve(),
        args.rounds,
        args.cpu,
    )
    print(
        json.dumps(
            {
                "status": manifest["status"],
                "routes": len(manifest["routes"]),
                "output": str(args.output.resolve()),
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
