#!/usr/bin/env python3
"""Describe and validate the supplemental cold first-commit PPT route.
The frozen PPT probe's normal timed route warms the package before measuring.
This consumer adapter gives the already-built probe a separately named route
for the one cold operation that the normal ABBA route cannot measure.  It does
not alter the probe, the production API, or the 28-route qualification
manifest.  Execution is opt-in; the default operation only validates the
reviewed binary/fixture identity and prints the future command.
"""
from __future__ import annotations
import argparse
import hashlib
import json
import math
from pathlib import Path
import subprocess
import sys
from typing import Any, Mapping, Sequence
SCHEMA = "litchi.ppt.lazy-native.cold-first-commit.v1"
PROBE_SCHEMA = "litchi.ppt.lazy-native.probe.v1"
ROUTE = "primary_first_commit_cold"
UNDERLYING_ROUTE = "primary_noop_commit"
CPU = 29
FIXTURE = {"bytes": 14_848, "sha256": "02fa44ca76903943e80a3a2da1f9dbd326c7a4a047540c3aa9cf211e34ef617c"}
BINARY_IDENTITIES = {
    "baseline-normal": {
        "bytes": 12_131_056,
        "sha256": "ffcf3aebd0adca8ec26047f732e58026506829562cf51480056315a0a598416a",
    },
    "baseline-allocator": {
        "bytes": 12_138_552,
        "sha256": "0d045e2c6a36e83d4d36fb9525b814849ec39b3d06a0e4fe02fcc9ed94cb85d7",
    },
    "candidate-normal": {
        "bytes": 12_371_352,
        "sha256": "ac72404c98bf37f5a7b9433c01880547a4b34d24a8d346dc359ecc765770ff1a",
    },
    "candidate-allocator": {
        "bytes": 12_378_664,
        "sha256": "1c69d5d7627a32a633147a33afafaf3c368cbc75a75306c3da079ce8495a8390",
    },
}
ARGS = (
    "--mode",
    "timed",
    "--route",
    UNDERLYING_ROUTE,
    "--iterations",
    "1",
    "--samples",
    "1",
    "--warmups",
    "0",
    "--max-iterations",
    "1",
)
HEX64 = set("0123456789abcdef")
class ColdRouteError(ValueError):
    """The fixed cold-route contract was not satisfied."""
def _strict_json(raw: bytes, label: str) -> Any:
    def unique(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise ColdRouteError(f"{label} contains duplicate key {key!r}")
            result[key] = value
        return result
    def reject_constant(value: str) -> Any:
        raise ColdRouteError(f"{label} contains non-finite JSON constant {value!r}")
    try:
        value = json.loads(
            raw.decode("utf-8"),
            object_pairs_hook=unique,
            parse_constant=reject_constant,
        )
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ColdRouteError(f"{label} is not strict JSON: {error}") from error
    def finite(item: Any) -> None:
        if isinstance(item, float) and not math.isfinite(item):
            raise ColdRouteError(f"{label} contains a non-finite number")
        if isinstance(item, dict):
            for child in item.values():
                finite(child)
        elif isinstance(item, list):
            for child in item:
                finite(child)
    finite(value)
    return value
def _digest(path: Path) -> dict[str, Any]:
    if path.is_symlink() or not path.is_file():
        raise ColdRouteError(f"expected a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
            size += len(chunk)
    return {"path": str(path.resolve()), "bytes": size, "sha256": digest.hexdigest()}
def _check_digest(actual: Mapping[str, Any], expected: Mapping[str, Any], label: str) -> dict[str, Any]:
    if actual.get("bytes") != expected["bytes"] or actual.get("sha256") != expected["sha256"]:
        raise ColdRouteError(f"{label} identity differs from the reviewed packet")
    if not isinstance(actual.get("sha256"), str) or len(actual["sha256"]) != 64 or set(actual["sha256"]) - HEX64:
        raise ColdRouteError(f"{label} does not carry a lowercase SHA-256")
    return dict(actual)
def check_fixture(path: Path) -> dict[str, Any]:
    return _check_digest(_digest(path), {"path": str(path.resolve()), **FIXTURE}, "cold fixture")
def check_binary(path: Path, role: str) -> dict[str, Any]:
    if role not in BINARY_IDENTITIES:
        raise ColdRouteError(f"unknown reviewed binary role: {role}")
    identity = _check_digest(_digest(path), BINARY_IDENTITIES[role], f"{role} binary")
    if not path.stat().st_mode & 0o111:
        raise ColdRouteError(f"{role} binary is not executable: {path}")
    return identity
def command(binary: Path) -> list[str]:
    return ["taskset", "--cpu-list", str(CPU), str(binary), *ARGS]
def validate_report(value: Mapping[str, Any], fixture: Mapping[str, Any], *, allocator: bool) -> dict[str, Any]:
    expected = {"schema", "mode", "route", "samples", "fixture", "postcondition"}
    if set(value) != expected:
        raise ColdRouteError("cold probe report fields differ from the fixed route")
    if value.get("schema") != PROBE_SCHEMA or value.get("mode") != "timed" or value.get("route") != UNDERLYING_ROUTE:
        raise ColdRouteError("cold probe report route identity differs")
    report_fixture = value.get("fixture")
    if not isinstance(report_fixture, dict) or set(report_fixture) != {"key", "output"} or report_fixture.get("key") != "default":
        raise ColdRouteError("cold probe report is not bound to the default fixture")
    if report_fixture.get("output") != {"bytes": fixture["bytes"], "sha256": fixture["sha256"]}:
        raise ColdRouteError("cold probe fixture digest differs")
    samples = value.get("samples")
    if not isinstance(samples, list) or len(samples) != 1:
        raise ColdRouteError("cold route requires exactly one measured sample")
    sample = samples[0]
    if not isinstance(sample, dict) or set(sample) != {"elapsed_ns", "iterations", "allocation"}:
        raise ColdRouteError("cold sample fields differ")
    if type(sample.get("elapsed_ns")) is not int or sample["elapsed_ns"] <= 0 or sample.get("iterations") != 1:
        raise ColdRouteError("cold sample is not one positive iteration")
    allocation = sample.get("allocation")
    if not isinstance(allocation, dict):
        raise ColdRouteError("cold allocation observation is malformed")
    if allocator:
        required = {
            "status", "scope", "allocation_calls", "deallocation_calls", "reallocation_calls",
            "failed_allocation_calls", "allocated_bytes", "deallocated_bytes", "live_bytes_before",
            "live_bytes_after", "peak_live_bytes_before", "peak_live_bytes_after",
        }
        if set(allocation) != required or allocation.get("status") != "measured" or allocation.get("scope") != "operation_global_system_allocator":
            raise ColdRouteError("cold allocator observation is not measured")
        if any(type(allocation[name]) is not int or allocation[name] < 0 for name in required - {"status", "scope"}):
            raise ColdRouteError("cold allocator counters are invalid")
        if allocation["live_bytes_after"] - allocation["live_bytes_before"] != allocation["allocated_bytes"] - allocation["deallocated_bytes"]:
            raise ColdRouteError("cold allocator live-byte conservation failed")
        if allocation["peak_live_bytes_before"] < allocation["live_bytes_before"] or allocation["peak_live_bytes_after"] < max(allocation["peak_live_bytes_before"], allocation["live_bytes_after"]):
            raise ColdRouteError("cold allocator peak bounds failed")
    elif allocation != {"status": "unavailable", "scope": "operation_global_system_allocator"}:
        raise ColdRouteError("normal cold probe unexpectedly reported allocation metrics")
    proof = value.get("postcondition")
    if not isinstance(proof, dict):
        raise ColdRouteError("cold postcondition is missing")
    if (
        proof.get("entry") != "commit"
        or proof.get("source_exact") is not True
        or proof.get("changed") is not False
        or proof.get("patch_noop") is not True
        or proof.get("operation_count") != 0
        or proof.get("patch_before_exact") is not True
        or proof.get("patch_after_exact") is not True
    ):
        raise ColdRouteError("cold first commit did not prove an exact no-op")
    return dict(value)
def run(binary: Path, fixture: Path, role: str, *, execute: bool) -> dict[str, Any]:
    binary_identity = check_binary(binary, role)
    fixture_identity = check_fixture(fixture)
    result: dict[str, Any] = {
        "schema": SCHEMA,
        "route": ROUTE,
        "underlying_route": UNDERLYING_ROUTE,
        "role": role,
        "cpu": CPU,
        "warmups": 0,
        "iterations": 1,
        "samples": 1,
        "binary": binary_identity,
        "fixture": fixture_identity,
        "command": command(binary),
        "claim_scope": "supplemental cold first-commit observation; not part of the 28-route ABBA acceptance claim",
    }
    if execute:
        completed = subprocess.run(
            result["command"],
            cwd=binary.parent,
            check=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=300,
        )
        if completed.returncode != 0:
            raise ColdRouteError(f"cold probe exited {completed.returncode}: {completed.stderr.strip()}")
        report = _strict_json(completed.stdout.encode(), "cold probe output")
        if not isinstance(report, dict):
            raise ColdRouteError("cold probe output must be an object")
        result["report"] = validate_report(report, FIXTURE, allocator=role.endswith("allocator"))
    return result
def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--binary", type=Path, required=True)
    parser.add_argument("--fixture", type=Path, required=True)
    parser.add_argument("--role", choices=sorted(BINARY_IDENTITIES), required=True)
    parser.add_argument("--execute", action="store_true", help="run the fixed cold route; omitted for source-only planning")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args(argv)
    try:
        result = run(args.binary.resolve(), args.fixture.resolve(), args.role, execute=args.execute)
        encoded = json.dumps(result, indent=2, sort_keys=True) + "\n"
        if args.output:
            if args.output.exists() or args.output.is_symlink():
                raise ColdRouteError(f"refusing to overwrite output: {args.output}")
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(encoded, encoding="utf-8")
        else:
            print(encoded, end="")
        return 0
    except (OSError, ColdRouteError, subprocess.TimeoutExpired) as error:
        print(f"PPT cold first-commit route failed: {error}", file=sys.stderr)
        return 1
if __name__ == "__main__":
    raise SystemExit(main())
