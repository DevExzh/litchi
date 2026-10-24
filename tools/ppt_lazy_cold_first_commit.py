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
ABBA_ORDER = ("a1", "b1", "b2", "a2")
ROLE_ENDPOINT = {"a1": "baseline", "a2": "baseline", "b1": "candidate", "b2": "candidate"}
PAIR_EVENT_KEYS = {"label", "endpoint", "binary", "args", "guard", "target", "event"}
DEFAULT_ROUNDS = 2
MAX_ROUNDS = 2
PUBLIC_RUNNER_BYTES = 86_075
PUBLIC_RUNNER_SHA256 = "88f10be52bf02a8d5210a831d38af4358738ea35aea956341d7d4e8af1342977"
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


def _public_runner() -> Any:
    path = Path(__file__).with_name("ppt_lazy_native_qualification.py").resolve()
    actual = _digest(path)
    if actual["bytes"] != PUBLIC_RUNNER_BYTES or actual["sha256"] != PUBLIC_RUNNER_SHA256:
        raise ColdRouteError("public PPT qualification runner identity differs")
    tools = str(path.parent)
    if tools not in sys.path:
        sys.path.insert(0, tools)
    import importlib.util

    spec = importlib.util.spec_from_file_location("ppt_lazy_native_qualification_cold_pair", path)
    if spec is None or spec.loader is None:
        raise ColdRouteError("cannot load public PPT qualification runner")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _aggregate(samples: Sequence[Mapping[str, Any]]) -> dict[str, int]:
    if not samples:
        raise ColdRouteError("cold pair has no measured operation samples")
    elapsed = []
    for sample in samples:
        if type(sample.get("elapsed_ns")) is not int or sample["elapsed_ns"] <= 0:
            raise ColdRouteError("cold pair sample elapsed_ns is invalid")
        if sample.get("iterations") != 1:
            raise ColdRouteError("cold pair sample is not one operation")
        elapsed.append(sample["elapsed_ns"])
    total = sum(elapsed)
    count = len(elapsed)
    return {"operation_total_ns": total, "operation_count": count, "operation_mean_ns": total // count}


def _validate_event_record(row: Mapping[str, Any], event: Any) -> None:
    if not isinstance(event, dict) or set(event) != PAIR_EVENT_KEYS:
        raise ColdRouteError("cold pair event provenance fields differ")
    expected_label = f"cold-{row['kind']}-round-{row['round']}-{row['role']}"
    if event.get("label") != expected_label or event.get("endpoint") != row["endpoint"]:
        raise ColdRouteError("cold pair event route identity differs")
    if event.get("args") != list(ARGS):
        raise ColdRouteError("cold pair event arguments differ from the fixed route")
    binary = event.get("binary")
    if not isinstance(binary, dict) or set(binary) != {"path", "bytes", "sha256"}:
        raise ColdRouteError("cold pair event binary identity is malformed")
    if not isinstance(binary["path"], str) or type(binary["bytes"]) is not int or binary["bytes"] <= 0:
        raise ColdRouteError("cold pair event binary identity is invalid")
    if not isinstance(binary["sha256"], str) or len(binary["sha256"]) != 64 or set(binary["sha256"]) - HEX64:
        raise ColdRouteError("cold pair event binary digest is invalid")
    guard = event.get("guard")
    if not isinstance(guard, dict) or set(guard) != {
        "before", "live", "after", "memory_limit_bytes", "max_output_bytes", "max_load"
    }:
        raise ColdRouteError("cold pair event guard is malformed")
    target = event.get("target")
    if not isinstance(target, dict) or set(target) != {"path", "bytes_before", "bytes_after", "hard_limit_bytes"}:
        raise ColdRouteError("cold pair event target accounting is malformed")
    child_event = event.get("event")
    if not isinstance(child_event, dict) or set(child_event) != {
        "label", "command", "cwd", "returncode", "wall_ns", "target", "stdout", "stderr", "monitor"
    }:
        raise ColdRouteError("cold pair child event is malformed")


def validate_pair_report(value: Mapping[str, Any], *, rounds: int) -> dict[str, Any]:
    """Validate the fixed ABBA aggregate without setup time or a median."""

    required = {"schema", "route", "abba_order", "rounds", "samples", "aggregate"}
    allowed = required | {
        "underlying_route",
        "warmups",
        "iterations",
        "commands",
        "manifest",
        "build_input_projection",
        "claim_scope",
        "execution",
    }
    if not required.issubset(value) or set(value) - allowed:
        raise ColdRouteError("cold pair report fields or route differ from the fixed contract")
    if value.get("schema") != SCHEMA or value.get("route") != ROUTE:
        raise ColdRouteError("cold pair report fields or route differ from the fixed contract")
    if value.get("abba_order") != list(ABBA_ORDER) or value.get("rounds") != rounds:
        raise ColdRouteError("cold pair report is not the bounded symmetric ABBA workload")
    rows = value.get("samples")
    if not isinstance(rows, list) or len(rows) != rounds * len(ABBA_ORDER) * 2:
        raise ColdRouteError("cold pair report does not contain both binary roles for every ABBA position")
    expected_positions = tuple(
        (round_number, expected_role, expected_kind)
        for round_number in range(rounds)
        for expected_role in ABBA_ORDER
        for expected_kind in ("normal", "allocator")
    )
    grouped: dict[tuple[str, str], list[Mapping[str, Any]]] = {}
    for index, row_value in enumerate(rows):
        row = row_value if isinstance(row_value, dict) else None
        row_keys = {"round", "role", "endpoint", "kind", "sample", "event"}
        if row is None or set(row) != row_keys:
            raise ColdRouteError(f"cold pair sample {index} fields differ")
        round_index = row["round"]
        role = row["role"]
        endpoint = row["endpoint"]
        kind = row["kind"]
        if type(round_index) is not int or not 0 <= round_index < rounds:
            raise ColdRouteError("cold pair sample round is invalid")
        if role not in ABBA_ORDER or ROLE_ENDPOINT[role] != endpoint or kind not in {"normal", "allocator"}:
            raise ColdRouteError("cold pair sample role/endpoint/kind is invalid")
        if (round_index, role, kind) != expected_positions[index]:
            raise ColdRouteError("cold pair samples are not in the fixed ABBA order")
        sample = row["sample"]
        if not isinstance(sample, dict) or set(sample) != {"elapsed_ns", "iterations", "allocation"}:
            raise ColdRouteError("cold pair sample must be the operation timing observation")
        _validate_event_record(row, row["event"])
        grouped.setdefault((endpoint, kind), []).append(sample)
    aggregate = value.get("aggregate")
    if not isinstance(aggregate, dict) or set(aggregate) != {"baseline", "candidate"}:
        raise ColdRouteError("cold pair aggregate endpoints differ")
    for endpoint in ("baseline", "candidate"):
        endpoint_aggregate = aggregate[endpoint]
        if not isinstance(endpoint_aggregate, dict) or set(endpoint_aggregate) != {"normal", "allocator"}:
            raise ColdRouteError("cold pair aggregate roles differ")
        for kind in ("normal", "allocator"):
            expected_aggregate = _aggregate(grouped.get((endpoint, kind), []))
            if endpoint_aggregate[kind] != expected_aggregate:
                raise ColdRouteError("cold pair aggregate is not the sum/count of operation timings")
    return dict(value)


def run_pair(
    manifest: Path,
    projection: Path,
    output: Path,
    *,
    rounds: int = DEFAULT_ROUNDS,
    execute: bool = False,
) -> dict[str, Any]:
    if type(rounds) is not int or not 1 <= rounds <= MAX_ROUNDS:
        raise ColdRouteError(f"cold pair rounds must be between 1 and {MAX_ROUNDS}")
    runner = _public_runner()
    manifest_path = manifest.resolve()
    projection_path = projection.resolve()
    manifest_value = runner._load_manifest(manifest_path, plan=False)
    if manifest_value["build"]["fixtures"]["default"]["bytes"] != FIXTURE["bytes"] or manifest_value["build"]["fixtures"]["default"]["sha256"] != FIXTURE["sha256"]:
        raise ColdRouteError("cold pair default fixture differs from the reviewed fixture")
    lifecycle = runner.validate_inputs(manifest_path, projection_path)
    endpoints = {name: lifecycle[name] for name in ("baseline", "candidate")}
    roots = [endpoints[endpoint]["root"] for endpoint in ("baseline", "candidate")]
    resources = [
        output.resolve(),
        *[
            endpoints[endpoint][target]
            for endpoint in endpoints
            for target in ("target_dir", "allocator_target_dir")
        ],
    ]
    if any(runner._overlap(resource, root) for resource in resources for root in roots):
        raise ColdRouteError("cold pair output or target directory overlaps a source worktree")
    for index, first in enumerate(resources):
        for second in resources[index + 1 :]:
            if runner._overlap(first, second):
                raise ColdRouteError("cold pair output and target directories overlap")
    commands = {
        endpoint: {
            kind: [str(endpoints[endpoint][f"{kind}_binary"]["path"]), *ARGS]
            for kind in ("normal", "allocator")
        }
        for endpoint in endpoints
    }
    result: dict[str, Any] = {
        "schema": SCHEMA,
        "route": ROUTE,
        "underlying_route": UNDERLYING_ROUTE,
        "abba_order": list(ABBA_ORDER),
        "rounds": rounds,
        "warmups": 0,
        "iterations": 1,
        "samples": [],
        "commands": commands,
        "manifest": lifecycle["manifest"],
        "build_input_projection": lifecycle["build_input_projection"],
        "claim_scope": "symmetric bounded cold operation observations; timing is the probe's operation elapsed_ns only",
        "execution": "not_run",
    }
    if not execute:
        return result
    out_dir = runner._prepare_output(output.resolve(), roots)
    manifest_identity = lifecycle["manifest"]
    protocol = dict(manifest_value["protocol"])
    protocol["cpu"] = CPU
    for round_index in range(rounds):
        for role in ABBA_ORDER:
            endpoint_name = ROLE_ENDPOINT[role]
            endpoint = endpoints[endpoint_name]
            runner._recheck(
                endpoint,
                lifecycle["probe_sources"],
                manifest_path,
                manifest_identity,
                f"cold-before-{round_index}-{role}",
                projection_path,
                lifecycle["build_input_projection"],
            )
            for kind in ("normal", "allocator"):
                binary = Path(endpoint[f"{kind}_binary"]["path"])
                target = endpoint["target_dir"] if kind == "normal" else endpoint["allocator_target_dir"]
                report, event = runner._run_child(
                    endpoint,
                    protocol,
                    out_dir,
                    f"cold-{kind}-round-{round_index}-{role}",
                    binary,
                    target,
                    ARGS,
                )
                if not isinstance(event, dict) or "event" not in event:
                    raise ColdRouteError("cold child did not return event provenance")
                try:
                    runner._validate_child_event(
                        event["event"],
                        endpoint=endpoint,
                        out_dir=out_dir,
                        label=f"cold-{kind}-round-{round_index}-{role}",
                        binary=binary,
                        target_dir=target,
                        args=ARGS,
                        protocol=protocol,
                    )
                except RuntimeError as error:
                    raise ColdRouteError(f"cold child event provenance failed: {error}") from error
                checked = validate_report(report, FIXTURE, allocator=kind == "allocator")
                result["samples"].append({
                    "round": round_index,
                    "role": role,
                    "endpoint": endpoint_name,
                    "kind": kind,
                    "sample": checked["samples"][0],
                    "event": event,
                })
            runner._recheck(
                endpoint,
                lifecycle["probe_sources"],
                manifest_path,
                manifest_identity,
                f"cold-after-{round_index}-{role}",
                projection_path,
                lifecycle["build_input_projection"],
            )
    result["aggregate"] = {
        endpoint: {
            kind: _aggregate([
                row["sample"] for row in result["samples"]
                if row["endpoint"] == endpoint and row["kind"] == kind
            ])
            for kind in ("normal", "allocator")
        }
        for endpoint in ("baseline", "candidate")
    }
    result["execution"] = "native_executed"
    validate_pair_report(result, rounds=rounds)
    return result
def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--projection", type=Path, required=True)
    parser.add_argument("--out-dir", type=Path, required=True)
    parser.add_argument("--rounds", type=int, default=DEFAULT_ROUNDS)
    parser.add_argument("--execute", action="store_true", help="run the fixed pair; omitted for source-only planning")
    args = parser.parse_args(argv)
    try:
        result = run_pair(
            args.manifest.resolve(),
            args.projection.resolve(),
            args.out_dir.resolve(),
            rounds=args.rounds,
            execute=args.execute,
        )
        args.out_dir.mkdir(parents=True, exist_ok=True)
        output = args.out_dir / ("report.json" if args.execute else "plan.json")
        if output.exists() or output.is_symlink():
            raise ColdRouteError(f"refusing to overwrite output: {output}")
        output.write_bytes((json.dumps(result, indent=2, sort_keys=True) + "\n").encode())
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0
    except (OSError, ColdRouteError, subprocess.TimeoutExpired) as error:
        print(f"PPT cold first-commit route failed: {error}", file=sys.stderr)
        return 1
if __name__ == "__main__":
    raise SystemExit(main())
