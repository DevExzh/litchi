#!/usr/bin/env python3
"""Opt-in route measurements for the 0484 replayable DOCX stream.

The existing :mod:`measure` driver remains the default deterministic lane.
This module owns a separate protocol for the deterministic, memory-store, and
file-store routes, plus a bounded deterministic one-factor inventory for
input, sink, and compression profiles.  It reuses the deterministic report
shell and sample validators without changing their module globals, then
applies the route, storage, and profile checks to every measured sample.

No command in this module builds or captures anything merely by being
imported.  The root coordinator chooses when to freeze the route protocol and
when to run its explicitly listed processes.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
from typing import Any, Iterable

import measure as base

try:
    import corpus_oracle
except ImportError:  # pragma: no cover - useful only for an incomplete checkout
    corpus_oracle = None  # type: ignore[assignment]

from common import ENV, ENV_KEYS, REPO, ROOT, TEMP, meta, now, sha, write


ROUTE_SCHEMA = "docx-replayable-tail-append-route-measurement-v1"
REPORT_SCHEMA = base.REPORT_SCHEMA
ROUTE_PROTOCOL_FILE = "route-protocol.json"
ROUTE_PROTOCOL_VERSION = 1
MACHINE_FILE = "machine.json"
ENVIRONMENT_RECORDER = "record_environment.py"
ROUTE_NAMES = ("deterministic", "memory_store", "file_store")
INPUT_MODE = "owned"
SINK_WRITE_BYTES = base.SINK_WRITE_BYTES
COMPRESSION = "current"
SOURCE_CONTRACTS = {
    "owned": "profile_owned_positional_read_at_requested_returned_fixed_histograms",
    "file": "profile_file_source_positional_read_at_requested_returned_fixed_histograms",
    "short-read": "profile_short_read_positional_read_at_requested_returned_fixed_histograms",
    "latency": "profile_latency_positional_read_at_requested_returned_fixed_histograms",
}
INPUT_BACKINGS = {"owned": "owned", "file": "file", "short-read": "owned", "latency": "owned"}
INPUT_IDENTITY_VALIDATION = {
    "owned": "setup_fingerprint_outside_timing",
    "file": "setup_and_post_sample_fingerprint_outside_timing",
    "short-read": "setup_fingerprint_outside_timing",
    "latency": "setup_fingerprint_outside_timing",
}
INPUT_RANGE_BYTES = 64 * 1024
LATENCY_DELAY_US = 50
LATENCY_OVERHEAD_US = 10
LATENCY_BYTES_PER_SECOND = 100_000_000
FILE_REFERENCE_MAX_BYTES = 4 * 1024 * 1024
FILE_REFERENCE_MAX_PATCH_BYTES = 64 * 1024
MEMORY_CAPACITY_PROVENANCE = "exact_reserve_equal_to_ceiling"
# The largest selected authored vector is 64 near-limit paragraphs with the
# one-chunk framing: its encoded XML is 6,748,480 bytes.  Eight MiB leaves a
# finite, explicit margin while remaining distinguishable from the old
# deterministic ``max_replay_bytes = 1`` sentinel.
STORE_REPLAY_MAX_BYTES = 8 * 1024 * 1024
DETERMINISTIC_REPLAY_MAX_BYTES = 1
MEMORY_REPLAY_SYNC = "none"
FILE_REPLAY_SYNC = "data"
FORMAL_SAMPLES = base.FORMAL_SAMPLES
FORMAL_WARMUPS = base.FORMAL_WARMUPS
PILOT_SAMPLES = base.PILOT_SAMPLES
PILOT_WARMUPS = base.PILOT_WARMUPS
ROLES = base.ROLES
REPEATS = base.REPEATS
CASES = tuple(dict(case) for case in base.CASES)
CASE_BY_LABEL = {case["label"]: case for case in CASES}

# These are inventory declarations for later one-factor arms.  They are
# intentionally not folded into the initial 120-process store protocol.
SEPARATE_AXIS_ARMS = {
    "input": {
        "modes": ["file", "short-read", "latency"],
        "workloads": ["s131072-a64-short-c64", "s64-a16384-short-c64", "s64-a64-short-c64"],
    },
    "sink": {
        "write_bytes": [512, SINK_WRITE_BYTES, 64 * 1024],
        "workloads": ["s131072-a64-short-c64", "s64-a16384-short-c64", "s64-a64-short-c64"],
    },
    "compression": {
        "modes": ["current", "store", "deflate"],
        "workloads": ["s131072-a64-short-c64", "s64-a16384-short-c64", "s64-a64-short-c64"],
    },
}

AXIS_WORKLOAD_LABELS = (
    "s131072-a64-short-c64",
    "s64-a16384-short-c64",
    "s64-a64-short-c64",
)


def _axis_arm(axis: str, value: Any, workload: str) -> dict[str, Any]:
    """Return one explicit one-factor arm with every process axis bound."""

    if workload not in CASE_BY_LABEL:
        raise RuntimeError(f"axis workload is not in the selected base matrix: {workload}")
    arm = {
        "axis": axis,
        "value": value,
        "workload": workload,
        "provider": "deterministic",
        "cli_provider": "deterministic",
        "report_provider": base.AUTHORED_PROVIDER,
        "input_mode": "owned",
        "input_backing": "owned",
        "input_identity_validation": INPUT_IDENTITY_VALIDATION["owned"],
        "input_file": None,
        "input_max_range_bytes": None,
        "input_delay_us": 0,
        "input_overhead_us": 0,
        "input_bytes_per_second": None,
        "sink_write_bytes": SINK_WRITE_BYTES,
        "compression": COMPRESSION,
        "replay_max_bytes": DETERMINISTIC_REPLAY_MAX_BYTES,
        "replay_sync": "none",
        "expected_authored_opens": base.EXPECTED_AUTHORED_OPENS,
        "expected_replay_opens": 0,
        "storage_profile": {
            "provider": "deterministic",
            "replay_max_bytes": DETERMINISTIC_REPLAY_MAX_BYTES,
            "replay_sync": "none",
            "replay_store": "none",
        },
    }
    if axis == "input":
        if value not in ("file", "short-read", "latency"):
            raise RuntimeError(f"unknown input arm {value!r}")
        arm["input_mode"] = value
        arm["input_backing"] = INPUT_BACKINGS[value]
        if value == "file":
            arm["input_file"] = f"axis-input/{workload}-source.docx"
        elif value == "short-read":
            arm["input_max_range_bytes"] = INPUT_RANGE_BYTES
        else:
            arm["input_max_range_bytes"] = INPUT_RANGE_BYTES
            arm["input_delay_us"] = LATENCY_DELAY_US
            arm["input_overhead_us"] = LATENCY_OVERHEAD_US
            arm["input_bytes_per_second"] = LATENCY_BYTES_PER_SECOND
    elif axis == "sink":
        if type(value) is not int or value not in (512, SINK_WRITE_BYTES, 64 * 1024):
            raise RuntimeError(f"unknown sink arm {value!r}")
        arm["sink_write_bytes"] = value
    elif axis == "compression":
        if value not in ("current", "store", "deflate"):
            raise RuntimeError(f"unknown compression arm {value!r}")
        arm["compression"] = value
    else:
        raise RuntimeError(f"unknown one-factor axis {axis!r}")
    arm["source_contract"] = SOURCE_CONTRACTS[arm["input_mode"]]
    arm["input_identity_validation"] = INPUT_IDENTITY_VALIDATION[arm["input_mode"]]
    arm["compression_profile"] = {"name": arm["compression"], "source_raw_members_preserved": True}
    arm["input_profile"] = {
        "mode": arm["input_mode"],
        "backing": arm["input_backing"],
        "file": arm["input_file"],
        "max_range_bytes": arm["input_max_range_bytes"],
        "delay_us": arm["input_delay_us"],
        "overhead_us": arm["input_overhead_us"],
        "bytes_per_second": arm["input_bytes_per_second"],
    }
    arm["label"] = f"axis-{axis}-{value}-{workload}"
    return arm


AXIS_ARMS = tuple(
    _axis_arm(axis, value, workload)
    for axis, values in (
        ("input", ("file", "short-read", "latency")),
        ("sink", (512, SINK_WRITE_BYTES, 64 * 1024)),
        ("compression", ("current", "store", "deflate")),
    )
    for workload in AXIS_WORKLOAD_LABELS
    for value in values
)
AXIS_ARM_BY_LABEL = {arm["label"]: arm for arm in AXIS_ARMS}


class RouteMeasureError(base.MeasureError):
    """A route protocol, receipt, or report failed closed validation."""


def fail(message: str) -> None:
    raise RouteMeasureError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


@dataclass(frozen=True)
class RouteSpec:
    name: str
    cli_provider: str
    report_provider: str
    expected_authored_opens: int
    expected_replay_opens: int
    replay_max_bytes: int
    replay_sync: str
    store: bool


ROUTES = (
    RouteSpec(
        name="deterministic",
        cli_provider="deterministic",
        report_provider=base.AUTHORED_PROVIDER,
        expected_authored_opens=base.EXPECTED_AUTHORED_OPENS,
        expected_replay_opens=0,
        replay_max_bytes=DETERMINISTIC_REPLAY_MAX_BYTES,
        replay_sync="none",
        store=False,
    ),
    RouteSpec(
        name="memory_store",
        cli_provider="memory-store",
        report_provider="memory_explicit_replay_store",
        expected_authored_opens=0,
        expected_replay_opens=4,
        replay_max_bytes=STORE_REPLAY_MAX_BYTES,
        replay_sync=MEMORY_REPLAY_SYNC,
        store=True,
    ),
    RouteSpec(
        name="file_store",
        cli_provider="file-store",
        report_provider="file_explicit_replay_store",
        expected_authored_opens=0,
        expected_replay_opens=4,
        replay_max_bytes=STORE_REPLAY_MAX_BYTES,
        replay_sync=FILE_REPLAY_SYNC,
        store=True,
    ),
)
ROUTE_BY_NAME = {route.name: route for route in ROUTES}

_STORE_COUNTER_FIELDS = (
    "producer_invocations",
    "store_prepare_calls",
    "store_append_calls",
    "store_appended_bytes",
    "store_finish_calls",
)


def _route(name: str) -> RouteSpec:
    try:
        return ROUTE_BY_NAME[name]
    except KeyError as error:
        fail(f"unknown route {name!r}")
        raise error


def _route_case(case: dict[str, Any]) -> dict[str, Any]:
    """Copy a base case and bind every initial route axis explicitly."""

    value = dict(case)
    value.update(
        input_mode=INPUT_MODE,
        input_backing=INPUT_BACKINGS[INPUT_MODE],
        input_file=None,
        input_max_range_bytes=None,
        input_delay_us=0,
        input_overhead_us=0,
        input_bytes_per_second=None,
        input_identity_validation=INPUT_IDENTITY_VALIDATION[INPUT_MODE],
        source_contract=SOURCE_CONTRACTS[INPUT_MODE],
        sink_write_bytes=SINK_WRITE_BYTES,
        compression=COMPRESSION,
    )
    return value


ROUTE_CASES = tuple(_route_case(case) for case in CASES)
ROUTE_CASE_BY_LABEL = {case["label"]: case for case in ROUTE_CASES}


def _script_hashes() -> dict[str, str]:
    helper = ROOT / "corpus_oracle.py"
    require(helper.is_file(), f"independent corpus oracle is missing: {helper}")
    recorder = ROOT / ENVIRONMENT_RECORDER
    require(recorder.is_file(), f"environment recorder is missing: {recorder}")
    return {
        **{name: sha(ROOT / name) for name in base.DRIVER_SCRIPTS},
        "measure.py": sha(ROOT / "measure.py"),
        "corpus_oracle.py": sha(helper),
        "measure_routes.py": sha(Path(__file__)),
        ENVIRONMENT_RECORDER: sha(recorder),
    }


def _validate_machine_record(value: Any) -> None:
    value = base._object(value, "machine")
    require(value.get("schema") == "docx-stream-route-machine-v1", "machine schema differs")
    require(value.get("driver_sha256") == sha(ROOT / ENVIRONMENT_RECORDER), "machine recorder binding differs")
    _exact_int(value.get("selected_cpu"), base.CPU, "machine.selected_cpu")
    affinity = value.get("coordinator_affinity")
    require(isinstance(affinity, list) and all(type(cpu) is int for cpu in affinity)
            and base.CPU in affinity, "machine affinity does not include the selected CPU")
    require(value.get("environment") == {key: ENV[key] for key in ENV_KEYS}, "machine build environment differs")
    scratch = base._object(value.get("scratch"), "machine.scratch")
    require(scratch.get("path") == str(TEMP.resolve()), "machine scratch capability differs")
    commands = base._object(value.get("commands"), "machine.commands")
    for name in ("cpu", "rustc", "cargo", "time", "scratch_mount"):
        record = base._object(commands.get(name), f"machine.commands.{name}")
        require(record.get("status") == "pass", f"machine {name} observation failed")
        _exact_int(record.get("exit_code"), 0, f"machine.commands.{name}.exit_code")
        for stream in ("stdout", "stderr"):
            output = record.get(stream)
            require(isinstance(output, str), f"machine {name} {stream} missing")
            require(hashlib.sha256(output.encode("utf-8")).hexdigest() == record.get(f"{stream}_sha256"),
                    f"machine {name} {stream} custody differs")
    cache = base._object(value.get("cache_policy"), "machine.cache_policy")
    require(cache.get("cold_cache_claim") is False, "machine cache claim differs from warm route protocol")


def _machine_binding(*, required: bool) -> dict[str, Any]:
    path = ROOT / MACHINE_FILE
    if not path.is_file():
        require(not required, f"machine inventory is missing: {path}; run {ENVIRONMENT_RECORDER}")
        return {"path": MACHINE_FILE, "sha256": None, "status": "missing"}
    _validate_machine_record(base.read(path))
    return {"path": MACHINE_FILE, "sha256": sha(path), "status": "ready"}


def _route_json(route: RouteSpec) -> dict[str, Any]:
    return {
        "route": route.name,
        "cli_provider": route.cli_provider,
        "report_provider": route.report_provider,
        "expected_authored_opens": route.expected_authored_opens,
        "expected_replay_opens": route.expected_replay_opens,
        "replay_max_bytes": route.replay_max_bytes,
        "replay_sync": route.replay_sync,
        "store": route.store,
        "input_mode": INPUT_MODE,
        "input_backing": INPUT_BACKINGS[INPUT_MODE],
        "input_file": None,
        "input_max_range_bytes": None,
        "input_delay_us": 0,
        "input_overhead_us": 0,
        "input_bytes_per_second": None,
        "input_identity_validation": INPUT_IDENTITY_VALIDATION[INPUT_MODE],
        "source_contract": SOURCE_CONTRACTS[INPUT_MODE],
        "input_profile": {
            "mode": INPUT_MODE,
            "backing": INPUT_BACKINGS[INPUT_MODE],
            "file": None,
            "max_range_bytes": None,
            "delay_us": 0,
            "overhead_us": 0,
            "bytes_per_second": None,
        },
        "sink_write_bytes": SINK_WRITE_BYTES,
        "compression": COMPRESSION,
        "compression_profile": {"name": COMPRESSION, "source_raw_members_preserved": True},
        "storage_profile": {
            "provider": route.name,
            "replay_max_bytes": route.replay_max_bytes,
            "replay_sync": route.replay_sync,
            "replay_store": route.name if route.store else "none",
        },
    }


def _run_inventory(*, pilot: bool) -> list[dict[str, Any]]:
    values: list[dict[str, Any]] = []
    for route in ROUTES:
        for role in ROLES:
            for case in ROUTE_CASES:
                values.append({
                    "route": route.name,
                    "role": role,
                    "case": case["label"],
                    "input_mode": INPUT_MODE,
                    "input_backing": INPUT_BACKINGS[INPUT_MODE],
                    "input_file": None,
                    "input_max_range_bytes": None,
                    "input_delay_us": 0,
                    "input_overhead_us": 0,
                    "input_bytes_per_second": None,
                    "input_identity_validation": INPUT_IDENTITY_VALIDATION[INPUT_MODE],
                    "source_contract": SOURCE_CONTRACTS[INPUT_MODE],
                    "sink_write_bytes": SINK_WRITE_BYTES,
                    "compression": COMPRESSION,
                    "replay_max_bytes": route.replay_max_bytes,
                    "replay_sync": route.replay_sync,
                    "storage_profile": {
                        "provider": route.name,
                        "replay_max_bytes": route.replay_max_bytes,
                        "replay_sync": route.replay_sync,
                        "replay_store": route.name if route.store else "none",
                    },
                    "compression_profile": {"name": COMPRESSION, "source_raw_members_preserved": True},
                    "input_profile": {
                        "mode": INPUT_MODE,
                        "backing": INPUT_BACKINGS[INPUT_MODE],
                        "file": None,
                        "max_range_bytes": None,
                        "delay_us": 0,
                        "overhead_us": 0,
                        "bytes_per_second": None,
                    },
                    "label": f"{'pilot-' if pilot else ''}{route.name}-{role}-{case['label']}",
                })
    if not pilot:
        formal: list[dict[str, Any]] = []
        for repeat, sequence in ((1, values), (2, tuple(reversed(values)))):
            for item in sequence:
                formal.append({**item, "repeat": repeat, "label": f"r{repeat}-{item['label']}"})
        return formal
    return values


def _axis_case(arm: dict[str, Any]) -> dict[str, Any]:
    case = dict(CASE_BY_LABEL[arm["workload"]])
    case.update(
        input_mode=arm["input_mode"],
        sink_write_bytes=arm["sink_write_bytes"],
        compression=arm["compression"],
    )
    return case


def _axis_inventory(*, pilot: bool) -> list[dict[str, Any]]:
    values: list[dict[str, Any]] = []
    for arm in AXIS_ARMS:
        for role in ROLES:
            values.append({
                **dict(arm),
                "role": role,
                "arm": arm["label"],
                "case": arm["workload"],
                "label": f"{'pilot-' if pilot else ''}{arm['label']}-{role}",
            })
    if pilot:
        return values
    formal: list[dict[str, Any]] = []
    for repeat, sequence in ((1, values), (2, tuple(reversed(values)))):
        for item in sequence:
            formal.append({**item, "repeat": repeat, "label": f"r{repeat}-{item['label']}"})
    return formal


MISSING_INTERSECTIONS = (
    {
        "kind": "provider-axis",
        "providers": ["memory_store", "file_store"],
        "axis": "input",
        "values": ["file", "short-read", "latency"],
        "workloads": list(AXIS_WORKLOAD_LABELS),
        "roles": list(ROLES),
        "repeats": list(REPEATS),
        "reason": "store-provider input intersections are not in the bounded one-factor inventory",
    },
    {
        "kind": "provider-axis",
        "providers": ["memory_store", "file_store"],
        "axis": "sink",
        "values": [512, SINK_WRITE_BYTES, 64 * 1024],
        "workloads": list(AXIS_WORKLOAD_LABELS),
        "roles": list(ROLES),
        "repeats": list(REPEATS),
        "reason": "store-provider sink intersections are not in the bounded one-factor inventory",
    },
    {
        "kind": "provider-axis",
        "providers": ["memory_store", "file_store"],
        "axis": "compression",
        "values": ["current", "store", "deflate"],
        "workloads": list(AXIS_WORKLOAD_LABELS),
        "roles": list(ROLES),
        "repeats": list(REPEATS),
        "reason": "store-provider compression intersections are not in the bounded one-factor inventory",
    },
    {
        "kind": "axis-cross",
        "axes": ["input", "sink"],
        "values": {"input": ["file", "short-read", "latency"], "sink": [512, SINK_WRITE_BYTES, 64 * 1024]},
        "workloads": list(AXIS_WORKLOAD_LABELS),
        "providers": ["deterministic"],
        "reason": "the separate inventory contains one-factor arms only",
    },
    {
        "kind": "axis-cross",
        "axes": ["input", "compression"],
        "values": {"input": ["file", "short-read", "latency"], "compression": ["current", "store", "deflate"]},
        "workloads": list(AXIS_WORKLOAD_LABELS),
        "providers": ["deterministic"],
        "reason": "the separate inventory contains one-factor arms only",
    },
    {
        "kind": "axis-cross",
        "axes": ["sink", "compression"],
        "values": {"sink": [512, SINK_WRITE_BYTES, 64 * 1024], "compression": ["current", "store", "deflate"]},
        "workloads": list(AXIS_WORKLOAD_LABELS),
        "providers": ["deterministic"],
        "reason": "the separate inventory contains one-factor arms only",
    },
    {
        "kind": "workload",
        "workloads": sorted(set(CASE_BY_LABEL) - set(AXIS_WORKLOAD_LABELS)),
        "providers": ["deterministic"],
        "reason": "separate axes are limited to the three declared request-overhead workloads",
    },
)


def _protocol_value() -> dict[str, Any]:
    formal = _run_inventory(pilot=False)
    pilots = _run_inventory(pilot=True)
    axis_formal = _axis_inventory(pilot=False)
    axis_pilots = _axis_inventory(pilot=True)
    return {
        "schema": ROUTE_SCHEMA,
        "version": ROUTE_PROTOCOL_VERSION,
        "change": 484,
        "frozen_utc": now(),
        "claim_authorized": False,
        "performance_claim": "none",
        "comparison": "none; route-specific bounded-stream enabler measurements",
        "scope": "three authored replay routes over the existing ten cases plus a separate deterministic one-factor inventory over three workloads",
        "route_axis": [_route_json(route) for route in ROUTES],
        "initial_axes": {
            "input_mode": [INPUT_MODE],
            "sink_write_bytes": [SINK_WRITE_BYTES],
            "compression": [COMPRESSION],
        },
        "separate_axis_arms": SEPARATE_AXIS_ARMS,
        "axis_arms": [dict(arm) for arm in AXIS_ARMS],
        "missing_intersections": [dict(item) for item in MISSING_INTERSECTIONS],
        "old_deterministic_max_replay_bytes": DETERMINISTIC_REPLAY_MAX_BYTES,
        "store_replay_max_bytes": STORE_REPLAY_MAX_BYTES,
        "cpu": base.CPU,
        "source_counts": list(base.SOURCE_COUNTS),
        "authored_counts": list(base.AUTHORED_COUNTS),
        "chunk_modes": list(base.CHUNK_MODES),
        "text_modes": list(base.TEXT_MODES),
        "lifecycle": list(base.LIFECYCLE),
        "source_contract": SOURCE_CONTRACTS[INPUT_MODE],
        "source_contracts": dict(SOURCE_CONTRACTS),
        "initial_source_contract": SOURCE_CONTRACTS[INPUT_MODE],
        "sink_contract": base.SINK_CONTRACT,
        "formal": {"samples": FORMAL_SAMPLES, "warmups": FORMAL_WARMUPS},
        "pilot": {"samples": PILOT_SAMPLES, "warmups": PILOT_WARMUPS},
        "roles": list(ROLES),
        "repeats": list(REPEATS),
        "cases": [dict(case) for case in ROUTE_CASES],
        "expected_formal_processes": len(formal),
        "expected_pilot_processes": len(pilots),
        "expected_axis_formal_processes": len(axis_formal),
        "expected_axis_pilot_processes": len(axis_pilots),
        "formal_runs": formal,
        "pilot_runs": pilots,
        "axis_formal_runs": axis_formal,
        "axis_pilot_runs": axis_pilots,
        "environment": {key: ENV[key] for key in ENV_KEYS},
        "machine": _machine_binding(required=False),
        "scripts": _script_hashes(),
        "binary": {
            "name": base.BINARY_NAME,
            "cargo_manifest": "tools/perf-baseline/Cargo.toml",
            "release_profile": True,
            "allocator_feature": "allocator-metrics",
        },
    }


def protocol_path() -> Path:
    return ROOT / ROUTE_PROTOCOL_FILE


def load_protocol() -> tuple[dict[str, Any], str]:
    path = protocol_path()
    require(path.is_file(), f"route measurement protocol is missing: {path}")
    value = base.read(path)
    require(isinstance(value, dict), f"{ROUTE_PROTOCOL_FILE} must be an object")
    require(
        value.get("schema") == ROUTE_SCHEMA and value.get("version") == ROUTE_PROTOCOL_VERSION,
        "route protocol schema/version differs",
    )
    expected = _protocol_value()
    for key in (
        "route_axis", "initial_axes", "separate_axis_arms", "axis_arms", "missing_intersections",
        "source_counts", "authored_counts",
        "chunk_modes", "text_modes", "lifecycle", "source_contract", "source_contracts",
        "initial_source_contract", "sink_contract", "formal",
        "pilot", "roles", "repeats", "cases", "formal_runs", "pilot_runs", "axis_formal_runs",
        "axis_pilot_runs", "expected_formal_processes", "expected_pilot_processes",
        "expected_axis_formal_processes", "expected_axis_pilot_processes", "binary",
    ):
        require(value.get(key) == expected[key], f"route protocol {key} differs from driver")
    require(value.get("old_deterministic_max_replay_bytes") == DETERMINISTIC_REPLAY_MAX_BYTES, "old deterministic replay ceiling differs")
    require(value.get("store_replay_max_bytes") == STORE_REPLAY_MAX_BYTES, "store replay ceiling differs")
    require(value.get("machine") == _machine_binding(required=True), "machine inventory changed after protocol freeze")
    require(value.get("scripts") == expected["scripts"], "route measurement helpers changed after protocol freeze")
    return value, base._json_hash(path)


def freeze() -> None:
    path = protocol_path()
    require(not path.exists(), f"refusing to replace existing route protocol: {path}")
    _machine_binding(required=True)
    require((ROOT / ENVIRONMENT_RECORDER).is_file(), f"environment recorder is missing: {ROOT / ENVIRONMENT_RECORDER}")
    value = _protocol_value()
    write(path, value)
    print(
        f"wrote {path} with {len(value['route_axis'])} routes, "
        f"{len(value['formal_runs'])} formal processes, and {len(value['pilot_runs'])} pilots"
    )


def _positive_int(value: Any, path: str) -> int:
    require(type(value) is int and value > 0, f"{path}: expected a positive integer")
    return value


def _exact_int(value: Any, expected: int, path: str) -> int:
    require(type(value) is int, f"{path}: expected an integer")
    require(value == expected, f"{path}: expected {expected}")
    return value


def _nonnegative_int(value: Any, path: str) -> int:
    require(type(value) is int and value >= 0, f"{path}: expected a nonnegative integer")
    return value


def _optional_nonnegative_int(value: Any, path: str) -> int | None:
    if value is None:
        return None
    return _nonnegative_int(value, path)


def _sha256(value: Any, path: str) -> str:
    require(isinstance(value, str) and base.SHA256.fullmatch(value) is not None, f"{path}: invalid SHA-256")
    return value


def _field(record: dict[str, Any], name: str, path: str) -> Any:
    if name in record:
        return record[name]
    fail(f"{path}.{name}: missing route evidence")
    return None


def _reject_alias(record: dict[str, Any], name: str, path: str) -> None:
    require(name not in record, f"{path}.{name}: non-canonical alias is not accepted")


def _check_route_histogram(value: Any, path: str, calls: int, observed_bytes: int) -> None:
    # Reuse the base bucket lower/upper proof.  Replay readers do not expose a
    # largest item, so the aggregate is the independently checked quantity.
    try:
        base._check_histogram(value, path, calls, observed_bytes=observed_bytes)
    except base.MeasureError as error:
        fail(str(error))


def _check_corpus_case(observed: dict[str, Any], path: str) -> None:
    require(corpus_oracle is not None, f"{path}: independent corpus oracle is unavailable")
    try:
        corpus_oracle.validate_case(observed)
    except (KeyError, TypeError, ValueError) as error:
        fail(f"{path}: independent corpus oracle rejected report: {error}")


def _check_compression_profile(observed: dict[str, Any], compression: str, path: str) -> None:
    source = base._object(observed.get("source"), f"{path}.source")
    members = source.get("members")
    require(isinstance(members, list), f"{path}.source.members: compression profile cannot be authenticated")
    document = [member for member in members if isinstance(member, dict) and member.get("path") == "word/document.xml"]
    require(len(document) == 1, f"{path}.source.members: exact main document member is required")
    expected_method = {"store": "Store", "deflate": "Deflate"}.get(compression)
    if expected_method is not None:
        require(
            document[0].get("compression_method") == expected_method,
            f"{path}.source.members[word/document.xml].compression_method: differs from {compression} profile",
        )


def _check_replay_observation(
    sample: dict[str, Any],
    spec: RouteSpec,
    proof: dict[str, Any],
    path: str,
) -> None:
    replay = sample.get("replay")
    if not spec.store:
        # The deterministic route intentionally has no store observation.  Its
        # authored counters are checked by the shared validator.
        require(replay is None, f"{path}.replay: deterministic route must not report store counters")
        return
    replay = base._object(replay, f"{path}.replay")
    _reject_alias(replay, "file", f"{path}.replay")
    require(replay.get("route") == spec.name, f"{path}.replay.route: route identity differs")
    authored = base._object(proof.get("authored"), f"{path}.proof.authored")
    encoded = _positive_int(authored.get("encoded_xml_bytes"), f"{path}.proof.authored.encoded_xml_bytes")
    expected_replays = spec.expected_replay_opens
    require(encoded <= spec.replay_max_bytes, f"{path}.proof.authored.encoded_xml_bytes: exceeds selected store ceiling")

    required_fields = (
        "route", "producer_invocations", "store_prepare_calls", "store_append_calls",
        "store_appended_bytes", "store_finish_calls", "replay_opens", "replay_read_calls",
        "replay_requested_bytes", "replay_returned_bytes", "replay_finish_calls",
        "replay_sha256_checks", "request_histogram", "returned_histogram",
        "retained_logical_bytes", "retained_capacity_bytes", "retained_capacity_provenance",
        "file_logical_bytes", "file_allocated_bytes", "file_write_calls", "file_sync_calls",
        "file_cleanup_verified", "seal_sha256_checks", "cleanup_sha256_checks",
        "durable_reference_kind", "durable_reference_bytes", "durable_reference_sha256",
    )
    for field in required_fields:
        require(field in replay, f"{path}.replay.{field}: missing route evidence")
    for field in _STORE_COUNTER_FIELDS:
        _nonnegative_int(_field(replay, field, f"{path}.replay"), f"{path}.replay.{field}")
    require(_field(replay, "producer_invocations", path) == 1, f"{path}.replay.producer_invocations: expected one producer pass")
    require(_field(replay, "store_prepare_calls", path) == 1, f"{path}.replay.store_prepare_calls: expected one prepare")
    require(_field(replay, "store_finish_calls", path) == 1, f"{path}.replay.store_finish_calls: expected one finish")
    require(_field(replay, "store_append_calls", path) > 0, f"{path}.replay.store_append_calls: append evidence is missing")
    require(_field(replay, "store_appended_bytes", path) == encoded, f"{path}.replay.store_appended_bytes: differs from authored proof")

    replay_opens = _positive_int(replay.get("replay_opens"), f"{path}.replay.replay_opens")
    require(replay_opens == expected_replays, f"{path}.replay.replay_opens: expected {expected_replays}")
    read_calls = _positive_int(replay.get("replay_read_calls"), f"{path}.replay.replay_read_calls")
    returned = _nonnegative_int(replay.get("replay_returned_bytes"), f"{path}.replay.replay_returned_bytes")
    require(returned == encoded * expected_replays, f"{path}.replay.replay_returned_bytes: differs from four replay passes")
    checks = _positive_int(_field(replay, "replay_sha256_checks", f"{path}.replay"), f"{path}.replay.replay_sha256_checks")
    require(checks == expected_replays, f"{path}.replay.replay_sha256_checks: expected one proof check per replay pass")
    finish = _positive_int(replay.get("replay_finish_calls"), f"{path}.replay.replay_finish_calls")
    require(finish == expected_replays, f"{path}.replay.replay_finish_calls: every reader must finish")

    requested = _positive_int(replay.get("replay_requested_bytes"), f"{path}.replay.replay_requested_bytes")
    require(requested >= returned, f"{path}.replay.replay_requested_bytes: below returned bytes")
    _check_route_histogram(replay["request_histogram"], f"{path}.replay.request_histogram", read_calls, requested)
    _check_route_histogram(replay["returned_histogram"], f"{path}.replay.returned_histogram", read_calls, returned)

    logical_value = replay.get("retained_logical_bytes")
    capacity = replay.get("retained_capacity_bytes")
    file_logical = replay.get("file_logical_bytes")
    file_allocated = replay.get("file_allocated_bytes")
    file_writes = replay.get("file_write_calls")
    file_syncs = replay.get("file_sync_calls")
    cleanup = replay.get("file_cleanup_verified")
    if spec.name == "memory_store":
        logical = _positive_int(logical_value, f"{path}.replay.retained_logical_bytes")
        require(logical == encoded, f"{path}.replay.retained_logical_bytes: differs from authored proof")
        capacity = _positive_int(capacity, f"{path}.replay.retained_capacity_bytes")
        require(capacity == spec.replay_max_bytes, f"{path}.replay.retained_capacity_bytes: must equal configured exact reservation")
        require(
            replay.get("retained_capacity_provenance") == MEMORY_CAPACITY_PROVENANCE,
            f"{path}.replay.retained_capacity_provenance: exact reservation provenance is missing",
        )
        require(file_logical is None and file_allocated is None, f"{path}.replay: memory route reported file storage")
        require(file_writes is None and file_syncs is None and cleanup is None, f"{path}.replay: memory route reported file cleanup fields")
        _exact_int(replay.get("seal_sha256_checks"), 0, f"{path}.replay.seal_sha256_checks")
        _exact_int(replay.get("cleanup_sha256_checks"), 0, f"{path}.replay.cleanup_sha256_checks")
        expected_kind = "none"
    else:
        require("retained_logical_bytes" in replay and logical_value is None, f"{path}.replay.retained_logical_bytes: file route must report null external retention")
        require(capacity is None, f"{path}.replay.retained_capacity_bytes: file route must not claim memory capacity")
        require(replay.get("retained_capacity_provenance") is None, f"{path}.replay.retained_capacity_provenance: file route reported memory reservation")
        require(_nonnegative_int(file_logical, f"{path}.replay.file_logical_bytes") == encoded, f"{path}.replay.file_logical_bytes: differs from authored proof")
        _optional_nonnegative_int(file_allocated, f"{path}.replay.file_allocated_bytes")
        require(_positive_int(file_writes, f"{path}.replay.file_write_calls") > 0, f"{path}.replay.file_write_calls: file append evidence is missing")
        syncs = _nonnegative_int(file_syncs, f"{path}.replay.file_sync_calls")
        require(syncs == 1, f"{path}.replay.file_sync_calls: data-sync file route must sync exactly once")
        require(cleanup is True, f"{path}.replay.file_cleanup_verified: cleanup was not verified")
        require(
            _positive_int(replay.get("seal_sha256_checks"), f"{path}.replay.seal_sha256_checks") == 1,
            f"{path}.replay.seal_sha256_checks: expected one seal proof",
        )
        require(
            _positive_int(replay.get("cleanup_sha256_checks"), f"{path}.replay.cleanup_sha256_checks") == 1,
            f"{path}.replay.cleanup_sha256_checks: expected one cleanup proof",
        )
        expected_kind = "file"
    reference_kind = replay.get("durable_reference_kind")
    require(reference_kind == expected_kind, f"{path}.replay.durable_reference_kind: route reference identity differs")
    reference_bytes = _nonnegative_int(replay.get("durable_reference_bytes"), f"{path}.replay.durable_reference_bytes")
    reference_hash = replay.get("durable_reference_sha256")
    if spec.name == "memory_store" and reference_kind == "none":
        require(reference_bytes == 0, f"{path}.replay.durable_reference_bytes: absent memory reference must be zero")
        require(reference_hash is None, f"{path}.replay.durable_reference_sha256: absent memory reference must be null")
    else:
        require(
            0 < reference_bytes <= FILE_REFERENCE_MAX_BYTES
            and reference_bytes <= FILE_REFERENCE_MAX_PATCH_BYTES,
            f"{path}.replay.durable_reference_bytes: reference token is outside absolute/patch bounds",
        )
        _sha256(reference_hash, f"{path}.replay.durable_reference_sha256")


def check_route_report(
    report: Path,
    role: str,
    case: dict[str, Any],
    spec: RouteSpec | str,
    *,
    samples: int,
    warmups: int,
    binary: dict[str, Any],
    argv: list[str],
    replay_dir: Path | None = None,
) -> dict[str, Any]:
    """Validate one route report with the shared deterministic shell first."""

    if isinstance(spec, str):
        spec = _route(spec)
    require(case.get("input_mode") == INPUT_MODE, "route case input mode differs from initial protocol")
    require(case.get("sink_write_bytes") == SINK_WRITE_BYTES, "route case sink size differs from initial protocol")
    require(case.get("compression") == COMPRESSION, "route case compression differs from initial protocol")
    try:
        value = base.check_report_shell(
            report,
            role,
            case,
            samples=samples,
            warmups=warmups,
            binary=binary,
            argv=argv,
            expected_sink_write_bytes=SINK_WRITE_BYTES,
            expected_authored_opens=spec.expected_authored_opens,
            expected_authored_provider=spec.report_provider,
            expected_source_contract=SOURCE_CONTRACTS[INPUT_MODE],
            expected_max_replay_bytes=spec.replay_max_bytes if spec.store else None,
        )
    except base.MeasureError as error:
        fail(str(error))
    config = base._object(value.get("config"), f"{report}.config")
    _reject_alias(config, "compression_mode", f"{report}.config")
    require(config.get("provider") == spec.name, f"{report}.config.provider: route identity differs")
    require(config.get("input_mode") == INPUT_MODE, f"{report}.config.input_mode: differs from protocol")
    require(config.get("input_storage_kind") == INPUT_BACKINGS[INPUT_MODE], f"{report}.config.input_storage_kind: differs from protocol")
    require(config.get("input_identity_validation") == INPUT_IDENTITY_VALIDATION[INPUT_MODE], f"{report}.config.input_identity_validation: differs from protocol")
    require(config.get("input_max_range_bytes") is None, f"{report}.config.input_max_range_bytes: initial route must be unbounded owned input")
    _exact_int(config.get("input_delay_us"), 0, f"{report}.config.input_delay_us")
    _exact_int(config.get("input_overhead_us"), 0, f"{report}.config.input_overhead_us")
    require(config.get("input_bytes_per_second") is None, f"{report}.config.input_bytes_per_second: initial route must not throttle input")
    _exact_int(config.get("sink_write_bytes"), SINK_WRITE_BYTES, f"{report}.config.sink_write_bytes")
    require(config.get("compression") == COMPRESSION, f"{report}.config.compression: differs from protocol")
    _exact_int(config.get("expected_replay_opens"), spec.expected_replay_opens, f"{report}.config.expected_replay_opens")
    _exact_int(config.get("replay_max_bytes"), spec.replay_max_bytes, f"{report}.config.replay_max_bytes")
    require(config.get("replay_sync") == spec.replay_sync, f"{report}.config.replay_sync: differs from protocol")
    if spec.name == "file_store":
        require(replay_dir is not None, f"{report}: file route requires an explicit replay directory")
        require(config.get("replay_dir") == str(replay_dir.resolve()), f"{report}.config.replay_dir: caller capability differs")
    else:
        require(config.get("replay_dir") is None, f"{report}.config.replay_dir: non-file route must not infer a path")

    observed = base._object(value["cases"][0], f"{report}.cases[0]")
    _reject_alias(observed, "compression_mode", f"{report}.cases[0]")
    require(observed.get("provider") == spec.name, f"{report}.cases[0].provider: route identity differs")
    _exact_int(observed.get("replay_max_bytes"), spec.replay_max_bytes, f"{report}.cases[0].replay_max_bytes")
    require(observed.get("input_mode") == INPUT_MODE, f"{report}.cases[0].input_mode: differs from protocol")
    require(observed.get("input_storage_kind") == INPUT_BACKINGS[INPUT_MODE], f"{report}.cases[0].input_storage_kind: differs from protocol")
    require(observed.get("input_identity_validation") == INPUT_IDENTITY_VALIDATION[INPUT_MODE], f"{report}.cases[0].input_identity_validation: differs from protocol")
    _exact_int(observed.get("sink_write_bytes"), SINK_WRITE_BYTES, f"{report}.cases[0].sink_write_bytes")
    require(observed.get("compression") == COMPRESSION, f"{report}.cases[0].compression: differs from protocol")
    _check_corpus_case(observed, f"{report}.cases[0]")
    _check_compression_profile(observed, COMPRESSION, f"{report}.cases[0]")
    proof = base._object(observed.get("proof"), f"{report}.cases[0].proof")
    for index, item in enumerate(observed["samples"]):
        sample = base._object(item, f"{report}.cases[0].samples[{index}]")
        _check_replay_observation(sample, spec, proof, f"{report}.cases[0].samples[{index}]")
    return value


# Keep a private spelling parallel to the deterministic driver's historical
# helper for callers that use the acceptance functions directly.
_check_route_report = check_route_report


def _axis_argv(
    binary: dict[str, Any],
    case: dict[str, Any],
    arm: dict[str, Any],
    *,
    samples: int,
    warmups: int,
    report: Path,
    resource: Path,
) -> list[str]:
    """Build the exact one-factor command line from the sealed arm."""

    argv = [
        "/usr/bin/time", "-v", "-o", str(resource),
        "/usr/bin/taskset", "-c", str(base.CPU), binary["path"],
        "--source-counts", str(case["source_count"]),
        "--authored-counts", str(case["authored_count"]),
        "--chunks", case["chunk_mode"],
        "--text", case["text_mode"],
        "--samples", str(samples),
        "--warmups", str(warmups),
        "--sink-write", str(arm["sink_write_bytes"]),
        "--authored-provider", arm["cli_provider"],
        "--replay-max-bytes", str(arm["replay_max_bytes"]),
        "--replay-sync", arm["replay_sync"],
        "--compression", arm["compression"],
        "--input-mode", arm["input_mode"],
    ]
    if arm["input_mode"] == "file":
        path = (ROOT / arm["input_file"]).resolve()
        argv.extend(["--input-file", str(path)])
    elif arm["input_mode"] in ("short-read", "latency"):
        argv.extend(["--input-max-range", str(arm["input_max_range_bytes"])])
    if arm["input_mode"] == "latency":
        argv.extend([
            "--input-delay-us", str(arm["input_delay_us"]),
            "--input-overhead-us", str(arm["input_overhead_us"]),
            "--input-bytes-per-second", str(arm["input_bytes_per_second"]),
        ])
    argv.extend(["--json", str(report)])
    return argv


def _check_axis_report(
    report: Path,
    role: str,
    arm: dict[str, Any],
    *,
    samples: int,
    warmups: int,
    binary: dict[str, Any],
    argv: list[str],
    input_metadata: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Validate a deterministic one-factor arm against its explicit profile."""

    case = _axis_case(arm)
    try:
        value = base.check_report_shell(
            report,
            role,
            case,
            samples=samples,
            warmups=warmups,
            binary=binary,
            argv=argv,
            expected_sink_write_bytes=arm["sink_write_bytes"],
            expected_authored_opens=arm["expected_authored_opens"],
            expected_authored_provider=arm["report_provider"],
            expected_source_contract=arm["source_contract"],
        )
    except base.MeasureError as error:
        fail(str(error))
    config = base._object(value.get("config"), f"{report}.config")
    _reject_alias(config, "compression_mode", f"{report}.config")
    for field, expected in (
        ("provider", arm["provider"]),
        ("input_mode", arm["input_mode"]),
        ("input_storage_kind", arm["input_backing"]),
        ("input_identity_validation", arm["input_identity_validation"]),
        ("input_max_range_bytes", arm["input_max_range_bytes"]),
        ("input_delay_us", arm["input_delay_us"]),
        ("input_overhead_us", arm["input_overhead_us"]),
        ("input_bytes_per_second", arm["input_bytes_per_second"]),
        ("sink_write_bytes", arm["sink_write_bytes"]),
        ("compression", arm["compression"]),
        ("expected_replay_opens", arm["expected_replay_opens"]),
        ("replay_max_bytes", arm["replay_max_bytes"]),
        ("replay_sync", arm["replay_sync"]),
    ):
        if isinstance(expected, int):
            _exact_int(config.get(field), expected, f"{report}.config.{field}")
        else:
            require(config.get(field) == expected, f"{report}.config.{field}: differs from one-factor arm")
    require(config.get("replay_dir") is None, f"{report}.config.replay_dir: deterministic axis arm must not use a file store")
    observed = base._object(value["cases"][0], f"{report}.cases[0]")
    _reject_alias(observed, "compression_mode", f"{report}.cases[0]")
    for field, expected in (
        ("provider", arm["provider"]),
        ("compression", arm["compression"]),
        ("input_mode", arm["input_mode"]),
        ("input_storage_kind", arm["input_backing"]),
        ("input_identity_validation", arm["input_identity_validation"]),
        ("sink_write_bytes", arm["sink_write_bytes"]),
        ("replay_max_bytes", arm["replay_max_bytes"]),
    ):
        if isinstance(expected, int):
            _exact_int(observed.get(field), expected, f"{report}.cases[0].{field}")
        else:
            require(observed.get(field) == expected, f"{report}.cases[0].{field}: differs from one-factor arm")
    _check_corpus_case(observed, f"{report}.cases[0]")
    _check_compression_profile(observed, arm["compression"], f"{report}.cases[0]")
    if arm["input_mode"] == "file":
        require(input_metadata is not None, f"{report}: file input metadata is missing from capture receipt")
        source = base._object(observed.get("source"), f"{report}.cases[0].source")
        require(source.get("archive_bytes") == input_metadata["bytes"], f"{report}.cases[0].source.archive_bytes: staged input differs")
        require(source.get("archive_sha256") == input_metadata["sha256"], f"{report}.cases[0].source.archive_sha256: staged input differs")
    else:
        require(input_metadata is None, f"{report}: non-file input has file metadata")
    proof = base._object(observed.get("proof"), f"{report}.cases[0].proof")
    for index, item in enumerate(observed["samples"]):
        sample = base._object(item, f"{report}.cases[0].samples[{index}]")
        _check_replay_observation(sample, ROUTE_BY_NAME["deterministic"], proof, f"{report}.cases[0].samples[{index}]")
    return value


def _route_attempt_root(attempt: str) -> Path:
    base._attempt(attempt)
    path = ROOT / "route-attempts" / attempt
    path.mkdir(parents=True, exist_ok=True)
    return path


def _route_binary_destination(attempt: str, role: str) -> Path:
    return TEMP / "0484" / "routes" / attempt / role / base.BINARY_NAME


def _route_build_command(role: str) -> list[str]:
    command = base._build_command(role)
    return command


def build(role: str, attempt: str) -> None:
    require(role in ROLES, f"unknown build role {role}")
    _, protocol_sha256 = load_protocol()
    attempt_root = _route_attempt_root(attempt)
    record_path = attempt_root / f"build-{role}.json"
    require(not record_path.exists(), f"refusing to replace route build record: {record_path}")
    gate_path = base._run_gate(f"route-build-{role}", attempt, _route_build_command(role))
    gate = base.read(gate_path)
    require(isinstance(gate, dict), f"{gate_path}: gate receipt must be an object")
    require(gate.get("exit_code") == 0 and gate.get("source_unchanged") is True, f"{gate_path}: build did not pass custody")
    source_before = gate.get("source_before")
    source_after = gate.get("source_after")
    require(isinstance(source_before, dict) and isinstance(source_after, dict), f"{gate_path}: source identity is missing")
    require(source_before == source_after, f"{gate_path}: source identity changed during build")
    origin = REPO / "tools" / "perf-baseline" / "target" / "release" / base.BINARY_NAME
    origin_binary = base._binary_metadata(origin, f"route-build-{role}.origin")
    destination = _route_binary_destination(attempt, role)
    require(not destination.exists(), f"refusing to replace copied route binary: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(origin, destination)
    copied_binary = base._binary_metadata(destination, f"route-build-{role}.copy")
    require(copied_binary["bytes"] == origin_binary["bytes"] and copied_binary["sha256"] == origin_binary["sha256"], f"route-build-{role}: copied binary differs from release output")
    record = {
        "schema": base.BUILD_SCHEMA,
        "version": 1,
        "attempt": attempt,
        "role": role,
        "protocol": {"path": ROUTE_PROTOCOL_FILE, "sha256": protocol_sha256},
        "command": _route_build_command(role),
        "gate": {"path": gate_path.relative_to(ROOT).as_posix(), "sha256": sha(gate_path)},
        "source_before": source_before,
        "source_after": source_after,
        "source_unchanged": True,
        "binary": copied_binary,
        "original_binary": origin_binary,
        "copied_utc": now(),
        "environment": {key: ENV[key] for key in ENV_KEYS},
    }
    write(record_path, record)
    print(f"built route {role}: {record_path}")


def _load_builds(attempt: str, protocol_sha256: str) -> dict[str, dict[str, Any]]:
    attempt_root = _route_attempt_root(attempt)
    result: dict[str, dict[str, Any]] = {}
    for role in ROLES:
        path = attempt_root / f"build-{role}.json"
        require(path.is_file(), f"{role} route build record is missing: {path}")
        value = base.read(path)
        require(isinstance(value, dict) and value.get("schema") == base.BUILD_SCHEMA, f"{path}: build schema differs")
        require(value.get("attempt") == attempt and value.get("role") == role, f"{path}: build identity differs")
        protocol = base._object(value.get("protocol"), f"{path}.protocol")
        require(protocol.get("path") == ROUTE_PROTOCOL_FILE and protocol.get("sha256") == protocol_sha256, f"{path}: route protocol binding differs")
        require(value.get("source_unchanged") is True and value.get("source_before") == value.get("source_after"), f"{path}: source custody failed")
        binary = base._object(value.get("binary"), f"{path}.binary")
        actual = base._binary_metadata(Path(binary.get("path", "")), f"{path}.binary")
        require(actual == binary, f"{path}: copied binary metadata changed")
        result[role] = value
    require(result["normal"]["source_after"] == result["allocator"]["source_after"], "route normal/allocator source identities differ")
    return result


def _route_identity(kind: str, route: RouteSpec, role: str, case: dict[str, Any], repeat: int | None) -> dict[str, Any]:
    stem = f"{route.name}-{role}-{case['label']}"
    label = f"{kind}-{stem}" if repeat is None else f"r{repeat}-{stem}"
    return {
        "kind": kind,
        "label": label,
        "route": route.name,
        "role": role,
        "case": case["label"],
        "input_mode": INPUT_MODE,
        "input_backing": INPUT_BACKINGS[INPUT_MODE],
        "input_file": None,
        "input_max_range_bytes": None,
        "input_delay_us": 0,
        "input_overhead_us": 0,
        "input_bytes_per_second": None,
        "input_identity_validation": INPUT_IDENTITY_VALIDATION[INPUT_MODE],
        "source_contract": SOURCE_CONTRACTS[INPUT_MODE],
        "sink_write_bytes": SINK_WRITE_BYTES,
        "compression": COMPRESSION,
        "replay_max_bytes": route.replay_max_bytes,
        "replay_sync": route.replay_sync,
        "storage_profile": {
            "provider": route.name,
            "replay_max_bytes": route.replay_max_bytes,
            "replay_sync": route.replay_sync,
            "replay_store": route.name if route.store else "none",
        },
        "compression_profile": {"name": COMPRESSION, "source_raw_members_preserved": True},
        "input_profile": {
            "mode": INPUT_MODE,
            "backing": INPUT_BACKINGS[INPUT_MODE],
            "file": None,
            "max_range_bytes": None,
            "delay_us": 0,
            "overhead_us": 0,
            "bytes_per_second": None,
        },
        "source_count": case["source_count"],
        "authored_count": case["authored_count"],
        "chunk_mode": case["chunk_mode"],
        "text_mode": case["text_mode"],
        "repeat": repeat,
    }


def _route_run_directory(attempt: str, identity: dict[str, Any]) -> Path:
    root = ROOT / ("route-pilots" if identity["kind"] == "pilot" else "route-captures") / attempt
    directory = root / identity["label"]
    require(not directory.exists(), f"refusing to replace route run directory: {directory}")
    directory.mkdir(parents=True)
    return directory


def _route_argv(
    binary: dict[str, Any],
    case: dict[str, Any],
    spec: RouteSpec,
    *,
    samples: int,
    warmups: int,
    report: Path,
    resource: Path,
    replay_dir: Path | None,
) -> list[str]:
    argv = [
        "/usr/bin/time", "-v", "-o", str(resource),
        "/usr/bin/taskset", "-c", str(base.CPU), binary["path"],
        *case["cli"], "--samples", str(samples), "--warmups", str(warmups),
        "--authored-provider", spec.cli_provider,
        "--replay-max-bytes", str(spec.replay_max_bytes),
        "--replay-sync", spec.replay_sync,
        "--input-mode", INPUT_MODE,
        "--compression", COMPRESSION,
        "--json", str(report),
    ]
    if spec.name == "file_store":
        require(replay_dir is not None, "file route requires replay directory")
        argv.extend(["--replay-dir", str(replay_dir.resolve())])
    return argv


def capture_one(
    attempt: str,
    route_name: str,
    role: str,
    case_label: str,
    *,
    pilot: bool,
    repeat: int | None,
) -> Path:
    spec = _route(route_name)
    require(role in ROLES, f"unknown role {role}")
    require((pilot and repeat is None) or (not pilot and repeat in REPEATS), "route repeat identity is invalid")
    case = dict(ROUTE_CASE_BY_LABEL[case_label])
    _, protocol_sha256 = load_protocol()
    builds = _load_builds(attempt, protocol_sha256)
    identity = _route_identity("pilot" if pilot else "formal", spec, role, case, repeat)
    directory = _route_run_directory(attempt, identity)
    samples = PILOT_SAMPLES if pilot else FORMAL_SAMPLES
    warmups = PILOT_WARMUPS if pilot else FORMAL_WARMUPS
    binary = builds[role]["binary"]
    report = directory / "report.json"
    resource = directory / "resource.txt"
    stdout = directory / "stdout.txt"
    stderr = directory / "stderr.txt"
    replay_dir = directory / "replay" if spec.name == "file_store" else None
    if replay_dir is not None:
        replay_dir.mkdir()
    argv = _route_argv(binary, case, spec, samples=samples, warmups=warmups, report=report, resource=resource, replay_dir=replay_dir)
    started = {
        "schema": "docx-replayable-tail-append-route-capture-v1",
        "version": 1,
        "status": "running",
        "attempt": attempt,
        "run": identity,
        "protocol": {"path": ROUTE_PROTOCOL_FILE, "sha256": protocol_sha256},
        "machine": _machine_binding(required=True),
        "build": {"path": (Path("route-attempts") / attempt / f"build-{role}.json").as_posix(), "sha256": base._json_hash(_route_attempt_root(attempt) / f"build-{role}.json")},
        "binary": binary,
        "argv": argv,
        "cwd": str(REPO),
        "environment": {key: ENV[key] for key in ENV_KEYS},
        "started_utc": now(),
    }
    write(directory / "started.json", started)
    launch_error: str | None = None
    try:
        with stdout.open("xb") as out, stderr.open("xb") as err:
            process = subprocess.run(argv, cwd=REPO, env=ENV, stdout=out, stderr=err, check=False)
        exit_code: int | None = process.returncode
    except (OSError, subprocess.SubprocessError) as error:
        exit_code = None
        launch_error = f"{type(error).__name__}: {error}"
    artifacts = {path.name: meta(path) for path in (stdout, stderr, resource, report) if path.is_file()}
    missing = [path.name for path in (stdout, stderr, resource, report) if not path.is_file()]
    validation_error: str | None = None
    if exit_code == 0:
        try:
            check_route_report(report, role, case, spec, samples=samples, warmups=warmups, binary=binary, argv=argv, replay_dir=replay_dir)
            if replay_dir is not None:
                require(not any(replay_dir.iterdir()), f"{report}: replay directory was not cleaned")
        except (base.MeasureError, OSError, ValueError) as error:
            validation_error = str(error)
    passed = exit_code == 0 and validation_error is None and not missing
    finished = dict(started, status="pass" if passed else "failed", exit_code=exit_code, finished_utc=now(), artifacts=artifacts, missing_artifacts=missing)
    if launch_error is not None:
        finished["launch_error"] = launch_error
    if validation_error is not None:
        finished["validation_error"] = validation_error
    write(directory / "receipt.json", finished)
    if not passed:
        detail = launch_error or validation_error or f"exit {exit_code}"
        fail(f"{identity['label']} failed ({detail}); receipt retained")
    print(f"captured {identity['label']} ({samples} samples, {warmups} warmups)")
    return directory / "receipt.json"


def _axis_identity(kind: str, arm: dict[str, Any], role: str, repeat: int | None) -> dict[str, Any]:
    stem = f"{arm['label']}-{role}"
    label = f"{kind}-{stem}" if repeat is None else f"r{repeat}-{stem}"
    return {
        **dict(arm),
        "kind": kind,
        "label": label,
        "axis": arm["axis"],
        "value": arm["value"],
        "role": role,
        "case": arm["workload"],
        "repeat": repeat,
    }


def _axis_run_directory(attempt: str, identity: dict[str, Any]) -> Path:
    root = ROOT / ("axis-pilots" if identity["kind"] == "pilot" else "axis-captures") / attempt
    directory = root / identity["label"]
    require(not directory.exists(), f"refusing to replace axis run directory: {directory}")
    directory.mkdir(parents=True)
    return directory


def _axis_input_metadata(arm: dict[str, Any]) -> dict[str, Any] | None:
    if arm["input_mode"] != "file":
        require(arm["input_file"] is None, "non-file axis arm has an input file binding")
        return None
    path = (ROOT / arm["input_file"]).resolve()
    require(path.is_file() and not path.is_symlink(), f"axis file input is missing or not a regular file: {path}")
    details = meta(path)
    return {
        "path": arm["input_file"],
        "absolute_path": str(path),
        "bytes": details["bytes"],
        "sha256": details["sha256"],
        "identity": "prepared_file_capability_fingerprint_must_match_report_source_archive",
    }


def capture_axis_one(
    attempt: str,
    arm_label: str,
    role: str,
    *,
    pilot: bool,
    repeat: int | None,
) -> Path:
    require(arm_label in AXIS_ARM_BY_LABEL, f"unknown one-factor arm {arm_label!r}")
    require(role in ROLES, f"unknown role {role}")
    require((pilot and repeat is None) or (not pilot and repeat in REPEATS), "axis repeat identity is invalid")
    arm = dict(AXIS_ARM_BY_LABEL[arm_label])
    input_metadata = _axis_input_metadata(arm)
    case = _axis_case(arm)
    _, protocol_sha256 = load_protocol()
    builds = _load_builds(attempt, protocol_sha256)
    identity = _axis_identity("pilot" if pilot else "formal", arm, role, repeat)
    directory = _axis_run_directory(attempt, identity)
    samples = PILOT_SAMPLES if pilot else FORMAL_SAMPLES
    warmups = PILOT_WARMUPS if pilot else FORMAL_WARMUPS
    binary = builds[role]["binary"]
    report = directory / "report.json"
    resource = directory / "resource.txt"
    stdout = directory / "stdout.txt"
    stderr = directory / "stderr.txt"
    argv = _axis_argv(binary, case, arm, samples=samples, warmups=warmups, report=report, resource=resource)
    started = {
        "schema": "docx-replayable-tail-append-axis-capture-v1",
        "version": 1,
        "status": "running",
        "attempt": attempt,
        "run": identity,
        "protocol": {"path": ROUTE_PROTOCOL_FILE, "sha256": protocol_sha256},
        "machine": _machine_binding(required=True),
        "build": {
            "path": (Path("route-attempts") / attempt / f"build-{role}.json").as_posix(),
            "sha256": base._json_hash(_route_attempt_root(attempt) / f"build-{role}.json"),
        },
        "binary": binary,
        "argv": argv,
        "axis": dict(arm),
        "input_file": input_metadata,
        "storage_profile": dict(arm["storage_profile"]),
        "compression_profile": dict(arm["compression_profile"]),
        "cwd": str(REPO),
        "environment": {key: ENV[key] for key in ENV_KEYS},
        "started_utc": now(),
    }
    write(directory / "started.json", started)
    launch_error: str | None = None
    try:
        with stdout.open("xb") as out, stderr.open("xb") as err:
            process = subprocess.run(argv, cwd=REPO, env=ENV, stdout=out, stderr=err, check=False)
        exit_code: int | None = process.returncode
    except (OSError, subprocess.SubprocessError) as error:
        exit_code = None
        launch_error = f"{type(error).__name__}: {error}"
    artifacts = {path.name: meta(path) for path in (stdout, stderr, resource, report) if path.is_file()}
    missing = [path.name for path in (stdout, stderr, resource, report) if not path.is_file()]
    validation_error: str | None = None
    if exit_code == 0:
        try:
            _check_axis_report(
                report,
                role,
                arm,
                samples=samples,
                warmups=warmups,
                binary=binary,
                argv=argv,
                input_metadata=input_metadata,
            )
        except (base.MeasureError, OSError, ValueError) as error:
            validation_error = str(error)
    passed = exit_code == 0 and validation_error is None and not missing
    finished = dict(
        started,
        status="pass" if passed else "failed",
        exit_code=exit_code,
        finished_utc=now(),
        artifacts=artifacts,
        missing_artifacts=missing,
    )
    if launch_error is not None:
        finished["launch_error"] = launch_error
    if validation_error is not None:
        finished["validation_error"] = validation_error
    write(directory / "receipt.json", finished)
    if not passed:
        detail = launch_error or validation_error or f"exit {exit_code}"
        fail(f"{identity['label']} failed ({detail}); receipt retained")
    print(f"captured {identity['label']} ({samples} samples, {warmups} warmups)")
    return directory / "receipt.json"


def _runs(protocol: dict[str, Any], *, pilot: bool) -> Iterable[dict[str, Any]]:
    key = "pilot_runs" if pilot else "formal_runs"
    expected = _run_inventory(pilot=pilot)
    runs = protocol.get(key)
    require(isinstance(runs, list) and runs == expected, f"route protocol {key} inventory differs")
    return runs


def run_all(attempt: str, *, pilot: bool) -> None:
    protocol, _ = load_protocol()
    for run in _runs(protocol, pilot=pilot):
        capture_one(
            attempt,
            run["route"],
            run["role"],
            run["case"],
            pilot=pilot,
            repeat=None if pilot else run["repeat"],
        )


def _axis_runs(protocol: dict[str, Any], *, pilot: bool) -> Iterable[dict[str, Any]]:
    key = "axis_pilot_runs" if pilot else "axis_formal_runs"
    expected = _axis_inventory(pilot=pilot)
    runs = protocol.get(key)
    require(isinstance(runs, list) and runs == expected, f"route protocol {key} inventory differs")
    return runs


def run_axis_all(attempt: str, *, pilot: bool) -> None:
    protocol, _ = load_protocol()
    for run in _axis_runs(protocol, pilot=pilot):
        capture_axis_one(
            attempt,
            run["arm"],
            run["role"],
            pilot=pilot,
            repeat=None if pilot else run["repeat"],
        )


def plan() -> None:
    print(json.dumps(_protocol_value(), indent=2, sort_keys=True))


def parser() -> argparse.ArgumentParser:
    command = argparse.ArgumentParser(description=__doc__)
    subcommands = command.add_subparsers(dest="command", required=True)
    subcommands.add_parser("freeze", help="write the immutable route protocol")
    subcommands.add_parser("plan", help="print the route protocol without writing it")
    for name, help_text in (
        ("build-normal", "build and retain the normal route binary"),
        ("build-allocator", "build and retain the allocator route binary"),
        ("capture-all", "run all formal route processes"),
        ("pilot-all", "run all excluded route pilot processes"),
        ("axis-capture-all", "run all formal one-factor axis processes"),
        ("axis-pilot-all", "run all excluded one-factor axis pilot processes"),
    ):
        sub = subcommands.add_parser(name, help=help_text)
        sub.add_argument("--attempt", required=True)
    for name, help_text in (("capture", "run one formal route process"), ("pilot", "run one excluded route pilot process")):
        sub = subcommands.add_parser(name, help=help_text)
        sub.add_argument("--attempt", required=True)
        sub.add_argument("--route", choices=ROUTE_NAMES, required=True)
        sub.add_argument("--role", choices=ROLES, required=True)
        sub.add_argument("--case", choices=tuple(ROUTE_CASE_BY_LABEL), required=True)
        if name == "capture":
            sub.add_argument("--repeat", type=int, choices=REPEATS, required=True)
    for name, help_text in (("axis-capture", "run one formal one-factor axis process"), ("axis-pilot", "run one excluded one-factor axis process")):
        sub = subcommands.add_parser(name, help=help_text)
        sub.add_argument("--attempt", required=True)
        sub.add_argument("--arm", choices=tuple(AXIS_ARM_BY_LABEL), required=True)
        sub.add_argument("--role", choices=ROLES, required=True)
        if name == "axis-capture":
            sub.add_argument("--repeat", type=int, choices=REPEATS, required=True)
    return command


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    try:
        if args.command == "freeze":
            freeze()
        elif args.command == "plan":
            plan()
        elif args.command == "build-normal":
            build("normal", base._attempt(args.attempt))
        elif args.command == "build-allocator":
            build("allocator", base._attempt(args.attempt))
        elif args.command == "capture-all":
            with base._cpu_lock():
                run_all(base._attempt(args.attempt), pilot=False)
        elif args.command == "pilot-all":
            with base._cpu_lock():
                run_all(base._attempt(args.attempt), pilot=True)
        elif args.command == "axis-capture-all":
            with base._cpu_lock():
                run_axis_all(base._attempt(args.attempt), pilot=False)
        elif args.command == "axis-pilot-all":
            with base._cpu_lock():
                run_axis_all(base._attempt(args.attempt), pilot=True)
        elif args.command == "capture":
            with base._cpu_lock():
                capture_one(base._attempt(args.attempt), args.route, args.role, args.case, pilot=False, repeat=args.repeat)
        elif args.command == "pilot":
            with base._cpu_lock():
                capture_one(base._attempt(args.attempt), args.route, args.role, args.case, pilot=True, repeat=None)
        elif args.command == "axis-capture":
            with base._cpu_lock():
                capture_axis_one(base._attempt(args.attempt), args.arm, args.role, pilot=False, repeat=args.repeat)
        elif args.command == "axis-pilot":
            with base._cpu_lock():
                capture_axis_one(base._attempt(args.attempt), args.arm, args.role, pilot=True, repeat=None)
        else:
            fail(f"unknown command {args.command!r}")
    except (base.MeasureError, OSError, ValueError, subprocess.SubprocessError) as error:
        print(f"measure_routes.py: FAIL: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
