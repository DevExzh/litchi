#!/usr/bin/env python3
"""Independent read-only replay bound to the final 0924 attribution summary."""
from __future__ import annotations

import argparse
import collections
import hashlib
import json
import pathlib
import subprocess

DEFAULT_SUMMARY = pathlib.Path("/tmp/litchi-performance-0924/stack-attribution-v2.json")
SCHEMA = "litchi.performance.0924.heaptrack-stack-attribution.v2"
READ_LIMITED = "litchi_opc::phys_pkg::read_limited"
ROUTES = ("generic_read_at", "owned_vec", "owned_reader")
VARIANTS = ("baseline", "candidate")


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def digest(path: pathlib.Path) -> dict[str, object]:
    digest_value = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest_value.update(block)
            size += len(block)
    return {"path": str(path), "bytes": size, "sha256": digest_value.hexdigest()}


def parse(path: pathlib.Path):
    """Parse Heaptrack's event tables independently of the attribution tool."""
    strings = [""]
    instructions = [()]
    traces = [()]
    allocations = []
    bytes_by_stack = collections.Counter()
    calls_by_stack = collections.Counter()
    events = 0
    requested = 0

    process = subprocess.Popen(
        ["zstdcat", str(path)],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    require(process.stdout is not None, f"zstdcat stdout unavailable: {path}")
    assert process.stdout is not None
    for line_number, line in enumerate(process.stdout, 1):
        tag, _, rest = line.rstrip("\n").partition(" ")
        if tag == "s":
            fields = rest.split(" ", 1)
            require(len(fields) == 2, f"{path}:{line_number}: malformed string")
            strings.append(fields[1])
        elif tag == "i":
            fields = rest.split()
            instructions.append(tuple(strings[int(fields[index], 16)] for index in range(2, len(fields), 3)))
        elif tag == "t":
            instruction, parent = (int(value, 16) for value in rest.split())
            require(instruction < len(instructions), f"{path}:{line_number}: instruction reference")
            require(parent < len(traces), f"{path}:{line_number}: trace reference")
            traces.append(instructions[instruction] + traces[parent])
        elif tag == "a":
            size, trace = (int(value, 16) for value in rest.split())
            require(trace < len(traces), f"{path}:{line_number}: allocation trace reference")
            allocations.append((size, trace))
        elif tag == "+":
            allocation = int(rest, 16)
            require(allocation < len(allocations), f"{path}:{line_number}: allocation event reference")
            size, trace = allocations[allocation]
            stack = traces[trace]
            bytes_by_stack[stack] += size
            calls_by_stack[stack] += 1
            events += 1
            requested += size
    error = process.stderr.read() if process.stderr is not None else ""
    status = process.wait()
    require(status == 0, f"zstdcat failed for {path}: {error.strip()}")
    return events, requested, bytes_by_stack, calls_by_stack


def compare_value(actual: object, expected: object, label: str) -> None:
    require(actual == expected, f"{label}: expected {expected!r}, got {actual!r}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--summary", type=pathlib.Path, default=DEFAULT_SUMMARY)
    args = parser.parse_args()
    summary_path = args.summary.resolve()
    summary_descriptor = digest(summary_path)
    data = json.loads(summary_path.read_text(encoding="utf-8"))
    compare_value(data.get("schema"), SCHEMA, "summary schema")
    compare_value(data.get("raw_only"), True, "summary raw_only")
    profiles = data.get("profiles")
    require(isinstance(profiles, list) and len(profiles) == 25, "final summary must contain 25 profiles")

    # Bind and hash every raw descriptor before parsing any Heaptrack input.
    bound = []
    profile_ids = set()
    raw_paths = set()
    raw_hashes = set()
    for profile in profiles:
        profile_id = str(profile.get("id"))
        require(profile_id not in profile_ids, f"duplicate profile id: {profile_id}")
        profile_ids.add(profile_id)
        raw = profile.get("raw")
        require(isinstance(raw, dict), f"{profile_id}: missing raw descriptor")
        raw_path = pathlib.Path(str(raw.get("path"))).resolve()
        expected_bytes = int(raw.get("bytes"))
        expected_sha = str(raw.get("sha256"))
        actual = digest(raw_path)
        compare_value(actual["bytes"], expected_bytes, f"{profile_id}: raw byte size")
        compare_value(actual["sha256"], expected_sha, f"{profile_id}: raw SHA-256")
        require(raw_path not in raw_paths, f"duplicate raw path: {raw_path}")
        require(expected_sha not in raw_hashes, f"duplicate raw SHA-256: {profile_id}")
        raw_paths.add(raw_path)
        raw_hashes.add(expected_sha)
        bound.append((profile, raw_path, actual))
    require(len(raw_paths) == len(raw_hashes) == 25, "raw descriptors are not unique")

    # Only after all descriptors are bound and verified do we parse the profiles.
    parsed = []
    for profile, raw_path, actual in bound:
        events, requested, bytes_by_stack, calls_by_stack = parse(raw_path)
        compare_value(events, profile["allocation_event_count"], f"{profile['id']}: event count")
        compare_value(requested, profile["requested_bytes"], f"{profile['id']}: requested bytes")
        parsed.append((profile, bytes_by_stack, calls_by_stack))

    groups = {}
    group_checks = {}
    for route in ROUTES:
        for variant in VARIANTS:
            matching = [item for item in parsed if item[0]["route"] == route and item[0]["variant"] == variant]
            bytes_by_stack = collections.Counter()
            calls_by_stack = collections.Counter()
            for _, profile_bytes, profile_calls in matching:
                bytes_by_stack.update(profile_bytes)
                calls_by_stack.update(profile_calls)
            groups[route, variant] = (bytes_by_stack, calls_by_stack, len(matching))
            key = f"{route}:{variant}"
            expected = data["groups"][key]
            actual = {
                "profiles": len(matching),
                "requested_bytes": sum(bytes_by_stack.values()),
                "allocation_event_count": sum(calls_by_stack.values()),
                "read_limited_bytes": sum(value for stack, value in bytes_by_stack.items() if any(READ_LIMITED in frame for frame in stack)),
                "read_limited_calls": sum(value for stack, value in calls_by_stack.items() if any(READ_LIMITED in frame for frame in stack)),
            }
            for field, value in actual.items():
                compare_value(expected[field], value, f"{key}.{field}")
            group_checks[key] = actual

    pair_checks = {}
    for variant in VARIANTS:
        reader_bytes, reader_calls, reader_profiles = groups["owned_reader", variant]
        vec_bytes, vec_calls, vec_profiles = groups["owned_vec", variant]
        keys = set(reader_bytes) | set(vec_bytes)
        reader_only_keys = set(reader_bytes) - set(vec_bytes)
        vec_only_keys = set(vec_bytes) - set(reader_bytes)
        primary_keys = {
            stack for stack in reader_only_keys if any(READ_LIMITED in frame for frame in stack)
        }
        common_rows = [
            (stack, reader_bytes[stack] - vec_bytes[stack], reader_calls[stack] - vec_calls[stack])
            for stack in reader_bytes.keys() & vec_bytes.keys()
            if reader_bytes[stack] != vec_bytes[stack] or reader_calls[stack] != vec_calls[stack]
        ]
        reader_only_bytes = sum(reader_bytes[stack] for stack in reader_only_keys)
        vec_only_bytes = sum(vec_bytes[stack] for stack in vec_only_keys)
        reader_only_calls = sum(reader_calls[stack] for stack in reader_only_keys)
        vec_only_calls = sum(vec_calls[stack] for stack in vec_only_keys)
        primary_bytes = sum(reader_bytes[stack] for stack in primary_keys)
        primary_calls = sum(reader_calls[stack] for stack in primary_keys)
        reader_only_non_primary = reader_only_bytes - primary_bytes
        route_setup_net = reader_only_non_primary - vec_only_bytes
        common_delta_bytes = sum(delta for _, delta, _ in common_rows)
        common_delta_calls = sum(delta for _, _, delta in common_rows)
        total_delta = sum(reader_bytes.values()) - sum(vec_bytes.values())
        accounted = primary_bytes + route_setup_net + common_delta_bytes
        require(reader_only_non_primary == vec_only_bytes, f"{variant}: route setup bytes do not cancel")
        require(accounted == total_delta, f"{variant}: exact stack accounting does not conserve")
        actual = {
            "reader_profiles": reader_profiles,
            "vec_profiles": vec_profiles,
            "reader_requested_bytes": sum(reader_bytes.values()),
            "vec_requested_bytes": sum(vec_bytes.values()),
            "reader_allocation_events": sum(reader_calls.values()),
            "vec_allocation_events": sum(vec_calls.values()),
            "total_delta_bytes": total_delta,
            "primary_read_limited_bytes": primary_bytes,
            "primary_read_limited_calls": primary_calls,
            "primary_stack_count": len(primary_keys),
            "reader_route_only_bytes": reader_only_bytes,
            "reader_route_only_calls": reader_only_calls,
            "reader_route_only_non_read_limited_bytes": reader_only_non_primary,
            "vec_route_only_bytes": vec_only_bytes,
            "vec_route_only_calls": vec_only_calls,
            "route_equivalent_setup_net_bytes": route_setup_net,
            "common_stack_delta_bytes": common_delta_bytes,
            "common_stack_delta_calls": common_delta_calls,
            "accounting_check": {
                "primary_plus_route_equivalent_plus_common": accounted,
                "total_delta": total_delta,
                "pass": True,
                "parent_stack_rows_added": False,
                "each_plus_event_assigned_once": True,
            },
            "union_stack_count": len(keys),
        }
        expected = data["owned_reader_vs_owned_vec"][variant]
        for field, value in actual.items():
            if field == "union_stack_count":
                continue
            compare_value(expected[field], value, f"{variant}.{field}")
        pair_checks[variant] = actual

    output = {
        "status": "verified",
        "summary_schema": data["schema"],
        "summary": summary_descriptor,
        "profiles": len(parsed),
        "unique_raw_sha256": len(raw_hashes),
        "raw_descriptors_verified_before_parse": True,
        "group_checks": group_checks,
        "pair_checks": pair_checks,
    }
    print(json.dumps(output, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
