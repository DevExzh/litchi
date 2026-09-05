#!/usr/bin/env python3
"""Run adversarial mutations against the independent 0428 report verifier.

The original report is opened read-only. Every candidate is written below a
temporary directory, and every rejection must be a structured INVALID error
rather than an interpreter traceback or arbitrary process crash.
"""

from __future__ import annotations

import argparse
import copy
import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Callable

U64_MAX = (1 << 64) - 1
CACHE_COUNTERS = {
    "hits", "cold_loads", "waiter_joins", "successful_loads", "failed_loads",
    "evictions", "bypasses", "oversized_bypasses", "allocation_bypasses",
    "budget_reservation_failures",
}
CACHE_FIELDS = CACHE_COUNTERS | {
    "retained_entries", "retained_bytes", "in_flight_loads", "budget_managed",
    "budget_memory_used", "budget_cache_reserved_bytes", "budget_memory_limit",
    "budget_input_bytes_used", "budget_input_bytes_limit",
    "budget_output_bytes_used", "budget_output_bytes_limit",
    "budget_work_used", "budget_work_limit", "budget_objects_used",
    "budget_objects_limit", "budget_catalog_reserved_objects",
    "budget_cache_reserved_objects",
}
BUDGET_FIELDS = {
    "memory_used", "memory_limit", "objects_used", "objects_limit",
    "input_bytes_used", "input_bytes_limit", "output_bytes_used",
    "output_bytes_limit", "work_used", "work_limit", "depth_used", "depth_limit",
}
Mutation = Callable[[dict[str, Any]], None]


def walk(value: Any) -> list[dict[str, Any]]:
    found: list[dict[str, Any]] = []
    if isinstance(value, dict):
        found.append(value)
        for child in value.values():
            found.extend(walk(child))
    elif isinstance(value, list):
        for child in value:
            found.extend(walk(child))
    return found


def cache_objects(report: dict[str, Any]) -> list[dict[str, Any]]:
    return [
        current for current in walk(report)
        if len({key.lower() for key in current} & CACHE_FIELDS) >= 4
    ]


def availability_objects(report: dict[str, Any]) -> list[dict[str, Any]]:
    return [
        current for current in walk(report)
        if any(key.lower() in {"availability", "state"} for key in current)
        and (
            any(key.lower() in CACHE_FIELDS for key in current)
            or any(key.lower() in {"diagnostics", "snapshot", "cache_diagnostics"} for key in current)
        )
    ]


def budget_objects(report: dict[str, Any]) -> list[dict[str, Any]]:
    return [
        current for current in walk(report)
        if "availability" not in {key.lower() for key in current}
        and "depth_used" in {key.lower() for key in current}
        and len({key.lower().removeprefix("budget_") for key in current} & BUDGET_FIELDS) >= 4
    ]


def first_key(report: dict[str, Any], names: set[str]) -> tuple[dict[str, Any], str] | None:
    for current in walk(report):
        for key in current:
            if key.lower() in names:
                return current, key
    return None


def mutate_missing_top_level(report: dict[str, Any]) -> None:
    report.pop("schema", None)


def mutate_bool_as_integer(report: dict[str, Any]) -> None:
    report["samples"] = True


def mutate_u64_overflow(report: dict[str, Any]) -> None:
    target = first_key(report, {"samples"})
    if target is None:
        raise ValueError("report has no bounded count")
    target[0][target[1]] = U64_MAX + 1


def mutate_missing_cache_gauge(report: dict[str, Any]) -> None:
    candidates = cache_objects(report)
    if not candidates:
        raise ValueError("report has no available cache object")
    for current in candidates:
        if "retained_bytes" in current:
            del current["retained_bytes"]
            return
        if "budget_cache_reserved_bytes" in current:
            del current["budget_cache_reserved_bytes"]
            return
    raise ValueError("cache objects have no removable gauge")


def mutate_unavailable_fabricated_zero(report: dict[str, Any]) -> None:
    for current in availability_objects(report):
        state_key = next(
            (key for key in current if key.lower() in {"availability", "state"}),
            None,
        )
        if state_key is None or str(current[state_key]).lower() not in {"unavailable", "omitted"}:
            continue
        current["retained_bytes"] = 0
        return
    candidates = cache_objects(report)
    if not candidates:
        raise ValueError("report has no cache object")
    current = candidates[0]
    current["availability"] = "unavailable"
    current["reason"] = "mutation"
    current["retained_bytes"] = 0


def mutate_counter_delta(report: dict[str, Any]) -> None:
    for current in walk(report):
        before = next((key for key in current if key.lower() in {"before", "start"}), None)
        after = next((key for key in current if key.lower() in {"after", "end"}), None)
        delta = next(
            (key for key in current if key.lower() in {"delta", "counter_delta", "checked_delta"}),
            None,
        )
        if before is None or after is None or delta is None:
            continue
        if isinstance(current[delta], dict):
            key = next((name for name in CACHE_COUNTERS if name in current[delta]), None)
            if key is not None:
                current[delta][key] = current[delta][key] + 1
                return
    candidates = cache_objects(report)
    if len(candidates) < 2:
        raise ValueError("report has no checked counter interval")
    # A wrapped event counter is represented as zero after a positive value.
    positive = next((current for current in candidates if any(
        isinstance(current.get(key), int) and not isinstance(current.get(key), bool)
        and current[key] > 0 for key in CACHE_COUNTERS
    )), None)
    if positive is None:
        raise ValueError("report has no positive event counter")
    key = next(key for key in CACHE_COUNTERS if positive.get(key, 0) > 0)
    positive[key] = 0


def mutate_scope(report: dict[str, Any]) -> None:
    target = first_key(report, {"cache_scope", "budget_scope", "source_io_scope"})
    if target is None:
        raise ValueError("report has no declared diagnostic scope")
    current, key = target
    current[key] = "wrong_scope"


def mutate_availability_state(report: dict[str, Any]) -> None:
    candidates = availability_objects(report)
    if not candidates:
        raise ValueError("report has no explicit availability object")
    current = candidates[0]
    key = next(
        key for key in current if key.lower() in {"availability", "state"}
    )
    current[key] = "bogus"


def mutate_missing_budget_gauge(report: dict[str, Any]) -> None:
    candidates = budget_objects(report)
    if not candidates:
        raise ValueError("report has no budget snapshot")
    for current in candidates:
        if "memory_used" in current:
            del current["memory_used"]
            return
        if "budget_memory_used" in current:
            del current["budget_memory_used"]
            return
    raise ValueError("budget snapshots have no Memory gauge")


def mutate_refusal_io(report: dict[str, Any]) -> None:
    scenario = report.get("scenario")
    if scenario == "one-under":
        target = first_key(report, {"payload_read_calls"})
        if target is None:
            raise ValueError("one-under report has no payload-read counter")
        if target[0][target[1]] is None:
            raise ValueError("one-under payload-read counter is unavailable")
        target[0][target[1]] = 1
        return
    if scenario in {"exact-admission", "oversized-bypass"}:
        target = first_key(report, {"payload_read_calls"})
        if target is None:
            raise ValueError("image report has no payload-read counter")
        if target[0][target[1]] is None:
            raise ValueError("image payload-read counter is unavailable")
        target[0][target[1]] = 0
        return
    if scenario == "pinned-eviction":
        target = first_key(report, {"pinned_bypass_verified"})
        if target is None:
            raise ValueError("pinned report has no bypass oracle")
        target[0][target[1]] = False
        return
    if scenario == "repeated-publication":
        target = first_key(report, {"typed_resource_refusal"})
        if target is None:
            raise ValueError("repeated report has no refusal oracle")
        target[0][target[1]] = False
        return
    target = first_key(
        report,
        {"payload_read_calls", "payload_reads", "target_payload_read_calls"},
    )
    if target is None:
        # Lifecycle rows have no selected-image refusal counter. Exercise the
        # same no-fabrication boundary on an unavailable source-read point.
        for current in walk(report):
            state = next(
                (key for key in current if key.lower() in {"availability", "state"}),
                None,
            )
            if state is not None and str(current[state]).lower() in {"unavailable", "omitted"}:
                for key in current:
                    if key.lower() in {"read_calls", "read_bytes", "source_read_calls", "source_read_bytes"}:
                        current[key] = 1
                        return
        raise ValueError("report has no payload-read or unavailable source-read oracle")
    current, key = target
    current[key] = 1


def mutate_expected_limit(report: dict[str, Any]) -> None:
    target = first_key(
        report,
        {
            "cache_max_bytes", "max_cache_bytes", "capacity_bytes",
            "memory_limit", "budget_memory_limit", "max_memory",
        },
    )
    if target is None:
        raise ValueError("report has no configured limit")
    current, key = target
    if not isinstance(current[key], int) or isinstance(current[key], bool):
        raise ValueError("configured limit is not numeric")
    current[key] = 0


def mutate_output_gate(report: dict[str, Any]) -> None:
    names = {
        "exact_output_verified", "exact_payload_verified", "typed_resource_refusal",
        "repeated_output_identities_verified", "eviction_verified",
    }
    target = None
    for current in walk(report):
        for key, value in current.items():
            if key.lower() in names and value is True:
                target = (current, key)
                break
        if target is not None:
            break
    if target is None:
        raise ValueError("report has no output gate")
    current, key = target
    current[key] = False


def mutate_output_hash(report: dict[str, Any]) -> None:
    if "expected_output_sha256" not in report:
        raise ValueError("report has no expected output digest")
    report["expected_output_sha256"] = "0" * 63 + "z"


def duplicate_json_key(original: bytes) -> bytes:
    text = original.decode("utf-8")
    opening = text.find("{")
    if opening < 0:
        raise ValueError("report JSON has no object root")
    return (
        text[: opening + 1]
        + '\n  "schema": "pptx_cache_retention_v1",\n'
        + '  "schema": "pptx_cache_retention_v1",'
        + text[opening + 1 :]
    ).encode("utf-8")


def nonfinite_json_constant(original: bytes) -> bytes:
    text = original.decode("utf-8")
    mutated, count = re.subn(
        r'("samples"\s*:\s*)[0-9]+', r"\1NaN", text, count=1
    )
    if count != 1:
        raise ValueError("report JSON has no samples number")
    return mutated.encode("utf-8")


def mutate_hidden_live_cache(report: dict[str, Any]) -> None:
    for phase in report["samples_raw"][0]["phases"]:
        cache = phase["source_cache"]
        if cache["availability"] == "available":
            for key in cache:
                cache[key] = None
            cache["availability"] = "unavailable"
            cache["unavailable_reason"] = "mutation hides an existing owner"
            return
    raise ValueError("no live source cache")


def mutate_caller_budget_mismatch(report: dict[str, Any]) -> None:
    for phase in report["samples_raw"][0]["phases"]:
        if phase["source_cache"]["availability"] == "available":
            phase["source_budget"]["memory_used"] += 1
            return
    raise ValueError("no live source budget")


MUTATIONS: tuple[tuple[str, Mutation], ...] = (
    ("hidden_live_cache", mutate_hidden_live_cache),
    ("caller_budget_mismatch", mutate_caller_budget_mismatch),
    ("missing_top_level_field", mutate_missing_top_level),
    ("bool_as_integer", mutate_bool_as_integer),
    ("u64_overflow", mutate_u64_overflow),
    ("missing_cache_gauge", mutate_missing_cache_gauge),
    ("unavailable_fabricated_zero", mutate_unavailable_fabricated_zero),
    ("checked_counter_delta", mutate_counter_delta),
    ("scope", mutate_scope),
    ("availability_state", mutate_availability_state),
    ("missing_budget_gauge", mutate_missing_budget_gauge),
    ("refusal_payload_io", mutate_refusal_io),
    ("configured_limit", mutate_expected_limit),
    ("output_gate", mutate_output_gate),
    ("output_hash_shape", mutate_output_hash),
)
RAW_MUTATIONS = ("duplicate_json_key", "nonfinite_json_constant")


def run_verifier(
    verifier: Path, report: Path, protocol: Path | None
) -> subprocess.CompletedProcess[str]:
    command = [sys.executable, str(verifier), "--report", str(report)]
    if protocol is not None:
        command.extend(["--protocol", str(protocol)])
    return subprocess.run(command, capture_output=True, text=True, check=False)


def structured_rejection(result: subprocess.CompletedProcess[str]) -> bool:
    return result.returncode != 0 and result.stderr.lstrip().startswith("INVALID:")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path", nargs="?", type=Path, help="cache-retention report JSON")
    parser.add_argument("--report", dest="report_option", type=Path, help="cache-retention report JSON")
    parser.add_argument(
        "--verifier",
        type=Path,
        default=Path(__file__).with_name("verify-report.py"),
    )
    parser.add_argument("--protocol", type=Path, default=None)
    args = parser.parse_args(argv)
    path = args.report_option or args.path
    if path is None:
        parser.error("a report path is required")
    try:
        original_bytes = path.read_bytes()
        original = json.loads(original_bytes.decode("utf-8"))
        if not isinstance(original, dict):
            raise ValueError("report root must be an object")
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as exc:
        print(f"INVALID INPUT: {exc}", file=sys.stderr)
        return 1

    baseline = run_verifier(args.verifier, path, args.protocol)
    if baseline.returncode != 0:
        print("BASELINE REJECTED", file=sys.stderr)
        if baseline.stderr:
            print(baseline.stderr.rstrip(), file=sys.stderr)
        return 1

    results: list[dict[str, Any]] = []
    with tempfile.TemporaryDirectory(prefix="litchi-0428-report-probes-") as temporary:
        directory = Path(temporary)
        for name, mutation in MUTATIONS:
            candidate = copy.deepcopy(original)
            try:
                mutation(candidate)
            except (KeyError, TypeError, ValueError, IndexError, StopIteration) as exc:
                print(f"MUTATION {name} COULD NOT BE APPLIED: {exc}", file=sys.stderr)
                return 1
            target = directory / f"{name}.json"
            target.write_text(json.dumps(candidate, indent=2) + "\n", encoding="utf-8")
            result = run_verifier(args.verifier, target, args.protocol)
            rejected = structured_rejection(result)
            results.append({"mutation": name, "rejected": rejected})
            if not rejected:
                print(f"MUTATION DID NOT PRODUCE STRUCTURED INVALID: {name}", file=sys.stderr)
                if result.stderr:
                    print(result.stderr.rstrip(), file=sys.stderr)
                return 1
        for name in RAW_MUTATIONS:
            target = directory / f"{name}.json"
            raw = (
                duplicate_json_key(original_bytes)
                if name == "duplicate_json_key"
                else nonfinite_json_constant(original_bytes)
            )
            target.write_bytes(raw)
            result = run_verifier(args.verifier, target, args.protocol)
            rejected = structured_rejection(result)
            results.append({"mutation": name, "rejected": rejected})
            if not rejected:
                print(f"MUTATION DID NOT PRODUCE STRUCTURED INVALID: {name}", file=sys.stderr)
                if result.stderr:
                    print(result.stderr.rstrip(), file=sys.stderr)
                return 1

    if path.read_bytes() != original_bytes:
        print("ORIGINAL REPORT CHANGED", file=sys.stderr)
        return 1
    print(json.dumps({"status": "pass", "mutations": results}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
