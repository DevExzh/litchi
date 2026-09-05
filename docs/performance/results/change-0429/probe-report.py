#!/usr/bin/env python3
"""Exercise the independent 0429 verifier with adversarial report mutations.

The source report is never modified.  Every mutation is written to a private
temporary file and must produce the verifier's structured ``INVALID:`` result.
Both provider-lifecycle and native-image-lifecycle reports are supported.
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
    "budget_output_bytes_used", "budget_output_bytes_limit", "budget_work_used",
    "budget_work_limit", "budget_objects_used", "budget_objects_limit",
    "budget_catalog_reserved_objects", "budget_cache_reserved_objects",
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


def first_key(report: dict[str, Any], names: set[str]) -> tuple[dict[str, Any], str] | None:
    for current in walk(report):
        for key in current:
            if key.lower() in names:
                return current, key
    return None


def cache_objects(report: dict[str, Any]) -> list[dict[str, Any]]:
    return [
        current for current in walk(report)
        if len({key.lower() for key in current} & CACHE_FIELDS) >= 5
    ]


def read_objects(report: dict[str, Any]) -> list[dict[str, Any]]:
    return [
        current for current in walk(report)
        if "availability" in current and "logical_calls" in current and "returned_bytes" in current
    ]


def mutate_missing_top_field(report: dict[str, Any]) -> None:
    report.pop("schema", None)


def mutate_bool_as_integer(report: dict[str, Any]) -> None:
    report["samples"] = True


def mutate_u64_overflow(report: dict[str, Any]) -> None:
    report["samples"] = U64_MAX + 1


def mutate_provider_config_mislabel(report: dict[str, Any]) -> None:
    config = report["provider_config"]
    provider = report["provider"]
    config["provider"] = "file" if provider != "file" else "bytes"


def mutate_duration_arithmetic(report: dict[str, Any]) -> None:
    target = first_key(report, {"api_sum_ns"})
    if target is None:
        raise ValueError("report has no API duration sum")
    current, key = target
    if not isinstance(current[key], int) or isinstance(current[key], bool):
        raise ValueError("API duration sum is not an integer")
    current[key] += 1


def mutate_cap_bytes(report: dict[str, Any]) -> None:
    config = report["provider_config"]
    if report["provider"] == "range":
        config["max_range_bytes"] = 0
        if "adapter_max_range_bytes" in config:
            config["adapter_max_range_bytes"] = 0
    else:
        config["max_range_bytes"] = 1


def mutate_live_cache_missing(report: dict[str, Any]) -> None:
    for current in cache_objects(report):
        if current.get("availability") == "available":
            current["availability"] = "unavailable"
            current["unavailable_reason"] = "mutation removed a live cache owner"
            return
    raise ValueError("report has no available cache owner")


def mutate_budget_mismatch(report: dict[str, Any]) -> None:
    rows = report.get("samples_raw")
    if not isinstance(rows, list):
        raise ValueError("report has no lifecycle rows")
    for row in rows:
        phases = row.get("phases") if isinstance(row, dict) else None
        if not isinstance(phases, list):
            continue
        for phase in phases:
            if not isinstance(phase, dict):
                continue
            cache = phase.get("source_cache")
            budget = phase.get("source_budget")
            if isinstance(cache, dict) and cache.get("availability") == "available" and isinstance(budget, dict):
                budget["memory_used"] = budget.get("memory_used", 0) + 1
                return
    raise ValueError("report has no live source cache and caller budget")


def mutate_payload_output_oracle(report: dict[str, Any]) -> None:
    if report.get("schema") == "pptx_provider_lifecycle_v1":
        report["expected_output_bytes"] += 1
    else:
        report["payload_oracle"]["payload_bytes"] += 1


def mutate_revision(report: dict[str, Any]) -> None:
    report["source_revision"] = "z" * 40


def mutate_phase_order(report: dict[str, Any]) -> None:
    phases = report.get("phases")
    if not isinstance(phases, list) or len(phases) < 2:
        raise ValueError("report has fewer than two phase descriptions")
    phases[0], phases[1] = phases[1], phases[0]


def mutate_checked_delta(report: dict[str, Any]) -> None:
    for current in walk(report):
        delta = current.get("delta")
        if not isinstance(delta, dict) or not isinstance(delta.get("logical_calls"), int):
            continue
        delta["logical_calls"] += 1
        return
    for current in walk(report):
        event = current.get("event_delta")
        if isinstance(event, dict) and isinstance(event.get("hits"), int):
            event["hits"] += 1
            return
    raise ValueError("report has no checked cache or read delta")


def mutate_request_histogram(report: dict[str, Any]) -> None:
    # Baseline and dropped-owner points deliberately serialize this field as
    # null.  Mutate the first available, nonempty cumulative histogram so the
    # candidate remains a meaningful counter inconsistency for every lane.
    for current in walk(report):
        counts = current.get("request_size_counts")
        if not isinstance(counts, list) or not counts:
            continue
        if not any(isinstance(item, int) and not isinstance(item, bool) and item > 0 for item in counts):
            continue
        counts[0] += 1
        return
    raise ValueError("report has no available nonempty request-size histogram")


def mutate_unavailable_fabricated_zero(report: dict[str, Any]) -> None:
    for current in cache_objects(report):
        if current.get("availability") == "unavailable":
            current["retained_bytes"] = 0
            return
    raise ValueError("report has no unavailable cache point")


def mutate_missing_cache_gauge(report: dict[str, Any]) -> None:
    for current in cache_objects(report):
        if current.get("availability") == "available" and "retained_bytes" in current:
            del current["retained_bytes"]
            return
    raise ValueError("report has no available cache gauge")


def mutate_output_gate(report: dict[str, Any]) -> None:
    names = {"exact_output_verified", "descriptor_verified"}
    target = first_key(report, names)
    if target is None:
        raise ValueError("report has no row output gate")
    target[0][target[1]] = False


def mutate_duration_scope(report: dict[str, Any]) -> None:
    report["timing_scope"] = "wrong_scope"


def mutate_source_read_availability(report: dict[str, Any]) -> None:
    for current in read_objects(report):
        if current.get("availability") == "available":
            current["availability"] = "unavailable"
            current["unavailable_reason"] = "mutation hid caller source"
            return
    raise ValueError("report has no available source-read point")


def mutate_oracle_identity(report: dict[str, Any]) -> None:
    if report.get("schema") == "pptx_provider_lifecycle_v1":
        digest = report["corpus_manifest"]["archive_sha256"]
        report["corpus_manifest"]["archive_sha256"] = ("0" if digest[0] != "0" else "1") + digest[1:]
    else:
        report["payload_oracle"]["relationship_id"] = "rId-mutated"


def mutate_descriptor_oracle(report: dict[str, Any]) -> None:
    if report.get("schema") == "pptx_native_image_lifecycle_v1":
        report["payload_oracle"]["shape_id"] += 1
    else:
        report["expected_output_bytes"] += 1


def mutate_output_hash(report: dict[str, Any]) -> None:
    target = first_key(report, {"output_sha256", "sha256"})
    if target is None:
        raise ValueError("report has no output digest")
    target[0][target[1]] = "0" * 64


def mutate_owner_gate(report: dict[str, Any]) -> None:
    if report.get("schema") == "pptx_native_image_lifecycle_v1":
        for phase in report["phases"]:
            if phase["label"] == "drop_slide":
                phase["live_owners"] = "returned payload was discarded"
                return
        raise ValueError("native report has no drop-slide phase")
    report["destination_editor_consumed_during_publish"] = False


def mutate_bool_counter(report: dict[str, Any]) -> None:
    for current in cache_objects(report):
        if "hits" in current:
            current["hits"] = True
            return
    raise ValueError("report has no cache counter")


def mutate_unknown_field(report: dict[str, Any]) -> None:
    report["unexpected_probe_field"] = 1


MUTATIONS: tuple[tuple[str, Mutation], ...] = (
    ("missing_top_field", mutate_missing_top_field),
    ("bool_as_integer", mutate_bool_as_integer),
    ("u64_overflow", mutate_u64_overflow),
    ("provider_config_mislabel", mutate_provider_config_mislabel),
    ("duration_arithmetic", mutate_duration_arithmetic),
    ("cap_bytes", mutate_cap_bytes),
    ("live_cache_missing", mutate_live_cache_missing),
    ("budget_mismatch", mutate_budget_mismatch),
    ("payload_output_oracle", mutate_payload_output_oracle),
    ("revision_invalid", mutate_revision),
    ("phase_order", mutate_phase_order),
    ("checked_delta", mutate_checked_delta),
    ("request_histogram", mutate_request_histogram),
    ("unavailable_fabricated_zero", mutate_unavailable_fabricated_zero),
    ("missing_cache_gauge", mutate_missing_cache_gauge),
    ("output_gate", mutate_output_gate),
    ("duration_scope", mutate_duration_scope),
    ("source_read_availability", mutate_source_read_availability),
    ("oracle_identity", mutate_oracle_identity),
    ("descriptor_oracle", mutate_descriptor_oracle),
    ("output_hash", mutate_output_hash),
    ("owner_lifetime_gate", mutate_owner_gate),
    ("bool_cache_counter", mutate_bool_counter),
    ("unknown_field", mutate_unknown_field),
)


def duplicate_json_key(original: bytes) -> bytes:
    text = original.decode("utf-8")
    opening = text.find("{")
    if opening < 0:
        raise ValueError("report JSON has no object root")
    schema = re.search(r'"schema"\s*:\s*"([^"]+)"', text)
    value = schema.group(1) if schema else "pptx_provider_lifecycle_v1"
    insertion = f'\n  "schema": {json.dumps(value)},\n  "schema": {json.dumps(value)},'
    return (text[: opening + 1] + insertion + text[opening + 1 :]).encode("utf-8")


def nonfinite_json_constant(original: bytes) -> bytes:
    text = original.decode("utf-8")
    mutated, count = re.subn(r'("samples"\s*:\s*)[0-9]+', r"\1NaN", text, count=1)
    if count != 1:
        raise ValueError("report JSON has no samples number")
    return mutated.encode("utf-8")


def run_verifier(verifier: Path, report: Path, protocol: Path | None) -> subprocess.CompletedProcess[str]:
    command = [sys.executable, str(verifier), "--report", str(report)]
    if protocol is not None:
        command.extend(["--protocol", str(protocol)])
    return subprocess.run(command, capture_output=True, text=True, check=False)


def structured_rejection(result: subprocess.CompletedProcess[str]) -> bool:
    return result.returncode != 0 and result.stderr.lstrip().startswith("INVALID:")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path", nargs="?", type=Path)
    parser.add_argument("--report", dest="report_option", type=Path)
    parser.add_argument("--verifier", type=Path, default=Path(__file__).with_name("verify-report.py"))
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
    with tempfile.TemporaryDirectory(prefix="litchi-0429-report-probes-") as temporary:
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
        for name, raw_mutation in (
            ("duplicate_json_key", duplicate_json_key),
            ("nonfinite_json_constant", nonfinite_json_constant),
        ):
            try:
                mutated_bytes = raw_mutation(original_bytes)
            except (UnicodeError, ValueError) as exc:
                print(f"RAW MUTATION {name} COULD NOT BE APPLIED: {exc}", file=sys.stderr)
                return 1
            target = directory / f"{name}.json"
            target.write_bytes(mutated_bytes)
            result = run_verifier(args.verifier, target, args.protocol)
            rejected = structured_rejection(result)
            results.append({"mutation": name, "rejected": rejected})
            if not rejected:
                print(f"RAW MUTATION DID NOT PRODUCE STRUCTURED INVALID: {name}", file=sys.stderr)
                if result.stderr:
                    print(result.stderr.rstrip(), file=sys.stderr)
                return 1

    print(json.dumps({"status": "probes-valid", "mutations": results}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
