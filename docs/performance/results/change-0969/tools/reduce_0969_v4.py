#!/usr/bin/env python3
"""Offline 0969 reducer for retained runner packets.

No executable, Cargo command, source mutation, or workload is started.  A
complete noisy A/A stage yields ``readiness_fail``/uninterpretable; missing or
invalid A/A evidence yields ``incomplete``.  Formal packets are reduced only
after A/A passes.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import sys
from pathlib import Path
from typing import Any

from protocol_common_v4 import (
    DECISION_SCHEMA,
    ProtocolError,
    RUNNER_SCHEMA,
    check_packet,
    check_plan,
    decision_from_results,
    load_json,
    reduce_aa_packets,
    reduce_real,
    validate_go,
    write_json_exclusive,
)

SCHEMA = "litchi.performance.0969.reduction.v4"


def packet_paths(manifest_path: Path, manifest: dict[str, Any]) -> list[Path]:
    raw = manifest.get("packets")
    if not isinstance(raw, list):
        raise ProtocolError("capture manifest packets must be a list")
    result: list[Path] = []
    seen: set[Path] = set()
    for item in raw:
        if not isinstance(item, str):
            raise ProtocolError("capture manifest packet path is not a string")
        path = Path(item)
        if not path.is_absolute():
            path = manifest_path.parent / path
        path = path.resolve()
        if path in seen:
            raise ProtocolError(f"duplicate packet path in manifest: {path}")
        seen.add(path)
        result.append(path)
    return result


def load_packets(manifest_path: Path) -> tuple[dict[str, Any], list[dict[str, Any]], list[str]]:
    manifest = load_json(manifest_path)
    packets: list[dict[str, Any]] = []
    errors: list[str] = []
    for path in packet_paths(manifest_path, manifest):
        try:
            packets.append(load_json(path))
        except ProtocolError as exc:
            errors.append(str(exc))
    return manifest, packets, errors


def reduce_capture(plan_path: Path, manifest_path: Path) -> dict[str, Any]:
    # Reduction is admitted only for the final, capture-authorized binding.
    # This rechecks the source/review/binary/root-GO/A/A receipt custody before
    # any packet arithmetic is accepted.
    plan_path = plan_path.resolve()
    plan = check_plan(plan_path, allow_unbound=False, for_capture=True)
    manifest, packets, errors = load_packets(manifest_path.resolve())
    if manifest.get("schema") != RUNNER_SCHEMA or manifest.get("no_retry") is not True:
        raise ProtocolError("capture manifest schema/retry binding drifted")
    root_go = manifest.get("root_go") or {}
    if not isinstance(root_go, dict) or not isinstance(root_go.get("path"), str) or not isinstance(root_go.get("sha256"), str):
        raise ProtocolError("capture manifest root GO binding is missing")
    if root_go.get("sha256") != hashlib.sha256(Path(root_go["path"]).read_bytes()).hexdigest():
        raise ProtocolError("capture manifest root GO SHA mismatch")
    validate_go(
        Path(root_go["path"]),
        plan_path,
        plan,
        binary_roles=("native_baseline", "native_candidate", "allocator_baseline", "allocator_candidate"),
    )
    manifest_plan = manifest.get("plan")
    if manifest_plan and Path(manifest_plan).resolve() != plan_path:
        raise ProtocolError("capture manifest plan binding mismatch")
    checked: list[dict[str, Any]] = []
    for packet in packets:
        try:
            check_packet(packet, plan)
            checked.append(packet)
        except ProtocolError as exc:
            errors.append(str(exc))
    aa_packets = [packet for packet in checked if packet.get("role") == "aa_baseline"]
    aa = reduce_aa_packets(aa_packets, plan)
    result: dict[str, Any] = {
        "schema": SCHEMA,
        "plan": str(plan_path.resolve()),
        "manifest": str(manifest_path.resolve()),
        "packet_count": len(packets),
        "raw_packets_retained": True,
        "errors": errors,
        "aa_readiness": aa,
        "native": None,
        "allocator": None,
        "perf_decision": None,
    }
    if errors:
        # Packet validation failures are incomplete evidence.  A valid complete
        # A/A noise failure is handled separately below.
        result["perf_decision"] = {
            "schema": DECISION_SCHEMA,
            "status": "incomplete",
            "interpretation": "no_decision",
            "reason": "packet validation or custody failure",
        }
        result["status"] = "incomplete"
        return result
    if aa.get("status") != "pass":
        result["perf_decision"] = decision_from_results(aa, None, None)
        result["status"] = result["perf_decision"]["status"]
        return result

    real_packets = [packet for packet in checked if packet.get("role") in {"native_baseline", "native_candidate", "allocator_baseline", "allocator_candidate"}]
    try:
        native, allocator = reduce_real(real_packets, plan)
    except (ProtocolError, KeyError, TypeError, ValueError) as exc:
        result["errors"].append(f"formal packets incomplete: {exc}")
        result["perf_decision"] = {
            "schema": DECISION_SCHEMA,
            "status": "incomplete",
            "interpretation": "no_decision",
            "reason": "formal packets are missing or invalid",
        }
        result["status"] = "incomplete"
        return result
    result["native"] = native
    result["allocator"] = allocator
    result["perf_decision"] = decision_from_results(aa, native, allocator)
    result["status"] = result["perf_decision"]["status"]
    return result


def self_test() -> int:
    # The runner self-test owns packet-shape and A/A stage semantics.  This
    # check guards the decision classification independently of packet launch.
    from protocol_common_v4 import child_metric, decision_from_results, expected_cells

    empty = {"status": "incomplete"}
    if decision_from_results(empty, None, None)["status"] != "incomplete":
        raise SystemExit("missing A/A classification failed")
    if decision_from_results({"status": "readiness_fail"}, None, None)["status"] != "readiness_fail":
        raise SystemExit("complete noisy A/A classification failed")
    if decision_from_results({"status": "pass"}, None, None)["status"] != "incomplete":
        raise SystemExit("formal missing classification failed")
    allocation_packet = {
        "packet_id": "allocator-shape",
        "report": {"samples": [{"allocation": {"allocated_bytes": 100, "peak_live_bytes": 50, "allocation_calls": 4, "retained_bytes": 2}}]},
        "rss": {"kb": 7},
    }
    if child_metric(allocation_packet, "allocated_bytes") != 100.0 or child_metric(allocation_packet, "peak_live_bytes") != 50.0 or child_metric(allocation_packet, "rss_kb") != 7.0:
        raise SystemExit("allocator metric field mapping failed")
    native_row = {"noninferior": True, "p50": {"bootstrap": {"median": 1.0, "upper": 1.0}}}
    allocator_row = {
        "per_repeat_guard_pass": True,
        "allocation_calls_median_guard_pass": True,
        "allocated_bytes": {"bootstrap": {"median": 0.95}},
    }
    native = {cell: native_row for cell in expected_cells()}
    allocator = {cell: allocator_row for cell in expected_cells()}
    if decision_from_results({"status": "pass"}, native, allocator)["status"] != "qualified":
        raise SystemExit("benefit/safety qualification classification failed")
    allocator["docnohf/e1/zero"] = {**allocator_row, "per_repeat_guard_pass": False}
    if decision_from_results({"status": "pass"}, native, allocator)["status"] != "rejected":
        raise SystemExit("per-repeat allocator hard gate classification failed")
    print(json.dumps({"schema": SCHEMA, "status": "self_test_pass", "aa_missing": "incomplete", "aa_noise": "readiness_fail"}))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--plan", type=Path)
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    if args.self_test:
        return self_test()
    if args.plan is None or args.manifest is None or args.output is None:
        print("--plan, --manifest, and --output are required unless --self-test", file=sys.stderr)
        return 2
    try:
        result = reduce_capture(args.plan, args.manifest)
        write_json_exclusive(args.output.resolve(), result)
        print(json.dumps({"schema": SCHEMA, "status": result["status"], "output": str(args.output.resolve()), "decision": result["perf_decision"]}))
        return 0 if result["status"] in {"qualified", "rejected", "readiness_fail", "incomplete"} else 2
    except (OSError, ProtocolError, ValueError, KeyError, TypeError, json.JSONDecodeError) as exc:
        print(f"reducer rejected input: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
