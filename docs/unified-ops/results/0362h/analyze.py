#!/usr/bin/env python3
"""Fail-closed analysis for calibrated File qualification receipts.

The runner records one outer ABBA block per process invocation. This analyzer
checks raw receipts and host snapshots before reducing paired totals; embedded
probe JSON is never treated as a substitute for those artifacts.
"""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
import statistics
import sys

MIN_SAMPLE_NS = 150_000_000
SHA256 = set("0123456789abcdef")
CHANGED_SUFFIXES = {
    "set_commit", "set_publication", "batch_commit", "batch_publication",
    "import_commit", "import_publication", "merge_commit", "merge_publication",
}
EXACT_NOOP_SUFFIXES = {"open_get", "save", "repeated_save"}


def fail(message: str) -> None:
    raise SystemExit(f"0362h receipt error: {message}")


def load(path: Path):
    try:
        return json.loads(path.read_text())
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{path}: {error}")


def digest(path: Path) -> tuple[int, str]:
    try:
        size = path.stat().st_size
        hasher = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                hasher.update(block)
        return size, hasher.hexdigest()
    except (OSError, IsADirectoryError) as error:
        fail(f"cannot hash {path}: {error}")


def check_artifact(value, label: str) -> dict:
    if not isinstance(value, dict):
        fail(f"{label} artifact is not an object")
    path = value.get("path")
    expected_size = value.get("bytes")
    expected_hash = value.get("sha256")
    if not isinstance(path, str) or not isinstance(expected_size, int) or expected_size < 0:
        fail(f"{label} artifact has an invalid path/size")
    if not isinstance(expected_hash, str) or len(expected_hash) != 64 or set(expected_hash) - SHA256:
        fail(f"{label} artifact has an invalid hash")
    actual_size, actual_hash = digest(Path(path))
    if actual_size != expected_size or actual_hash != expected_hash:
        fail(f"{label} artifact changed after capture")
    return value


def check_endpoint(value, label: str) -> dict:
    if not isinstance(value, dict):
        fail(f"{label} provenance is missing")
    check_artifact(value.get("manifest"), f"{label}.manifest")
    for name in ("binary", "source", "lock", "probe", "fixture"):
        check_artifact(value.get(name), f"{label}.{name}")
    if not isinstance(value.get("source_commit"), str) or not value["source_commit"]:
        fail(f"{label}.source_commit is missing")
    dependency = value.get("dependency_projection_sha256")
    if not isinstance(dependency, str) or len(dependency) != 64 or set(dependency) - SHA256:
        fail(f"{label}.dependency_projection_sha256 is invalid")
    return value


def check_host(path: Path, label: str) -> None:
    value = load(path)
    if not isinstance(value, dict) or value.get("blockers") != []:
        fail(f"{label} has host blockers")
    if not isinstance(value.get("processes"), list) or not isinstance(value.get("allowed"), list):
        fail(f"{label} host snapshot is incomplete")


def check_receipt(
    value: dict,
    label: str,
    expected_route: str,
    expected_mode: str,
    expected_calls: int | None,
    expected_binary: dict,
) -> None:
    receipt_path = value.get("_receipt")
    snapshots = value.get("_host_snapshots")
    if not isinstance(receipt_path, str) or not isinstance(snapshots, list) or not snapshots:
        fail(f"{label} lacks raw receipt/host snapshot identities")
    receipt = load(Path(receipt_path))
    if not isinstance(receipt, dict) or receipt.get("schema") != "litchi_file_probe_receipt.v2":
        fail(f"{label} receipt schema is invalid")
    if receipt.get("route") != expected_route or receipt.get("mode") != expected_mode:
        fail(f"{label} receipt route/mode identity is invalid")
    if receipt.get("exit_code") != 0 or receipt.get("overlap") is not None or receipt.get("guard_error") is not None:
        fail(f"{label} receipt records a failed process/guard")
    check_artifact(receipt.get("binary"), f"{label}.receipt.binary")
    if receipt["binary"]["sha256"] != expected_binary["sha256"]:
        fail(f"{label} receipt used the wrong endpoint binary")
    command = receipt.get("command")
    if not isinstance(command, list) or not command or str(expected_binary["path"]) not in command:
        fail(f"{label} receipt command does not identify its binary")
    check_artifact(receipt.get("stdout"), f"{label}.receipt.stdout")
    check_artifact(receipt.get("stderr"), f"{label}.receipt.stderr")
    if receipt.get("label") != Path(receipt_path).name.removesuffix(".receipt.json"):
        fail(f"{label} receipt label does not match its path")
    if expected_mode == "measure" and ("--route" not in command or expected_route not in command):
        fail(f"{label} receipt command route is not recorded")
    receipt_snapshots = receipt.get("host_snapshots")
    if receipt_snapshots != snapshots:
        fail(f"{label} receipt and probe host identities differ")
    for index, snapshot in enumerate(snapshots):
        if not isinstance(snapshot, dict):
            fail(f"{label} host identity {index} is malformed")
        check_artifact(snapshot, f"{label}.host[{index}].artifact")
        check_host(Path(snapshot["path"]), f"{label}.host[{index}]")
    if expected_calls is not None and receipt.get("calls") != expected_calls:
        fail(f"{label} receipt call identity differs from shared calibration")


def integer(row: dict, key: str, label: str) -> int:
    value = row.get(key)
    if not isinstance(value, int) or value <= 0:
        fail(f"{label}.{key} is not a positive integer")
    return value


def hash_text(value: object, label: str) -> str:
    if not isinstance(value, str) or len(value) != 64 or set(value) - SHA256:
        fail(f"{label} is not a well-formed SHA-256 hash")
    return value


def phase(row: dict, label: str, calls: int) -> dict:
    samples = row.get("samples")
    if not isinstance(samples, list) or not samples or not all(isinstance(value, int) and value > 0 for value in samples):
        fail(f"{label}.samples is malformed")
    sample_calls = integer(row, "calls", label)
    if sample_calls != calls:
        fail(f"{label}.calls={sample_calls} differs from shared calls={calls}")
    total = integer(row, "total_ns", label)
    if total != sum(samples):
        fail(f"{label}.total_ns does not equal the raw samples")
    if min(samples) < MIN_SAMPLE_NS:
        fail(f"{label} has a sample below {MIN_SAMPLE_NS} ns")
    operation_count = integer(row, "operation_count", label)
    if operation_count % (calls * len(samples)) != 0:
        fail(f"{label}.operation_count is not a whole multiple of logical calls")
    operations_per_call = operation_count // (calls * len(samples))
    return {
        "samples_ns": samples,
        "total_ns": total,
        "calls": calls * len(samples),
        "operation_count": operation_count,
        "operations_per_call": operations_per_call,
        "per_operation_ns": total / operation_count,
    }


def semantic(value: dict, label: str, suffix: str) -> None:
    validation = value.get("validation")
    if not isinstance(validation, dict):
        fail(f"{label}.validation is missing")
    for key in (
        "semantic_state_valid", "source_state_preserved", "semantic_equal",
        "source_state_exact", "leaf_repeated_equal", "facade_repeated_equal",
        "exact_noop",
    ):
        if validation.get(key) is not True:
            fail(f"{label}.validation.{key} is not true")
    if suffix == "open_get":
        if validation.get("read_marker_equal") is not True:
            fail(f"{label}.validation.read_marker_equal is not true")
        leaf_marker = validation.get("leaf_read_marker_len")
        facade_marker = validation.get("facade_read_marker_len")
        if (
            not isinstance(leaf_marker, int)
            or leaf_marker < 0
            or not isinstance(facade_marker, int)
            or facade_marker < 0
            or leaf_marker != facade_marker
        ):
            fail(f"{label}.validation native/facade marker lengths differ")
    for side in ("leaf", "facade"):
        if not isinstance(validation.get(f"{side}_changed"), bool):
            fail(f"{label}.validation.{side}_changed is malformed")
        actual_changed = validation.get(f"{side}_output_changed")
        if not isinstance(actual_changed, bool):
            fail(f"{label}.validation.{side}_output_changed is malformed")
        if suffix in CHANGED_SUFFIXES and actual_changed is not True:
            fail(f"{label}.validation.{side} did not change output bytes")
        if suffix in CHANGED_SUFFIXES and validation[f"{side}_changed"] is not True:
            fail(f"{label}.validation.{side}_changed is false for a changed route")
    output_keys = (
        "source_leaf_output", "source_facade_output", "leaf_output", "facade_output",
        "leaf_repeat", "facade_repeat",
    )
    for key in output_keys:
        output = validation.get(key)
        if not isinstance(output, dict) or not isinstance(output.get("bytes"), int) or output["bytes"] <= 0:
            fail(f"{label}.validation.{key} is malformed")
        hash_text(output.get("sha256"), f"{label}.validation.{key}.sha256")
        hash_text(output.get("source_sha256"), f"{label}.validation.{key}.source_sha256")
    source_hashes = {validation[key]["source_sha256"] for key in output_keys}
    if len(source_hashes) != 1:
        fail(f"{label}.validation source identity differs between outputs")
    for owner in ("leaf", "facade"):
        before = hash_text(
            validation.get(f"source_{owner}_before_sha256"),
            f"{label}.validation.source_{owner}_before_sha256",
        )
        after = hash_text(
            validation.get(f"source_{owner}_after_sha256"),
            f"{label}.validation.source_{owner}_after_sha256",
        )
        if before != after:
            fail(f"{label}.validation source {owner} owner hash changed")
    for side in ("leaf", "facade"):
        first = validation[f"{side}_output"]
        repeated = validation[f"{side}_repeat"]
        if first["bytes"] != repeated["bytes"] or first["sha256"] != repeated["sha256"]:
            fail(f"{label}.validation.{side} repeated output is not byte-identical")
    if suffix in EXACT_NOOP_SUFFIXES and validation["exact_noop"] is not True:
        fail(f"{label}.validation exact no-op was not proven")


def main() -> None:
    if len(sys.argv) != 2:
        fail("usage: analyze.py OUTPUT_DIR")
    root = Path(sys.argv[1])
    manifest = load(root / "manifest.json")
    if not isinstance(manifest, dict) or manifest.get("schema") != "litchi_file_qualification.v3":
        fail("manifest schema is not v3")
    routes = manifest.get("routes")
    if not isinstance(routes, list) or not routes or not all(isinstance(route, str) for route in routes):
        fail("manifest.routes is empty")
    if len(set(routes)) != len(routes):
        fail("manifest.routes contains duplicates")
    rounds = manifest.get("rounds")
    if not isinstance(rounds, int) or rounds < 5:
        fail("fewer than five outer ABBA blocks were requested")
    if manifest.get("target_ns") != 500_000_000 or manifest.get("minimum_sample_ns") != MIN_SAMPLE_NS:
        fail("manifest calibration thresholds are not the qualified values")

    before_binary = check_artifact(manifest.get("before"), "before binary")
    after_binary = check_artifact(manifest.get("after"), "after binary")
    provenance = manifest.get("provenance")
    if not isinstance(provenance, dict):
        fail("manifest.provenance is missing")
    before_endpoint = check_endpoint(provenance.get("before"), "before")
    after_endpoint = check_endpoint(provenance.get("after"), "after")
    shared = provenance.get("shared")
    if not isinstance(shared, dict):
        fail("manifest.provenance.shared is missing")
    shared_names = {
        "dependency_projection_sha256": "dependency_projection_sha256",
        "lock_sha256": "lock",
        "probe_sha256": "probe",
        "fixture_sha256": "fixture",
    }
    for key, endpoint_key in shared_names.items():
        before_value = before_endpoint[endpoint_key] if endpoint_key == "dependency_projection_sha256" else before_endpoint[endpoint_key]["sha256"]
        after_value = after_endpoint[endpoint_key] if endpoint_key == "dependency_projection_sha256" else after_endpoint[endpoint_key]["sha256"]
        if shared.get(key) != before_value or shared.get(key) != after_value:
            fail(f"shared provenance {key} differs between endpoint manifests")
    if before_endpoint["binary"]["sha256"] != before_binary["sha256"] or after_endpoint["binary"]["sha256"] != after_binary["sha256"]:
        fail("endpoint manifests do not identify the selected binaries")

    seen_receipts: set[str] = set()
    route_lists = manifest.get("route_lists")
    if not isinstance(route_lists, dict):
        fail("manifest.route_lists is missing")
    for side, binary in (("before", before_binary), ("after", after_binary)):
        route_value = route_lists.get(side)
        if not isinstance(route_value, dict):
            fail(f"{side} route-list receipt is missing")
        check_receipt(route_value, f"{side}.route-list", "*", "list", None, binary)
        receipt_path = route_value.get("_receipt")
        if not isinstance(receipt_path, str) or receipt_path in seen_receipts:
            fail(f"{side}.route-list reuses a raw receipt")
        seen_receipts.add(receipt_path)
    calibration_rows = manifest.get("calibration")
    if not isinstance(calibration_rows, list) or {row.get("route") for row in calibration_rows if isinstance(row, dict)} != set(routes):
        fail("calibration does not cover every listed route")
    for calibration in calibration_rows:
        if not isinstance(calibration, dict):
            fail("calibration row is malformed")
        route = calibration.get("route")
        common = integer(calibration, "common_calls", str(route))
        rows = calibration.get("rows")
        if not isinstance(rows, list) or len(rows) != 2:
            fail(f"{route}: calibration must contain before and after")
        for row, binary, side in zip(rows, (before_binary, after_binary), ("before", "after")):
            if not isinstance(row, dict) or integer(row, "common_calls", str(route)) <= 0:
                fail(f"{route}: malformed {side} calibration")
            if row.get("minimum_elapsed_ns", 0) < MIN_SAMPLE_NS:
                fail(f"{route}: {side} calibration is below the minimum sample")
            check_receipt(row, f"{route}.{side}.calibration", str(route), "calibrate", None, binary)
            receipt_path = row.get("_receipt")
            if receipt_path in seen_receipts:
                fail(f"{route}.{side}.calibration reuses a raw receipt")
            seen_receipts.add(receipt_path)
        if common != max(row["common_calls"] for row in rows):
            fail(f"{route}: common call count is not the shared maximum")

    abba = manifest.get("abba")
    if not isinstance(abba, list) or len(abba) != len(routes):
        fail("abba does not contain one row per route")
    result = {"schema": "litchi_file_qualification_analysis.v3", "routes": {}, "manifest": manifest}
    for route_row in abba:
        if not isinstance(route_row, dict):
            fail("abba route row is malformed")
        route = route_row.get("route")
        if route not in routes or route in result["routes"]:
            fail(f"unknown or duplicate route {route!r}")
        calls = integer(route_row, "calls", str(route))
        blocks = route_row.get("blocks")
        if not isinstance(blocks, list) or len(blocks) < rounds:
            fail(f"{route}: fewer than {rounds} outer ABBA blocks")
        paired = []
        for block_index, block in enumerate(blocks):
            if not isinstance(block, dict) or block.get("block") != block_index:
                fail(f"{route} block {block_index}: identity is invalid")
            legs = block.get("legs")
            if not isinstance(legs, list) or len(legs) != 4 or [leg.get("side") for leg in legs] != ["before", "after", "after", "before"]:
                fail(f"{route} block {block_index}: outer order is invalid")
            normalized = {"before": [], "after": []}
            for leg_index, leg in enumerate(legs):
                value = leg.get("value")
                if not isinstance(value, dict) or value.get("route") != route or value.get("calls") != calls:
                    fail(f"{route} block {block_index} leg {leg_index}: identity/shared call mismatch")
                side = leg["side"]
                binary = before_binary if side == "before" else after_binary
                label = f"{route}.block{block_index}.leg{leg_index}"
                check_receipt(value, label, route, "measure", calls, binary)
                receipt_path = value["_receipt"]
                if receipt_path in seen_receipts:
                    fail(f"{label} reuses a raw receipt")
                seen_receipts.add(receipt_path)
                semantic(value, label, route.rsplit("/", 1)[-1])
                inner = value.get("blocks")
                if not isinstance(inner, list) or len(inner) != 1:
                    fail(f"{label}: each outer leg must contain exactly one inner block")
                inner_block = inner[0]
                if inner_block.get("order") != ["leaf", "facade", "facade", "leaf"]:
                    fail(f"{label}: invalid inner order")
                leaf = inner_block.get("leaf")
                facade = inner_block.get("facade")
                if not isinstance(leaf, list) or not isinstance(facade, list) or len(leaf) != 2 or len(facade) != 2:
                    fail(f"{label}: incomplete inner ABBA")
                leaf_rows = [phase(item, f"{label}.leaf{index}", calls) for index, item in enumerate(leaf)]
                facade_rows = [phase(item, f"{label}.facade{index}", calls) for index, item in enumerate(facade)]
                normalized[side].append({
                    "leaf_ns": sum(item["total_ns"] for item in leaf_rows),
                    "leaf_operations": sum(item["operation_count"] for item in leaf_rows),
                    "facade_ns": sum(item["total_ns"] for item in facade_rows),
                    "facade_operations": sum(item["operation_count"] for item in facade_rows),
                })
            before = normalized["before"]
            after = normalized["after"]
            before_leaf_ns = before[0]["leaf_ns"] + before[1]["leaf_ns"]
            before_leaf_ops = before[0]["leaf_operations"] + before[1]["leaf_operations"]
            after_leaf_ns = after[0]["leaf_ns"] + after[1]["leaf_ns"]
            after_leaf_ops = after[0]["leaf_operations"] + after[1]["leaf_operations"]
            before_facade_ns = before[0]["facade_ns"] + before[1]["facade_ns"]
            before_facade_ops = before[0]["facade_operations"] + before[1]["facade_operations"]
            after_facade_ns = after[0]["facade_ns"] + after[1]["facade_ns"]
            after_facade_ops = after[0]["facade_operations"] + after[1]["facade_operations"]
            paired.append({
                "block": block_index,
                "before_leaf_per_operation_ns": before_leaf_ns / before_leaf_ops,
                "after_leaf_per_operation_ns": after_leaf_ns / after_leaf_ops,
                "leaf_after_before": (after_leaf_ns / after_leaf_ops) / (before_leaf_ns / before_leaf_ops),
                "before_per_operation_ns": before_facade_ns / before_facade_ops,
                "after_per_operation_ns": after_facade_ns / after_facade_ops,
                "after_before": (after_facade_ns / after_facade_ops) / (before_facade_ns / before_facade_ops),
            })
        result["routes"][route] = {
            "blocks": paired,
            "after_before_median": statistics.median(item["after_before"] for item in paired),
            "after_before_range": [min(item["after_before"] for item in paired), max(item["after_before"] for item in paired)],
        }
    output = root / "analysis.json"
    output.write_text(json.dumps(result, indent=2) + "\n")
    print(json.dumps({"routes": len(routes), "analysis": str(output)}))


if __name__ == "__main__":
    main()
