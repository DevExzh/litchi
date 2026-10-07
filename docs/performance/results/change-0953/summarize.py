#!/usr/bin/env python3
"""Reduce immutable 0953 native-admission-v3 results to bounded diagnostics.

This is an offline reader. It never starts the benchmark, perf, or a decoder.
It verifies the four expected admission result files, checks sample/period
conservation, normalizes symbol offsets for exclusive-leaf ranking, and keeps
nested stack rows explicitly non-additive.
"""
from __future__ import annotations

import argparse
import collections
import hashlib
import json
import os
import re
import stat
import sys
from pathlib import Path
from typing import Any

SCHEMA = "litchi.performance.0953.native-admission.v3"
PLAN_SCHEMA = "litchi.performance.0953.native-publish-sampling-plan.v1"
EXPECTED_RAW_GATE = {
    "owner_samples_min": 100,
    "fully_resolved_owner_samples_min": 100,
    "owner_period_min": 100_000,
    "lost_records": 0,
}
OFFSET = re.compile(r"\+0x[0-9a-fA-F]+$")


class ReductionError(RuntimeError):
    pass


def require(value: bool, message: str) -> None:
    if not value:
        raise ReductionError(message)


def integer(value: Any, label: str) -> int:
    require(type(value) is int, f"{label} is not an integer")
    return value


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    st = path.stat()
    return {
        "path": str(path),
        "bytes": st.st_size,
        "sha256": sha(path),
        "mode": stat.S_IMODE(st.st_mode),
        "executable": bool(st.st_mode & stat.S_IXUSR),
    }


def json_file(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise ReductionError(f"invalid JSON {path}: {exc}") from exc


def desc_equal(actual: dict[str, Any], expected: dict[str, Any], label: str) -> None:
    require(isinstance(actual, dict), f"{label} descriptor missing")
    for key in ("bytes", "sha256", "mode", "executable"):
        require(actual.get(key) == expected.get(key), f"{label} descriptor {key} differs")


def normalize_symbol(symbol: Any) -> str:
    require(isinstance(symbol, str) and symbol, "ranking symbol is missing")
    return OFFSET.sub("", symbol)


def symbol_category(symbol: str) -> str:
    folded = symbol.casefold()
    if any(token in folded for token in ("zlib", "deflate", "compress", "zip")):
        return "compression"
    if "xml" in folded or "audit_source_part_replacement" in folded:
        return "xml_audit"
    return "other"


def add_rank(target: dict[str, dict[str, Any]], item: dict[str, Any], *, nested: bool) -> None:
    symbol = normalize_symbol(item.get("symbol"))
    period = integer(item.get("period"), f"{symbol} period")
    require(period >= 0, f"{symbol} period is negative")
    entry = target.setdefault(symbol, {"symbol": symbol, "period": 0,
                                       "rows": 0, "original_symbols": set()})
    entry["period"] += period
    entry["rows"] += 1
    entry["original_symbols"].add(item["symbol"])
    if nested:
        entry["nested"] = True


def finalize_ranking(values: dict[str, dict[str, Any]], limit: int | None = None) -> list[dict[str, Any]]:
    ordered = sorted(values.values(), key=lambda item: (-item["period"], item["symbol"]))
    result = []
    for item in ordered[:limit]:
        result.append({
            "symbol": item["symbol"],
            "period": item["period"],
            "rows": item["rows"],
            "original_symbols": sorted(item["original_symbols"]),
            **({"nested_nonadditive": True} if item.get("nested") else {}),
        })
    return result


def verify_plan(packet: Path, path: Path) -> dict[str, Any]:
    path = path.resolve()
    try:
        path.relative_to(packet.resolve())
    except ValueError as exc:
        raise ReductionError(f"bound plan is outside packet: {path}") from exc
    plan = json_file(path)
    require(plan.get("schema") == PLAN_SCHEMA, "plan schema differs")
    raw = plan.get("first_child_admission", {}).get("raw_gate")
    require(isinstance(raw, dict), "plan raw gate is missing")
    for key, expected in EXPECTED_RAW_GATE.items():
        require(type(raw.get(key)) is int and raw.get(key) == expected,
                f"plan raw gate {key} differs")
    return {"descriptor": descriptor(path), "raw_gate": raw}


def verify_result(packet: Path, path: Path, expected_index: int, plan: dict[str, Any]) -> dict[str, Any]:
    result = json_file(path)
    require(result.get("schema") == SCHEMA, f"{path.name} schema differs")
    require(result.get("status") == "pass", f"{path.name} was not admitted")
    row = result.get("row")
    require(isinstance(row, dict), f"{path.name} row is missing")
    require(row.get("index") == expected_index, f"{path.name} index differs")
    shape = row.get("shape")
    require(isinstance(shape, str) and shape, f"{path.name} shape is missing")
    repeat = integer(row.get("repeat"), f"{path.name} repeat")
    require(repeat in (0, 1), f"{path.name} repeat is unexpected")
    desc_equal(result.get("plan"), plan["descriptor"], f"{path.name} plan")
    binary = result.get("binary")
    runner = result.get("runner")
    require(isinstance(binary, dict) and isinstance(runner, dict), f"{path.name} custody is incomplete")
    if plan.get("binary") is None:
        plan["binary"] = binary
        plan["runner"] = runner
    else:
        desc_equal(binary, plan["binary"], f"{path.name} binary")
        desc_equal(runner, plan["runner"], f"{path.name} runner")
    profile = result.get("profile")
    require(isinstance(profile, dict) and profile.get("admitted") is True,
            f"{path.name} profile was not admitted")
    gate = profile.get("gate")
    require(isinstance(gate, dict) and all(gate.get(k) is True for k in (
        "lost_records_zero", "fully_resolved_owner_samples_min", "owner_period_min")),
        f"{path.name} raw profile gate failed")
    all_values = profile.get("all")
    owner = profile.get("owner_qualified")
    unknown = profile.get("unknown_or_unqualified")
    require(isinstance(all_values, dict) and isinstance(owner, dict) and isinstance(unknown, dict),
            f"{path.name} profile totals are incomplete")
    all_samples = integer(all_values.get("samples"), f"{path.name} all samples")
    all_period = integer(all_values.get("period"), f"{path.name} all period")
    owner_samples = integer(owner.get("samples"), f"{path.name} owner samples")
    owner_period = integer(owner.get("period"), f"{path.name} owner period")
    unknown_samples = integer(unknown.get("samples"), f"{path.name} unknown samples")
    unknown_period = integer(unknown.get("period"), f"{path.name} unknown period")
    require(all_samples == owner_samples + unknown_samples,
            f"{path.name} sample conservation failed")
    require(all_period == owner_period + unknown_period,
            f"{path.name} period conservation failed")
    require(integer(profile.get("lost_records"), f"{path.name} lost records") == 0,
            f"{path.name} lost records are nonzero")
    leaves = profile.get("exclusive_leaf_ranking")
    nested = profile.get("nested_compression_xml_rows")
    require(isinstance(leaves, list) and isinstance(nested, list),
            f"{path.name} stack rankings are incomplete")
    leaf_values: dict[str, dict[str, Any]] = {}
    nested_values: dict[str, dict[str, Any]] = {}
    for item in leaves:
        require(isinstance(item, dict), f"{path.name} leaf row malformed")
        add_rank(leaf_values, item, nested=False)
    for item in nested:
        require(isinstance(item, dict), f"{path.name} nested row malformed")
        add_rank(nested_values, item, nested=True)
    leaf_period = sum(item["period"] for item in leaf_values.values())
    require(leaf_period == owner_period, f"{path.name} exclusive leaf conservation failed")
    mmap_binding = result.get("mmap_binding")
    elf = result.get("elf")
    require(isinstance(mmap_binding, dict) and isinstance(elf, dict), f"{path.name} ELF binding missing")
    require(mmap_binding.get("build_id") == elf.get("build_id"), f"{path.name} build ID differs")
    return {
        "input": descriptor(path),
        "row": row,
        "child": result.get("child"),
        "all": {"samples": all_samples, "period": all_period},
        "owner_qualified": {"samples": owner_samples, "period": owner_period},
        "unknown_or_unqualified": {"samples": unknown_samples, "period": unknown_period},
        "fractions": {
            "owner_samples": owner_samples / all_samples if all_samples else None,
            "owner_period": owner_period / all_period if all_period else None,
            "unknown_samples": unknown_samples / all_samples if all_samples else None,
            "unknown_period": unknown_period / all_period if all_period else None,
        },
        "conservation": {"samples": True, "period": True, "exclusive_leaf_period": True},
        "exclusive_leaf_period": leaf_period,
        "nested_period_sum_nonadditive": sum(item["period"] for item in nested_values.values()),
        "exclusive_leaf_ranking": finalize_ranking(leaf_values),
        "nested_compression_xml_ranking": finalize_ranking(nested_values),
        "interesting_exclusive_leaf": finalize_ranking(
            {k: v for k, v in leaf_values.items() if symbol_category(k) != "other"}),
        "interesting_nested": finalize_ranking(
            {k: v for k, v in nested_values.items() if symbol_category(k) != "other"}),
    }


def write_exclusive(path: Path, value: Any) -> None:
    require(not path.exists(), f"refusing to overwrite output: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8", newline="") as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write("\n")
        stream.flush()
        os.fsync(stream.fileno())


def reduce_packet(packet: Path, output: Path, indexes: list[int]) -> dict[str, Any]:
    packet = packet.resolve()
    first_path = packet / "analysis" / f"native-admission-{indexes[0]:02d}-v1.json"
    first = json_file(first_path)
    bound_plan = first.get("plan")
    require(isinstance(bound_plan, dict), "first result plan descriptor is missing")
    candidates = []
    for candidate in sorted(packet.glob("plan-0953-v*.json")):
        try:
            actual = descriptor(candidate)
        except ReductionError:
            continue
        try:
            desc_equal(bound_plan, actual, f"candidate plan {candidate.name}")
        except ReductionError:
            continue
        candidates.append(candidate)
    require(len(candidates) == 1, f"expected one packet-local plan matching results, got {len(candidates)}")
    plan = verify_plan(packet, candidates[0])
    rows = []
    for index in indexes:
        matches = sorted((packet / "analysis").glob(f"native-admission-{index:02d}-v1.json"))
        require(len(matches) == 1, f"expected exactly one result for index {index}")
        rows.append(verify_result(packet, matches[0], index, plan))
    seen = {(row["row"]["shape"], row["row"]["repeat"]) for row in rows}
    require(len(seen) == len(rows), "duplicate shape/repeat result")
    plan_binary = plan["binary"]
    plan_runner = plan["runner"]
    aggregates = {"all": {"samples": 0, "period": 0},
                  "owner_qualified": {"samples": 0, "period": 0},
                  "unknown_or_unqualified": {"samples": 0, "period": 0}}
    leaf_values: dict[str, dict[str, Any]] = {}
    nested_values: dict[str, dict[str, Any]] = {}
    repeat_values: dict[int, dict[str, int]] = collections.defaultdict(
        lambda: {"rows": 0, "samples": 0, "period": 0, "owner_samples": 0, "owner_period": 0})
    for row in rows:
        for key in aggregates:
            for metric in ("samples", "period"):
                aggregates[key][metric] += row[key][metric]
        rep = repeat_values[row["row"]["repeat"]]
        rep["rows"] += 1
        rep["samples"] += row["all"]["samples"]
        rep["period"] += row["all"]["period"]
        rep["owner_samples"] += row["owner_qualified"]["samples"]
        rep["owner_period"] += row["owner_qualified"]["period"]
        for item in row["exclusive_leaf_ranking"]:
            add_rank(leaf_values, item, nested=False)
        for item in row["nested_compression_xml_ranking"]:
            add_rank(nested_values, item, nested=True)
    for values in repeat_values.values():
        values["owner_period_fraction"] = values["owner_period"] / values["period"] if values["period"] else None
        values["owner_sample_fraction"] = values["owner_samples"] / values["samples"] if values["samples"] else None
    total_period = aggregates["all"]["period"]
    total_samples = aggregates["all"]["samples"]
    owner_period = aggregates["owner_qualified"]["period"]
    owner_samples = aggregates["owner_qualified"]["samples"]
    return {
        "schema": "litchi.performance.0953.native-admission-summary.v1",
        "status": "pass",
        "scope": "Descriptive cycles:u owner attribution from four admitted children; no timing, speedup, allocation, RSS, or adoption claim.",
        "packet": str(packet),
        "plan": plan,
        "expected_indexes": indexes,
        "result_count": len(rows),
        "results": rows,
        "aggregates": {
            **aggregates,
            "fractions": {
                "owner_samples": owner_samples / total_samples if total_samples else None,
                "owner_period": owner_period / total_period if total_period else None,
                "unknown_samples": aggregates["unknown_or_unqualified"]["samples"] / total_samples if total_samples else None,
                "unknown_period": aggregates["unknown_or_unqualified"]["period"] / total_period if total_period else None,
            },
        },
        "per_repeat": {str(key): value for key, value in sorted(repeat_values.items())},
        "exclusive_leaf_ranking_normalized_function": finalize_ranking(leaf_values, limit=50),
        "nested_compression_xml_ranking_normalized_function": finalize_ranking(nested_values, limit=50),
        "diagnostic_targets": {
            "exclusive_leaf": [item for item in finalize_ranking(leaf_values, limit=50)
                               if symbol_category(item["symbol"]) != "other"],
            "nested_nonadditive": [item for item in finalize_ranking(nested_values, limit=50)
                                    if symbol_category(item["symbol"]) != "other"],
            "category_definition": {
                "compression": "symbol contains zlib, deflate, compress, or zip",
                "xml_audit": "symbol contains xml or audit_source_part_replacement",
                "other": "all remaining symbols",
            },
        },
        "conservation": {
            "all_samples": sum(row["all"]["samples"] for row in rows) == aggregates["all"]["samples"],
            "all_period": sum(row["all"]["period"] for row in rows) == aggregates["all"]["period"],
            "owner_period_equals_exclusive_leaf_period": owner_period == sum(item["period"] for item in leaf_values.values()),
            "nested_rows_are_nonadditive": True,
        },
        "source_results": [row["input"] for row in rows],
        "binary_binding": plan_binary,
        "runner_binding": plan_runner,
        "claim_limits": [
            "Owner fractions are cycles:u sample-period fractions only.",
            "Nested stack rows overlap and are not added to exclusive leaf totals.",
            "No native uninstrumented speedup, latency, allocation, RSS, or adoption inference.",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--packet", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--indexes", type=int, nargs="+", default=[0, 1, 2, 3])
    args = parser.parse_args()
    try:
        output = args.output if args.output.is_absolute() else args.packet / args.output
        result = reduce_packet(args.packet, output, args.indexes)
        write_exclusive(output.resolve(), result)
        print(json.dumps({"status": "pass", "output": str(output.resolve()),
                          "sha256": sha(output.resolve())}, sort_keys=True))
        return 0
    except (ReductionError, OSError, ValueError, KeyError) as exc:
        print(f"summarize_0953_v1: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
