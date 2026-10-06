#!/usr/bin/env python3
"""Independent offline checker for the 0933 allocator reports.

This script reads an already completed external packet. It never starts a
workload, reads a binary, or changes packet files. It independently checks
report identity, retention-cap/topology facts, allocator gauge conservation,
and construction deltas at the after_snapshot_drop checkpoint. It emits a
small immutable summary when --output is supplied.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import statistics
import sys
from pathlib import Path

MAX_JSON = 1 << 20
PHASES = ("after_snapshot_drop", "commits_retained", "after_release_all_but_last", "after_release_last", "after_commit_drop")
MONOTONIC = ("allocation_calls", "deallocation_calls", "reallocation_calls", "failed_allocation_calls", "allocated_bytes", "deallocated_bytes", "peak_live_bytes")
SECTIONS = (16, 8192)
OWNERSHIPS = ("shared_snapshot_clones", "independent_sources")
# Frozen expected identity fixtures copied from the retained 0931 exact/c08
# reports. They are per inventory size because the synthetic allocator input
# and emitted output differ between 16 and 8192 sections.
ALLOCATOR_FIXTURES = {
    16: {
        "source_archive_bytes": 1052,
        "source_archive_sha256": "ac0e4a784d75f16c3c3d7f203e6f2a92b505b28d7db965b7a4ff54ad61f52726",
        "expected_output_bytes": [1078] * 8,
        "expected_output_sha256": [
            "cb2a24195fcf4a7c47644cf38714fe9320a84c576176451a4672c49129d1c0a1",
            "908871772e7902e5c57c1735cba761c0e8e5290ff442bc0618ece118b7ec6fb5",
            "51bd0f53a939a2e0815a70d275a903ea74f0eb1934dbe5896902b5190feb87d5",
            "fe6cde5e7713a2dfe624a3ade7844fc230dd20c70e99b11e0975a95cd8c0b2b0",
            "cb2a24195fcf4a7c47644cf38714fe9320a84c576176451a4672c49129d1c0a1",
            "908871772e7902e5c57c1735cba761c0e8e5290ff442bc0618ece118b7ec6fb5",
            "51bd0f53a939a2e0815a70d275a903ea74f0eb1934dbe5896902b5190feb87d5",
            "fe6cde5e7713a2dfe624a3ade7844fc230dd20c70e99b11e0975a95cd8c0b2b0",
        ],
        "required_retained_base_bytes": 7280,
    },
    8192: {
        "source_archive_bytes": 25498,
        "source_archive_sha256": "92b5b61c00568e1b104b8096d1ce06ca5bc247cacf037561b3752e3657bd67f3",
        "expected_output_bytes": [25545] * 8,
        "expected_output_sha256": [
            "e087fe9e71eaabad6adafb630f0165b9be6f8a0106018e8f3b2fc6e536471caf",
            "21311b30407f4169981e49975dee475a1424edfd53b0314b2949f68fa916ce0b",
            "338c76b4d93da473c2d2251044cc419528c04a71322f1cb150d12dbe42fa8049",
            "4823bf871699999c7d646080da83dd080798ba0b1fff4a922d41aee9903a4546",
            "e087fe9e71eaabad6adafb630f0165b9be6f8a0106018e8f3b2fc6e536471caf",
            "21311b30407f4169981e49975dee475a1424edfd53b0314b2949f68fa916ce0b",
            "338c76b4d93da473c2d2251044cc419528c04a71322f1cb150d12dbe42fa8049",
            "4823bf871699999c7d646080da83dd080798ba0b1fff4a922d41aee9903a4546",
        ],
        "required_retained_base_bytes": 3580192,
    },
}

class CheckError(RuntimeError):
    pass

def need(ok: bool, message: str) -> None:
    if not ok:
        raise CheckError(message)

def sha256_path(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()

def read_json(path: Path) -> dict:
    need(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    need(path.stat().st_size < MAX_JSON, f"JSON >= 1 MiB: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise CheckError(f"cannot parse {path}: {exc}") from exc
    need(isinstance(value, dict), f"JSON root is not an object: {path}")
    return value

def integer(value, label: str, *, nonnegative: bool = True) -> int:
    need(type(value) is int and (value >= 0 if nonnegative else value > 0), f"{label}: invalid integer {value!r}")
    return value

def gauge(value: dict, label: str) -> None:
    need(isinstance(value, dict), f"{label}: missing gauge")
    for name in ("allocation_calls", "deallocation_calls", "reallocation_calls", "failed_allocation_calls", "allocated_bytes", "deallocated_bytes", "live_bytes", "peak_live_bytes"):
        integer(value.get(name), f"{label}.{name}")
    need(value["allocated_bytes"] >= value["deallocated_bytes"], f"{label}: allocator underflow")
    need(value["live_bytes"] == value["allocated_bytes"] - value["deallocated_bytes"], f"{label}: live-byte balance mismatch")
    need(value["peak_live_bytes"] >= value["live_bytes"], f"{label}: peak below live bytes")
    need(value.get("overflowed") is False and value.get("observer_invalid") is False, f"{label}: invalid/overflowed observer")
    region = value.get("region_peak_live_bytes")
    need(region is None or (type(region) is int and region >= 0), f"{label}: invalid region peak")

def report_facts(row: dict, report: dict, label: str) -> dict:
    need(report.get("schema") == "docx_section_retention_v1", f"{label}: schema mismatch")
    for field, expected in (("section_count", row["sections"]), ("concurrent_commits", row["concurrent_commits"]), ("retention_policy", row["policy"]), ("ownership_mode", row["ownership"]), ("source_revision", row.get("source_revision"))):
        if expected is not None:
            need(report.get(field) == expected, f"{label}: {field} mismatch")
    fixture = ALLOCATOR_FIXTURES.get(row["sections"])
    need(isinstance(fixture, dict), f"{label}: no frozen identity fixture for section count")
    need(report.get("source_archive_bytes") == fixture["source_archive_bytes"],
         f"{label}: source archive size differs from frozen fixture")
    need(report.get("source_archive_sha256") == fixture["source_archive_sha256"],
         f"{label}: source archive hash differs from frozen fixture")
    need(report.get("expected_output_bytes") == fixture["expected_output_bytes"],
         f"{label}: expected output sizes differ from frozen fixture")
    need(report.get("expected_output_sha256") == fixture["expected_output_sha256"],
         f"{label}: expected output hashes differ from frozen fixture")
    admission = report.get("retention_admission")
    need(isinstance(admission, dict), f"{label}: missing retention admission")
    required = integer(admission.get("required_retained_base_bytes"), f"{label}.required_retained_base_bytes", nonnegative=False)
    need(required == fixture["required_retained_base_bytes"],
         f"{label}: retained-base requirement differs from frozen fixture")
    need(admission.get("boundary_checks_verified") is True, f"{label}: boundary checks not verified")
    need(admission.get("probe_ceiling_bytes") == 64 * 1024 * 1024, f"{label}: probe ceiling mismatch")
    need(admission.get("zero_cap_bytes") == 0 and admission.get("zero_retained_base_bytes") == 0, f"{label}: zero cap mismatch")
    need(admission.get("exact_cap_bytes") == required and admission.get("exact_retained_base_bytes") == required, f"{label}: exact cap mismatch")
    need(admission.get("one_under_cap_bytes") == required - 1 and admission.get("one_under_retained_base_bytes") == 0, f"{label}: one-under cap mismatch")
    default_retained = required if required <= 1_048_576 else 0
    need(admission.get("default_cap_bytes") == 1_048_576 and admission.get("default_retained_base_bytes") == default_retained, f"{label}: default cap mismatch")
    cap = report.get("retention_ceiling_bytes")
    need(cap == required, f"{label}: selected cap mismatch")
    need(report.get("default_retention_ceiling_bytes") == 1_048_576, f"{label}: default cap constant mismatch")
    scenarios = report.get("scenarios")
    need(isinstance(scenarios, list) and len(scenarios) == 1, f"{label}: scenario cardinality mismatch")
    scenario = scenarios[0]
    need(scenario.get("ownership_mode") == row["ownership"] and scenario.get("section_count") == row["sections"] and scenario.get("concurrent_commits") == row["concurrent_commits"] and scenario.get("retention_ceiling_bytes") == cap, f"{label}: scenario identity mismatch")
    each = scenario.get("retained_base_bytes_each")
    need(isinstance(each, list) and len(each) == row["concurrent_commits"], f"{label}: retained vector shape mismatch")
    need(all(type(v) is int and v >= 0 for v in each), f"{label}: retained vector values")
    need(each == [required] * row["concurrent_commits"], f"{label}: retained vector mismatch")
    need(scenario.get("logical_retained_base_bytes_sum") == sum(each), f"{label}: retained sum mismatch")
    need(scenario.get("retained_commit_count") == sum(v != 0 for v in each), f"{label}: retained count mismatch")
    pair_count = row["concurrent_commits"] * (row["concurrent_commits"] - 1) // 2
    expected_shared = pair_count if row["ownership"] == "shared_snapshot_clones" else 0
    need(scenario.get("source_inventory_pair_count") == pair_count and scenario.get("source_inventory_shared_pairs") == expected_shared, f"{label}: source topology mismatch")
    for field in ("source_inventory_sharing_verified", "measured_commit_shape_verified", "measured_output_verified", "measured_inverse_verified", "preflight_output_verified", "preflight_inverse_verified"):
        need(scenario.get(field) is True, f"{label}: {field} not verified")
    expected_validation = 1 if row["ownership"] == "shared_snapshot_clones" else row["concurrent_commits"]
    need(scenario.get("measured_commit_validation_count") == expected_validation, f"{label}: validation count mismatch")
    expected_scope = "one_original_lineage_commit_plus_all_preflight_commits" if row["ownership"] == "shared_snapshot_clones" else "every_distinct_commit"
    need(scenario.get("measured_validation_scope") == expected_scope, f"{label}: validation scope mismatch")
    gauges = scenario.get("gauges")
    need(isinstance(gauges, dict), f"{label}: gauges missing")
    base = gauges.get("baseline")
    gauge(base, f"{label}.baseline")
    sequence = [base]
    for phase in PHASES:
        entry = gauges.get(phase)
        need(isinstance(entry, dict), f"{label}.{phase}: phase missing")
        absolute = entry.get("absolute")
        gauge(absolute, f"{label}.{phase}.absolute")
        if phase == "commits_retained":
            need(absolute.get("region_peak_live_bytes") is None, f"{label}.{phase}: held checkpoint has a region peak")
        else:
            region = absolute.get("region_peak_live_bytes")
            need(type(region) is int and region >= absolute["live_bytes"], f"{label}.{phase}: endpoint region peak invalid")
        need(entry.get("live_bytes_delta_from_baseline") == absolute["live_bytes"] - base["live_bytes"], f"{label}.{phase}: live delta mismatch")
        sequence.append(absolute)
    for index, (prev, current) in enumerate(zip(sequence, sequence[1:])):
        for field in MONOTONIC:
            need(current[field] >= prev[field], f"{label}: {field} regressed at transition {index}")
        expected_live = prev["live_bytes"] + (current["allocated_bytes"] - prev["allocated_bytes"]) - (current["deallocated_bytes"] - prev["deallocated_bytes"])
        need(current["live_bytes"] == expected_live, f"{label}: conservation failed at transition {index}")
    need(gauges.get("returned_to_baseline") is True, f"{label}: returned_to_baseline false")
    need(sequence[-1]["live_bytes"] == base["live_bytes"], f"{label}: final live bytes differ from baseline")
    def delta(phase: str, field: str) -> int:
        return gauges[phase]["absolute"][field] - base[field]
    return {
        "allocated_bytes_delta_after_snapshot_drop": delta("after_snapshot_drop", "allocated_bytes"),
        "allocation_calls_delta_after_snapshot_drop": delta("after_snapshot_drop", "allocation_calls"),
        "allocated_bytes_delta_commits_retained": delta("commits_retained", "allocated_bytes"),
        "allocation_calls_delta_commits_retained": delta("commits_retained", "allocation_calls"),
        "live_bytes_delta_after_snapshot_drop": delta("after_snapshot_drop", "live_bytes"),
        "release_live_bytes_delta": delta("after_commit_drop", "live_bytes"),
        "logical_retained_base_bytes": scenario["logical_retained_base_bytes_sum"],
        "retained_commit_count": scenario["retained_commit_count"],
        "required_retained_base_bytes": required,
        "fixture_source_archive_bytes": fixture["source_archive_bytes"],
        "fixture_source_archive_sha256": fixture["source_archive_sha256"],
        "fixture_expected_output_bytes": fixture["expected_output_bytes"],
        "fixture_expected_output_sha256": fixture["expected_output_sha256"],
    }

def local_descriptor(packet: Path, value: dict, label: str) -> Path:
    need(isinstance(value, dict) and isinstance(value.get("path"), str), f"{label}: descriptor missing")
    raw = Path(value["path"])
    need(not raw.is_absolute() and raw.parts and all(part not in ("", ".", "..") for part in raw.parts),
         f"{label}: unsafe packet path")
    root = packet.resolve()
    path = (root / raw).resolve()
    need(path.is_file() and not path.is_symlink(), f"{label}: packet file missing")
    need(root in path.parents, f"{label}: packet path escaped root")
    need(path.stat().st_size == value.get("bytes") and sha256_path(path) == value.get("sha256"),
         f"{label}: packet descriptor changed")
    return path

def child_environment(manifest: dict, row: dict) -> dict:
    env = dict(manifest["child_environment"])
    if row["lane"] == "normal":
        env["LITCHI_DOCX_INGRESS_ROUTE"] = row["route"]
    return env

def check_terminal_identity(packet: Path, manifest: dict, row: dict, child: dict) -> Path:
    terminal_path = local_descriptor(packet, child["terminal"], f"terminal {row['id']}")
    terminal = read_json(terminal_path)
    need(terminal.get("schedule") == row, f"terminal schedule changed: {row['id']}")
    need(terminal.get("exit_code") == 0 and terminal.get("deadline_exceeded") is False
         and terminal.get("runner_error") is None and terminal.get("process_group_closed") is True,
         f"terminal status failed: {row['id']}")
    need(terminal.get("deadline_seconds") == row["deadline_seconds"],
         f"terminal deadline changed: {row['id']}")
    version = row["version"]
    need(terminal.get("cwd") == manifest["source"][version]["source_root"],
         f"terminal cwd changed: {row['id']}")
    need(terminal.get("source_revision") == manifest["source"][version]["source_revision"],
         f"terminal source revision changed: {row['id']}")
    need(terminal.get("environment") == child_environment(manifest, row),
         f"terminal environment changed: {row['id']}")
    need(terminal.get("binary") == manifest["build"][row["lane"]][version]["binary"],
         f"terminal binary changed: {row['id']}")
    need(terminal.get("source_receipt_sha256") == manifest["source"][version]["receipt"]["sha256"],
         f"terminal source receipt changed: {row['id']}")
    need(terminal.get("runner_sha256") == manifest["scripts"]["runner"]["sha256"],
         f"terminal runner binding changed: {row['id']}")
    need(terminal.get("manifest_sha256") == sha256_path(packet / "manifest.json"),
         f"terminal manifest binding changed: {row['id']}")
    rss = integer(terminal.get("rss_bytes"), f"terminal rss {row['id']}")
    need(rss == child.get("rss_bytes"), f"terminal/state RSS differs: {row['id']}")
    for name in ("stdout", "stderr", "time"):
        local_descriptor(packet, terminal[name], f"{row['id']} {name}")
    report_path = local_descriptor(packet, terminal["report"], f"report {row['id']}")
    need(terminal["report"] == child.get("report"), f"terminal/report descriptor differs: {row['id']}")
    validation_path = local_descriptor(packet, child["validation"], f"validation {row['id']}")
    validation = read_json(validation_path)
    need(validation.get("status") == "passed" and validation.get("report") == child.get("report"),
         f"validation failed: {row['id']}")
    return report_path

def check(packet: Path) -> dict:
    packet = packet.resolve()
    manifest = read_json(packet / "manifest.json")
    rows = manifest.get("schedule", {}).get("rows")
    need(isinstance(rows, list) and len(rows) == 72, "manifest schedule rows missing/incomplete")
    state = read_json(packet / "capture-state.json")
    need(state.get("status") == "completed" and state.get("children_expected") == len(rows),
         "packet state is not a completed 72-child cohort")
    children = state.get("children")
    need(isinstance(children, list) and len(children) == len(rows), "packet child state is incomplete")
    children_by_id = {child.get("id"): child for child in children}
    need(len(children_by_id) == len(rows) and set(children_by_id) == {row["id"] for row in rows},
         "packet child IDs are not unique/frozen")
    for row in rows:
        child = children_by_id[row["id"]]
        need(child.get("status") == "passed", f"child is not passed: {row['id']}")
        check_terminal_identity(packet, manifest, row, child)
    alloc_rows = [r for r in rows if r.get("lane") == "allocator"]
    need(len(alloc_rows) == 24, f"allocator row count {len(alloc_rows)} != 24")
    cells: dict[tuple, dict[str, list[dict]]] = {}
    for row in alloc_rows:
        child_id = row["id"]
        child = children_by_id[child_id]
        report_path = packet / "children" / child_id / "report.json"
        report = read_json(report_path)
        row_with_revision = dict(row, source_revision=manifest["source"][row["version"]]["source_revision"])
        facts = report_facts(row_with_revision, report, child_id)
        cells.setdefault((row["sections"], row["ownership"]), {}).setdefault(row["version"], []).append(facts)
    output_cells = []
    for section in SECTIONS:
        for ownership in OWNERSHIPS:
            key = (section, ownership)
            versions = cells.get(key, {})
            need(set(versions) == {"baseline", "candidate"}, f"cell versions missing: {key}")
            for version in versions:
                need(len(versions[version]) == 3, f"cell repetition count: {key}/{version}")
            base, cand = versions["baseline"], versions["candidate"]
            def metrics(items):
                return {
                    "after_snapshot_drop_allocated_bytes_delta": statistics.median(x["allocated_bytes_delta_after_snapshot_drop"] for x in items),
                    "after_snapshot_drop_allocation_calls_delta": statistics.median(x["allocation_calls_delta_after_snapshot_drop"] for x in items),
                    "commits_retained_allocated_bytes_delta": statistics.median(x["allocated_bytes_delta_commits_retained"] for x in items),
                    "commits_retained_allocation_calls_delta": statistics.median(x["allocation_calls_delta_commits_retained"] for x in items),
                    "release_live_bytes_delta": statistics.median(x["release_live_bytes_delta"] for x in items),
                }
            baseline, candidate = metrics(base), metrics(cand)
            row = {"section_count": section, "ownership": ownership, "replications_per_version": 3, "baseline": baseline, "candidate": candidate}
            bytes_base = baseline["after_snapshot_drop_allocated_bytes_delta"]
            bytes_candidate = candidate["after_snapshot_drop_allocated_bytes_delta"]
            calls_base = baseline["after_snapshot_drop_allocation_calls_delta"]
            calls_candidate = candidate["after_snapshot_drop_allocation_calls_delta"]
            row["requested_bytes_reduction_fraction"] = (bytes_base - bytes_candidate) / bytes_base if bytes_base else None
            row["allocation_calls_reduction_fraction"] = (calls_base - calls_candidate) / calls_base if calls_base else None
            row["allocation_calls_strictly_decreased"] = calls_candidate < calls_base
            row["allocation_gate_pass"] = bool(
                bytes_base > 0
                and (bytes_base - bytes_candidate) / bytes_base >= 0.20
                and calls_candidate < calls_base
            )
            output_cells.append(row)
    return {
        "schema": "litchi.performance.0933.independent-allocator-check.v2",
        "status": "passed",
        "terminal_identity_rows": len(rows),
        "frozen_allocator_fixtures": ALLOCATOR_FIXTURES,
        "packet": str(packet),
        "packet_manifest_sha256": sha256_path(packet / "manifest.json"),
        "allocator_rows": len(alloc_rows),
        "cells": output_cells,
        "allocation_gate_pass": all(cell["allocation_gate_pass"] for cell in output_cells),
        "gate_definition": {"checkpoint": "after_snapshot_drop absolute counter delta from baseline", "minimum_requested_bytes_reduction_fraction": 0.20, "allocation_calls_condition": "candidate strictly less than baseline", "no_rss_or_peak_substitution": True},
        "checker_sha256": sha256_path(Path(__file__).resolve()),
    }

def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--packet", type=Path, required=True)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    try:
        result = check(args.packet)
        encoded = json.dumps(result, indent=2, sort_keys=True) + "\n"
        if args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            with args.output.open("x", encoding="utf-8") as f:
                f.write(encoded)
        print(encoded, end="")
        return 0
    except Exception as exc:
        print(f"independent allocator check failed: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1

if __name__ == "__main__":
    raise SystemExit(main())
