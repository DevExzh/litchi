#!/usr/bin/env python3
"""Independent second replay of the sealed 0735 packet.

This file intentionally shares no code with ``analyze.py``.  It checks the
same source/fixture/oracle/capture custody and recomputes process quantiles,
paired changes, deterministic bootstrap intervals, and allocation deltas
before comparing its result with ``analysis.json``.
"""

from __future__ import annotations

import hashlib
import json
import math
import random
import statistics
from pathlib import Path
from typing import Any


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
CAPTURES = PACKET / "captures"
HEX = set("0123456789abcdef")
CASES = ("primary", "secondary")
LANES = ("native", "allocation")
VARIANTS = ("baseline", "candidate")
ALLOC_FIELDS = ("allocated_bytes", "deallocated_bytes", "allocation_calls", "peak_live_bytes", "retained_bytes")
TIMING_FIELDS = ("p50", "mean", "p95", "p99", "maximum")
SEED = 7335
RESAMPLES = 10_000
THRESHOLD = 5.0
ANCESTOR_ORACLE_SHA = "34eeca72612a6a9e7352225a969ae0b5631fa818a248b19f0a556a578c591ebf"
ANCESTOR_REFERENCE_SHA = "51459c3fea40603ce335cf66ff9f0df07ad2be758b22d2c1830dc7be4199cb2a"
STATIC = (
    "schema_version", "case", "format", "operation", "scope", "phase_contract", "input",
    "policy", "policy_applied", "policy_application_scope", "policy_argument_effect",
    "policy_contract", "allocation_ownership_contract", "directory_metadata_fields",
    "source_sha256", "expected_output_sha256", "replacements_sha256", "source_inventory",
    "expected_output_inventory", "replacements", "changed_length_proof", "expected_oracle",
    "oracle_controls",
)


class AuditError(Exception):
    """Independent replay failure."""


def need(value: bool, message: str) -> None:
    if not value:
        raise AuditError(message)


def read(path: Path) -> Any:
    need(path.is_file() and not path.is_symlink(), f"missing file: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise AuditError(f"invalid JSON {path}: {error}") from error


def sha(path: Path) -> str:
    need(path.is_file() and not path.is_symlink(), f"missing or symlinked file: {path}")
    return hashlib.sha256(path.read_bytes()).hexdigest()


def digest(value: Any, label: str) -> str:
    need(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
         f"{label}: invalid SHA-256")
    return value


def integer(value: Any, label: str, minimum: int | None = 0) -> int:
    need(isinstance(value, int) and not isinstance(value, bool), f"{label}: invalid integer")
    if minimum is not None:
        need(value >= minimum, f"{label}: below {minimum}")
    return value


def rel(value: Any, label: str) -> str:
    need(isinstance(value, str) and value and not Path(value).is_absolute()
         and ".." not in Path(value).parts, f"{label}: unsafe path")
    return value


def basename(value: Any, label: str) -> str:
    value = rel(value, label)
    need("/" not in value and "\\" not in value, f"{label}: not a basename")
    return value


def stats(values: list[int | float]) -> dict[str, int | float]:
    need(values, "empty statistics")
    ordered = sorted(values)
    return {"n": len(ordered), "p50": statistics.median(ordered),
            "mean": math.fsum(ordered) / len(ordered),
            "p95": ordered[max(0, math.ceil(.95 * len(ordered)) - 1)],
            "p99": ordered[max(0, math.ceil(.99 * len(ordered)) - 1)],
            "maximum": ordered[-1]}


def all_true(value: Any, label: str) -> None:
    if isinstance(value, bool):
        need(value, f"{label}: false witness")
    elif isinstance(value, dict):
        for key, item in value.items():
            all_true(item, f"{label}.{key}")
    elif isinstance(value, list):
        for index, item in enumerate(value):
            all_true(item, f"{label}[{index}]")


def oracle_guard(value: Any, label: str) -> None:
    need(isinstance(value, dict) and value.get("oracle_ok") is True
         and value.get("failure_reasons") == [], f"{label}: oracle failed")
    all_true(value, label)


def packet_plan() -> dict[str, Any]:
    value = read(PACKET / "plan.json")
    need(isinstance(value, dict) and value.get("cpu") == 12, "plan or CPU changed")
    need(value.get("native_samples") == 50 and value.get("native_warmups") == 3,
         "native sampling changed")
    need(value.get("allocation_samples") == 1 and value.get("allocation_warmups") == 0,
         "allocation sampling changed")
    schedule = value.get("schedule")
    need(isinstance(schedule, list) and len(schedule) == 48, "schedule count changed")
    expected = {
        ("native", cycle, repeat, case, variant)
        for cycle in range(3) for repeat in range(3) for case in CASES for variant in VARIANTS
    } | {
        ("allocation", 0, repeat, case, variant)
        for repeat in range(3) for case in CASES for variant in VARIANTS
    }
    actual = set()
    for index, row in enumerate(schedule):
        need(isinstance(row, dict), f"schedule row {index}: not an object")
        key = tuple(row.get(name) for name in ("lane", "cycle", "repeat", "case", "variant"))
        need(key in expected and key not in actual, f"schedule row {index}: invalid/duplicate")
        actual.add(key)
    need(actual == expected, "schedule matrix changed")
    normalized = dict(value)
    normalized.update(samples=value["native_samples"], warmups=value["native_warmups"],
                      allocation_processes=12, native_processes=36)
    return normalized


def packet_cases() -> dict[str, dict[str, Any]]:
    rows = read(PACKET / "cases.json")
    need(isinstance(rows, list) and [row.get("id") for row in rows] == list(CASES),
         "case order changed")
    result = {}
    for row in rows:
        case_id = row.get("id")
        need(case_id in CASES and case_id not in result, f"invalid case: {case_id!r}")
        path = rel(row.get("path"), f"{case_id} fixture path")
        fixture = ROOT / path
        need(fixture.is_file() and not fixture.is_symlink()
             and fixture.stat().st_size == row.get("bytes")
             and sha(fixture) == row.get("sha256"), f"{case_id}: fixture custody changed")
        digest(row.get("sha256"), f"{case_id}: fixture hash")
        need(row.get("case") in ("ppt45543", "ppt-secondary"), f"{case_id}: CLI selector changed")
        result[case_id] = row
    return result


def map_custody(mapping: Any, base: Path, label: str, alternatives: tuple[Path, ...] = ()) -> dict[str, str]:
    need(isinstance(mapping, dict) and mapping, f"{label}: empty custody map")
    result = {}
    for raw, expected in mapping.items():
        path = rel(raw, f"{label} path")
        expected = digest(expected, f"{label} {path}")
        targets = (base / path,) + tuple(root / path for root in alternatives)
        need(any(target.is_file() and not target.is_symlink() and sha(target) == expected
                 for target in targets), f"{label}: changed {path}")
        result[path] = expected
    return result


def quality(build: dict[str, Any], variant: str) -> None:
    name = rel(build.get("quality"), f"{variant} quality path")
    path = PACKET / name
    need(sha(path) == digest(build.get("quality_sha256"), f"{variant} quality hash"),
         f"{variant}: quality manifest changed")
    value = read(path)
    need(isinstance(value, dict) and value.get("source") == build.get("source")
         and value.get("probe") == build.get("probe"), f"{variant}: quality custody changed")
    runs = value.get("runs")
    need(isinstance(runs, list) and len(runs) == 5, f"{variant}: quality gate count changed")
    manifest = str(PACKET / "probe" / "Cargo.toml")
    common = ["--manifest-path", manifest, "--release", "--offline", "--locked"]
    expected_commands = [
        ["cargo", "fmt", "--manifest-path", manifest, "--", "--check"],
        ["cargo", "test", *common, "--lib"],
        ["cargo", "clippy", *common, "--all-targets", "--", "-D", "warnings"],
        ["cargo", "doc", *common, "--no-deps"],
        ["cargo", "build", *common, "--bins"],
    ]
    for index, row in enumerate(runs):
        need(row.get("exit_code") == 0 and isinstance(row.get("command"), list),
             f"{variant}: quality gate {index} failed")
        need(row.get("command") == expected_commands[index],
             f"{variant}: quality command {index} changed")
        output = rel(row.get("output"), f"{variant}: quality output")
        need(sha(PACKET / output) == digest(row.get("sha256"), f"{variant}: quality output hash"),
             f"{variant}: quality output changed")


def build(variant: str) -> dict[str, Any]:
    value = read(PACKET / f"{variant}-build.json")
    need(isinstance(value, dict), f"{variant}: build receipt changed")
    source = map_custody(value.get("source"), ROOT, f"{variant} source",
                         (PACKET / "source-archive" / "before", PACKET / "source-archive" / "after"))
    probe = map_custody(value.get("probe"), PACKET, f"{variant} probe")
    need(value.get("source") == source and value.get("probe") == probe,
         f"{variant}: build maps changed")
    quality(value, variant)
    binaries = value.get("binaries")
    need(isinstance(binaries, dict) and set(binaries) == set(LANES), f"{variant}: binary map changed")
    for lane in LANES:
        row = binaries[lane]
        need(isinstance(row, dict) and isinstance(row.get("path"), str)
             and Path(row["path"]).is_absolute(), f"{variant}/{lane}: binary path changed")
        integer(row.get("bytes"), f"{variant}/{lane}: binary size", 1)
        digest(row.get("sha256"), f"{variant}/{lane}: binary hash")
    return value


def binary_custody(builds: dict[str, dict[str, Any]]) -> None:
    expected = [builds[variant]["binaries"][lane] for variant in VARIANTS for lane in LANES]
    cleanup_path = PACKET / "cleanup.json"
    cleanup = read(cleanup_path) if cleanup_path.is_file() else None
    missing = []
    for row in expected:
        path = Path(row["path"])
        if path.is_file():
            need(not path.is_symlink() and path.stat().st_size == row["bytes"]
                 and sha(path) == row["sha256"], f"binary changed: {path}")
        else:
            need(not path.exists() and not path.is_symlink(), f"invalid binary path: {path}")
            missing.append(row)
    if missing:
        need(isinstance(cleanup, dict) and cleanup.get("removed") is True,
             "missing binaries lack cleanup receipt")
        actual = cleanup.get("binaries", cleanup.get("identities"))
        need(sorted(actual or [], key=lambda x: json.dumps(x, sort_keys=True))
             == sorted(expected, key=lambda x: json.dumps(x, sort_keys=True)),
             "cleanup does not identify all four binaries")


def source_bridge(builds: dict[str, dict[str, Any]]) -> None:
    before, after = builds["baseline"]["source"], builds["candidate"]["source"]
    need(set(before) == set(after), "source file inventory changed")
    base = read(PACKET / "base.json")
    owned = rel(base.get("owned_file"), "owned source")
    need(before.get(owned) == base.get("before_sha256"), "baseline source does not match base")
    changed = [name for name in before if before[name] != after[name]]
    need(changed == [owned], f"source changed outside owned file: {changed}")
    need(Path(ROOT / owned).is_file() and sha(ROOT / owned) == after[owned],
         "candidate source is not live")
    for name in before:
        if name != owned:
            need(before[name] == after[name], f"unowned source drifted: {name}")
    need(sha(PACKET / "source-archive" / "before" / owned) == before[owned],
         "before archive changed")
    need(sha(PACKET / "source-archive" / "after" / owned) == after[owned],
         "after archive changed")
    need(builds["baseline"]["probe"] == builds["candidate"]["probe"],
         "probe custody differs between variants")


def source_review(builds: dict[str, dict[str, Any]]) -> None:
    value = read(PACKET / "source-review.json")
    need(value.get("schema_version") == 1
         and value.get("status") == "approved_for_measurement"
         and value.get("blocking_findings") == [], "source review is not approved")
    base = read(PACKET / "base.json")
    need(value.get("base_head") == base.get("head"), "source review base head changed")
    files = value.get("files")
    need(isinstance(files, dict)
         and {"baseline", "candidate", "live_formatted", "implementation_notes"} <= set(files),
         "source review file map incomplete")
    for key in ("baseline", "candidate", "live_formatted", "implementation_notes"):
        row = files[key]
        path = rel(row.get("path"), f"source review {key} path")
        target = ROOT / path
        need(sha(target) == digest(row.get("sha256"), f"source review {key} hash"),
             f"source review {key} changed")
    owned = base["owned_file"]
    need(files["baseline"]["sha256"] == builds["baseline"]["source"][owned]
         and files["live_formatted"]["sha256"] == builds["candidate"]["source"][owned],
         "source review does not bind both builds")
    equivalent = value.get("source_equivalence")
    need(isinstance(equivalent, dict)
         and equivalent.get("live_matches_candidate_after_rustfmt") is True
         and equivalent.get("production_files_reviewed") == 1
         and equivalent.get("production_files_edited_by_reviewer") == 0,
         "source review equivalence changed")
    need(value.get("constraints_sha256") == sha(PACKET / "constraints.json"),
         "source review constraints binding changed")


def freeze(builds: dict[str, dict[str, Any]]) -> None:
    value = read(PACKET / "freeze.json")
    bindings = value.get("bindings", value) if isinstance(value, dict) else None
    need(isinstance(bindings, dict) and bindings, "freeze is empty")
    binaries = {Path(builds[v]["binaries"][lane]["path"]): builds[v]["binaries"][lane]
                for v in VARIANTS for lane in LANES}
    cleanup = read(PACKET / "cleanup.json") if (PACKET / "cleanup.json").is_file() else None
    for raw, expected in bindings.items():
        need(isinstance(raw, str) and Path(raw).is_absolute(), f"unsafe freeze path: {raw!r}")
        expected = digest(expected, f"freeze {raw}")
        path = Path(raw)
        if path in binaries and not path.is_file():
            need(isinstance(cleanup, dict) and cleanup.get("removed") is True,
                 f"frozen binary missing cleanup: {raw}")
            need(expected == binaries[path]["sha256"], f"frozen binary hash changed: {raw}")
        else:
            need(path.is_file() and not path.is_symlink() and sha(path) == expected,
                 f"frozen input changed: {raw}")
    need(set(binaries) <= {Path(raw) for raw in bindings}, "freeze omits binary")


def preflight() -> None:
    value = read(PACKET / "preflight.json")
    need(value.get("status") == "passed" and value.get("freeze_sha256") == sha(PACKET / "freeze.json"),
         "preflight binding changed")
    scripts = value.get("scripts")
    if scripts is not None:
        for name in ("preflight.py", "analyze.py", "audit.py"):
            need(name in scripts and sha(PACKET / name) == digest(scripts[name], f"preflight {name}"),
                 f"preflight script changed: {name}")


def oracles() -> dict[str, dict[str, Any]]:
    value = read(PACKET / "oracle.json")
    need(isinstance(value, dict) and set(value) == set(CASES), "oracle map changed")
    for case_id in CASES:
        row = value[case_id]
        need(isinstance(row, dict) and isinstance(row.get("expected"), dict),
             f"{case_id}: oracle row changed")
        digest(row.get("sha256"), f"{case_id}: oracle hash")
        output = rel(row.get("qualification_output"), f"{case_id}: oracle output")
        need(sha(PACKET / output) == row["sha256"], f"{case_id}: oracle output changed")
        need(row["expected"] == read(PACKET / output), f"{case_id}: expected report differs from output")
        oracle_guard(row["expected"].get("expected_oracle"), f"{case_id}: expected oracle")
    ancestor_wrapper = ROOT / "docs/performance/results/change-0731/oracle.json"
    need(sha(ancestor_wrapper) == ANCESTOR_ORACLE_SHA, "sealed 0731 oracle wrapper changed")
    ancestor = read(ancestor_wrapper)
    reference = ROOT / ancestor["reference"]
    need(sha(reference) == ANCESTOR_REFERENCE_SHA == ancestor["sha256"],
         "sealed 0731 oracle reference changed")
    for key in STATIC:
        need(value["primary"]["expected"].get(key) == ancestor["expected"].get(key),
             f"primary oracle differs from sealed 0731 field: {key}")
    old = read(ROOT / "docs/performance/results/change-0728/oracle-contract.json")
    need(value["primary"]["expected"]["expected_oracle"]["semantic_witness"]
         == old["ppt45543"]["semantic_witness"], "primary witness differs from 0728")
    return value


def report(value: Any, spec: dict[str, Any], case: dict[str, Any], expected: dict[str, Any],
           plan: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    need(isinstance(value, dict), "capture report is not an object")
    lane = spec["lane"]
    n = plan["native_samples"] if lane == "native" else plan["allocation_samples"]
    warmups = plan["native_warmups"] if lane == "native" else plan["allocation_warmups"]
    for key in STATIC:
        need(value.get(key) == expected.get(key), f"{lane}/{spec['case']}: static field {key} changed")
    need(value.get("input") == case["path"] and value.get("source_sha256") == case["sha256"],
         f"{lane}/{spec['case']}: source identity changed")
    need(value.get("timing_claim") is (lane == "native")
         and value.get("allocator_instrumented") is (lane == "allocation"),
         f"{lane}/{spec['case']}: mode changed")
    need(value.get("samples_requested") == n and value.get("warmups") == warmups,
         f"{lane}/{spec['case']}: sample header changed")
    oracle_guard(value.get("expected_oracle"), f"{lane}/{spec['case']}: expected")
    controls = value.get("oracle_controls")
    need(isinstance(controls, list) and controls == expected.get("oracle_controls"),
         f"{lane}/{spec['case']}: controls changed")
    for control in controls:
        need(control.get("status") == "rejected" and control.get("rejected") is True
             and control.get("failure_reasons"), f"{lane}/{spec['case']}: accepted control")
    samples = value.get("samples")
    need(isinstance(samples, list) and len(samples) == n, f"{lane}/{spec['case']}: sample count")
    times = []
    alloc = None
    for index, item in enumerate(samples):
        need(item.get("index") == index and item.get("output_sha256") == expected["expected_output_sha256"]
             and item.get("output_inventory") == expected["expected_output_inventory"]
             and item.get("oracle") == expected["expected_oracle"],
             f"{lane}/{spec['case']}: sample {index} identity")
        oracle_guard(item.get("oracle"), f"{lane}/{spec['case']}: sample {index}")
        if lane == "native":
            phase = item.get("phase_ns")
            need(isinstance(phase, dict) and set(phase) == {"whole_ns"}
                 and "allocations" not in item, f"{lane}/{spec['case']}: timing schema")
            times.append(integer(phase["whole_ns"], "whole_ns", 1))
        else:
            value_alloc = item.get("allocations")
            need(isinstance(value_alloc, dict) and set(value_alloc) == {"whole"}
                 and "phase_ns" not in item, f"{lane}/{spec['case']}: allocation schema")
            whole = value_alloc["whole"]
            need(isinstance(whole, dict) and set(whole) == set(ALLOC_FIELDS),
                 f"{lane}/{spec['case']}: allocation fields")
            current = {key: integer(whole[key], key, None if key == "retained_bytes" else 0)
                       for key in ALLOC_FIELDS}
            alloc = current if alloc is None else alloc
            need(current == alloc, f"{lane}/{spec['case']}: allocation drift")
    return (stats(times), {}) if lane == "native" else ({}, alloc or {})


def qualification(receipt_name: str, values: dict[str, dict[str, Any]], cases: dict[str, dict[str, Any]],
                  label: str) -> None:
    value = read(PACKET / receipt_name)
    need(value.get("status") in ("passed", "pass"), f"{label}: qualification failed")
    files = value.get("files")
    need(isinstance(files, dict) and files, f"{label}: qualification files empty")
    for raw, expected in files.items():
        path = rel(raw, f"{label}: file")
        need(sha(PACKET / path) == digest(expected, f"{label}: file hash"), f"{label}: file changed")
    reports = []
    for raw in files:
        path = PACKET / raw
        if path.suffix == ".json" and ".receipt." not in path.name and path.name != "manifest.json":
            reports.append(read(path))
    need(len(reports) >= 4, f"{label}: qualification reports incomplete")
    for item in reports:
        case_id = next((key for key in CASES if values[key]["expected"]["case"] == item.get("case")), None)
        need(case_id is not None, f"{label}: unknown qualification case")
        lane = "allocation" if item.get("allocator_instrumented") is True else "native"
        plan = {"native_samples": item.get("samples_requested") if lane == "native" else 50,
                "native_warmups": item.get("warmups") if lane == "native" else 3,
                "allocation_samples": item.get("samples_requested") if lane == "allocation" else 1,
                "allocation_warmups": item.get("warmups") if lane == "allocation" else 0}
        report(item, {"lane": lane, "case": case_id}, cases[case_id], values[case_id]["expected"], plan)


def qualified_processes(receipt_name: str, values: dict[str, dict[str, Any]],
                        cases: dict[str, dict[str, Any]], builds: dict[str, dict[str, Any]],
                        label: str, variants: set[str]) -> None:
    """Replay every retained qualification report and its command receipt."""
    receipt = read(PACKET / receipt_name)
    need(receipt.get("status") in ("passed", "pass") and isinstance(receipt.get("files"), dict),
         f"{label}: receipt changed")
    files = receipt["files"]
    for raw, expected_hash in files.items():
        path = rel(raw, f"{label}: file")
        need(sha(PACKET / path) == digest(expected_hash, f"{label}: file hash"),
             f"{label}: file changed: {path}")
    reports = []
    for raw in files:
        path = PACKET / raw
        if path.suffix == ".json" and ".receipt." not in path.name and path.name != "manifest.json":
            reports.append((raw, read(path)))
    need(len(reports) == len(CASES) * len(LANES) * len(variants),
         f"{label}: qualification report count changed")
    seen = set()
    for raw, item in reports:
        case_id = next((name for name in CASES if values[name]["expected"]["case"] == item.get("case")), None)
        need(case_id is not None, f"{label}: unknown case in {raw}")
        variant = "baseline" if Path(raw).name.startswith("baseline-") else "candidate"
        need(variant in variants, f"{label}: unexpected variant in {raw}")
        lane = "allocation" if item.get("allocator_instrumented") is True else "native"
        samples = item.get("samples_requested")
        warmups = item.get("warmups")
        need(samples in (1, 50) and warmups in (0, 3), f"{label}: sample plan changed: {raw}")
        qplan = {"native_samples": samples, "native_warmups": warmups,
                 "allocation_samples": samples, "allocation_warmups": warmups}
        spec = {"lane": lane, "case": case_id, "variant": variant}
        report(item, spec, cases[case_id], values[case_id]["expected"], qplan)
        command = ["taskset", "-c", "12", builds[variant]["binaries"][lane]["path"],
                   "--case", cases[case_id]["case"], "--input", cases[case_id]["path"],
                   "--operation", "format", "--samples", str(samples), "--warmups", str(warmups)]
        report_path = PACKET / raw
        stem = report_path.name[:-5]
        receipt_path = report_path.with_name(stem + ".receipt.json")
        stderr_path = report_path.with_name(stem + ".stderr")
        need(str(receipt_path.relative_to(PACKET)) in files
             and str(stderr_path.relative_to(PACKET)) in files,
             f"{label}: receipt or stderr missing for {raw}")
        process_receipt = read(receipt_path)
        need(process_receipt.get("command") == command and process_receipt.get("exit_code") == 0
             and process_receipt.get("output") == raw
             and process_receipt.get("sha256") == sha(report_path)
             and process_receipt.get("stderr_sha256") == sha(stderr_path),
             f"{label}: command receipt changed for {raw}")
        seen.add((case_id, lane, variant))
    need(seen == {(case_id, lane, variant) for case_id in CASES for lane in LANES
                  for variant in variants}, f"{label}: qualification matrix incomplete")


def capture(plan: dict[str, Any], builds: dict[str, dict[str, Any]], cases: dict[str, dict[str, Any]],
            values: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    manifest = read(CAPTURES / "manifest.json")
    need(manifest.get("status") == "complete"
         and manifest.get("freeze_sha256") == sha(PACKET / "freeze.json")
         and manifest.get("preflight_sha256") == sha(PACKET / "preflight.json"),
         "capture binding changed")
    runs = manifest.get("runs")
    need(isinstance(runs, list) and len(runs) == 48, "capture count changed")
    seen = set()
    rows = []
    for index, row in enumerate(runs):
        need(isinstance(row, dict), f"capture row {index}: not object")
        spec = dict(row.get("spec", {}))
        for key in ("lane", "cycle", "repeat", "case", "variant"):
            spec.setdefault(key, row.get(key))
        planned = plan["schedule"][index]
        need(all(spec.get(key) == planned.get(key)
                 for key in ("lane", "cycle", "repeat", "case", "variant")),
             f"capture row {index}: schedule changed")
        key = tuple(spec.get(name) for name in ("lane", "cycle", "repeat", "case", "variant"))
        need(key not in seen, f"capture row {index}: duplicate")
        seen.add(key)
        need(row.get("exit_code") == 0, f"capture row {index}: child failure")
        case = cases[spec["case"]]
        binary = builds[spec["variant"]]["binaries"][spec["lane"]]["path"]
        samples = plan["native_samples"] if spec["lane"] == "native" else plan["allocation_samples"]
        warmups = plan["native_warmups"] if spec["lane"] == "native" else plan["allocation_warmups"]
        command = ["taskset", "-c", str(plan["cpu"]), binary, "--case", case["case"],
                   "--input", case["path"], "--operation", "format", "--samples", str(samples),
                   "--warmups", str(warmups)]
        need(row.get("command") == command, f"capture row {index}: command changed")
        output_name = row.get("output", row.get("output_basename", row.get("outputbasename")))
        output_name = basename(output_name, f"capture row {index}: output")
        output = CAPTURES / output_name
        stderr_name = row.get("stderr", output_name + ".stderr")
        stderr_name = basename(stderr_name, f"capture row {index}: stderr")
        stderr = CAPTURES / stderr_name
        need(output.is_file() and stderr.is_file(), f"capture row {index}: missing output")
        need(row.get("sha256") == sha(output) and row.get("stderr_sha256") == sha(stderr),
             f"capture row {index}: raw hash changed")
        timing, allocation = report(read(output), spec, case, values[spec["case"]]["expected"], plan)
        rows.append({"lane": spec["lane"], "cycle": spec["cycle"], "repeat": spec["repeat"],
                     "case": spec["case"], "variant": spec["variant"], "timing": timing,
                     "allocations": allocation,
                     "output_sha256": values[spec["case"]]["expected"]["expected_output_sha256"]})
    need(seen == {tuple(row.get(name) for name in ("lane", "cycle", "repeat", "case", "variant"))
                  for row in plan["schedule"]}, "capture matrix incomplete")
    disk = {path.name for path in CAPTURES.iterdir() if path.is_file() and not path.is_symlink()}
    required = {row.get("output", row.get("output_basename", row.get("outputbasename"))) for row in runs}
    required |= {row.get("stderr", row.get("output", row.get("output_basename", row.get("outputbasename"))) + ".stderr")
                for row in runs}
    need(disk == required | {"manifest.json"}, "capture directory has unmanifested files")
    return rows


def change(before: float, after: float) -> dict[str, Any]:
    delta = after - before
    percent = None if before == 0 else delta / abs(before) * 100.0
    return {"before": before, "after": after, "delta": delta, "percent": percent,
            "flag_gt_5pct": (delta != 0 if percent is None else abs(percent) > THRESHOLD)}


def bootstrap(values: list[float], reducer: str, rng: random.Random) -> dict[str, Any] | None:
    if any(value is None for value in values):
        return None
    need(len(values) == 9, "bootstrap pair count changed")
    output = []
    for _ in range(RESAMPLES):
        selected = [values[rng.randrange(9)] for _ in range(9)]
        output.append(statistics.median(selected) if reducer == "median" else math.fsum(selected) / 9)
    output.sort()
    return {"seed": SEED, "resamples": RESAMPLES, "reducer": reducer,
            "percentile_2_5": output[math.ceil(.025 * len(output)) - 1],
            "percentile_97_5": output[math.ceil(.975 * len(output)) - 1]}


def paired(rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    index = {(row["cycle"], row["repeat"], row["case"], row["variant"]): row
             for row in rows if row["lane"] == "native"}
    pairs = []
    rng = random.Random(SEED)
    for case_id in CASES:
        for cycle in range(3):
            for repeat in range(3):
                left = index[(cycle, repeat, case_id, "baseline")]["timing"]
                right = index[(cycle, repeat, case_id, "candidate")]["timing"]
                differences = {field: change(float(left[field]), float(right[field])) for field in TIMING_FIELDS}
                pairs.append({"case": case_id, "cycle": cycle, "repeat": repeat,
                              "baseline": left, "candidate": right,
                              "differences": differences,
                              "flags_gt_5pct": [field for field, value in differences.items()
                                                 if value["flag_gt_5pct"]]})
    summaries = []
    for case_id in CASES:
        selected = [row for row in pairs if row["case"] == case_id]
        metrics = {}
        for field in ("p50", "mean"):
            values = [row["differences"][field]["percent"] for row in selected]
            need(all(value is not None for value in values), f"{case_id}/{field}: zero baseline")
            numeric = [float(value) for value in values]
            metrics[field] = {"pair_percentages": numeric,
                              "median_percent": statistics.median(numeric),
                              "minimum_percent": min(numeric), "maximum_percent": max(numeric),
                              "flags_gt_5pct": [i for i, value in enumerate(numeric) if abs(value) > THRESHOLD],
                              "bootstrap_median": bootstrap(numeric, "median", rng),
                              "bootstrap_mean": bootstrap(numeric, "mean", rng)}
        summaries.append({"case": case_id, "pairs": 9, "metrics": metrics,
                          "review_flags": sorted({field for row in selected for field in row["flags_gt_5pct"]})})
    return pairs, summaries


def allocations(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    output = []
    for case_id in CASES:
        selected = {(row["repeat"], row["variant"]): row for row in rows
                    if row["lane"] == "allocation" and row["case"] == case_id}
        need(set(selected) == {(repeat, variant) for repeat in range(3) for variant in VARIANTS},
             f"allocation matrix incomplete: {case_id}")
        fields = {}
        for field in ALLOC_FIELDS:
            before = [selected[(repeat, "baseline")]["allocations"][field] for repeat in range(3)]
            after = [selected[(repeat, "candidate")]["allocations"][field] for repeat in range(3)]
            paired_rows = [change(float(left), float(right)) for left, right in zip(before, after)]
            fields[field] = {"baseline_values": before, "candidate_values": after,
                             "baseline_stats": stats(before), "candidate_stats": stats(after),
                             "same_repeat": paired_rows,
                             "mean_comparison": change(math.fsum(before) / 3, math.fsum(after) / 3),
                             "review_flags": [i for i, row in enumerate(paired_rows)
                                              if row["flag_gt_5pct"]]}
        output.append({"case": case_id, "repeats": 3, "fields": fields,
                       "review_flags": sorted({field for field, row in fields.items()
                                                if row["review_flags"] or row["mean_comparison"]["flag_gt_5pct"]})})
    return output


def close(actual: Any, expected: Any, path: str = "analysis") -> None:
    if isinstance(actual, bool) or isinstance(expected, bool):
        need(actual is expected, f"{path}: boolean differs")
    elif isinstance(actual, dict) or isinstance(expected, dict):
        need(isinstance(actual, dict) and isinstance(expected, dict) and set(actual) == set(expected),
             f"{path}: object shape differs")
        for key in expected:
            close(actual[key], expected[key], f"{path}.{key}")
    elif isinstance(actual, list) or isinstance(expected, list):
        need(isinstance(actual, list) and isinstance(expected, list) and len(actual) == len(expected),
             f"{path}: list shape differs")
        for index, (left, right) in enumerate(zip(actual, expected)):
            close(left, right, f"{path}[{index}]")
    elif isinstance(actual, (int, float)) and isinstance(expected, (int, float)):
        need(math.isclose(float(actual), float(expected), rel_tol=1e-12, abs_tol=1e-9),
             f"{path}: number differs")
    else:
        need(actual == expected, f"{path}: value differs")


def main() -> int:
    plan = packet_plan()
    cases = packet_cases()
    constraints = read(PACKET / "constraints.json")
    for raw, expected in constraints.items():
        path = rel(raw, "constraint")
        need(sha(ROOT / path) == digest(expected, f"constraint {path}"), f"constraint changed: {path}")
    builds = {variant: build(variant) for variant in VARIANTS}
    root_quality = read(PACKET / "quality.json")
    need(root_quality.get("source") == builds["candidate"]["source"]
         and root_quality.get("probe") == builds["candidate"]["probe"]
         and isinstance(root_quality.get("runs"), list) and len(root_quality["runs"]) == 7,
         "root quality receipt changed")
    owner = ["-p", "litchi-ppt", "--release", "--offline", "--locked"]
    feature = ["--features", "performance-diagnostics"]
    quality_commands = [
        ["cargo", "fmt", "-p", "litchi-ppt", "--", "--check"],
        ["cargo", "check", *owner, "--no-default-features"],
        ["cargo", "test", *owner, *feature, "--all-targets"],
        ["cargo", "clippy", *owner, *feature, "--all-targets", "--", "-D", "warnings"],
        ["cargo", "test", *owner, *feature, "--doc"],
        ["cargo", "doc", *owner, *feature, "--no-deps"],
        ["python3", "tools/check_crate_boundaries.py"],
    ]
    for index, row in enumerate(root_quality["runs"]):
        need(row.get("command") == quality_commands[index], f"root quality command {index} changed")
        need(row.get("exit_code") == 0 and sha(PACKET / rel(row.get("output"), "root quality output"))
             == digest(row.get("sha256"), "root quality hash"), "root quality output changed")
    binary_custody(builds)
    source_bridge(builds)
    source_review(builds)
    freeze(builds)
    preflight()
    oracle = oracles()
    # The baseline-before receipt is checked independently from the final
    # candidate qualification receipt; both retain raw file hashes and exact
    # child command receipts.
    before = read(PACKET / "before.json")
    need(before.get("source") == builds["baseline"]["source"]
         and before.get("build_sha256") == sha(PACKET / "baseline-build.json"),
         "before source/build binding changed")
    need(before.get("status") == "passed" and before.get("candidate_order") == [
        "test-data/poi/test-data/slideshow/41246-1.ppt",
        "test-data/office-interop/libreoffice-resaved/45543-transition-litchi.ppt",
        "test-data/ole/ppt/SampleShow.ppt",
    ], "secondary qualification order changed")
    qualified_processes("before.json", oracle, cases, builds, "before", {"baseline"})
    qualification = read(PACKET / "qualification.json")
    need(qualification.get("builds") == {
        variant: sha(PACKET / f"{variant}-build.json") for variant in VARIANTS
    }, "candidate qualification build binding changed")
    qualified_processes("qualification.json", oracle, cases, builds, "qualification",
                        set(VARIANTS))
    rows = capture(plan, builds, cases, oracle)
    need(len([row for row in rows if row["lane"] == "native"]) == 36
         and len([row for row in rows if row["lane"] == "allocation"]) == 12,
         "capture lanes changed")
    native_pairs, summaries = paired(rows)
    allocation_rows = allocations(rows)
    independent = {
        "schema_version": 1,
        "packet": "change-0735",
        "status": "passed",
        "disposition": "root review required; no automatic acceptance",
        "matrix": {"native_processes": 36, "allocation_processes": 12, "cases": list(CASES),
                   "variants": list(VARIANTS), "native_samples": 50, "native_warmups": 3,
                   "allocation_samples": 1, "allocation_warmups": 0},
        "statistics": {"timing_unit": "ns", "quantiles": "midpoint p50; nearest-rank p95/p99",
                        "timing_fields": list(TIMING_FIELDS), "allocation_fields": list(ALLOC_FIELDS),
                        "paired_threshold_percent": THRESHOLD},
        "native_processes": [row for row in rows if row["lane"] == "native"],
        "native_pairs": native_pairs,
        "native_case_summaries": summaries,
        "allocation_processes": [row for row in rows if row["lane"] == "allocation"],
        "allocation_comparisons": allocation_rows,
        "review_flags": {"native": sorted({field for row in native_pairs for field in row["flags_gt_5pct"]}),
                         "allocation": sorted({field for row in allocation_rows for field in row["review_flags"]})},
        "limitations": [
            "nine paired process repeats per case are the independent comparison units; samples within a process are not independent processes",
            "the matrix is one public PPT format workflow exercising staged record edits on two fixtures and does not establish cold-I/O, RSS, concurrency, or broad CRUD behavior",
            "all retained samples, tails, raw allocation values, and review flags are reported; no selective rerun or automatic acceptance is performed",
        ],
    }
    analysis = read(PACKET / "analysis.json")
    close(analysis, independent)
    record = {"schema_version": 1, "status": "passed", "analysis_match": True,
              "analysis_sha256": sha(PACKET / "analysis.json"),
              "freeze_sha256": sha(PACKET / "freeze.json"),
              "capture_sha256": sha(CAPTURES / "manifest.json"),
              "processes": 48, "native_pairs": native_pairs, "allocation_comparisons": allocation_rows}
    (PACKET / "audit.json").write_text(json.dumps(record, indent=2, sort_keys=True) + "\n",
                                        encoding="utf-8")
    print("PASS independent 0735 custody, oracle, 48-process replay, paired bootstrap, and allocation audit")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (AuditError, OSError, KeyError, TypeError, ValueError) as error:
        print(f"FAIL: {error}")
        raise SystemExit(1)
