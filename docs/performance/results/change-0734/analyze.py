#!/usr/bin/env python3
"""Validate and summarize the 0734 PPT owned-stream comparison.

The program is deliberately a pure packet reader.  It does not start Cargo or
either probe binary.  Before reading a timing number it checks fixture,
source, probe, binary, oracle, command, and capture custody.  The paired
statistics are then computed from the retained process reports; process
samples are never treated as independent process repeats.
"""

from __future__ import annotations

import hashlib
import json
import math
import random
import statistics
from pathlib import Path
from typing import Any, Iterable


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
CAPTURES = P / "captures"
HEX = set("0123456789abcdef")
CASES = ("primary", "secondary")
LANES = ("native", "allocation")
VARIANTS = ("baseline", "candidate")
ALLOC_FIELDS = (
    "allocated_bytes",
    "deallocated_bytes",
    "allocation_calls",
    "peak_live_bytes",
    "retained_bytes",
)
TIMING_FIELDS = ("p50", "mean", "p95", "p99", "maximum")
BOOTSTRAP_COUNT = 10_000
BOOTSTRAP_SEED = 7334
PAIR_THRESHOLD_PERCENT = 5.0
ANCESTOR_ORACLE_SHA = "34eeca72612a6a9e7352225a969ae0b5631fa818a248b19f0a556a578c591ebf"
ANCESTOR_REFERENCE_SHA = "51459c3fea40603ce335cf66ff9f0df07ad2be758b22d2c1830dc7be4199cb2a"


class Failure(Exception):
    """Malformed, incomplete, or unqualified evidence."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise Failure(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise Failure(f"invalid JSON {path}: {error}") from error


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing or symlinked file: {path}")
    return sha_bytes(path.read_bytes())


def digest(value: Any, label: str) -> str:
    require(
        isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
        f"{label} is not a lowercase SHA-256 digest",
    )
    return value


def integer(value: Any, label: str, minimum: int | None = 0) -> int:
    require(isinstance(value, int) and not isinstance(value, bool), f"{label} is not an integer")
    if minimum is not None:
        require(value >= minimum, f"{label} is below {minimum}")
    return value


def safe_relative(value: Any, label: str) -> str:
    require(
        isinstance(value, str)
        and value
        and not Path(value).is_absolute()
        and ".." not in Path(value).parts,
        f"{label} is unsafe: {value!r}",
    )
    return value


def safe_capture_name(value: Any, label: str) -> str:
    value = safe_relative(value, label)
    require("/" not in value and "\\" not in value, f"{label} is not a basename")
    return value


def packet_path(raw: str) -> Path:
    value = Path(raw)
    return value if value.is_absolute() else P / value


def root_path(raw: str) -> Path:
    value = Path(raw)
    return value if value.is_absolute() else ROOT / value


def stats(values: Iterable[int | float]) -> dict[str, int | float]:
    values = list(values)
    require(values, "statistics received no values")
    ordered = sorted(values)
    return {
        "n": len(ordered),
        "p50": statistics.median(ordered),
        "mean": math.fsum(ordered) / len(ordered),
        "p95": ordered[max(0, math.ceil(0.95 * len(ordered)) - 1)],
        "p99": ordered[max(0, math.ceil(0.99 * len(ordered)) - 1)],
        "maximum": ordered[-1],
    }


def all_oracle_booleans(value: Any, label: str) -> None:
    """Reject an oracle that quietly contains a false nested witness."""
    if isinstance(value, bool):
        require(value, f"{label} contains false boolean")
    elif isinstance(value, dict):
        for key, nested in value.items():
            all_oracle_booleans(nested, f"{label}.{key}")
    elif isinstance(value, list):
        for index, nested in enumerate(value):
            all_oracle_booleans(nested, f"{label}[{index}]")


def oracle_guard(value: Any, label: str) -> None:
    require(isinstance(value, dict), f"{label} is not an object")
    require(value.get("oracle_ok") is True, f"{label}.oracle_ok is false")
    require(value.get("failure_reasons") == [], f"{label} has failure reasons")
    all_oracle_booleans(value, label)


def load_plan() -> dict[str, Any]:
    plan = read_json(P / "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("cpu") == 12, "CPU binding changed")
    require(plan.get("native_samples") == 50 and plan.get("native_warmups") == 3,
            "native sample plan changed")
    require(plan.get("allocation_samples") == 1 and plan.get("allocation_warmups") == 0,
            "allocation sample plan changed")
    # Keep the packet's explicit names while exposing the short aliases used
    # internally by the report validator.
    plan = dict(plan)
    plan["samples"] = plan["native_samples"]
    plan["warmups"] = plan["native_warmups"]
    plan["cycles"] = 3
    plan["repeats"] = 3
    plan["cases"] = list(CASES)
    plan["variants"] = list(VARIANTS)
    plan["native_processes"] = 36
    plan["allocation_processes"] = 12
    schedule = plan.get("schedule")
    require(isinstance(schedule, list) and len(schedule) == 48, "schedule is not 48 rows")
    seen: set[tuple[str, int, int, str, str]] = set()
    for index, row in enumerate(schedule):
        require(isinstance(row, dict), f"schedule row {index} is not an object")
        key = (row.get("lane"), row.get("cycle"), row.get("repeat"), row.get("case"), row.get("variant"))
        require(key not in seen, f"duplicate schedule row {index}")
        seen.add(key)
        lane, cycle, repeat, case, variant = key
        require(lane in LANES and case in CASES and variant in VARIANTS,
                f"invalid schedule row {index}: {row!r}")
        require(isinstance(cycle, int) and not isinstance(cycle, bool),
                f"schedule row {index} cycle is invalid")
        require(isinstance(repeat, int) and not isinstance(repeat, bool),
                f"schedule row {index} repeat is invalid")
        if lane == "native":
            require(cycle in range(3) and repeat in range(3),
                    f"native schedule row {index} outside 3x3 matrix")
        else:
            require(cycle == 0 and repeat in range(3),
                    f"allocation schedule row {index} outside 3 repeats")
    expected = {
        ("native", cycle, repeat, case, variant)
        for cycle in range(3)
        for repeat in range(3)
        for case in CASES
        for variant in VARIANTS
    }
    expected |= {
        ("allocation", 0, repeat, case, variant)
        for repeat in range(3)
        for case in CASES
        for variant in VARIANTS
    }
    require(seen == expected, "schedule does not cover the fixed native/allocation matrix")
    return plan


def load_cases() -> dict[str, dict[str, Any]]:
    rows = read_json(P / "cases.json")
    require(isinstance(rows, list) and len(rows) == 2, "cases is not a two-row list")
    result: dict[str, dict[str, Any]] = {}
    for row in rows:
        require(isinstance(row, dict), "case row is not an object")
        case_id = row.get("id")
        require(case_id in CASES and case_id not in result, f"invalid case id: {case_id!r}")
        case_name = row.get("case")
        require(case_name in ("ppt45543", "ppt-secondary"), f"{case_id}: CLI case changed")
        path = safe_relative(row.get("path"), f"{case_id} fixture path")
        fixture = ROOT / path
        require(fixture.is_file() and not fixture.is_symlink(), f"missing fixture: {fixture}")
        require(row.get("bytes") == fixture.stat().st_size, f"{case_id}: fixture size changed")
        require(row.get("sha256") == sha(fixture), f"{case_id}: fixture digest changed")
        digest(row.get("sha256"), f"{case_id}: fixture digest")
        result[case_id] = row
    require(list(result) == list(CASES), "case order changed")
    return result


def verify_constraints() -> None:
    constraints = read_json(P / "constraints.json")
    require(isinstance(constraints, dict) and constraints, "constraints are empty")
    for raw, expected in constraints.items():
        relative = safe_relative(raw, "constraint path")
        digest(expected, f"constraint {relative}")
        target = ROOT / relative
        require(target.is_file() and not target.is_symlink(), f"missing constraint: {relative}")
        require(sha(target) == expected, f"constraint changed: {relative}")


def verify_source_map(mapping: Any, base: Path, label: str,
                      alternatives: Iterable[Path] = ()) -> dict[str, str]:
    require(isinstance(mapping, dict) and mapping, f"{label} custody map is empty")
    checked: dict[str, str] = {}
    for raw, expected in mapping.items():
        relative = safe_relative(raw, f"{label} path")
        expected = digest(expected, f"{label} {relative}")
        candidates = [base / relative]
        candidates.extend(root / relative for root in alternatives)
        require(any(path.is_file() and not path.is_symlink() and sha(path) == expected
                    for path in candidates), f"{label} changed or missing: {relative}")
        checked[relative] = expected
    return checked


def expected_quality_commands(build: dict[str, Any]) -> list[list[str]]:
    manifest = str(P / "probe" / "Cargo.toml")
    common = ["--manifest-path", manifest, "--release", "--offline", "--locked"]
    return [
        ["cargo", "fmt", "--manifest-path", manifest, "--", "--check"],
        ["cargo", "test", *common, "--lib"],
        ["cargo", "clippy", *common, "--all-targets", "--", "-D", "warnings"],
        ["cargo", "doc", *common, "--no-deps"],
        ["cargo", "build", *common, "--bins"],
    ]


def verify_quality(build: dict[str, Any], variant: str) -> None:
    quality_name = safe_relative(build.get("quality"), f"{variant} quality path")
    quality_path = P / quality_name
    require(sha(quality_path) == digest(build.get("quality_sha256"), f"{variant} quality digest"),
            f"{variant} quality manifest digest changed")
    quality = read_json(quality_path)
    require(isinstance(quality, dict), f"{variant} quality manifest is not an object")
    require(quality.get("source") == build.get("source") and quality.get("probe") == build.get("probe"),
            f"{variant} quality custody differs from build")
    runs = quality.get("runs")
    require(isinstance(runs, list) and len(runs) == 5, f"{variant} quality command count changed")
    expected_commands = expected_quality_commands(build)
    for index, row in enumerate(runs):
        require(isinstance(row, dict) and row.get("exit_code") == 0,
                f"{variant} quality command {index} failed")
        require(row.get("command") == expected_commands[index],
                f"{variant} quality command {index} changed")
        output = safe_relative(row.get("output"), f"{variant} quality output {index}")
        output_path = P / output
        require(sha(output_path) == digest(row.get("sha256"), f"{variant} quality output {index}"),
                f"{variant} quality output {index} digest changed")


def verify_root_quality(builds: dict[str, dict[str, Any]]) -> None:
    """Check the final owner gates that bind the candidate matrix."""
    value = read_json(P / "quality.json")
    require(isinstance(value, dict), "root quality receipt is not an object")
    candidate = builds["candidate"]
    require(value.get("source") == candidate["source"] and value.get("probe") == candidate["probe"],
            "root quality custody differs from candidate build")
    runs = value.get("runs")
    require(isinstance(runs, list) and len(runs) == 7, "root quality command count changed")
    owner = ["-p", "litchi-ppt", "--release", "--offline", "--locked"]
    feature = ["--features", "performance-diagnostics"]
    expected_commands = [
        ["cargo", "fmt", "-p", "litchi-ppt", "--", "--check"],
        ["cargo", "check", *owner, "--no-default-features"],
        ["cargo", "test", *owner, *feature, "--all-targets"],
        ["cargo", "clippy", *owner, *feature, "--all-targets", "--", "-D", "warnings"],
        ["cargo", "test", *owner, *feature, "--doc"],
        ["cargo", "doc", *owner, *feature, "--no-deps"],
        ["python3", "tools/check_crate_boundaries.py"],
    ]
    for index, row in enumerate(runs):
        require(isinstance(row, dict) and row.get("exit_code") == 0,
                f"root quality command {index} failed")
        require(row.get("command") == expected_commands[index],
                f"root quality command {index} changed")
        output = safe_relative(row.get("output"), f"root quality output {index}")
        require(sha(P / output) == digest(row.get("sha256"), f"root quality output {index}"),
                f"root quality output {index} digest changed")


def load_build(variant: str) -> dict[str, Any]:
    build = read_json(P / f"{variant}-build.json")
    require(isinstance(build, dict), f"{variant} build receipt is not an object")
    source = verify_source_map(build.get("source"), ROOT, f"{variant} source",
                               (P / "source-archive" / "before",
                                P / "source-archive" / "after"))
    probe = verify_source_map(build.get("probe"), P, f"{variant} probe")
    require(build.get("source") == source and build.get("probe") == probe,
            f"{variant} custody maps changed during validation")
    verify_quality(build, variant)
    binaries = build.get("binaries")
    require(isinstance(binaries, dict) and set(binaries) == set(LANES),
            f"{variant} binary lanes changed")
    for lane in LANES:
        row = binaries[lane]
        require(isinstance(row, dict) and isinstance(row.get("path"), str)
                and Path(row["path"]).is_absolute(), f"{variant}/{lane} binary path changed")
        integer(row.get("bytes"), f"{variant}/{lane} binary bytes", 1)
        digest(row.get("sha256"), f"{variant}/{lane} binary digest")
    return build


def verify_binary_receipts(builds: dict[str, dict[str, Any]]) -> None:
    expected = [builds[v]["binaries"][lane] for v in VARIANTS for lane in LANES]
    cleanup_path = P / "cleanup.json"
    cleanup = read_json(cleanup_path) if cleanup_path.is_file() else None
    missing = []
    for row in expected:
        path = Path(row["path"])
        if path.is_file():
            require(not path.is_symlink() and path.stat().st_size == row["bytes"]
                    and sha(path) == row["sha256"], f"binary identity changed: {path}")
        else:
            require(not path.exists() and not path.is_symlink(), f"binary path invalid: {path}")
            missing.append(row)
    if missing:
        require(isinstance(cleanup, dict) and cleanup.get("removed") is True,
                "missing binaries have no cleanup receipt")
        cleaned = cleanup.get("binaries", cleanup.get("identities"))
        require(sorted(cleaned or [], key=lambda x: json.dumps(x, sort_keys=True))
                == sorted(expected, key=lambda x: json.dumps(x, sort_keys=True)),
                "cleanup binary identities are not exact")


def verify_source_bridge(builds: dict[str, dict[str, Any]]) -> None:
    baseline, candidate = (builds[v]["source"] for v in VARIANTS)
    require(set(baseline) == set(candidate), "baseline/candidate source file sets differ")
    base = read_json(P / "base.json")
    owned = base.get("owned_file")
    safe_relative(owned, "owned source path")
    require(digest(base.get("before_sha256"), "base before source digest") == baseline.get(owned),
            "baseline owned source does not match base")
    changed = [path for path in baseline if baseline[path] != candidate[path]]
    require(changed == [owned], f"source changed outside owned file: {changed}")
    require(candidate[owned] != baseline[owned], "candidate owned source did not change")
    for path in baseline:
        if path != owned:
            require(baseline[path] == candidate[path], f"source drifted: {path}")
    before = P / "source-archive" / "before" / owned
    after = P / "source-archive" / "after" / owned
    require(before.is_file() and sha(before) == baseline[owned], "before source archive changed")
    require(after.is_file() and sha(after) == candidate[owned], "after source archive changed")
    require(sha(ROOT / owned) == candidate[owned], "live candidate source changed")
    probe = builds["baseline"]["probe"]
    require(probe == builds["candidate"]["probe"], "probe source is not byte-identical")


def verify_source_review(builds: dict[str, dict[str, Any]]) -> None:
    review = read_json(P / "source-review.json")
    require(isinstance(review, dict) and review.get("schema_version") == 1
            and review.get("status") == "approved_for_measurement"
            and review.get("blocking_findings") == [], "source review is not approved")
    base = read_json(P / "base.json")
    require(review.get("base_head") == base.get("head"), "source review base head changed")
    files = review.get("files")
    require(isinstance(files, dict) and {"baseline", "candidate", "live_formatted", "implementation_notes"} <= set(files),
            "source review file map is incomplete")
    for key in ("baseline", "candidate", "live_formatted", "implementation_notes"):
        row = files[key]
        require(isinstance(row, dict), f"source review row changed: {key}")
        path = row.get("path")
        relative = safe_relative(path, f"source review {key} path")
        expected = digest(row.get("sha256"), f"source review {key} hash")
        target = ROOT / relative if relative.startswith("docs/") else ROOT / relative
        require(sha(target) == expected, f"source review {key} changed")
    owned = base["owned_file"]
    require(files["baseline"]["sha256"] == builds["baseline"]["source"][owned],
            "source review baseline differs from build")
    require(files["live_formatted"]["sha256"] == builds["candidate"]["source"][owned],
            "source review live candidate differs from build")
    equivalence = review.get("source_equivalence")
    require(isinstance(equivalence, dict)
            and equivalence.get("live_matches_candidate_after_rustfmt") is True
            and equivalence.get("production_files_reviewed") == 1
            and equivalence.get("production_files_edited_by_reviewer") == 0,
            "source review equivalence changed")
    require(review.get("constraints_sha256") == sha(P / "constraints.json"),
            "source review constraints binding changed")


def verify_freeze(builds: dict[str, dict[str, Any]]) -> dict[str, str]:
    frozen = read_json(P / "freeze.json")
    bindings = frozen.get("bindings", frozen) if isinstance(frozen, dict) else None
    require(isinstance(bindings, dict) and bindings, "freeze bindings are empty")
    binary_paths = {Path(builds[v]["binaries"][lane]["path"]): builds[v]["binaries"][lane]
                    for v in VARIANTS for lane in LANES}
    cleanup = read_json(P / "cleanup.json") if (P / "cleanup.json").is_file() else None
    for raw, expected in bindings.items():
        require(isinstance(raw, str) and Path(raw).is_absolute(), f"freeze path is not absolute: {raw!r}")
        expected = digest(expected, f"freeze {raw}")
        path = Path(raw)
        if path in binary_paths and not path.is_file():
            require(isinstance(cleanup, dict) and cleanup.get("removed") is True,
                    f"missing frozen binary lacks cleanup receipt: {raw}")
            require(expected == binary_paths[path]["sha256"], f"frozen binary digest changed: {raw}")
        else:
            require(path.is_file() and not path.is_symlink() and sha(path) == expected,
                    f"frozen input changed: {raw}")
    require(set(binary_paths) <= {Path(raw) for raw in bindings}, "freeze omits a binary")
    return {str(k): v for k, v in bindings.items()}


def verify_preflight() -> dict[str, Any]:
    value = read_json(P / "preflight.json")
    require(isinstance(value, dict) and value.get("status") == "passed", "preflight did not pass")
    require(value.get("freeze_sha256") == sha(P / "freeze.json"), "preflight freeze binding changed")
    scripts = value.get("scripts")
    if scripts is not None:
        require(isinstance(scripts, dict), "preflight scripts are not an object")
        for name in ("preflight.py", "analyze.py", "audit.py"):
            require(name in scripts and sha(P / name) == digest(scripts[name], f"preflight {name}"),
                    f"preflight script binding changed: {name}")
    return value


def oracle_rows() -> dict[str, dict[str, Any]]:
    value = read_json(P / "oracle.json")
    require(isinstance(value, dict) and set(value) == set(CASES), "oracle case map changed")
    rows: dict[str, dict[str, Any]] = {}
    for case_id in CASES:
        row = value[case_id]
        require(isinstance(row, dict), f"oracle {case_id} is not an object")
        expected = row.get("expected")
        require(isinstance(expected, dict), f"oracle {case_id} expected report is missing")
        digest(row.get("sha256"), f"oracle {case_id} digest")
        output_name = safe_relative(row.get("qualification_output"),
                                    f"oracle {case_id} qualification output")
        output = P / output_name
        require(sha(output) == row["sha256"], f"oracle {case_id} qualification output changed")
        require(expected.get("case") in ("ppt45543", "ppt-secondary")
                and expected.get("operation") == "format", f"oracle {case_id} identity changed")
        require(expected == read_json(output),
                f"oracle {case_id} expected report differs from qualification output")
        oracle_guard(expected.get("expected_oracle"), f"oracle {case_id} expected oracle")
        rows[case_id] = row
    # The primary report inherits the sealed 0731/0728 oracle contract.  Pin
    # both the wrapper and its referenced one-sample report so a new fixture or
    # regenerated oracle cannot silently become the comparison baseline.
    ancestor_wrapper = ROOT / "docs/performance/results/change-0731/oracle.json"
    require(sha(ancestor_wrapper) == ANCESTOR_ORACLE_SHA,
            "sealed 0731 oracle wrapper changed")
    ancestor = read_json(ancestor_wrapper)
    reference = root_path(ancestor["reference"])
    require(sha(reference) == ANCESTOR_REFERENCE_SHA == ancestor["sha256"],
            "sealed 0731 oracle reference changed")
    ancestor_expected = ancestor["expected"]
    for key in STATIC_KEYS:
        require(rows["primary"]["expected"].get(key) == ancestor_expected.get(key),
                f"primary oracle differs from sealed 0731 field: {key}")
    witness = rows["primary"]["expected"]["expected_oracle"].get("semantic_witness")
    contract = read_json(ROOT / "docs/performance/results/change-0728/oracle-contract.json")
    require(witness == contract["ppt45543"]["semantic_witness"],
            "primary semantic witness differs from sealed 0728 oracle")
    return rows


STATIC_KEYS = (
    "schema_version", "case", "format", "operation", "scope", "phase_contract", "input",
    "policy", "policy_applied", "policy_application_scope", "policy_argument_effect",
    "policy_contract", "allocation_ownership_contract", "directory_metadata_fields",
    "source_sha256", "expected_output_sha256", "replacements_sha256", "source_inventory",
    "expected_output_inventory", "replacements", "changed_length_proof", "expected_oracle",
    "oracle_controls",
)


def validate_static_report(report: dict[str, Any], expected: dict[str, Any], case: dict[str, Any],
                           label: str, lane: str, samples: int, warmups: int) -> None:
    require(isinstance(report, dict), f"{label}: report is not an object")
    for key in STATIC_KEYS:
        require(report.get(key) == expected.get(key), f"{label}: static field {key} changed")
    require(report.get("input") == case["path"] and report.get("source_sha256") == case["sha256"],
            f"{label}: fixture identity changed")
    require(report.get("timing_claim") is (lane == "native"), f"{label}: timing mode changed")
    require(report.get("allocator_instrumented") is (lane == "allocation"),
            f"{label}: allocator mode changed")
    require(report.get("samples_requested") == samples and report.get("warmups") == warmups,
            f"{label}: sample header changed")
    oracle_guard(report.get("expected_oracle"), f"{label}: expected oracle")
    controls = report.get("oracle_controls")
    require(isinstance(controls, list) and controls == expected.get("oracle_controls"),
            f"{label}: oracle controls changed")
    for index, control in enumerate(controls):
        require(isinstance(control, dict) and control.get("rejected") is True
                and control.get("status") == "rejected"
                and isinstance(control.get("failure_reasons"), list)
                and control["failure_reasons"], f"{label}: control {index} accepted")


def validate_report(report: dict[str, Any], row: dict[str, Any], case: dict[str, Any],
                    expected: dict[str, Any], plan: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    lane = row["lane"]
    samples_requested = plan["samples"] if lane == "native" else plan["allocation_samples"]
    warmups = plan["warmups"] if lane == "native" else plan["allocation_warmups"]
    label = f"{lane}/c{row['cycle']}/r{row['repeat']}/{row['case']}/{row['variant']}"
    validate_static_report(report, expected, case, label, lane, samples_requested, warmups)
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == samples_requested,
            f"{label}: sample count changed")
    timings: list[int] = []
    allocation: dict[str, int] | None = None
    expected_hash = expected.get("expected_output_sha256")
    expected_inventory = expected.get("expected_output_inventory")
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict) and sample.get("index") == index,
                f"{label}: sample {index} index changed")
        require(sample.get("output_sha256") == expected_hash
                and sample.get("output_inventory") == expected_inventory,
                f"{label}: sample {index} output identity changed")
        require(sample.get("oracle") == expected.get("expected_oracle"),
                f"{label}: sample {index} semantic oracle changed")
        oracle_guard(sample.get("oracle"), f"{label}: sample {index} oracle")
        if lane == "native":
            phase = sample.get("phase_ns")
            require(isinstance(phase, dict) and set(phase) == {"whole_ns"}
                    and "allocations" not in sample,
                    f"{label}: native sample shape changed")
            timings.append(integer(phase["whole_ns"], f"{label}: sample {index} whole_ns", 1))
        else:
            allocations = sample.get("allocations")
            require(isinstance(allocations, dict) and set(allocations) == {"whole"}
                    and "phase_ns" not in sample,
                    f"{label}: allocation sample shape changed")
            whole = allocations["whole"]
            require(isinstance(whole, dict) and set(whole) == set(ALLOC_FIELDS),
                    f"{label}: allocation fields changed")
            values: dict[str, int] = {}
            for key in ALLOC_FIELDS:
                minimum = None if key == "retained_bytes" else 0
                values[key] = integer(whole[key], f"{label}: {key}", minimum)
            if allocation is None:
                allocation = values
            require(values == allocation, f"{label}: repeated allocation values drifted")
    if lane == "native":
        return stats(timings), {}
    require(allocation is not None, f"{label}: allocation sample missing")
    return {}, allocation


def verify_qualification_files(receipt_name: str, expected: dict[str, dict[str, Any]],
                               builds: dict[str, dict[str, Any]], label: str) -> None:
    receipt = read_json(P / receipt_name)
    require(receipt.get("status") in ("passed", "pass"), f"{label} qualification did not pass")
    build_hashes = receipt.get("builds")
    if build_hashes is not None:
        require(build_hashes == {variant: sha(P / f"{variant}-build.json")
                                 for variant in VARIANTS},
                f"{label} qualification build bindings changed")
    files = receipt.get("files")
    require(isinstance(files, dict) and files, f"{label} qualification files are empty")
    for raw, expected_sha in files.items():
        relative = safe_relative(raw, f"{label} qualification path")
        require(sha(P / relative) == digest(expected_sha, f"{label} {relative}"),
                f"{label} qualification file changed: {relative}")
    # Every JSON report in the retained qualification set must carry the right
    # static oracle.  Each report also has a receipt that pins command, output,
    # stderr, and child exit status.
    reports = []
    for raw in files:
        path = P / raw
        if path.suffix == ".json" and ".receipt." not in path.name and path.name != "manifest.json":
            reports.append((raw, read_json(path)))
    require(reports, f"{label} has no qualification reports")
    seen: set[tuple[str, str, str]] = set()
    report_names: set[str] = set()
    for raw, report in reports:
        case_id = next((key for key, row in expected.items()
                        if row["expected"]["case"] == report.get("case")), None)
        require(case_id is not None, f"{label} report has unknown case: {raw}")
        lane = "allocation" if report.get("allocator_instrumented") is True else "native"
        case = load_cases()[case_id]
        exp = expected[case_id]["expected"]
        # Qualification intentionally has different sample counts for the
        # candidate; use the report header as its declared count after checking
        # the lane-specific shape and static identity.
        samples = report.get("samples_requested")
        warmups = report.get("warmups")
        require(samples in (1, 50) and warmups in (0, 3), f"{label} report sample plan changed: {raw}")
        variant = "baseline" if Path(raw).name.startswith("baseline-") else "candidate"
        require(Path(raw).name.startswith(("baseline-", "candidate-")),
                f"{label} report variant is not explicit: {raw}")
        fake_row = {"lane": lane, "cycle": 0, "repeat": 0,
                    "case": case_id, "variant": variant}
        fake_plan = {"samples": samples, "warmups": warmups,
                     "allocation_samples": samples, "allocation_warmups": warmups}
        validate_report(report, fake_row, case, exp, fake_plan)
        # The command shape is reconstructed from the report's own lane mode;
        # qualification deliberately uses 50/3 only for the baseline native
        # reports and 1/0 for all candidate reports and allocation reports.
        command_plan = {"cpu": 12, "samples": samples, "warmups": warmups,
                        "allocation_samples": samples, "allocation_warmups": warmups}
        command = expected_command(command_plan, builds, load_cases(), fake_row)
        report_path = P / raw
        stem = report_path.name[:-5]
        receipt_path = report_path.with_name(stem + ".receipt.json")
        stderr_path = report_path.with_name(stem + ".stderr")
        require(str(receipt_path.relative_to(P)) in files and str(stderr_path.relative_to(P)) in files,
                f"{label} receipt/stderr missing for {raw}")
        receipt_row = read_json(receipt_path)
        require(receipt_row.get("command") == command and receipt_row.get("exit_code") == 0,
                f"{label} command receipt changed: {raw}")
        require(receipt_row.get("output") == raw and receipt_row.get("sha256") == sha(report_path)
                and receipt_row.get("stderr_sha256") == sha(stderr_path),
                f"{label} raw receipt hashes changed: {raw}")
        seen.add((case_id, lane, variant))
        report_names.add(raw)
    require(seen == {(case_id, lane, variant) for case_id in CASES
                     for lane in LANES for variant in VARIANTS},
            f"{label} qualification matrix is incomplete")
    expected_files = set(report_names)
    for raw in report_names:
        stem = Path(raw).name[:-5]
        expected_files.update({str(Path(raw).with_name(stem + ".receipt.json")),
                               str(Path(raw).with_name(stem + ".stderr"))})
    require(set(files) == expected_files, f"{label} qualification file inventory changed")


def capture_path(name: Any, directory: Path, label: str) -> Path:
    value = safe_capture_name(name, label)
    path = directory / value
    require(path.parent == directory and path.is_file() and not path.is_symlink(),
            f"missing capture file: {path}")
    return path


def run_spec(row: dict[str, Any]) -> dict[str, Any]:
    spec = row.get("spec")
    if spec is None:
        return row
    require(isinstance(spec, dict), "capture spec is not an object")
    merged = dict(spec)
    for key, value in row.items():
        if key != "spec":
            merged.setdefault(key, value)
    return merged


def expected_command(plan: dict[str, Any], builds: dict[str, dict[str, Any]],
                     cases: dict[str, dict[str, Any]], spec: dict[str, Any]) -> list[str]:
    lane = spec["lane"]
    samples = plan["samples"] if lane == "native" else plan["allocation_samples"]
    warmups = plan["warmups"] if lane == "native" else plan["allocation_warmups"]
    binary = builds[spec["variant"]]["binaries"][lane]["path"]
    case = cases[spec["case"]]
    return ["taskset", "-c", str(plan["cpu"]), binary, "--case", case["case"],
            "--input", case["path"], "--operation", "format", "--samples", str(samples),
            "--warmups", str(warmups)]


def capture_matrix(plan: dict[str, Any], builds: dict[str, dict[str, Any]],
                   cases: dict[str, dict[str, Any]], oracle: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    manifest = read_json(CAPTURES / "manifest.json")
    require(isinstance(manifest, dict) and manifest.get("status") == "complete",
            "capture manifest is incomplete")
    require(manifest.get("freeze_sha256") == sha(P / "freeze.json"), "capture freeze binding changed")
    require(manifest.get("preflight_sha256") == sha(P / "preflight.json"), "capture preflight binding changed")
    runs = manifest.get("runs")
    require(isinstance(runs, list) and len(runs) == len(plan["schedule"]),
            "capture process count changed")
    schedule = plan["schedule"]
    rows: list[dict[str, Any]] = []
    seen_files: set[str] = set()
    seen_keys: set[tuple[str, int, int, str, str]] = set()
    for index, actual in enumerate(runs):
        require(isinstance(actual, dict), f"capture row {index} is not an object")
        spec = run_spec(actual)
        planned = schedule[index]
        key = (planned["lane"], planned["cycle"], planned["repeat"], planned["case"], planned["variant"])
        actual_key = (spec.get("lane"), spec.get("cycle"), spec.get("repeat"),
                      spec.get("case"), spec.get("variant"))
        require(actual_key == key and key not in seen_keys, f"capture schedule changed at row {index}")
        seen_keys.add(key)
        require(actual.get("exit_code") == 0, f"capture child failed at row {index}")
        command = expected_command(plan, builds, cases, spec)
        require(actual.get("command") == command, f"capture command changed at row {index}")
        output_name = actual.get("output", actual.get("output_basename", actual.get("outputbasename")))
        output = capture_path(output_name, CAPTURES, f"capture output {index}")
        stderr_name = actual.get("stderr", actual.get("stderr_basename", str(output_name) + ".stderr"))
        stderr = capture_path(stderr_name, CAPTURES, f"capture stderr {index}")
        require(output.name not in seen_files and stderr.name not in seen_files,
                f"capture file reused at row {index}")
        seen_files.update((output.name, stderr.name))
        require(actual.get("sha256") == sha(output), f"capture output digest changed at row {index}")
        require(actual.get("stderr_sha256") == sha(stderr), f"capture stderr digest changed at row {index}")
        report = read_json(output)
        timing, allocation = validate_report(report, spec, cases[spec["case"]],
                                             oracle[spec["case"]]["expected"], plan)
        rows.append({"lane": spec["lane"], "cycle": spec["cycle"], "repeat": spec["repeat"],
                     "case": spec["case"], "variant": spec["variant"],
                     "timing": timing, "allocations": allocation,
                     "output_sha256": report["expected_output_sha256"]})
    require(seen_keys == {
        (row["lane"], row["cycle"], row["repeat"], row["case"], row["variant"])
        for row in schedule
    }, "capture schedule is incomplete")
    actual_files = {path.name for path in CAPTURES.iterdir() if path.is_file() and not path.is_symlink()}
    require(actual_files == seen_files | {"manifest.json"}, "capture directory has unmanifested files")
    return rows


def percent_change(before: float, after: float) -> dict[str, Any]:
    delta = after - before
    percent = None if before == 0 else delta / abs(before) * 100.0
    flag = delta != 0 if percent is None else abs(percent) > PAIR_THRESHOLD_PERCENT
    return {"before": before, "after": after, "delta": delta, "percent": percent,
            "flag_gt_5pct": flag}


def bootstrap(values: list[float], reducer: str, rng: random.Random) -> dict[str, Any] | None:
    if any(value is None for value in values):
        return None
    require(len(values) == 9, "bootstrap requires nine process pairs")
    estimates = []
    for _ in range(BOOTSTRAP_COUNT):
        sample = [values[rng.randrange(len(values))] for _ in values]
        estimates.append(statistics.median(sample) if reducer == "median" else math.fsum(sample) / len(sample))
    estimates.sort()
    return {
        "seed": BOOTSTRAP_SEED,
        "resamples": BOOTSTRAP_COUNT,
        "reducer": reducer,
        "percentile_2_5": estimates[max(0, math.ceil(0.025 * len(estimates)) - 1)],
        "percentile_97_5": estimates[max(0, math.ceil(0.975 * len(estimates)) - 1)],
    }


def native_pairs(rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    by_key = {(row["cycle"], row["repeat"], row["case"], row["variant"]): row
              for row in rows if row["lane"] == "native"}
    pairs: list[dict[str, Any]] = []
    rng = random.Random(BOOTSTRAP_SEED)
    for case_id in CASES:
        for cycle in range(3):
            for repeat in range(3):
                before = by_key[(cycle, repeat, case_id, "baseline")]["timing"]
                after = by_key[(cycle, repeat, case_id, "candidate")]["timing"]
                differences = {field: percent_change(float(before[field]), float(after[field]))
                               for field in TIMING_FIELDS}
                pairs.append({"case": case_id, "cycle": cycle, "repeat": repeat,
                              "baseline": before, "candidate": after,
                              "differences": differences,
                              "flags_gt_5pct": [field for field, value in differences.items()
                                                 if value["flag_gt_5pct"]]})
    summaries: list[dict[str, Any]] = []
    for case_id in CASES:
        selected = [row for row in pairs if row["case"] == case_id]
        metrics: dict[str, Any] = {}
        for field in ("p50", "mean"):
            changes = [row["differences"][field]["percent"] for row in selected]
            require(all(value is not None for value in changes), f"{case_id}/{field}: zero baseline timing")
            numeric = [float(value) for value in changes]
            metrics[field] = {
                "pair_percentages": numeric,
                "median_percent": statistics.median(numeric),
                "minimum_percent": min(numeric),
                "maximum_percent": max(numeric),
                "flags_gt_5pct": [index for index, value in enumerate(numeric)
                                   if abs(value) > PAIR_THRESHOLD_PERCENT],
                "bootstrap_median": bootstrap(numeric, "median", rng),
                "bootstrap_mean": bootstrap(numeric, "mean", rng),
            }
        summaries.append({"case": case_id, "pairs": 9, "metrics": metrics,
                          "review_flags": sorted({field for row in selected
                                                   for field in row["flags_gt_5pct"]})})
    return pairs, summaries


def allocation_comparisons(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for case_id in CASES:
        selected = {(row["repeat"], row["variant"]): row for row in rows
                    if row["lane"] == "allocation" and row["case"] == case_id}
        require(set(selected) == {(repeat, variant) for repeat in range(3) for variant in VARIANTS},
                f"allocation matrix incomplete: {case_id}")
        fields: dict[str, Any] = {}
        for field in ALLOC_FIELDS:
            baseline = [selected[(repeat, "baseline")]["allocations"][field] for repeat in range(3)]
            candidate = [selected[(repeat, "candidate")]["allocations"][field] for repeat in range(3)]
            pairwise = [percent_change(float(a), float(b)) for a, b in zip(baseline, candidate)]
            left = math.fsum(baseline) / 3
            right = math.fsum(candidate) / 3
            fields[field] = {
                "baseline_values": baseline,
                "candidate_values": candidate,
                "baseline_stats": stats(baseline),
                "candidate_stats": stats(candidate),
                "same_repeat": pairwise,
                "mean_comparison": percent_change(left, right),
                "review_flags": [index for index, value in enumerate(pairwise)
                                 if value["flag_gt_5pct"]],
            }
        result.append({"case": case_id, "repeats": 3, "fields": fields,
                       "review_flags": sorted({field for field, value in fields.items()
                                                if value["review_flags"]
                                                or value["mean_comparison"]["flag_gt_5pct"]})})
    return result


def verify_before(before: dict[str, Any], builds: dict[str, dict[str, Any]],
                  cases: dict[str, dict[str, Any]], oracle: dict[str, dict[str, Any]]) -> None:
    require(before.get("status") == "passed", "before receipt did not pass")
    require(before.get("build_sha256") == sha(P / "baseline-build.json"),
            "before receipt is not bound to baseline build")
    require(before.get("source") == builds["baseline"]["source"],
            "before source custody differs from baseline build")
    files = before.get("files")
    require(isinstance(files, dict) and files, "before receipt files are empty")
    for raw, expected in files.items():
        relative = safe_relative(raw, "before receipt path")
        require(sha(P / relative) == digest(expected, f"before {relative}"),
                f"before file changed: {relative}")
    # Before contains exactly the baseline primary/secondary native and
    # allocation reports plus their process receipts and stderr files.
    reports = [raw for raw in files if raw.endswith(".json") and ".receipt." not in raw]
    require(len(reports) == 4, "before report count changed")
    seen: set[tuple[str, str]] = set()
    for raw in reports:
        report = read_json(P / raw)
        case_id = next(key for key in CASES if oracle[key]["expected"]["case"] == report["case"])
        lane = "allocation" if report["allocator_instrumented"] else "native"
        plan = {"samples": 50, "warmups": 3, "allocation_samples": 1, "allocation_warmups": 0}
        row = {"lane": lane, "cycle": 0, "repeat": 0, "case": case_id, "variant": "baseline"}
        validate_report(report, row, cases[case_id], oracle[case_id]["expected"], plan)
        command_plan = {"cpu": 12, "samples": 50 if lane == "native" else 1,
                        "warmups": 3 if lane == "native" else 0,
                        "allocation_samples": 1, "allocation_warmups": 0}
        command = expected_command(command_plan, builds, cases, row)
        report_path = P / raw
        stem = report_path.name[:-5]
        receipt_path = report_path.with_name(stem + ".receipt.json")
        stderr_path = report_path.with_name(stem + ".stderr")
        require(str(receipt_path.relative_to(P)) in files and str(stderr_path.relative_to(P)) in files,
                f"before receipt/stderr missing for {raw}")
        receipt = read_json(receipt_path)
        require(receipt.get("command") == command and receipt.get("exit_code") == 0
                and receipt.get("output") == raw and receipt.get("sha256") == sha(report_path)
                and receipt.get("stderr_sha256") == sha(stderr_path),
                f"before process receipt changed: {raw}")
        seen.add((case_id, lane))
    require(seen == {(case_id, lane) for case_id in CASES for lane in LANES},
            "before matrix is incomplete")
    expected_files = set(reports)
    for raw in reports:
        stem = Path(raw).name[:-5]
        expected_files.update({str(Path(raw).with_name(stem + ".receipt.json")),
                               str(Path(raw).with_name(stem + ".stderr"))})
    require(set(files) == expected_files, "before file inventory changed")
    order = before.get("candidate_order")
    require(isinstance(order, list) and order == [
        "test-data/poi/test-data/slideshow/41246-1.ppt",
        "test-data/office-interop/libreoffice-resaved/45543-transition-litchi.ppt",
        "test-data/ole/ppt/SampleShow.ppt",
    ] and cases["secondary"]["path"] == order[0],
            "secondary fixture selection order changed")


def main() -> None:
    plan = load_plan()
    cases = load_cases()
    verify_constraints()
    builds = {variant: load_build(variant) for variant in VARIANTS}
    verify_root_quality(builds)
    verify_binary_receipts(builds)
    verify_source_bridge(builds)
    verify_source_review(builds)
    verify_freeze(builds)
    verify_preflight()
    oracle = oracle_rows()
    before = read_json(P / "before.json")
    verify_before(before, builds, cases, oracle)
    verify_qualification_files("qualification.json", oracle, builds, "candidate")
    rows = capture_matrix(plan, builds, cases, oracle)
    native = [row for row in rows if row["lane"] == "native"]
    allocation = [row for row in rows if row["lane"] == "allocation"]
    require(len(native) == 36 and len(allocation) == 12, "capture lane counts changed")
    pairs, summaries = native_pairs(rows)
    allocations = allocation_comparisons(rows)
    result = {
        "schema_version": 1,
        "packet": "change-0734",
        "status": "passed",
        "disposition": "root review required; no automatic acceptance",
        "matrix": {"native_processes": len(native), "allocation_processes": len(allocation),
                   "cases": list(CASES), "variants": list(VARIANTS),
                   "native_samples": plan["samples"], "native_warmups": plan["warmups"],
                   "allocation_samples": plan["allocation_samples"],
                   "allocation_warmups": plan["allocation_warmups"]},
        "statistics": {"timing_unit": "ns", "quantiles": "midpoint p50; nearest-rank p95/p99",
                        "timing_fields": list(TIMING_FIELDS), "allocation_fields": list(ALLOC_FIELDS),
                        "paired_threshold_percent": PAIR_THRESHOLD_PERCENT},
        "native_processes": native,
        "native_pairs": pairs,
        "native_case_summaries": summaries,
        "allocation_processes": allocation,
        "allocation_comparisons": allocations,
        "review_flags": {"native": sorted({field for row in pairs for field in row["flags_gt_5pct"]}),
                         "allocation": sorted({field for row in allocations for field in row["review_flags"]})},
        "limitations": [
            "nine paired process repeats per case are the independent comparison units; samples within a process are not independent processes",
            "the matrix is one public PPT slide-removal workflow on two fixtures and does not establish cold-I/O, RSS, concurrency, or broad CRUD behavior",
            "all retained samples, tails, raw allocation values, and review flags are reported; no selective rerun or automatic acceptance is performed",
        ],
    }
    (P / "analysis.json").write_text(json.dumps(result, indent=2, sort_keys=True) + "\n",
                                       encoding="utf-8")
    print("PASS 0734 custody, exact PPT oracles, 36 native processes, 12 allocation processes, paired bootstrap statistics")


if __name__ == "__main__":
    try:
        main()
    except (Failure, OSError, KeyError, TypeError, ValueError) as error:
        raise SystemExit(f"FAIL: {error}")
