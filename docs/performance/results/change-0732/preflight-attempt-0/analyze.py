#!/usr/bin/env python3
"""Offline validation and statistics for the 0732 PPT phase packet.

This program consumes the retained packet and native JSON reports.  It never
starts Cargo, the probe, or a profiler.  The packet is intentionally checked
before timings are read: a fast number from a report with a changed fixture,
binary, route, or semantic oracle is not evidence.
"""

from __future__ import annotations

import hashlib
import json
import math
import statistics
from pathlib import Path
from typing import Any, Iterable


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
CAPTURES = P / "captures"
HEX = set("0123456789abcdef")
ROUTES = ("ordinary-opaque", "ordinary-split", "profiled-empty", "profiled-clock")
SPLIT_ROUTES = set(ROUTES[1:])
STAT_FIELDS = ("p50", "mean", "p95", "p99", "maximum")


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
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            f"{label} is not a lowercase SHA-256 digest")
    return value


def integer(value: Any, label: str, minimum: int = 0) -> int:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= minimum,
            f"{label} is not an integer >= {minimum}")
    return value


def safe_relative(value: Any, label: str) -> str:
    require(isinstance(value, str) and value and not Path(value).is_absolute()
            and ".." not in Path(value).parts, f"{label} is unsafe: {value!r}")
    return value


def stats(values: Iterable[int | float]) -> dict[str, float | int]:
    values = list(values)
    require(values, "statistics received no values")
    ordered = sorted(values)
    return {
        "n": len(ordered),
        "p50": statistics.median(ordered),
        "mean": statistics.fmean(ordered),
        "p95": ordered[max(0, math.ceil(0.95 * len(ordered)) - 1)],
        "p99": ordered[max(0, math.ceil(0.99 * len(ordered)) - 1)],
        "maximum": ordered[-1],
    }


def resolve_root_path(raw: str) -> Path:
    path = Path(raw)
    if path.is_absolute():
        return path
    return ROOT / path


def load_plan() -> dict[str, Any]:
    plan = read_json(P / "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("cpu") == 12, "CPU binding changed")
    require(plan.get("cycles") == 3 and plan.get("repeats") == 3,
            "cycle or repeat plan changed")
    require(plan.get("samples") == 50 and plan.get("warmups") == 3,
            "sample plan changed")
    require(plan.get("routes") == list(ROUTES), "route order changed")
    if "comparisons" in plan:
        require(plan.get("comparisons") == [[ROUTES[0], ROUTES[1]],
                                             [ROUTES[1], ROUTES[2]],
                                             [ROUTES[2], ROUTES[3]]],
                "comparison chain changed")
    require(plan.get("processes") == 36 and plan.get("measured_lifecycles") == 1800,
            "process or lifecycle count changed")
    schedule = plan.get("schedule")
    require(isinstance(schedule, list) and len(schedule) == 36, "schedule is not 36 cells")
    keys: set[tuple[int, int, str]] = set()
    for index, item in enumerate(schedule):
        require(isinstance(item, dict), f"schedule item {index} is not an object")
        key = (item.get("cycle"), item.get("repeat"), item.get("route"))
        require(isinstance(key[0], int) and isinstance(key[1], int)
                and key[0] in range(3) and key[1] in range(3)
                and key[2] in ROUTES, f"invalid schedule item {index}: {item!r}")
        require(key not in keys, f"duplicate schedule cell: {key!r}")
        keys.add(key)
    require(len(keys) == 36, "schedule does not cover all route cells")
    return plan


def load_case() -> dict[str, Any]:
    case = read_json(P / "case.json")
    require(isinstance(case, dict), "case is not an object")
    require(case.get("case") == "ppt45543" and case.get("format") == "ppt",
            "PPT case identity changed")
    path = safe_relative(case.get("path"), "case path")
    fixture = ROOT / path
    require(fixture.is_file() and not fixture.is_symlink(), f"missing PPT fixture: {fixture}")
    require(case.get("bytes") == fixture.stat().st_size, "PPT fixture size changed")
    require(case.get("sha256") == sha(fixture), "PPT fixture digest changed")
    return case


def verify_constraints() -> None:
    constraints = read_json(P / "constraints.json")
    require(isinstance(constraints, dict) and constraints, "constraints are empty")
    for raw, expected in constraints.items():
        relative = safe_relative(raw, "constraint path")
        digest(expected, f"constraint {relative}")
        target = ROOT / relative
        require(target.is_file() and not target.is_symlink(), f"missing constraint: {relative}")
        require(sha(target) == expected, f"constraint changed: {relative}")


def verify_quality(build: dict[str, Any]) -> dict[str, Any]:
    quality_name = safe_relative(build.get("quality"), "quality path")
    quality_sha = digest(build.get("quality_sha256"), "quality digest")
    quality_path = P / quality_name
    require(sha(quality_path) == quality_sha, "quality manifest digest changed")
    quality = read_json(quality_path)
    require(isinstance(quality, dict), "quality manifest is not an object")
    require(quality.get("source") == build.get("source"), "quality source custody differs")
    require(quality.get("probe") == build.get("probe"), "quality probe custody differs")
    runs = quality.get("runs")
    require(isinstance(runs, list) and len(runs) == 12, "quality command count changed")
    manifest_commands = expected_quality_commands()
    for index, row in enumerate(runs):
        require(isinstance(row, dict), f"quality row {index} is not an object")
        require(row.get("exit_code") == 0, f"quality command {index} failed")
        require(row.get("command") == manifest_commands[index],
                f"quality command {index} identity changed")
        output = safe_relative(row.get("output"), f"quality output {index}")
        output_path = P / output
        require(sha(output_path) == digest(row.get("sha256"), f"quality output {index}"),
                f"quality output {index} digest changed")
        require(isinstance(row.get("command"), list) and row["command"],
                f"quality command {index} is missing")
    return quality


def expected_quality_commands() -> list[list[str]]:
    owner = ["-p", "litchi-ppt", "--release", "--offline", "--locked"]
    feature = ["--features", "performance-diagnostics"]
    manifest = str(P / "probe" / "Cargo.toml")
    common = ["--manifest-path", manifest, "--release", "--offline", "--locked"]
    return [
        ["cargo", "fmt", "-p", "litchi-ppt", "--", "--check"],
        ["cargo", "check", *owner, "--no-default-features"],
        ["cargo", "test", *owner, *feature, "--all-targets"],
        ["cargo", "clippy", *owner, *feature, "--all-targets", "--", "-D", "warnings"],
        ["cargo", "test", *owner, *feature, "--doc"],
        ["cargo", "doc", *owner, *feature, "--no-deps"],
        ["cargo", "fmt", "--manifest-path", manifest, "--", "--check"],
        ["cargo", "test", *common, "--lib"],
        ["cargo", "clippy", *common, "--all-targets", "--", "-D", "warnings"],
        ["cargo", "doc", *common, "--no-deps"],
        ["cargo", "build", *common, "--bin", "ppt_phase_probe"],
        ["python3", "tools/check_crate_boundaries.py"],
    ]


def verify_source_map(mapping: Any, root: Path, label: str) -> dict[str, str]:
    require(isinstance(mapping, dict) and mapping, f"{label} custody map is empty")
    checked: dict[str, str] = {}
    for raw, expected in mapping.items():
        relative = safe_relative(raw, f"{label} path")
        expected = digest(expected, f"{label} {relative}")
        target = root / relative
        require(target.is_file() and not target.is_symlink(), f"{label} missing: {relative}")
        require(sha(target) == expected, f"{label} changed: {relative}")
        checked[relative] = expected
    return checked


def verify_build() -> tuple[dict[str, Any], dict[str, Any]]:
    build = read_json(P / "build.json")
    require(isinstance(build, dict), "build is not an object")
    source = verify_source_map(build.get("source"), ROOT, "build source")
    probe = verify_source_map(build.get("probe"), P, "build probe")
    require(build.get("source") == source and build.get("probe") == probe,
            "build custody map changed during validation")
    quality = verify_quality(build)
    binary = build.get("binary")
    require(isinstance(binary, dict), "build binary is not an object")
    binary_path = binary.get("path")
    require(isinstance(binary_path, str) and Path(binary_path).is_absolute(),
            "build binary path is not absolute")
    binary_bytes = integer(binary.get("bytes"), "build binary bytes", 1)
    binary_sha = digest(binary.get("sha256"), "build binary digest")
    path = Path(binary_path)
    cleanup = read_json(P / "cleanup.json") if (P / "cleanup.json").is_file() else None
    if path.is_file():
        require(not path.is_symlink() and path.stat().st_size == binary_bytes
                and sha(path) == binary_sha, "live binary identity changed")
    else:
        require(cleanup and cleanup.get("removed") is True,
                "missing binary has no cleanup receipt")
        require(cleanup.get("binaries") == [binary], "cleanup binary identity changed")
    return build, quality


def verify_source_review(build: dict[str, Any]) -> dict[str, Any]:
    review = read_json(P / "source-review.json")
    require(isinstance(review, dict), "source review is not an object")
    require(review.get("base_head") == read_json(P / "plan.json").get("base_head"),
            "source review base head changed")
    require(review.get("ordinary_commit_unchanged") is True,
            "ordinary PPT commit is not recorded unchanged")
    digest(review.get("ordinary_commit_sha256"), "ordinary commit digest")
    files = review.get("files")
    require(isinstance(files, dict) and files, "source review file map is empty")
    for raw, detail in files.items():
        relative = safe_relative(raw, "source review path")
        require(isinstance(detail, dict), f"source review row is not an object: {relative}")
        before = digest(detail.get("before_sha256"), f"source review before {relative}")
        after = digest(detail.get("after_sha256"), f"source review after {relative}")
        require(build["source"].get(relative) == after,
                f"build source does not match reviewed after source: {relative}")
        require(sha(ROOT / relative) == after, f"live reviewed source changed: {relative}")
        for side, expected in (("before", before), ("after", after)):
            archive = P / "source-archive" / side / relative
            if archive.exists():
                require(archive.is_file() and not archive.is_symlink()
                        and sha(archive) == expected,
                        f"source archive changed: {side}/{relative}")
    return review


def verify_freeze(build: dict[str, Any]) -> dict[str, str]:
    frozen = read_json(P / "freeze.json")
    require(isinstance(frozen, dict) and frozen, "freeze map is empty")
    binary_path = Path(build["binary"]["path"])
    cleanup = read_json(P / "cleanup.json") if (P / "cleanup.json").is_file() else None
    missing_binary_allowed = False
    for raw, expected in frozen.items():
        require(isinstance(raw, str) and Path(raw).is_absolute(),
                f"freeze path is not absolute: {raw!r}")
        expected = digest(expected, f"freeze {raw}")
        target = Path(raw)
        if target.is_file():
            require(not target.is_symlink() and sha(target) == expected,
                    f"frozen input changed: {raw}")
        else:
            require(target == binary_path, f"missing non-binary frozen input: {raw}")
            require(cleanup and cleanup.get("removed") is True
                    and cleanup.get("binaries") == [build["binary"]],
                    "missing frozen binary lacks exact cleanup witness")
            require(expected == build["binary"]["sha256"], "frozen binary digest changed")
            missing_binary_allowed = True
    require(any(Path(raw) == binary_path for raw in frozen),
            "freeze omits the measured binary")
    return frozen


def verify_preflight() -> dict[str, Any]:
    preflight = read_json(P / "preflight.json")
    require(isinstance(preflight, dict) and preflight.get("status") == "passed",
            "preflight did not pass")
    require(preflight.get("freeze_sha256") == sha(P / "freeze.json"),
            "preflight is not bound to freeze")
    scripts = preflight.get("scripts")
    require(isinstance(scripts, dict)
            and set(scripts) == {"preflight.py", "analyze.py", "audit.py"},
            "preflight script binding is incomplete")
    for name, expected in scripts.items():
        require(sha(P / name) == digest(expected, f"preflight {name}"),
                f"preflight script changed: {name}")
    return preflight


def verify_qualification(build: dict[str, Any], expected: dict[str, Any]) -> None:
    qualification = read_json(P / "qualification.json")
    require(isinstance(qualification, dict) and qualification.get("status") == "passed",
            "qualification did not pass")
    require(qualification.get("build_sha256") == sha(P / "build.json"),
            "qualification is not bound to build")
    files = qualification.get("files")
    require(isinstance(files, dict) and files, "qualification file map is empty")
    for raw, expected_sha in files.items():
        relative = safe_relative(raw, "qualification path")
        require(sha(P / relative) == digest(expected_sha, f"qualification {relative}"),
                f"qualification file changed: {relative}")
    manifest = read_json(P / "qualification" / "manifest.json")
    require(manifest.get("status") == "passed", "qualification manifest did not pass")
    runs = manifest.get("runs")
    require(isinstance(runs, list) and len(runs) == 4, "qualification route count changed")
    seen: set[str] = set()
    for row in runs:
        require(row.get("route") in ROUTES and row["route"] not in seen,
                "qualification route set changed")
        seen.add(row["route"])
        require(row.get("exit_code") == 0, "qualification route failed")
        output = P / "qualification" / row["output"]
        report = read_json(output)
        require(report.get("expected_oracle") == expected["expected_oracle"],
                f"qualification oracle changed: {row['route']}")
        require(report.get("expected_output_inventory") == expected["expected_output_inventory"],
                f"qualification output inventory changed: {row['route']}")
        require(len(report.get("samples", [])) == 1, "qualification sample count changed")
        sample = report["samples"][0]
        require(sample.get("output_sha256") == expected["expected_output_sha256"]
                and sample.get("output_inventory") == expected["expected_output_inventory"]
                and sample.get("oracle") == expected["expected_oracle"],
                f"qualification sample oracle changed: {row['route']}")
    require(seen == set(ROUTES), "qualification did not cover all routes")


def load_oracle() -> dict[str, Any]:
    wrapper = read_json(P / "oracle.json")
    require(isinstance(wrapper, dict), "oracle wrapper is not an object")
    reference = safe_relative(wrapper.get("reference"), "oracle reference")
    reference_path = ROOT / reference
    expected_sha = digest(wrapper.get("sha256"), "oracle reference digest")
    require(sha(reference_path) == expected_sha, "sealed oracle reference changed")
    expected = wrapper.get("expected")
    require(isinstance(expected, dict), "oracle expected report is not an object")
    require(expected.get("case") == "ppt45543" and expected.get("format") == "ppt"
            and expected.get("operation") == "format", "oracle identity changed")
    for key in ("source_inventory", "expected_output_inventory", "replacements",
                "changed_length_proof", "expected_oracle", "oracle_controls"):
        require(key in expected, f"oracle omits {key}")
    return wrapper


def oracle_booleans(value: Any, label: str) -> None:
    require(isinstance(value, dict), f"{label} is not an object")
    require(value.get("oracle_ok") is True, f"{label}.oracle_ok is false")
    require(value.get("failure_reasons") == [], f"{label} has failure reasons")
    for key, item in value.items():
        if isinstance(item, bool):
            require(item is True, f"{label}.{key} is false")
        elif isinstance(item, dict):
            # Semantic witnesses contain booleans in nested records in some
            # probe versions; check them without treating integer fields as bools.
            oracle_booleans_nested(item, f"{label}.{key}")
        elif isinstance(item, list):
            for index, nested in enumerate(item):
                if isinstance(nested, dict):
                    oracle_booleans_nested(nested, f"{label}.{key}[{index}]")


def oracle_booleans_nested(value: dict[str, Any], label: str) -> None:
    for key, item in value.items():
        if isinstance(item, bool):
            require(item is True, f"{label}.{key} is false")
        elif isinstance(item, dict):
            oracle_booleans_nested(item, f"{label}.{key}")
        elif isinstance(item, list):
            for index, nested in enumerate(item):
                if isinstance(nested, dict):
                    oracle_booleans_nested(nested, f"{label}.{key}[{index}]")


STATIC_REPORT_KEYS = (
    "case", "format", "operation", "input", "policy", "policy_applied",
    "policy_application_scope", "policy_argument_effect", "policy_contract",
    "directory_metadata_fields", "source_sha256",
    "expected_output_sha256", "replacements_sha256", "source_inventory",
    "expected_output_inventory", "replacements", "changed_length_proof",
    "expected_oracle", "oracle_controls",
)


def verify_report_header(report: dict[str, Any], row: dict[str, Any], expected: dict[str, Any],
                         case: dict[str, Any], plan: dict[str, Any]) -> None:
    label = f"c{row['cycle']}-r{row['repeat']}-{row['route']}"
    require(report.get("schema_version") == expected.get("schema_version") == 1,
            f"{label}: schema changed")
    for key in STATIC_REPORT_KEYS:
        require(report.get(key) == expected.get(key), f"{label}: {key} differs from oracle")
    require(report.get("mode") == "ppt_public_phase_attribution"
            and report.get("scope") == "public_ppt_open_edit_remove_commit_output_copy",
            f"{label}: phase probe identity changed")
    for key in ("phase_contract", "diagnostic_contract", "observer_contract",
                "allocation_ownership_contract"):
        require(isinstance(report.get(key), str) and report[key],
                f"{label}: missing {key}")
    require(report.get("input") == case["path"] and report.get("source_sha256") == case["sha256"],
            f"{label}: fixture identity changed")
    require(report.get("timing_claim") is True and report.get("allocator_instrumented") is False,
            f"{label}: timing/allocator mode changed")
    require(report.get("warmups") == plan["warmups"]
            and report.get("samples_requested") == plan["samples"],
            f"{label}: sample header changed")
    require(report.get("route") == row["route"], f"{label}: report route differs from manifest")
    oracle_booleans(report["expected_oracle"], f"{label}: expected oracle")
    controls = report["oracle_controls"]
    require(isinstance(controls, list) and controls == expected["oracle_controls"],
            f"{label}: oracle controls changed")
    for control in controls:
        require(control.get("rejected") is True and control.get("status") == "rejected"
                and isinstance(control.get("failure_reasons"), list)
                and control["failure_reasons"],
                f"{label}: corruption control was accepted")


def phase_map(sample: dict[str, Any], label: str) -> dict[str, int]:
    whole = sample.get("whole_ns")
    require(isinstance(whole, int) and not isinstance(whole, bool) and whole > 0,
            f"{label}: whole_ns is invalid")
    split = sample.get("split")
    if split is None:
        return {"whole_ns": whole}
    require(isinstance(split, dict), f"{label}: split is not an object")
    component_names = ("open_ns", "edit_ns", "remove_ns", "commit_ns", "output_copy_ns")
    result = {name: integer(split.get(name), f"{label}: {name}") for name in component_names}
    require(result["commit_ns"] > 0, f"{label}: commit_ns is not positive")
    split_sum = integer(split.get("split_sum_ns"), f"{label}: split_sum_ns")
    residual = integer(split.get("whole_residual_ns"), f"{label}: whole_residual_ns")
    require(split_sum == sum(result.values()), f"{label}: split_sum_ns differs from phase sum")
    require(split_sum + residual == whole,
            f"{label}: phase sum plus residual does not equal whole_ns")
    start = integer(split.get("commit_start_ns"), f"{label}: commit_start_ns")
    end = integer(split.get("commit_end_ns"), f"{label}: commit_end_ns", start)
    require(end <= whole and end - start == result["commit_ns"],
            f"{label}: commit window does not equal commit_ns")
    result["whole_residual_ns"] = residual
    return {"whole_ns": whole, **result}


def phase_endpoints(sample: dict[str, Any]) -> tuple[int | None, int | None]:
    split = sample.get("split")
    if not isinstance(split, dict):
        return None, None
    return split.get("commit_start_ns"), split.get("commit_end_ns")


def diagnostic_object(sample: dict[str, Any]) -> dict[str, Any] | None:
    value = sample.get("diagnostics")
    return value if isinstance(value, dict) else None


EXPECTED_PHASES = ("DocumentCommit", "BeforePayloadCapture", "EmbeddedOpen", "LiveDocumentRead",
                   "EmbeddedFinish", "UnrelatedStreamValidation", "PublicReopen",
                   "AfterPayloadCapture", "ArtifactHashBefore", "ArtifactHashAfter")
EXPECTED_EVENTS = tuple(value for phase in EXPECTED_PHASES for value in (f"started:{phase}", f"finished:{phase}"))


def validate_diagnostic(sample: dict[str, Any], phase: dict[str, int], label: str,
                        expected_names: tuple[str, ...] | None) -> tuple[tuple[str, ...], dict[str, Any]]:
    diag = diagnostic_object(sample)
    require(diag is not None and isinstance(diag.get("commit"), dict),
            f"{label}: profiled-clock diagnostic is missing")
    report = diag["commit"]
    require(report.get("event_count") == 20 and report.get("overflow") is False
            and report.get("balanced") is True and report.get("sequence_ok") is True
            and report.get("outcomes_ok") is True and report.get("timestamps_monotonic") is True,
            f"{label}: diagnostic validation flags changed")
    require(report.get("expected_phases") == list(EXPECTED_PHASES),
            f"{label}: expected diagnostic phases changed")
    events = report.get("events")
    require(isinstance(events, list) and len(events) == 20,
            f"{label}: diagnostic event count changed")
    names: list[str] = []
    previous = 0
    for index, event in enumerate(events):
        require(isinstance(event, dict), f"{label}: event {index} is not an object")
        phase_name = EXPECTED_PHASES[index // 2]
        kind = "started" if index % 2 == 0 else "finished"
        outcome = "started" if kind == "started" else "success"
        require(event.get("kind") == kind and event.get("phase") == phase_name
                and event.get("outcome") == outcome,
                f"{label}: event {index} sequence/outcome changed")
        timestamp = integer(event.get("t_ns"), f"{label}: event {index} timestamp")
        require(timestamp >= previous, f"{label}: event timestamps are not monotonic")
        previous = timestamp
        names.append(f"{kind}:{phase_name}")
    spans = report.get("spans")
    require(isinstance(spans, list) and len(spans) == 10,
            f"{label}: diagnostic span count changed")
    durations = 0
    phase_durations: dict[str, int] = {}
    for index, span in enumerate(spans):
        require(isinstance(span, dict) and span.get("phase") == EXPECTED_PHASES[index]
                and span.get("outcome") == "success", f"{label}: span {index} changed")
        start = integer(span.get("start_ns"), f"{label}: span {index} start")
        end = integer(span.get("finish_ns"), f"{label}: span {index} finish", start)
        duration = integer(span.get("duration_ns"), f"{label}: span {index} duration")
        require(duration == end - start, f"{label}: span {index} duration changed")
        require(start == events[index * 2]["t_ns"] and end == events[index * 2 + 1]["t_ns"],
                f"{label}: span {index} does not match events")
        durations += duration
        phase_durations[span["phase"]] = duration
    names = tuple(names)
    if expected_names is not None:
        require(names == expected_names, f"{label}: diagnostic event sequence changed")
    start, end = phase_endpoints(sample)
    require(start is not None and end is not None, f"{label}: commit window is missing")
    require(durations <= phase["commit_ns"], f"{label}: event durations exceed commit_ns")
    require(all(start <= event["t_ns"] <= end for event in events),
            f"{label}: diagnostic event escapes commit window")
    return names, {"commit_ns": phase["commit_ns"], "commit_start_ns": start,
                   "commit_end_ns": end, "event_count": 20,
                   "event_names": list(names), "event_durations_ns": durations,
                   "phase_durations_ns": phase_durations,
                   "whole_ns": phase["whole_ns"]}


def validate_sample(sample: dict[str, Any], index: int, row: dict[str, Any],
                    expected: dict[str, Any], plan: dict[str, Any],
                    diag_names: tuple[str, ...] | None) -> tuple[dict[str, int], tuple[str, ...] | None, dict[str, Any] | None]:
    label = f"{row['route']} sample {index} c{row['cycle']}r{row['repeat']}"
    require(sample.get("index") == index, f"{label}: sample index changed")
    require(sample.get("route") == row["route"], f"{label}: sample route changed")
    require(sample.get("output_sha256") == expected["expected_output_sha256"],
            f"{label}: output digest changed")
    require(sample.get("output_inventory") == expected["expected_output_inventory"],
            f"{label}: output inventory changed")
    require(sample.get("oracle") == expected["expected_oracle"], f"{label}: semantic oracle changed")
    oracle_booleans(sample["oracle"], f"{label}: sample oracle")
    phase = phase_map(sample, label)
    keys = set(phase)
    if row["route"] == "ordinary-opaque":
        require(keys == {"whole_ns"}, f"{label}: opaque route exposes split phases")
        require(diag_names is None and diagnostic_object(sample) is None
                and sample.get("observer_clock_control_ns") is None,
                f"{label}: opaque route contains diagnostics")
        return phase, diag_names, None
    require("whole_residual_ns" in phase and len(keys - {"whole_ns", "whole_residual_ns"}) >= 1,
            f"{label}: split route lacks phases or residual")
    if row["route"] == "profiled-clock":
        require(integer(sample.get("observer_clock_control_ns"),
                       f"{label}: observer clock control") >= 0,
                f"{label}: observer clock control is invalid")
        names, diagnostic = validate_diagnostic(sample, phase, label, diag_names)
        return phase, names, diagnostic
    require(diagnostic_object(sample) is None and sample.get("observer_clock_control_ns") is None,
            f"{label}: empty/ordinary split route contains diagnostics")
    return phase, diag_names, None


def expected_commands(plan: dict[str, Any], build: dict[str, Any], case: dict[str, Any]) -> dict[tuple[int, int, str], list[str]]:
    result: dict[tuple[int, int, str], list[str]] = {}
    binary = build["binary"]["path"]
    for item in plan["schedule"]:
        key = (item["cycle"], item["repeat"], item["route"])
        result[key] = [
            "taskset", "-c", str(plan["cpu"]), binary,
            "--route", item["route"], "--input", case["path"],
            "--samples", str(plan["samples"]), "--warmups", str(plan["warmups"]),
        ]
    return result


def verify_capture_manifest(plan: dict[str, Any], build: dict[str, Any], case: dict[str, Any],
                            expected: dict[str, Any]) -> tuple[list[dict[str, Any]], dict[tuple[int, int, str], dict[str, Any]]]:
    manifest = read_json(CAPTURES / "manifest.json")
    require(isinstance(manifest, dict) and manifest.get("status") == "complete",
            "capture manifest is not complete")
    require(manifest.get("freeze_sha256") == sha(P / "freeze.json"),
            "capture is not bound to freeze")
    preflight = verify_preflight()
    require(manifest.get("preflight_sha256") == sha(P / "preflight.json")
            and preflight.get("freeze_sha256") == sha(P / "freeze.json"),
            "capture is not bound to preflight")
    runs = manifest.get("runs")
    require(isinstance(runs, list) and len(runs) == 36, "capture process count changed")
    commands = expected_commands(plan, build, case)
    schedule_index = {
        (item["cycle"], item["repeat"], item["route"]): index
        for index, item in enumerate(plan["schedule"])
    }
    expected_by_key = commands
    seen: set[tuple[int, int, str]] = set()
    process_reports: list[dict[str, Any]] = []
    rows_by_key: dict[tuple[int, int, str], dict[str, Any]] = {}
    for index, row in enumerate(runs):
        require(isinstance(row, dict), f"capture row {index} is not an object")
        key = (row.get("cycle"), row.get("repeat"), row.get("route"))
        require(key in expected_by_key and key not in seen, f"invalid or duplicate capture row {index}")
        seen.add(key)
        require(index == schedule_index[key],
                f"capture row {index} changed schedule order")
        require(row.get("exit_code") == 0 and row.get("command") == commands[key],
                f"capture row {index} command or exit status changed")
        output_name = row.get("output")
        expected_name = f"c{key[0]}-r{key[1]}-{key[2]}.json"
        require(output_name == expected_name, f"capture row {index} output name changed")
        output_name = safe_relative(output_name, f"capture output {index}")
        output_path = CAPTURES / output_name
        stderr_path = CAPTURES / (output_name + ".stderr")
        require(sha(output_path) == digest(row.get("sha256"), f"capture output {index}"),
                f"capture output {index} digest changed")
        require(sha(stderr_path) == digest(row.get("stderr_sha256"), f"capture stderr {index}"),
                f"capture stderr {index} digest changed")
        report = read_json(output_path)
        verify_report_header(report, row, expected, case, plan)
        samples = report.get("samples")
        require(isinstance(samples, list) and len(samples) == plan["samples"],
                f"capture row {index} sample count changed")
        process_reports.append(dict(row=row, report=report))
        rows_by_key[key] = dict(row=row, report=report)
    require(seen == set(expected_by_key), "capture schedule is incomplete")
    files = {path.name for path in CAPTURES.iterdir() if path.is_file() and not path.is_symlink()}
    required_files = {"manifest.json"}
    for item in runs:
        required_files.add(item["output"])
        required_files.add(item["output"] + ".stderr")
    require(files == required_files, f"capture directory has unexpected files: {sorted(files ^ required_files)}")
    return process_reports, rows_by_key


def analyze_processes(plan: dict[str, Any], process_reports: list[dict[str, Any]], expected: dict[str, Any]) -> tuple[list[dict[str, Any]], tuple[str, ...]]:
    processes: list[dict[str, Any]] = []
    event_names: tuple[str, ...] | None = None
    for item in process_reports:
        row = item["row"]
        report = item["report"]
        phase_rows: list[dict[str, int]] = []
        diagnostics: list[dict[str, Any]] = []
        observer_controls: list[int] = []
        local_names: tuple[str, ...] | None = None
        for index, sample in enumerate(report["samples"]):
            phase, local_names, diagnostic = validate_sample(
                sample, index, row, expected, plan, local_names)
            phase_rows.append(phase)
            if diagnostic is not None:
                diagnostics.append(diagnostic)
            if row["route"] == "profiled-clock":
                observer_controls.append(integer(sample.get("observer_clock_control_ns"),
                                                 "profiled-clock observer control"))
        if row["route"] == "profiled-clock":
            require(local_names is not None, "profiled-clock has no diagnostic event names")
            if event_names is None:
                event_names = local_names
            require(local_names == event_names, "profiled-clock event names differ between processes")
        else:
            require(local_names is None, f"{row['route']} unexpectedly contains diagnostic names")
        phase_keys = sorted(set().union(*(set(value) for value in phase_rows)))
        timing = {key: stats(value[key] for value in phase_rows) for key in phase_keys}
        ratios = {}
        for key in phase_keys:
            if key != "whole_ns":
                ratios[key] = stats(value[key] / value["whole_ns"] for value in phase_rows)
        process = {
            "cycle": row["cycle"], "repeat": row["repeat"], "route": row["route"],
            "samples": plan["samples"], "timing_ns": timing,
            "phase_ratios": ratios,
        }
        if diagnostics:
            phase_duration_stats = {
                phase_name: stats(value["phase_durations_ns"][phase_name] for value in diagnostics)
                for phase_name in EXPECTED_PHASES
            }
            phase_whole_ratios = {
                phase_name: stats(value["phase_durations_ns"][phase_name] / value["whole_ns"]
                                  for value in diagnostics)
                for phase_name in EXPECTED_PHASES
            }
            phase_commit_ratios = {
                phase_name: stats(value["phase_durations_ns"][phase_name] / value["commit_ns"]
                                  for value in diagnostics)
                for phase_name in EXPECTED_PHASES
            }
            process["diagnostic"] = {
                "samples": len(diagnostics),
                "event_names": list(event_names or local_names or ()),
                "commit_ns": stats(value["commit_ns"] for value in diagnostics),
                "commit_window_ns": stats(value["commit_end_ns"] - value["commit_start_ns"] for value in diagnostics),
                "total_event_duration_ns": stats(value["event_durations_ns"] for value in diagnostics),
                "phase_durations_ns": phase_duration_stats,
                "phase_whole_ratios": phase_whole_ratios,
                "phase_commit_ratios": phase_commit_ratios,
                "commit_residual_ns": stats(value["commit_ns"] - value["event_durations_ns"]
                                             for value in diagnostics),
                "non_commit_ns": stats(value["whole_ns"] - value["commit_ns"] for value in diagnostics),
                "commit_whole_ratios": stats(value["commit_ns"] / value["whole_ns"] for value in diagnostics),
            }
            process["observer_clock_control_ns"] = stats(observer_controls)
        processes.append(process)
    require(event_names is not None and len(event_names) == 20,
            "no complete profiled-clock diagnostic event schema")
    return processes, event_names


def comparisons(plan: dict[str, Any], processes: list[dict[str, Any]]) -> list[dict[str, Any]]:
    by_key = {(row["cycle"], row["repeat"], row["route"]): row for row in processes}
    output: list[dict[str, Any]] = []
    for round_spec in ((cycle, repeat) for cycle in range(3) for repeat in range(3)):
        cycle, repeat = round_spec
        # The capture order is deliberately rotated to spread scheduler and
        # thermal position effects.  Comparisons use the same fixed semantic
        # chain in every round so route deltas retain one interpretation.
        for left, right in zip(ROUTES, ROUTES[1:]):
            a = by_key[(cycle, repeat, left)]["timing_ns"]
            b = by_key[(cycle, repeat, right)]["timing_ns"]
            delta = {}
            flags = []
            for field in STAT_FIELDS:
                left_value = a["whole_ns"][field]
                right_value = b["whole_ns"][field]
                percent = 100.0 * (right_value / left_value - 1.0) if left_value else 0.0
                delta[field] = {"left": left_value, "right": right_value, "percent": percent}
                if field in ("p50", "mean") and abs(percent) > 5.0:
                    flags.append(field)
            output.append({"cycle": cycle, "repeat": repeat, "left": left, "right": right,
                           "delta": delta, "observer_flags": flags})
    require(len(output) == 27, "adjacent route comparison count changed")
    return output


def main() -> None:
    plan = load_plan()
    case = load_case()
    verify_constraints()
    build, _quality = verify_build()
    verify_source_review(build)
    verify_freeze(build)
    oracle = load_oracle()
    expected = oracle["expected"]
    verify_qualification(build, expected)
    reports, _rows = verify_capture_manifest(plan, build, case, expected)
    processes, event_names = analyze_processes(plan, reports, expected)
    result = {
        "schema_version": 1,
        "status": "passed",
        "disposition": "native phase attribution only; no runtime optimization, output, or speedup claim",
        "process_count": len(processes),
        "measured_lifecycles": plan["measured_lifecycles"],
        "diagnostic_event_names": list(event_names),
        "processes": processes,
        "adjacent_route_comparisons": comparisons(plan, processes),
        "observer_threshold_percent": 5,
        "limitations": [
            "ordinary-opaque and ordinary-split are separate diagnostic routes and may differ in code generation, lifetimes, and compiler scheduling",
            "profiled-empty includes the split instrumentation and feature-enabled code shape even without diagnostic clock events",
            "profiled-clock phase spans describe the feature-gated diagnostic build and do not establish default-build behavior",
            "the fixture is one public PPT edit; results do not establish cold-I/O, allocation, RSS, concurrency, or broad CRUD behavior",
            "all route samples, tails, and observer flags are retained; no selective rerun or trimming is performed",
        ],
    }
    (P / "analysis.json").write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    print("PASS 36 native processes, exact PPT oracle, split phase sums, and 20-event diagnostic trace")


if __name__ == "__main__":
    try:
        main()
    except Failure as error:
        raise SystemExit(f"FAIL: {error}")
