#!/usr/bin/env python3
"""Independent replay of the 0732 evidence packet.

The audit deliberately duplicates the custody, route, oracle, timing, and
statistics checks instead of importing ``analyze.py``.  It is an offline
second calculation: no Cargo command, probe, or capture process is started.
"""

from __future__ import annotations

import hashlib
import json
import math
import statistics
from pathlib import Path
from typing import Any


PACKET = Path(__file__).resolve().parent
# Keep both names available: synthetic-preflight rebasing treats the primary
# packet root and its short alias as equivalent bindings.
P = PACKET
ROOT = PACKET.parents[3]
CAPTURES = PACKET / "captures"
HEX = set("0123456789abcdef")
ROUTES = ("ordinary-opaque", "ordinary-split", "profiled-empty", "profiled-clock")
SPLIT = set(ROUTES[1:])
FIELDS = ("p50", "mean", "p95", "p99", "maximum")


class AuditError(Exception):
    pass


def need(condition: bool, message: str) -> None:
    if not condition:
        raise AuditError(message)


def read(path: Path) -> Any:
    need(path.is_file() and not path.is_symlink(), f"missing file: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise AuditError(f"invalid JSON {path}: {error}") from error


def digest(value: Any, label: str) -> str:
    need(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
         f"{label}: invalid digest")
    return value


def sha(path: Path) -> str:
    need(path.is_file() and not path.is_symlink(), f"missing or symlinked: {path}")
    return hashlib.sha256(path.read_bytes()).hexdigest()


def rel(value: Any, label: str) -> str:
    need(isinstance(value, str) and value and not Path(value).is_absolute()
         and ".." not in Path(value).parts, f"{label}: unsafe path")
    return value


def integer(value: Any, label: str, minimum: int = 0) -> int:
    need(isinstance(value, int) and not isinstance(value, bool) and value >= minimum,
         f"{label}: invalid integer")
    return value


def st(values: list[int | float]) -> dict[str, int | float]:
    need(values, "empty statistic")
    values = sorted(values)
    return {"n": len(values), "p50": statistics.median(values), "mean": statistics.fmean(values),
            "p95": values[max(0, math.ceil(.95 * len(values)) - 1)],
            "p99": values[max(0, math.ceil(.99 * len(values)) - 1)],
            "maximum": values[-1]}


def plan() -> dict[str, Any]:
    p = read(PACKET / "plan.json")
    need(isinstance(p, dict), "plan is not an object")
    for key, value in {"cpu": 12, "cycles": 3, "repeats": 3, "samples": 50,
                       "warmups": 3, "routes": list(ROUTES), "processes": 36,
                       "measured_lifecycles": 1800}.items():
        need(p.get(key) == value, f"plan changed: {key}")
    if "comparisons" in p:
        need(p["comparisons"] == [[ROUTES[0], ROUTES[1]], [ROUTES[1], ROUTES[2]],
                                   [ROUTES[2], ROUTES[3]]], "comparison chain changed")
    schedule = p.get("schedule")
    need(isinstance(schedule, list) and len(schedule) == 36, "schedule count changed")
    seen = set()
    for item in schedule:
        need(isinstance(item, dict), "schedule item is not an object")
        key = (item.get("cycle"), item.get("repeat"), item.get("route"))
        need(key not in seen and key[0] in range(3) and key[1] in range(3)
             and key[2] in ROUTES, f"invalid schedule cell: {key!r}")
        seen.add(key)
    need(len(seen) == 36, "schedule does not cover all cells")
    return p


def case() -> dict[str, Any]:
    c = read(PACKET / "case.json")
    need(c.get("case") == "ppt45543" and c.get("format") == "ppt", "case identity changed")
    path = rel(c.get("path"), "case path")
    fixture = ROOT / path
    need(fixture.is_file() and c.get("bytes") == fixture.stat().st_size
         and c.get("sha256") == sha(fixture), "fixture custody changed")
    return c


def maps(build: dict[str, Any]) -> None:
    for name, root in (("source", ROOT), ("probe", P)):
        value = build.get(name)
        need(isinstance(value, dict) and value, f"build {name} map empty")
        for raw, expected in value.items():
            raw = rel(raw, f"build {name} path")
            need(sha(root / raw) == digest(expected, f"build {name} {raw}"),
                 f"build {name} changed: {raw}")


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


def custody(p: dict[str, Any], c: dict[str, Any]) -> dict[str, Any]:
    build = read(PACKET / "build.json")
    need(isinstance(build, dict), "build is not an object")
    maps(build)
    qpath = P / rel(build.get("quality"), "quality path")
    need(sha(qpath) == digest(build.get("quality_sha256"), "quality manifest"),
         "quality manifest changed")
    quality = read(qpath)
    need(quality.get("source") == build["source"] and quality.get("probe") == build["probe"],
         "quality source bindings changed")
    runs = quality.get("runs")
    need(isinstance(runs, list) and len(runs) == 12, "quality count changed")
    commands = expected_quality_commands()
    for n, row in enumerate(runs):
        need(row.get("exit_code") == 0 and row.get("command") == commands[n],
             f"quality {n} failed")
        out = P / rel(row.get("output"), f"quality {n} output")
        need(sha(out) == digest(row.get("sha256"), f"quality {n} digest"),
             f"quality {n} output changed")
    b = build.get("binary")
    need(isinstance(b, dict) and Path(b.get("path", "")).is_absolute(), "binary receipt changed")
    bp = Path(b["path"])
    if bp.is_file():
        need(bp.stat().st_size == b["bytes"] and sha(bp) == digest(b["sha256"], "binary"),
             "live binary changed")
    else:
        clean = read(PACKET / "cleanup.json")
        need(clean.get("removed") is True and clean.get("binaries") == [b],
             "binary cleanup witness changed")
    review = read(PACKET / "source-review.json")
    need(review.get("base_head") == p.get("base_head")
         and review.get("ordinary_commit_unchanged") is True,
         "ordinary commit review changed")
    digest(review.get("ordinary_commit_sha256"), "ordinary commit")
    for raw, item in review.get("files", {}).items():
        raw = rel(raw, "review file")
        need(sha(ROOT / raw) == digest(item.get("after_sha256"), f"review after {raw}"),
             f"reviewed source changed: {raw}")
        for side, field in (("before", "before_sha256"), ("after", "after_sha256")):
            archive = P / "source-archive" / side / raw
            if archive.exists():
                need(sha(archive) == digest(item.get(field), f"review {side} {raw}"),
                     f"review archive changed: {side}/{raw}")
    constraints = read(PACKET / "constraints.json")
    for raw, expected in constraints.items():
        raw = rel(raw, "constraint")
        need(sha(ROOT / raw) == digest(expected, f"constraint {raw}"), f"constraint changed: {raw}")
    frozen = read(PACKET / "freeze.json")
    need(isinstance(frozen, dict) and frozen, "freeze is empty")
    for raw, expected in frozen.items():
        need(isinstance(raw, str) and Path(raw).is_absolute(), "freeze path is not absolute")
        path = Path(raw)
        if path.is_file():
            need(sha(path) == digest(expected, f"freeze {raw}"), f"frozen input changed: {raw}")
        else:
            need(path == bp and digest(expected, f"freeze {raw}") == b["sha256"],
                 f"missing frozen input is not binary: {raw}")
    oracle = read(PACKET / "oracle.json")
    reference = ROOT / rel(oracle.get("reference"), "oracle reference")
    need(sha(reference) == digest(oracle.get("sha256"), "oracle reference"),
         "sealed oracle changed")
    expected = oracle["expected"]
    qual = read(PACKET / "qualification.json")
    need(qual.get("status") == "passed" and qual.get("build_sha256") == sha(PACKET / "build.json"),
         "qualification binding changed")
    for raw, value in qual.get("files", {}).items():
        raw = rel(raw, "qualification file")
        need(sha(PACKET / raw) == digest(value, f"qualification {raw}"),
             f"qualification changed: {raw}")
    return {"build": build, "expected": expected, "frozen": frozen}


def preflight() -> dict[str, Any]:
    value = read(P / "preflight.json")
    need(value.get("status") == "passed", "preflight did not pass")
    need(value.get("freeze_sha256") == sha(P / "freeze.json"),
         "preflight freeze binding changed")
    scripts = value.get("scripts")
    need(isinstance(scripts, dict)
         and set(scripts) == {"preflight.py", "analyze.py", "audit.py"},
         "preflight scripts are incomplete")
    for name, expected in scripts.items():
        need(sha(P / name) == digest(expected, f"preflight {name}"),
             f"preflight script changed: {name}")
    return value


STATIC = ("case", "format", "operation", "input", "policy", "policy_applied",
          "policy_application_scope", "policy_argument_effect", "policy_contract",
          "directory_metadata_fields", "source_sha256",
          "expected_output_sha256", "replacements_sha256", "source_inventory",
          "expected_output_inventory", "replacements", "changed_length_proof",
          "expected_oracle", "oracle_controls")


def all_true(value: Any, label: str) -> None:
    need(isinstance(value, dict) and value.get("oracle_ok") is True
         and value.get("failure_reasons") == [], f"{label}: oracle failed")
    for key, item in value.items():
        if isinstance(item, bool):
            need(item, f"{label}.{key} is false")
        elif isinstance(item, dict):
            all_nested(item, f"{label}.{key}")
        elif isinstance(item, list):
            for n, nested in enumerate(item):
                if isinstance(nested, dict):
                    all_nested(nested, f"{label}.{key}[{n}]")


def all_nested(value: dict[str, Any], label: str) -> None:
    for key, item in value.items():
        if isinstance(item, bool):
            need(item, f"{label}.{key} is false")
        elif isinstance(item, dict):
            all_nested(item, f"{label}.{key}")
        elif isinstance(item, list):
            for n, nested in enumerate(item):
                if isinstance(nested, dict):
                    all_nested(nested, f"{label}.{key}[{n}]")


def phase(sample: dict[str, Any], label: str) -> dict[str, int]:
    whole = integer(sample.get("whole_ns"), f"{label}.whole_ns", 1)
    split = sample.get("split")
    if split is None:
        return {"whole_ns": whole}
    need(isinstance(split, dict), f"{label}: split is not an object")
    names = ("open_ns", "edit_ns", "remove_ns", "commit_ns", "output_copy_ns")
    out = {name: integer(split.get(name), f"{label}.{name}") for name in names}
    need(out["commit_ns"] > 0, f"{label}: commit_ns is not positive")
    split_sum = integer(split.get("split_sum_ns"), f"{label}.split_sum_ns")
    residual = integer(split.get("whole_residual_ns"), f"{label}.whole_residual_ns")
    need(split_sum == sum(out.values()) and split_sum + residual == whole,
         f"{label}: phase sum does not equal whole")
    start = integer(split.get("commit_start_ns"), f"{label}.commit_start_ns")
    end = integer(split.get("commit_end_ns"), f"{label}.commit_end_ns", start)
    need(end <= whole and end - start == out["commit_ns"], f"{label}: commit window invalid")
    out["whole_residual_ns"] = residual
    return {"whole_ns": whole, **out}


def diag(sample: dict[str, Any]) -> dict[str, Any] | None:
    value = sample.get("diagnostics")
    return value if isinstance(value, dict) else None


EXPECTED_PHASES = ("DocumentCommit", "BeforePayloadCapture", "EmbeddedOpen", "LiveDocumentRead",
                   "EmbeddedFinish", "UnrelatedStreamValidation", "PublicReopen",
                   "AfterPayloadCapture", "ArtifactHashBefore", "ArtifactHashAfter")
EXPECTED_EVENTS = tuple(value for phase_name in EXPECTED_PHASES
                        for value in (f"started:{phase_name}", f"finished:{phase_name}"))


def check_diag(value: dict[str, Any], sample: dict[str, Any], phases: dict[str, int],
               names: tuple[str, ...] | None, label: str) -> tuple[tuple[str, ...], dict[str, Any]]:
    report = value.get("commit")
    need(isinstance(report, dict) and report.get("event_count") == 20
         and report.get("overflow") is False and report.get("balanced") is True
         and report.get("sequence_ok") is True and report.get("outcomes_ok") is True
         and report.get("timestamps_monotonic") is True,
         f"{label}: diagnostic flags changed")
    need(report.get("expected_phases") == list(EXPECTED_PHASES), f"{label}: expected phases changed")
    events = report.get("events")
    need(isinstance(events, list) and len(events) == 20, f"{label}: event count changed")
    got = []
    previous = 0
    for n, event in enumerate(events):
        phase_name = EXPECTED_PHASES[n // 2]
        kind = "started" if n % 2 == 0 else "finished"
        outcome = "started" if kind == "started" else "success"
        need(isinstance(event, dict) and event.get("kind") == kind
             and event.get("phase") == phase_name and event.get("outcome") == outcome,
             f"{label}: event {n} sequence/outcome changed")
        timestamp = integer(event.get("t_ns"), f"{label}: event {n} timestamp")
        need(timestamp >= previous, f"{label}: timestamps are not monotonic")
        previous = timestamp
        got.append(f"{kind}:{phase_name}")
    result = tuple(got)
    if names is not None:
        need(result == names, f"{label}: event sequence differs")
    spans = report.get("spans")
    need(isinstance(spans, list) and len(spans) == 10, f"{label}: span count changed")
    total = 0
    phase_durations = {}
    for n, span in enumerate(spans):
        need(isinstance(span, dict) and span.get("phase") == EXPECTED_PHASES[n]
             and span.get("outcome") == "success", f"{label}: span {n} changed")
        start = integer(span.get("start_ns"), f"{label}: span {n} start")
        end = integer(span.get("finish_ns"), f"{label}: span {n} finish", start)
        duration = integer(span.get("duration_ns"), f"{label}: span {n} duration")
        need(duration == end - start and start == events[2 * n]["t_ns"]
             and end == events[2 * n + 1]["t_ns"], f"{label}: span {n} mismatch")
        total += duration
        phase_durations[span["phase"]] = duration
    split = sample.get("split")
    need(isinstance(split, dict), f"{label}: split is missing")
    start = integer(split.get("commit_start_ns"), f"{label}: commit start")
    end = integer(split.get("commit_end_ns"), f"{label}: commit end", start)
    need(end - start == phases["commit_ns"] and end <= phases["whole_ns"]
         and total <= phases["commit_ns"], f"{label}: commit window invalid")
    need(all(start <= event["t_ns"] <= end for event in events),
         f"{label}: event outside commit window")
    return result, {
        "commit_ns": phases["commit_ns"],
        "commit_start_ns": start,
        "commit_end_ns": end,
        "event_count": 20,
        "event_names": list(result),
        "event_durations_ns": total,
        "phase_durations_ns": phase_durations,
        "whole_ns": phases["whole_ns"],
    }


def expected_command(plan_value: dict[str, Any], build: dict[str, Any], c: dict[str, Any], route: str) -> list[str]:
    return ["taskset", "-c", str(plan_value["cpu"]), build["binary"]["path"], "--route", route,
            "--input", c["path"], "--samples", str(plan_value["samples"]),
            "--warmups", str(plan_value["warmups"])]


def replay(plan_value: dict[str, Any], custody_value: dict[str, Any], c: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    expected = custody_value["expected"]
    manifest = read(CAPTURES / "manifest.json")
    need(manifest.get("status") == "complete"
         and manifest.get("freeze_sha256") == sha(PACKET / "freeze.json")
         and manifest.get("preflight_sha256") == sha(PACKET / "preflight.json"),
         "capture manifest status/binding changed")
    runs = manifest.get("runs")
    need(isinstance(runs, list) and len(runs) == 36, "capture run count changed")
    schedule = {(x["cycle"], x["repeat"], x["route"]): i for i, x in enumerate(plan_value["schedule"])}
    seen = set()
    process_rows = []
    names: tuple[str, ...] | None = None
    for index, row in enumerate(runs):
        key = (row.get("cycle"), row.get("repeat"), row.get("route"))
        need(key in schedule and key not in seen and schedule[key] == index, f"capture schedule row {index} changed")
        seen.add(key)
        route = key[2]
        output_name = f"c{key[0]}-r{key[1]}-{route}.json"
        need(row.get("output") == output_name and row.get("exit_code") == 0,
             f"capture row {index} identity changed")
        output = CAPTURES / rel(output_name, "capture output")
        stderr = CAPTURES / (output_name + ".stderr")
        need(sha(output) == digest(row.get("sha256"), f"capture {index}"), "capture digest changed")
        need(sha(stderr) == digest(row.get("stderr_sha256"), f"stderr {index}"), "stderr digest changed")
        report = read(output)
        need(report.get("route") == route, f"{route}: report route changed")
        need(row.get("command") == expected_command(plan_value, custody_value["build"], c, route),
             f"capture command {index} changed")
        for field in STATIC:
            need(report.get(field) == expected.get(field), f"{route}: {field} differs")
        need(report.get("schema_version") == 1 and report.get("timing_claim") is True
             and report.get("allocator_instrumented") is False
             and report.get("samples_requested") == 50 and report.get("warmups") == 3
             and report.get("mode") == "ppt_public_phase_attribution"
             and report.get("scope") == "public_ppt_open_edit_remove_commit_output_copy"
             and all(isinstance(report.get(key), str) and report[key]
                     for key in ("phase_contract", "diagnostic_contract", "observer_contract",
                                 "allocation_ownership_contract")),
             f"{route}: report header changed")
        controls = report.get("oracle_controls")
        need(isinstance(controls, list) and controls == expected["oracle_controls"],
             f"{route}: oracle controls changed")
        for control in controls:
            need(control.get("rejected") is True and control.get("status") == "rejected"
                 and isinstance(control.get("failure_reasons"), list)
                 and control["failure_reasons"],
                 f"{route}: corruption control was accepted")
        all_true(report["expected_oracle"], f"{route}: expected oracle")
        need(len(report["samples"]) == 50, f"{route}: sample count changed")
        phase_rows = []
        process_names = None
        diag_records = []
        observer_controls = []
        for n, sample in enumerate(report["samples"]):
            need(sample.get("route") == route, f"{route}: sample route changed")
            need(sample.get("index") == n and sample.get("output_sha256") == expected["expected_output_sha256"]
                 and sample.get("output_inventory") == expected["expected_output_inventory"]
                 and sample.get("oracle") == expected["expected_oracle"], f"{route}: sample {n} oracle changed")
            all_true(sample["oracle"], f"{route}: sample oracle")
            phases = phase(sample, f"{route} sample {n}")
            if route == ROUTES[0]:
                need(set(phases) == {"whole_ns"} and diag(sample) is None
                     and sample.get("observer_clock_control_ns") is None,
                     f"{route}: opaque shape changed")
            else:
                need("whole_residual_ns" in phases, f"{route}: residual is missing")
                if route == "profiled-clock":
                    value = diag(sample)
                    need(value is not None, f"{route}: missing trace")
                    process_names, diagnostic = check_diag(
                        value, sample, phases, process_names, f"{route} sample {n}")
                    diag_records.append(diagnostic)
                    observer_controls.append(integer(sample.get("observer_clock_control_ns"),
                                                     f"{route}: observer control"))
                else:
                    need(diag(sample) is None and sample.get("observer_clock_control_ns") is None,
                         f"{route}: unexpected trace")
            phase_rows.append(phases)
        if route == "profiled-clock":
            need(process_names is not None and len(process_names) == 20, "diagnostic schema missing")
            if names is None:
                names = process_names
            need(process_names == names, "diagnostic schema differs between processes")
        process = {"cycle": key[0], "repeat": key[1], "route": route,
                   "samples": 50,
                   "timing_ns": {field: st([x[field] for x in phase_rows])
                                 for field in sorted(set().union(*(set(x) for x in phase_rows)))},
                   "phase_ratios": {field: st([x[field] / x["whole_ns"] for x in phase_rows])
                                    for field in sorted(set().union(*(set(x) for x in phase_rows)) - {"whole_ns"})}}
        if diag_records:
            process["diagnostic"] = {
                "samples": len(diag_records), "event_names": list(names or process_names or ()),
                "commit_ns": st([x["commit_ns"] for x in diag_records]),
                "commit_window_ns": st([x["commit_end_ns"] - x["commit_start_ns"] for x in diag_records]),
                "total_event_duration_ns": st([x["event_durations_ns"] for x in diag_records]),
                "phase_durations_ns": {
                    phase_name: st([x["phase_durations_ns"][phase_name] for x in diag_records])
                    for phase_name in EXPECTED_PHASES
                },
                "phase_whole_ratios": {
                    phase_name: st([x["phase_durations_ns"][phase_name] / x["whole_ns"]
                                   for x in diag_records])
                    for phase_name in EXPECTED_PHASES
                },
                "phase_commit_ratios": {
                    phase_name: st([x["phase_durations_ns"][phase_name] / x["commit_ns"]
                                   for x in diag_records])
                    for phase_name in EXPECTED_PHASES
                },
                "commit_residual_ns": st([x["commit_ns"] - x["event_durations_ns"]
                                           for x in diag_records]),
                "non_commit_ns": st([x["whole_ns"] - x["commit_ns"] for x in diag_records]),
                "commit_whole_ratios": st([x["commit_ns"] / x["whole_ns"] for x in diag_records]),
            }
            process["observer_clock_control_ns"] = st(observer_controls)
        process_rows.append(process)
    need(len(seen) == 36 and names is not None, "capture matrix incomplete")
    expected_files = {"manifest.json"}
    for row in runs:
        expected_files.add(row["output"])
        expected_files.add(row["output"] + ".stderr")
    actual_files = {x.name for x in CAPTURES.iterdir() if x.is_file() and not x.is_symlink()}
    need(actual_files == expected_files, "capture directory inventory changed")
    adjacent = []
    by_key = {(x["cycle"], x["repeat"], x["route"]): x for x in process_rows}
    for cycle in range(3):
        for repeat in range(3):
            # Capture order rotates; semantic comparisons stay on the fixed
            # route chain for every matched round.
            for left, right in zip(ROUTES, ROUTES[1:]):
                a = by_key[(cycle, repeat, left)]["timing_ns"]["whole_ns"]
                b = by_key[(cycle, repeat, right)]["timing_ns"]["whole_ns"]
                delta = {}
                flags = []
                for field in FIELDS:
                    percent = 100.0 * (b[field] / a[field] - 1.0) if a[field] else 0.0
                    delta[field] = {"left": a[field], "right": b[field], "percent": percent}
                    if field in ("p50", "mean") and abs(percent) > 5.0:
                        flags.append(field)
                adjacent.append({"cycle": cycle, "repeat": repeat, "left": left, "right": right,
                                 "delta": delta, "observer_flags": flags})
    need(len(adjacent) == 27, "adjacent comparison count changed")
    return process_rows, adjacent


def main() -> None:
    p = plan()
    c = case()
    custody_value = custody(p, c)
    preflight()
    processes, adjacent = replay(p, custody_value, c)
    result = read(PACKET / "analysis.json")
    need(result.get("status") == "passed" and result.get("process_count") == 36,
         "analysis status/count changed")
    independent_names = next(
        row["diagnostic"]["event_names"] for row in processes if "diagnostic" in row
    )
    need(result.get("diagnostic_event_names") == independent_names
         and len(independent_names) == 20,
         "analysis diagnostic event names changed")
    need(result.get("processes") == processes, "independent process statistics differ")
    need(result.get("adjacent_route_comparisons") == adjacent,
         "independent adjacent comparisons differ")
    print("PASS independent 0732 custody, oracle, 36-process replay, phase sums, and statistics audit")


if __name__ == "__main__":
    try:
        main()
    except AuditError as error:
        raise SystemExit(f"FAIL: {error}")
