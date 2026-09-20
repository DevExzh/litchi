#!/usr/bin/env python3
"""Independent custody and statistical audit for the 0727 XLS diagnostic.

The 0727 run is a bounded replication of the four failed 0726 warm cells.  It
has no retention authority: this program reports the complete diagnostic and
returns one only when the frozen evidence is internally sound and every
configured timing comparison passes.  ``--allow-rejected`` preserves the
raw evidence and recomputed rows while returning one for a failed gate.

This file intentionally does not import or execute the primary analyzer,
probe, Cargo, a profiler, or a cleanup command.  It recomputes headers,
semantic outcomes, per-process statistics, corresponding-replicate pairs,
and timing flags from the frozen raw files.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import statistics
import subprocess
import sys
from pathlib import Path
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
PLAN = PACKET / "plan.json"
CAPTURES = PACKET / "captures"
BASELINE_REF = "959daa11e5c5e9d0efdb13fff814bbcc66284740"
SOURCE_ROOTS = ("litchi-cfb", "litchi-xls")
SOURCE_REL = "crates/litchi-xls/src/workbook/source.rs"
CANDIDATE_ARCHIVE_0726 = (
    ROOT / "docs/performance/results/change-0726/candidate-source" / SOURCE_REL
)
EXPECTED_BUILD_COMMAND = [
    "cargo", "build", "--manifest-path", "docs/performance/results/change-0686/probe/Cargo.toml",
    "--release", "--locked", "--offline",
]
HEX = set("0123456789abcdef")
METRIC_KEYS = ("open", "q1", "q2", "q3", "q8", "q3-to-q8-mean", "open-plus-eight")
STAT_KEYS = ("p50", "mean", "p95", "p99", "maximum")
PAIRS = (
    ("aa2/aa1", "aa1", "aa2"),
    ("b1/a1", "a1", "b1"),
    ("b2/a2", "a2", "b2"),
    ("a2/a1", "a1", "a2"),
    ("b2/b1", "b1", "b2"),
)
# These are the files the already-frozen capture driver binds.  The independent
# audit and the post-capture negative receipt are checked directly and sealed in
# the packet; they were intentionally not retroactively added to the freeze.
ALLOCATED_AUDIT_FILES = (
    "plan.json", "build.py", "capture.py", "analyze.py", "constraints.json",
    "environment.json", "builds.json", "build-restoration.json", "hypothesis.md",
)


class AuditError(Exception):
    """An evidence or gate failure."""


class Pending(Exception):
    """The requested draft or terminal evidence has not been produced yet."""

    def __init__(self, items: Iterable[str]):
        self.items = tuple(dict.fromkeys(items))
        super().__init__("evidence is pending")


def check(condition: bool, message: str) -> None:
    if not condition:
        raise AuditError(message)


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha(path: Path) -> str:
    check(path.is_file() and not path.is_symlink(), f"missing or symlinked file: {path}")
    return sha_bytes(path.read_bytes())


def read_json(path: Path) -> Any:
    check(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise AuditError(f"invalid JSON {path}: {error}") from error


def require_sha(value: Any, label: str) -> str:
    check(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
          f"{label} is not a lowercase SHA-256 digest")
    return value


def require_int(value: Any, label: str, minimum: int = 0) -> int:
    check(isinstance(value, int) and not isinstance(value, bool) and value >= minimum,
          f"{label} is not an integer >= {minimum}")
    return value


def relative(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def source_paths() -> list[Path]:
    paths: list[Path] = []
    for owner in SOURCE_ROOTS:
        paths.extend(
            path for path in (ROOT / "crates" / owner).rglob("*")
            if path.is_file() and path.suffix in {".rs", ".toml"}
        )
    return sorted(paths)


def working_source_map() -> dict[str, str]:
    return {relative(path): sha(path) for path in source_paths()}


def git_source_map(revision: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for path in source_paths():
        name = relative(path)
        try:
            value = subprocess.check_output(["git", "show", f"{revision}:{name}"], cwd=ROOT)
        except subprocess.CalledProcessError as error:
            raise AuditError(f"baseline does not contain source path: {revision}:{name}") from error
        result[name] = sha_bytes(value)
    return result


def load_plan() -> dict[str, Any]:
    plan = read_json(PLAN)
    check(isinstance(plan, dict), "plan is not an object")
    check(plan.get("change") == "0727-xls-warm-tail-replication", "plan identity changed")
    check(plan.get("baseline_revision") == BASELINE_REF, "baseline revision changed")
    check(plan.get("candidate_origin") == "change-0726", "candidate origin changed")
    check(plan.get("cpu") == 12 and plan.get("cycles") == 3, "CPU or cycle plan changed")
    check(plan.get("legs") == ["aa1", "aa2", "a1", "b1", "b2", "a2"], "leg order changed")
    check(plan.get("processes_per_cell_leg") == 3, "processes per cell changed")
    check(plan.get("samples") == 100 and plan.get("warmups") == 3 and plan.get("queries") == 8,
          "sample, warmup, or query plan changed")
    check(plan.get("timing_metrics") == list(METRIC_KEYS), "timing metric set/order changed")
    check(plan.get("total_processes") == 324, "total process count changed")
    check(plan.get("disposition") == "diagnostic-only; no retention or retroactive acceptance",
          "diagnostic disposition changed")
    expected_gates = {
        "timing_percent": 5.0,
        "timing_absolute_ns": 10.0,
        "absolute_exception_metrics": ["q3", "q8", "q3-to-q8-mean"],
        "outcomes_exact": True,
        "allocator_fixed_field": True,
        "zero_budget_and_refusal_parity": True,
        "budget_fence_metrics_exact": True,
    }
    check(plan.get("hard_gates") == expected_gates, "hard-gate thresholds changed")
    expected_cases = [
        {
            "case": "54016-stored-2097152", "path": "test-data/poi/test-data/spreadsheet/54016.xls",
            "sheet": 0, "row": 0, "column": 0, "budget": 2097152, "expected_status": "value",
            "mode": "owned", "focus_metric": "q3", "role": "failed", "id": "54016-stored-2097152-owned",
        },
        {
            "case": "54016-stored-2097152", "path": "test-data/poi/test-data/spreadsheet/54016.xls",
            "sheet": 0, "row": 0, "column": 0, "budget": 2097152, "expected_status": "value",
            "mode": "file", "focus_metric": "q3", "role": "failed", "id": "54016-stored-2097152-file",
        },
        {
            "case": "synthetic-70000-default", "path": "docs/performance/results/change-0686/fixtures/numeric-70000.xls",
            "sheet": 0, "row": 0, "column": 0, "budget": 2097152, "expected_status": "value",
            "mode": "owned", "focus_metric": "q3-to-q8-mean", "role": "failed", "id": "synthetic-70000-default-owned",
        },
        {
            "case": "45365-late", "path": "test-data/poi/test-data/spreadsheet/45365-2.xls",
            "sheet": 0, "row": 664, "column": 24, "budget": 2097152, "expected_status": "value",
            "mode": "file", "focus_metric": "q8", "role": "failed", "id": "45365-late-file",
        },
        {
            "case": "45365-first", "path": "test-data/poi/test-data/spreadsheet/45365-2.xls",
            "sheet": 0, "row": 0, "column": 0, "budget": 2097152, "expected_status": "value",
            "mode": "file", "focus_metric": "q8", "role": "matched-control", "id": "45365-first-file",
        },
        {
            "case": "54016-missing-1048576", "path": "test-data/poi/test-data/spreadsheet/54016.xls",
            "sheet": 0, "row": 0, "column": 108, "budget": 1048576, "expected_status": "missing",
            "mode": "owned", "focus_metric": "q8", "role": "benefit-control", "id": "54016-missing-1048576-owned",
        },
    ]
    check(plan.get("cases") == expected_cases, "diagnostic case matrix changed")
    return plan


def check_source_and_builds(plan: dict[str, Any]) -> dict[str, Any]:
    baseline = git_source_map(plan["baseline_revision"])
    baseline_record = read_json(PACKET / "source-baseline.json")
    candidate_record = read_json(PACKET / "source-candidate.json")
    check(baseline_record == baseline, "source-baseline.json differs from git baseline")
    check(set(candidate_record) == set(baseline), "candidate source census path set changed")
    changed = [name for name in sorted(baseline) if candidate_record[name] != baseline[name]]
    check(changed == [SOURCE_REL], f"candidate source delta changed: {changed}")
    candidate_path = PACKET / "sources/candidate" / SOURCE_REL
    baseline_path = PACKET / "sources/baseline" / SOURCE_REL
    check(sha(baseline_path) == baseline[SOURCE_REL], "archived baseline source differs")
    check(candidate_path.read_bytes() == CANDIDATE_ARCHIVE_0726.read_bytes(),
          "0727 candidate is not the exact archived 0726 source")
    check(sha(candidate_path) == candidate_record[SOURCE_REL], "candidate source census differs from archive")
    check(working_source_map() == baseline, "production source is not restored baseline")
    restoration = read_json(PACKET / "build-restoration.json")
    check(restoration == {"exact": True, "source_sha256": baseline},
          "build restoration receipt changed")

    rows = read_json(PACKET / "builds.json")
    check(isinstance(rows, list) and [row.get("phase") for row in rows] == ["baseline", "candidate"],
          "build receipt phase set/order changed")
    phase_hashes: dict[str, dict[str, Any]] = {}
    for row in rows:
        check(row.get("exit_code") == 0 and row.get("command") == EXPECTED_BUILD_COMMAND,
              f"{row.get('phase')} build command or exit status changed")
        phase = row["phase"]
        expected_map = baseline if phase == "baseline" else candidate_record
        source_manifest = PACKET / str(row.get("source_manifest"))
        check(source_manifest.is_file() and read_json(source_manifest) == expected_map,
              f"{phase} source build manifest differs")
        binary = Path(str(row.get("binary")))
        expected = require_sha(row.get("binary_sha256"), f"{phase} binary")
        check(require_int(row.get("bytes"), f"{phase} binary bytes", 1) ==
              (binary.stat().st_size if binary.is_file() else row["bytes"]),
              f"{phase} binary size receipt is invalid")
        phase_hashes[phase] = {"path": str(binary), "sha256": expected, "bytes": row["bytes"]}
    return {"baseline": baseline, "candidate": candidate_record, "binaries": phase_hashes}


def collect_witnesses() -> dict[str, tuple[str, int]]:
    path = PACKET / "cleanup.json"
    if not path.is_file():
        return {}
    value = read_json(path)
    result: dict[str, tuple[str, int]] = {}

    def walk(item: Any) -> None:
        if isinstance(item, dict):
            raw_path = item.get("path")
            digest = item.get("sha256", item.get("binary_sha256"))
            size = item.get("bytes", item.get("size"))
            if isinstance(raw_path, str) and isinstance(digest, str) and isinstance(size, int):
                result[str(Path(raw_path).resolve())] = (digest, size)
            for key, child in item.items():
                if isinstance(key, str) and key.startswith("/") and isinstance(child, str):
                    result[str(Path(key).resolve())] = (child, -1)
                walk(child)
        elif isinstance(item, list):
            for child in item:
                walk(child)
    walk(value)
    return result


def verify_binary(path: Path, digest: str, size: int, witnesses: dict[str, tuple[str, int]], label: str) -> str:
    if path.is_file():
        check(not path.is_symlink() and sha(path) == digest and path.stat().st_size == size,
              f"{label} live identity differs")
        return "live"
    witness = witnesses.get(str(path.resolve()))
    check(witness is not None and witness[0] == digest and (witness[1] in (-1, size)),
          f"{label} is absent without an exact cleanup witness")
    return "cleanup-witness"


def expected_bindings(plan: dict[str, Any]) -> dict[str, str]:
    paths: set[Path] = {PACKET / name for name in ALLOCATED_AUDIT_FILES}
    paths.update(PACKET.glob("source-*.json"))
    paths.update(path for path in (PACKET / "sources").rglob("*") if path.is_file())
    for owner in SOURCE_ROOTS:
        paths.update(
            path for path in (ROOT / "crates" / owner).rglob("*")
            if path.is_file() and path.suffix in {".rs", ".toml"}
        )
    probe = PACKET.parent / "change-0686/probe"
    paths.update((probe / name for name in ("Cargo.toml", "Cargo.lock")))
    paths.update(path for path in (probe / "src").rglob("*.rs"))
    paths.update(ROOT / case["path"] for case in plan["cases"])
    constraints = read_json(PACKET / "constraints.json")
    check(isinstance(constraints, dict), "constraints receipt is not an object")
    paths.update(ROOT / name for name in constraints)
    result: dict[str, str] = {}
    for path in sorted(paths):
        check(path.is_file() and not path.is_symlink(), f"binding path missing: {path}")
        result[relative(path)] = sha(path)
    return result


def check_negative_controls(draft: bool) -> dict[str, Any]:
    script_path = PACKET / "negative-checks.py"
    check(script_path.is_file(), "negative-checks.py is missing")
    script = script_path.read_text(encoding="utf-8")
    for token in (
        "mod.main()", "raw hash mismatch rejected", "outcome corruption with matching hash rejected",
        "duplicate process rejected", "command sample count corruption rejected", "phase relabel rejected",
        "end binding removed rejected", "warm exception", "warm failure", "workflow no exception",
    ):
        check(token in script, f"negative-checks.py lacks control: {token}")
    receipt_path = PACKET / "negative-checks.json"
    if not receipt_path.is_file():
        if draft:
            return {"status": "pending", "script_sha256": sha(script_path)}
        raise Pending(("negative-checks.json",))
    receipt = read_json(receipt_path)
    check(receipt.get("analyzer_sha256") == sha(PACKET / "analyze.py"),
          "negative-checks receipt is not bound to analyze.py")
    check(receipt.get("script_sha256") == sha(script_path),
          "negative-checks receipt is not bound to negative-checks.py")
    names = [
        "complete capture accepted", "raw hash mismatch rejected",
        "outcome corruption with matching hash rejected", "duplicate process rejected",
        "command sample count corruption rejected", "phase relabel rejected",
        "end binding removed rejected", "warm exception", "warm failure", "workflow no exception",
    ]
    checks = receipt.get("checks")
    check(isinstance(checks, list) and len(checks) == len(names), "negative control count changed")
    check([item.get("name") for item in checks] == names, "negative control order/names changed")
    check(all(item.get("accepted") is item.get("expected") for item in checks),
          "negative control did not produce its expected result")
    return {"status": "complete", "script_sha256": sha(script_path), "receipt_sha256": sha(receipt_path),
            "analyzer_sha256": receipt["analyzer_sha256"], "checks": len(checks)}


def check_freeze(plan: dict[str, Any], builds: dict[str, Any], draft: bool) -> tuple[dict[str, Any], dict[str, str]]:
    path = PACKET / "freeze.json"
    if not path.is_file():
        if draft:
            raise Pending(("freeze.json",))
        raise Pending(("freeze.json",))
    frozen = read_json(path)
    bindings = expected_bindings(plan)
    check(frozen.get("baseline_head") == plan["baseline_revision"], "freeze baseline head changed")
    check(frozen.get("bindings") == bindings, "freeze bindings differ from independently computed bindings")
    binary_rows = frozen.get("binaries")
    check(isinstance(binary_rows, list) and [row.get("phase") for row in binary_rows] == ["baseline", "candidate"],
          "freeze binary phase/order changed")
    witnesses = collect_witnesses()
    for row in binary_rows:
        phase = row["phase"]
        expected = builds["binaries"][phase]
        check(row.get("binary") == expected["path"] and row.get("binary_sha256") == expected["sha256"]
              and row.get("bytes") == expected["bytes"], f"freeze {phase} binary identity changed")
        verify_binary(Path(expected["path"]), expected["sha256"], expected["bytes"], witnesses, f"freeze {phase}")
    return frozen, bindings


def safe_capture_path(raw: Any, label: str) -> Path:
    check(isinstance(raw, str), f"{label} path is not a string")
    path = Path(raw)
    check(not path.is_absolute() and ".." not in path.parts, f"{label} path escapes packet")
    resolved = (CAPTURES / path).resolve()
    check(resolved == CAPTURES.resolve() or CAPTURES.resolve() in resolved.parents,
          f"{label} path is outside captures")
    return resolved


def expected_command(plan: dict[str, Any], case: dict[str, Any], binary: str) -> list[str]:
    return [
        "taskset", "-c", str(plan["cpu"]), binary,
        "--input", case["path"], "--budget", str(case["budget"]), "--mode", case["mode"],
        "--worksheet", str(case["sheet"]), "--row", str(case["row"]), "--column", str(case["column"]),
        "--queries", str(plan["queries"]), "--warmups", str(plan["warmups"]), "--samples", str(plan["samples"]),
    ]


def validate_raw(plan: dict[str, Any], case: dict[str, Any], report: dict[str, Any], label: str,
                 outcomes: dict[str, Any]) -> dict[str, list[float]]:
    check(report.get("schema_version") == 1 and report.get("probe") == "change-0686-xls-index-budget-retry",
          f"{label}: probe header changed")
    fixture = ROOT / case["path"]
    check(report.get("input_sha256") == sha(fixture) and report.get("input_bytes") == fixture.stat().st_size,
          f"{label}: input identity changed")
    check(Path(str(report.get("input_path"))).resolve() == fixture.resolve(), f"{label}: input path changed")
    for key, expected in (("mode", case["mode"]), ("worksheet", case["sheet"]), ("row", case["row"]),
                          ("column", case["column"]), ("max_query_index_bytes", case["budget"]),
                          ("queries", plan["queries"]), ("warmups", plan["warmups"]),
                          ("samples", plan["samples"]), ("fresh_owner_per_sample", True)):
        check(report.get(key) == expected, f"{label}: header {key} changed")
    records = report.get("records")
    check(isinstance(records, list) and len(records) == plan["samples"], f"{label}: record count changed")
    expected_samples = list(range(plan["warmups"], plan["warmups"] + plan["samples"]))
    check([record.get("sample") for record in records] == expected_samples, f"{label}: sample ordinals changed")
    arrays = {metric: [] for metric in METRIC_KEYS}
    first_outcomes: list[Any] | None = None
    for record in records:
        open_record = record.get("open")
        check(isinstance(open_record, dict) and open_record.get("outcome", {}).get("status") == "ok",
              f"{label}: open outcome changed")
        open_ns = require_int(open_record.get("elapsed_ns"), f"{label}: open elapsed")
        queries = record.get("queries")
        check(isinstance(queries, list) and len(queries) == plan["queries"], f"{label}: query count changed")
        check([query.get("ordinal") for query in queries] == list(range(plan["queries"])),
              f"{label}: query ordinals changed")
        check(record.get("all_queries_agree") is True, f"{label}: all_queries_agree changed")
        current_outcomes = []
        times: list[int] = []
        for query in queries:
            outcome = query.get("outcome")
            check(isinstance(outcome, dict) and outcome.get("status") == case["expected_status"],
                  f"{label}: semantic outcome changed")
            check(query.get("agrees_with_first") is True, f"{label}: agrees_with_first changed")
            elapsed = require_int(query.get("elapsed_ns"), f"{label}: query elapsed")
            current_outcomes.append(outcome)
            times.append(elapsed)
        if first_outcomes is None:
            first_outcomes = current_outcomes
        else:
            check(current_outcomes == first_outcomes, f"{label}: sample outcomes differ")
        # The primary analysis records the first query's semantic outcome per
        # cell.  Raw validation above still checks every query and every
        # process; this compact projection only mirrors that stable report
        # field for the independent equality check.
        prior = outcomes.setdefault(case["id"], current_outcomes[0])
        check(prior == current_outcomes[0], f"{label}: cross-process outcomes differ")
        values = [open_ns, times[0], times[1], times[2], times[7], sum(times[2:]) / 6.0,
                  open_ns + sum(times)]
        for metric, value in zip(METRIC_KEYS, values):
            arrays[metric].append(value)
    check(first_outcomes is not None, f"{label}: no outcomes")
    return arrays


def stat(values: list[float]) -> dict[str, float | int]:
    check(values, "empty timing vector")
    ordered = sorted(values)
    return {
        "n": len(values), "p50": statistics.median(values), "mean": statistics.mean(values),
        "p95": ordered[math.ceil(.95 * len(ordered)) - 1],
        "p99": ordered[math.ceil(.99 * len(ordered)) - 1],
        "minimum": ordered[0], "maximum": ordered[-1],
    }


def percent_change(before: float, after: float) -> float | None:
    if before == 0:
        return 0.0 if after == 0 else None
    return (after / before - 1.0) * 100.0


def timing_gate(metric: str, before: float, after: float, plan: dict[str, Any]) -> bool:
    percent = percent_change(before, after)
    exception = metric in set(plan["hard_gates"]["absolute_exception_metrics"])
    absolute = exception and after - before <= plan["hard_gates"]["timing_absolute_ns"]
    return (percent is not None and percent <= plan["hard_gates"]["timing_percent"]) or absolute


def audit_captures(plan: dict[str, Any], frozen: dict[str, Any], bindings: dict[str, str],
                   builds: dict[str, Any], allow_rejected: bool) -> tuple[dict[str, Any], list[dict[str, Any]], dict[str, Any]]:
    manifest_path = CAPTURES / "manifest.json"
    if not manifest_path.is_file():
        raise Pending(("captures/manifest.json",))
    manifest = read_json(manifest_path)
    check(manifest.get("status") == "complete", "capture manifest is not complete")
    check(manifest.get("freeze_sha256") == sha(PACKET / "freeze.json"), "capture freeze hash changed")
    check(manifest.get("bindings_start") == bindings and manifest.get("bindings_end") == bindings,
          "capture bindings changed during run")
    expected_runs = {
        (cycle, leg, case["id"], replicate)
        for cycle in range(plan["cycles"])
        for leg in plan["legs"]
        for case in plan["cases"]
        for replicate in range(plan["processes_per_cell_leg"])
    }
    runs = manifest.get("runs")
    check(isinstance(runs, list) and len(runs) == plan["total_processes"], "capture process count changed")
    cases = {case["id"]: case for case in plan["cases"]}
    binaries = {row["phase"]: row["binary"] for row in frozen["binaries"]}
    seen: set[tuple[int, str, str, int]] = set()
    process_rows: list[dict[str, Any]] = []
    values: dict[tuple[int, str, str, int], dict[str, dict[str, Any]]] = {}
    outcomes: dict[str, Any] = {}
    raw_paths: set[str] = set()
    for run in runs:
        check(isinstance(run, dict), "capture run is not an object")
        key = (run.get("cycle"), run.get("leg"), run.get("case"), run.get("replicate"))
        check(key in expected_runs and key not in seen, f"duplicate or unexpected run: {key}")
        seen.add(key)
        cycle, leg, case_id, replicate = key
        case = cases[case_id]
        expected_phase = "candidate" if leg in ("b1", "b2") else "baseline"
        check(run.get("phase") == expected_phase, f"{key}: phase label changed")
        status = run.get("status", run.get("exit_code"))
        check(status == 0, f"{key}: command status is not zero")
        binary = binaries[expected_phase]
        check(run.get("command") == expected_command(plan, case, binary), f"{key}: command changed")
        output = safe_capture_path(run.get("output"), f"{key} output")
        stderr = safe_capture_path(run.get("stderr"), f"{key} stderr")
        output_rel = relative(output)
        stderr_rel = relative(stderr)
        check(output_rel not in raw_paths and stderr_rel not in raw_paths, f"{key}: raw path reused")
        raw_paths.update((output_rel, stderr_rel))
        check(run.get("sha256") == sha(output), f"{key}: raw output hash changed")
        check(run.get("stderr_sha256") == sha(stderr), f"{key}: stderr hash changed")
        report = read_json(output)
        arrays = validate_raw(plan, case, report, str(output), outcomes)
        metrics = {metric: stat(array) for metric, array in arrays.items()}
        values[key] = metrics
        process_rows.append(dict(cycle=cycle, leg=leg, case=case_id, replicate=replicate, metrics=metrics))
    check(seen == expected_runs, "capture run matrix is incomplete")
    expected_raw = {
        relative(CAPTURES / f"c{cycle}-{leg}-{case_id}-r{replicate}.json")
        for cycle, leg, case_id, replicate in expected_runs
    }
    expected_raw |= {name + ".stderr" for name in expected_raw}
    check(raw_paths == expected_raw, "capture raw path set differs")
    on_disk_raw = {
        relative(path) for path in CAPTURES.rglob("*")
        if path.is_file() and path.name != "manifest.json"
    }
    check(on_disk_raw == expected_raw, "capture directory contains an unmanifested raw file")

    comparisons: list[dict[str, Any]] = []
    for cycle in range(plan["cycles"]):
        for case in plan["cases"]:
            for replicate in range(plan["processes_per_cell_leg"]):
                for pair, base_leg, candidate_leg in PAIRS:
                    for metric in METRIC_KEYS:
                        base = values[(cycle, base_leg, case["id"], replicate)][metric]
                        candidate = values[(cycle, candidate_leg, case["id"], replicate)][metric]
                        comparisons.append({
                            "cycle": cycle, "case": case["id"], "replicate": replicate,
                            "pair": pair, "metric": metric, "role": case["role"],
                            "focus": metric == case["focus_metric"],
                            "statistics": {
                                statistic_name: {
                                    "baseline": base[statistic_name], "candidate": candidate[statistic_name],
                                    "delta_ns": candidate[statistic_name] - base[statistic_name],
                                    "percent": percent_change(base[statistic_name], candidate[statistic_name]),
                                    "pass_": timing_gate(metric, base[statistic_name], candidate[statistic_name], plan),
                                }
                                for statistic_name in STAT_KEYS
                            },
                        })
    paired = [row for row in comparisons if row["pair"] in ("b1/a1", "b2/a2")]
    failed = [
        {"cycle": row["cycle"], "case": row["case"], "replicate": row["replicate"],
         "pair": row["pair"], "metric": row["metric"], "statistic": statistic_name,
         **row["statistics"][statistic_name]}
        for row in paired for statistic_name in ("p50", "mean")
        if not row["statistics"][statistic_name]["pass_"]
    ]
    summary = {
        "processes": len(process_rows), "owners": len(process_rows) * plan["samples"],
        "queries": len(process_rows) * plan["samples"] * plan["queries"],
        "failed_central_checks": len(failed),
        "failed_focus_checks": sum(
            not row["statistics"][name]["pass_"]
            for row in paired if row["focus"] and row["role"] == "failed"
            for name in ("p50", "mean")
        ),
        "paired_tail_flags": sum(
            row["statistics"][name]["percent"] > 5
            for row in paired for name in ("p95", "p99", "maximum")
            if row["statistics"][name]["percent"] is not None
        ),
    }
    expected_analysis = {
        "disposition": plan["disposition"], "summary": summary, "processes": process_rows,
        "comparisons": comparisons, "failed_checks": failed, "outcomes": outcomes,
    }
    analysis_path = PACKET / "analysis.json"
    if not analysis_path.is_file():
        raise Pending(("analysis.json",))
    analysis = read_json(analysis_path)
    check(analysis == expected_analysis, "primary analysis differs from independent raw recomputation")
    native_pass = not failed
    return manifest, comparisons, {
        "analysis": analysis, "summary": summary, "failed": failed,
        "native": native_pass, "all_rows": process_rows,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--draft", action="store_true", help="run pre-capture checks")
    parser.add_argument("--allow-rejected", "--report-rejected", dest="allow_rejected", action="store_true",
                        help="emit an internally valid rejected report and return one")
    args = parser.parse_args()
    try:
        plan = load_plan()
        builds = check_source_and_builds(plan)
        controls = check_negative_controls(args.draft)
        if controls.get("status") == "pending" and args.draft:
            print(json.dumps({"status": "DRAFT PASS", "controls": controls}, sort_keys=True))
        frozen, bindings = check_freeze(plan, builds, args.draft)
        # Every frozen binding is checked independently, including the audit
        # files and all source/probe/fixture/ADR inputs.
        for name, digest in bindings.items():
            check(sha(ROOT / name) == digest, f"frozen binding changed: {name}")
        if args.draft and not (CAPTURES / "manifest.json").is_file():
            print(json.dumps({"status": "DRAFT PASS", "controls": controls,
                              "bindings": len(bindings), "pending": ["captures/manifest.json"]}, sort_keys=True))
            return 0
        if not (CAPTURES / "manifest.json").is_file():
            raise Pending(("captures/manifest.json",))
        manifest, comparisons, result = audit_captures(plan, frozen, bindings, builds, args.allow_rejected)
        central_pass = result["native"]
        hard_gates = {
            "bindings": True, "raw_custody": True, "semantic_outcomes": True,
            "per_process_statistics": True, "timing": central_pass,
            "negative_controls": controls.get("status") == "complete",
        }
        report = {
            "packet": plan["change"], "baseline_revision": plan["baseline_revision"],
            "allow_rejected": args.allow_rejected, "manifest_status": manifest["status"],
            "processes": len(result["all_rows"]), "comparisons": len(comparisons),
            "summary": result["summary"], "failed_checks": result["failed"],
            "negative_controls": controls, "hard_gates": hard_gates,
        }
        # Keep the terminal audit replayable after the owned binaries are
        # removed.  Do not include live-vs-cleanup custody state, timestamps,
        # or other mutable environment details in this file.
        audit_record = {
            "packet": plan["change"], "baseline_revision": plan["baseline_revision"],
            "plan_sha256": sha(PLAN), "freeze_sha256": sha(PACKET / "freeze.json"),
            "manifest_sha256": sha(CAPTURES / "manifest.json"),
            "analysis_sha256": sha(PACKET / "analysis.json"),
            "source_baseline_sha256": sha(PACKET / "source-baseline.json"),
            "source_candidate_sha256": sha(PACKET / "source-candidate.json"),
            "build_restoration_sha256": sha(PACKET / "build-restoration.json"),
            "binary_sha256": {phase: builds["binaries"][phase]["sha256"] for phase in ("baseline", "candidate")},
            "summary": result["summary"], "processes": result["all_rows"],
            "comparisons": comparisons, "failed_checks": result["failed"],
            "outcomes": result["analysis"]["outcomes"],
            "negative_controls": controls, "hard_gates": hard_gates,
        }
        (PACKET / "audit.json").write_text(json.dumps(audit_record, indent=2, sort_keys=True) + "\n",
                                             encoding="utf-8")
        passed = all(hard_gates.values())
        label = "PASS" if passed else ("REJECTED" if args.allow_rejected else "FAIL")
        print(label, json.dumps({
            "packet": report["packet"], "processes": report["processes"],
            "comparisons": report["comparisons"],
            "failed_central_checks": report["summary"]["failed_central_checks"],
            "hard_gates": report["hard_gates"],
        }, sort_keys=True))
        return 0 if passed else 1
    except Pending as pending:
        print("PENDING: " + "; ".join(pending.items), file=sys.stderr)
        return 2
    except (AuditError, OSError, subprocess.CalledProcessError, KeyError, TypeError, ValueError, AssertionError) as error:
        print(f"FAIL: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
