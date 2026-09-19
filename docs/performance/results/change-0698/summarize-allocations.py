#!/usr/bin/env python3
"""Summarize measurement-only allocator runs for the 0698 workflow matrix.

Allocation-feature runs perturb timing and are reported as diagnostics.  The
driver records one baseline or candidate process per case, with three samples
per phase.  This script keeps those process results separate and emits only a
descriptive candidate-versus-baseline comparison when both manifests exist.
"""

import hashlib
import json
from pathlib import Path


P = Path(__file__).resolve().parent
PHASES = ("capture", "clone", "settext", "commit", "apply", "total")
METRICS = (
    "alloc_calls",
    "requested_bytes",
    "baseline_live_bytes",
    "peak_live_bytes",
    "current_live_bytes",
    "realloc_calls",
    "realloc_requested_bytes",
)


def manifest_paths():
    combined = P / "allocation-runs.json"
    if combined.exists():
        return [combined]
    paths = [P / f"allocation-runs-{suffix}.json" for suffix in ("baseline", "candidate")]
    return [path for path in paths if path.exists()]


def load_runs():
    paths = manifest_paths()
    if not paths:
        raise SystemExit("missing allocation run manifest")
    runs = []
    for path in paths:
        runs.extend(json.loads(path.read_text()))
    return runs


def output_path(value):
    path = Path(value)
    return path if path.is_absolute() else P / path


def parse(path):
    lines = path.read_text().splitlines()
    header_line = next((line for line in lines if line.startswith("sample\t")), None)
    if header_line is None:
        raise AssertionError(f"{path}: missing sample header")
    header = header_line.split("\t")
    rows = []
    for line in lines:
        fields = line.split("\t")
        if fields and fields[0].isdigit():
            if len(fields) != len(header):
                raise AssertionError(f"{path}: row/header width mismatch")
            rows.append(dict(zip(header, (int(value) for value in fields))))
    declared = int(next(line.split("\t")[1] for line in lines if line.startswith("samples\t")))
    if len(rows) != declared or [row["sample"] for row in rows] != list(range(declared)):
        raise AssertionError(f"{path}: invalid sample rows")
    metadata = {}
    for line in lines:
        fields = line.split("\t")
        if fields and not fields[0].isdigit() and not fields[0].startswith("sample"):
            metadata[fields[0]] = fields[1:]
    return header[1:], rows, metadata


def expected_cases():
    cases_path = P / "cases.json"
    if cases_path.exists():
        return {item["case"]: item["workflow"] for item in json.loads(cases_path.read_text())}
    return {}


def phase_metrics(rows, phase):
    return {
        metric: [row[f"{phase}_{metric}"] for row in rows]
        for metric in METRICS
    }


def comparison(case, workflow, phase, candidate, baseline):
    delta = {
        metric: [
            left - right
            for left, right in zip(
                candidate["metrics"][metric], baseline["metrics"][metric]
            )
        ]
        for metric in METRICS
    }
    peak_delta = [
        left - right
        for left, right in zip(candidate["peak_above_start"], baseline["peak_above_start"])
    ]
    net_delta = [
        left - right
        for left, right in zip(candidate["net_live_change"], baseline["net_live_change"])
    ]
    return {
        "case": case,
        "workflow": workflow,
        "phase": phase,
        "candidate_run_phase": "candidate",
        "baseline_run_phase": "baseline",
        "candidate": candidate,
        "baseline": baseline,
        "delta": delta,
        "peak_above_start_delta": peak_delta,
        "net_live_change_delta": net_delta,
    }


def main():
    runs = load_runs()
    cases = expected_cases()
    by_key = {}
    summary = []
    seen = set()
    for run in runs:
        if run.get("exit_code") != 0:
            raise AssertionError(f"allocation run failed: {run}")
        case = run["case"]
        if cases and case not in cases:
            raise AssertionError(f"unexpected allocation case {case}")
        run_phase = run.get("phase")
        if run_phase not in {"baseline", "candidate"}:
            raise AssertionError(f"allocation run lacks baseline/candidate phase: {run}")
        key = (case, run_phase)
        if key in seen:
            raise AssertionError(f"duplicate allocation case/run phase {key}")
        seen.add(key)
        path = output_path(run["output"])
        if hashlib.sha256(path.read_bytes()).hexdigest() != run["output_sha256"]:
            raise AssertionError(f"allocation output hash mismatch: {path}")
        columns, rows, metadata = parse(path)
        workflow = cases.get(case, run.get("workflow"))
        if run.get("workflow") and workflow != run["workflow"]:
            raise AssertionError(f"workflow mismatch for {case}: {workflow} vs {run['workflow']}")
        if metadata.get("workflow") and metadata["workflow"] != [workflow]:
            raise AssertionError(f"probe workflow mismatch for {case}: {metadata['workflow']}")
        for phase in PHASES:
            required = {f"{phase}_{metric}" for metric in METRICS}
            if not required.issubset(columns):
                raise AssertionError(f"{path}: missing allocation columns for {phase}")
            metrics = phase_metrics(rows, phase)
            record = {
                "case": case,
                "workflow": workflow,
                "run_phase": run_phase,
                "phase": phase,
                "metrics": metrics,
                "peak_above_start": [
                    row[f"{phase}_peak_live_bytes"] - row[f"{phase}_baseline_live_bytes"]
                    for row in rows
                ],
                "net_live_change": [
                    row[f"{phase}_current_live_bytes"] - row[f"{phase}_baseline_live_bytes"]
                    for row in rows
                ],
            }
            summary.append(record)
            by_key[(case, run_phase, phase)] = record

    present_cases = {record["case"] for record in summary}
    if cases and present_cases != set(cases):
        raise AssertionError(f"allocation case coverage is {present_cases}, expected {set(cases)}")
    present_run_phases = {record["run_phase"] for record in summary}
    if present_run_phases not in ({"baseline"}, {"baseline", "candidate"}):
        raise AssertionError(f"invalid allocation run phases: {present_run_phases}")
    for case in present_cases:
        case_phases = {record["run_phase"] for record in summary if record["case"] == case}
        if case_phases != present_run_phases:
            raise AssertionError(f"{case} has incomplete allocation process coverage")

    comparisons = []
    if present_run_phases == {"baseline", "candidate"}:
        for case in sorted(present_cases):
            workflow = cases.get(case)
            for phase in PHASES:
                comparisons.append(
                    comparison(
                        case,
                        workflow,
                        phase,
                        by_key[(case, "candidate", phase)],
                        by_key[(case, "baseline", phase)],
                    )
                )

    (P / "allocation-summary.json").write_text(json.dumps(summary, indent=2) + "\n")
    (P / "allocation-comparisons.json").write_text(json.dumps(comparisons, indent=2) + "\n")
    print(f"Allocation summary: {len(summary)} process/phase records")
    print(f"Allocation comparisons: {len(comparisons)} candidate-baseline process pairs")
    print("Allocation values are descriptive diagnostics; they are not native timing claims.")


if __name__ == "__main__":
    main()
