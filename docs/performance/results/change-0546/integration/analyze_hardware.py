#!/usr/bin/env python3
"""Validate the 0546 whole-child ``perf stat`` diagnostic lane.

The counters cover the complete fresh harness child, including corpus setup,
source-backed planning, publication, and output oracles.  This report binds
each row to the frozen source manifest, normal binary, run script and exact
perf command.  Grouped cycles/instructions/branches/branch-misses are used
for IPC and branch-miss diagnostics only when perf reports complete,
non-multiplexed coverage.  No hardware counter is a latency gate or an
operation-local planning counter.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import importlib.util
import json
import math
import sys
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
PLAN_PATH = HERE / "plan.json"
NUMERIC_PATH = HERE / "analyze.py"
_numeric_spec = importlib.util.spec_from_file_location(
    "litchi_xlsx_numeric_0546_for_hardware", NUMERIC_PATH
)
if _numeric_spec is None or _numeric_spec.loader is None:
    raise ImportError(f"cannot load numerical analyzer: {NUMERIC_PATH}")
NUMERIC = importlib.util.module_from_spec(_numeric_spec)
_numeric_spec.loader.exec_module(NUMERIC)
BASE = NUMERIC.BASE

GROUP_EVENTS = ("cycles", "instructions", "branches", "branch-misses")
EVENTS = GROUP_EVENTS + ("page-faults", "context-switches", "cpu-migrations")


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory evidence artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def sha(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def read_json(path: Path) -> Any:
    try:
        with path.open(encoding="utf-8") as stream:
            return json.load(stream)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {path}: {error}") from error


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(HERE))
    except ValueError as error:
        raise EvidenceError(f"path is outside the evidence directory: {path}") from error


def regular(path: Path, label: str) -> None:
    require(path.is_file() and not path.is_symlink(), f"{label} is not a regular file")


def plan_data() -> dict[str, Any]:
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict), "plan is not an object")
    hardware = plan.get("hardware")
    require(isinstance(hardware, dict), "plan.hardware is not an object")
    require(hardware.get("shapes") == ["medium", "dense-sparse"],
            "hardware shape matrix differs from the frozen plan")
    require(hardware.get("repeats") == 2 and hardware.get("warmup") == 0
            and hardware.get("samples") == 100,
            "hardware repeat, warmup or sample count differs from the frozen plan")
    expected_events = "{cycles,instructions,branches,branch-misses},page-faults,context-switches,cpu-migrations"
    require(hardware.get("events") == expected_events,
            "hardware event expression differs from the frozen plan")
    return plan


def expected_jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    hardware = plan["hardware"]
    return [
        {
            "name": f"hardware-r{repeat}-{shape}",
            "repeat": repeat,
            "shape": shape,
            "case": plan["primary"]["case"],
            "warmup": int(hardware["warmup"]),
            "samples": int(hardware["samples"]),
        }
        for repeat in range(1, int(hardware["repeats"]) + 1)
        for shape in hardware["shapes"]
    ]


def native_job(plan: dict[str, Any], repeat: int, shape: str) -> dict[str, Any]:
    primary = plan["primary"]
    return {
        "name": f"native-r{repeat}-primary-{shape}",
        "kind": "primary",
        "guard": None,
        "repeat": repeat,
        "case": primary["case"],
        "shape": shape,
        "warmup": int(primary["warmup"]),
        "samples": int(primary["samples"]),
    }


def hardware_job(job: dict[str, Any]) -> dict[str, Any]:
    return {
        "name": job["name"],
        "kind": "primary",
        "guard": None,
        "repeat": job["repeat"],
        "case": job["case"],
        "shape": job["shape"],
        "warmup": job["warmup"],
        "samples": job["samples"],
    }


def hardware_plan_for_report(plan: dict[str, Any]) -> dict[str, Any]:
    result = dict(plan)
    primary = dict(plan["primary"])
    primary["warmup"] = plan["hardware"]["warmup"]
    primary["samples"] = plan["hardware"]["samples"]
    result["primary"] = primary
    return result


def binary_metadata(stage: str, plan: dict[str, Any]) -> dict[str, Any]:
    folder = HERE / stage
    manifest = folder / "source-manifest.json"
    regular(manifest, relative(manifest))
    # The retained numerical analyzer checks the stage-local build receipt,
    # binary identity and cleanup custody chain.  It does not inspect the
    # current checkout, which is essential for retained baseline captures.
    try:
        binary = BASE.check_binary(folder, False)
    except (BASE.EvidenceError, EvidenceError, OSError, KeyError, TypeError) as error:
        raise EvidenceError(f"{stage} normal binary binding failed: {error}") from error
    require(binary["sha256"] == read_json(folder / "binary-normal.json")["sha256"],
            f"{stage} binary metadata is inconsistent")
    return {
        **binary,
        "source_manifest_sha256": sha(manifest),
        "stage": stage,
    }


def expected_command(stage: str, job: dict[str, Any], plan: dict[str, Any],
                     binary: dict[str, Any]) -> list[str]:
    folder = HERE / stage
    name = job["name"]
    return [
        "taskset", "-c", str(plan["cpu"]), "perf", "stat", "-x", ",",
        "-o", str(folder / f"{name}.csv"), "-e", plan["hardware"]["events"], "--",
        binary["path"], "--warmup", str(job["warmup"]),
        "--samples", str(job["samples"]), "--case", job["case"],
        "--xlsx-cell-crud-shape", job["shape"],
        "--json", str(folder / f"{name}.json"),
    ]


def validate_artifacts(folder: Path, receipt: dict[str, Any], expected: set[str],
                       label: str) -> dict[str, str]:
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict) and set(artifacts) == expected,
            f"{label} artifact inventory differs")
    for filename, digest in artifacts.items():
        require(isinstance(filename, str) and Path(filename).name == filename,
                f"{label} artifact is not stage-local: {filename!r}")
        path = folder / filename
        regular(path, f"{label}/{filename}")
        require(sha(path) == digest, f"{label}/{filename} digest differs")
    return {name: artifacts[name] for name in sorted(artifacts)}


def validate_receipt(stage: str, job: dict[str, Any], plan: dict[str, Any],
                     binary: dict[str, Any]) -> dict[str, Any]:
    folder = HERE / stage
    path = folder / f"{job['name']}.receipt.json"
    regular(path, relative(path))
    receipt = read_json(path)
    require(isinstance(receipt, dict), f"{relative(path)} is not an object")
    require(receipt.get("exit_code") == 0, f"{relative(path)} child failed")
    require(receipt.get("binary_sha256") == binary["sha256"],
            f"{relative(path)} binary binding differs")
    require(receipt.get("source_manifest_sha256") == binary["source_manifest_sha256"],
            f"{relative(path)} source binding differs")
    candidate_manifest = HERE / "candidate" / "source-manifest.json"
    expected_working = sha(candidate_manifest) if candidate_manifest.is_file() \
        else binary["source_manifest_sha256"]
    require(receipt.get("working_source_manifest_sha256") == expected_working,
            f"{relative(path)} working source binding differs")
    require(receipt.get("plan_sha256") == sha(PLAN_PATH),
            f"{relative(path)} plan binding differs")
    require(receipt.get("script_sha256") == sha(HERE / "run.py"),
            f"{relative(path)} run script binding differs")
    require(receipt.get("command") == expected_command(stage, job, plan, binary),
            f"{relative(path)} command differs from the frozen hardware command")
    artifacts = validate_artifacts(
        folder, receipt,
        {f"{job['name']}.csv", f"{job['name']}.json",
         f"{job['name']}.stdout", f"{job['name']}.stderr"},
        relative(path),
    )
    return {
        "file": relative(path),
        "sha256": sha(path),
        "start_utc": receipt.get("start_utc"),
        "end_utc": receipt.get("end_utc"),
        "artifacts": artifacts,
    }


def parse_events(path: Path) -> dict[str, dict[str, Any]]:
    label = relative(path)
    regular(path, label)
    events: dict[str, dict[str, Any]] = {}
    try:
        rows = csv.reader(path.read_text(encoding="utf-8", errors="replace").splitlines())
        for fields in rows:
            if not fields or not fields[0].strip() or fields[0].strip().startswith("#"):
                continue
            require(len(fields) >= 5, f"{label} malformed CSV row: {fields!r}")
            count, unit, event, runtime, running = (item.strip() for item in fields[:5])
            require(event in EVENTS, f"{label} contains unexpected event {event!r}")
            require(event not in events, f"{label} contains duplicate event {event!r}")
            unavailable = (
                not count or count.startswith("<") or
                not runtime or runtime.startswith("<") or
                not running or running.startswith("<")
            )
            if unavailable:
                events[event] = {"status": "unavailable", "unit": unit, "raw": fields}
                continue
            try:
                value = int(count.replace(",", ""))
                duration = int(runtime.replace(",", ""))
                coverage = float(running)
            except ValueError as error:
                raise EvidenceError(f"{label} has malformed measured row: {fields!r}") from error
            require(value >= 0 and duration > 0 and 0.0 <= coverage <= 100.0,
                    f"{label} has invalid measured values: {fields!r}")
            events[event] = {
                "status": "measured",
                "unit": unit,
                "value": value,
                "event_runtime_ns": duration,
                "running_percent": coverage,
            }
    except OSError as error:
        raise EvidenceError(f"cannot read {label}: {error}") from error
    require(set(events) == set(EVENTS),
            f"{label} event coverage differs: {sorted(set(EVENTS) - set(events))}")
    return events


def event_coverage(events: dict[str, dict[str, Any]]) -> dict[str, Any]:
    unavailable = [event for event in EVENTS if events[event]["status"] != "measured"]
    measured = [event for event in EVENTS if events[event]["status"] == "measured"]
    group_rows = [events[event] for event in GROUP_EVENTS]
    group_measured = all(row["status"] == "measured" for row in group_rows)
    runtimes = {row["event_runtime_ns"] for row in group_rows if row["status"] == "measured"}
    multiplexed = [
        event for event in EVENTS
        if events[event]["status"] == "measured"
        and (not math.isclose(events[event]["running_percent"], 100.0, abs_tol=1e-9)
             or (event in GROUP_EVENTS and len(runtimes) == 1
                 and events[event]["event_runtime_ns"] not in runtimes))
    ]
    # A grouped counter is complete only when every member ran for the same
    # duration and perf reports 100% time enabled/running for each member.
    same_runtime = group_measured and len(runtimes) == 1
    full_group_coverage = group_measured and same_runtime and not any(
        event in multiplexed for event in GROUP_EVENTS
    )
    result: dict[str, Any] = {
        "expected_events": list(EVENTS),
        "measured_events": measured,
        "unavailable_events": unavailable,
        "multiplexed_events": multiplexed,
        "group_events": list(GROUP_EVENTS),
        "group_measured": group_measured,
        "group_same_runtime": same_runtime,
        "group_full_coverage": full_group_coverage,
        "group_runtime_ns": next(iter(runtimes)) if same_runtime else None,
        "running_percent": {
            event: (events[event].get("running_percent")
                    if events[event]["status"] == "measured" else None)
            for event in EVENTS
        },
    }
    return result


def _native_identity(stage: str, repeat: int, shape: str, plan: dict[str, Any],
                     binary: dict[str, Any]) -> dict[str, Any]:
    folder = HERE / stage
    name = f"native-r{repeat}-primary-{shape}"
    job = native_job(plan, repeat, shape)
    native_receipt, row = BASE.check_receipt(folder, plan, job, binary, False)
    candidate_manifest = HERE / "candidate" / "source-manifest.json"
    allowed_working_manifests = {binary["source_manifest_sha256"]}
    if candidate_manifest.is_file():
        allowed_working_manifests.add(sha(candidate_manifest))
    require(native_receipt.get("working_source_manifest_sha256") in
            allowed_working_manifests,
            f"{stage}/{name} working source binding is outside retained manifests")
    return {
        "name": name,
        "identity": row["identity"],
        "identity_normalized": BASE.normalize_iteration_counts(row["identity"]),
        "report_sha256": sha(folder / f"{name}.json"),
        "receipt_sha256": sha(folder / f"{name}.receipt.json"),
    }


def analyze_capture(stage: str, job: dict[str, Any], plan: dict[str, Any],
                    binary: dict[str, Any]) -> dict[str, Any]:
    folder = HERE / stage
    receipt = validate_receipt(stage, job, plan, binary)
    report_path = folder / f"{job['name']}.json"
    csv_path = folder / f"{job['name']}.csv"
    native = _native_identity(stage, job["repeat"], job["shape"], plan, binary)
    report_plan = hardware_plan_for_report(plan)
    parsed = BASE.validate_result(
        read_json(report_path), report_plan, hardware_job(job), binary, False
    )
    normalized = BASE.normalize_iteration_counts(parsed["identity"])
    require(normalized == native["identity_normalized"],
            f"{job['name']} hardware/native logical result identity differs")
    events = parse_events(csv_path)
    coverage = event_coverage(events)
    row: dict[str, Any] = {
        "name": job["name"],
        "repeat": job["repeat"],
        "shape": job["shape"],
        "status": "measured" if coverage["group_full_coverage"]
        else "unavailable_for_group_claim",
        "receipt": receipt,
        "csv": {"file": relative(csv_path), "sha256": sha(csv_path)},
        "report": {"file": relative(report_path), "sha256": sha(report_path)},
        "native": native,
        "identity_equal": True,
        "events": events,
        "event_coverage": coverage,
        "scope": "whole_child",
    }
    if coverage["group_full_coverage"]:
        values = {event: events[event]["value"] for event in GROUP_EVENTS}
        require(values["cycles"] > 0, f"{job['name']} cycles counter is zero")
        require(values["branches"] >= values["branch-misses"],
                f"{job['name']} branch-miss count exceeds branch count")
        row["group_metrics"] = {
            "cycles": values["cycles"],
            "instructions": values["instructions"],
            "branches": values["branches"],
            "branch_misses": values["branch-misses"],
            "ipc": values["instructions"] / values["cycles"],
            "branch_miss_percent": (
                100.0 * values["branch-misses"] / values["branches"]
                if values["branches"] else 0.0
            ),
        }
    else:
        row["group_metrics"] = None
        row["reason"] = (
            "At least one grouped event is unavailable, multiplexed, or has "
            "incomplete running coverage; raw counters remain retained."
        )
    return row


def analyze_stage(stage: str, plan: dict[str, Any]) -> dict[str, Any]:
    folder = HERE / stage
    require(folder.is_dir(), f"{stage} stage directory is missing")
    binary = binary_metadata(stage, plan)
    captures = [analyze_capture(stage, job, plan, binary)
                for job in expected_jobs(plan)]
    return {
        "stage": stage,
        "status": "pass",
        "binary": {key: binary[key] for key in
                   ("sha256", "path", "bytes", "source_manifest_sha256")},
        "captures": captures,
        "capture_count": len(captures),
        "source_manifest_sha256": binary["source_manifest_sha256"],
        "validation": {
            "exact_source_binary_receipt_bindings": True,
            "event_set_complete_for_every_row": True,
            "native_logical_identity_parity": True,
            "whole_child_scope_retained": True,
        },
    }


def _percent_change(baseline: float, candidate: float) -> float | None:
    if baseline == 0.0:
        return 0.0 if candidate == 0.0 else None
    return (candidate / baseline - 1.0) * 100.0


def _counter_metric(baseline: dict[str, Any], candidate: dict[str, Any],
                   event: str) -> dict[str, Any]:
    left = baseline.get(event)
    right = candidate.get(event)
    if left is None or right is None or left.get("status") != "measured" \
            or right.get("status") != "measured":
        return {
            "status": "unavailable",
            "baseline": left,
            "candidate": right,
            "change_percent": None,
        }
    return {
        "status": "measured",
        "baseline": left["value"],
        "candidate": right["value"],
        "change_percent": _percent_change(float(left["value"]), float(right["value"])),
        "baseline_running_percent": left["running_percent"],
        "candidate_running_percent": right["running_percent"],
    }


def compare_stages(baseline: dict[str, Any], candidate: dict[str, Any]) -> dict[str, Any]:
    left = {(row["repeat"], row["shape"]): row for row in baseline["captures"]}
    right = {(row["repeat"], row["shape"]): row for row in candidate["captures"]}
    require(set(left) == set(right), "baseline/candidate hardware matrices differ")
    rows: list[dict[str, Any]] = []
    for key in sorted(left):
        base, cand = left[key], right[key]
        require(base["identity_equal"] and cand["identity_equal"],
                f"{key} hardware/native identity was not established")
        require(base["native"]["identity_normalized"] == cand["native"]["identity_normalized"],
                f"{key} baseline/candidate logical identity differs")
        metrics = {
            event: _counter_metric(base["events"], cand["events"], event)
            for event in EVENTS
        }
        group = None
        if base["group_metrics"] is not None and cand["group_metrics"] is not None:
            group = {
                metric: {
                    "baseline": base["group_metrics"][metric],
                    "candidate": cand["group_metrics"][metric],
                    "change_percent": _percent_change(
                        float(base["group_metrics"][metric]),
                        float(cand["group_metrics"][metric]),
                    ),
                }
                for metric in ("ipc", "branch_miss_percent")
            }
        rows.append({
            "repeat": key[0],
            "shape": key[1],
            "baseline": base["name"],
            "candidate": cand["name"],
            "event_metrics": metrics,
            "group_metrics": group,
            "baseline_coverage": base["event_coverage"],
            "candidate_coverage": cand["event_coverage"],
        })
    return {
        "rows": rows,
        "scope": "Whole fresh child, including corpus/setup/publication/oracles; diagnostic only.",
        "latency_gate": False,
        "operation_local_claim": False,
        "isolated_planning_counter_claim": False,
    }


def native_admission() -> dict[str, Any]:
    path = HERE / "comparison.json"
    regular(path, "0546 native comparison")
    value = read_json(path)
    require(isinstance(value, dict), "0546 native comparison is not an object")
    require(value.get("plan_sha256") == sha(PLAN_PATH),
            "0546 native comparison/plan binding differs")
    admission = value.get("native_admission")
    require(value.get("status") == "pass" and value.get("admission_status") ==
            "eligible-for-conditional-lanes" and isinstance(admission, dict)
            and admission.get("passed") is True,
            "0546 native gates did not authorize hardware diagnostics")
    return {
        "file": relative(path),
        "sha256": sha(path),
        "status": value["status"],
        "admission_status": value["admission_status"],
        "native_rows": len(admission.get("rows", [])),
    }


def render_markdown(document: dict[str, Any]) -> str:
    lines = [
        "# 0546 whole-child hardware diagnostics",
        "",
        "Counters cover the complete fresh child, including corpus setup, planning, publication, and output oracles.",
        "They remain diagnostics; they do not establish latency, operation-local cost, or an isolated planning counter.",
        "",
        f"Status: **{document['status']}**.",
        "",
    ]
    comparison = document.get("comparison")
    if isinstance(comparison, dict):
        lines.extend(("| Repeat | Shape | Cycles change | Instructions change | IPC change | Branch-miss change |", "| ---: | --- | ---: | ---: | ---: | ---: |"))
        for row in comparison["rows"]:
            event = row["event_metrics"]
            group = row.get("group_metrics") or {}
            def pct(value: Any) -> str:
                return "unavailable" if not isinstance(value, dict) or value.get("change_percent") is None \
                    else f"{value['change_percent']:.3f}%"
            lines.append(
                f"| {row['repeat']} | {row['shape']} | {pct(event['cycles'])} | "
                f"{pct(event['instructions'])} | {pct(group.get('ipc'))} | "
                f"{pct(group.get('branch_miss_percent'))} |"
            )
        lines.extend(("", "Every row retains the full event coverage and multiplexing record in the JSON report.", ""))
    for stage in ("baseline", "candidate"):
        value = document.get("stages", {}).get(stage)
        if not isinstance(value, dict):
            continue
        lines.append(f"## {stage}")
        lines.append("")
        lines.append(f"Capture rows: **{value.get('capture_count', 0)}**.")
        lines.append("")
        for row in value.get("captures", []):
            coverage = row["event_coverage"]
            lines.append(
                f"- `{row['name']}`: {row['status']}; grouped coverage "
                f"`{'complete' if coverage['group_full_coverage'] else 'incomplete'}`; "
                f"multiplexed events `{', '.join(coverage['multiplexed_events']) or 'none'}`."
            )
        lines.append("")
    return "\n".join(lines)


def _missing_stage_artifacts(stage: str, plan: dict[str, Any]) -> list[str]:
    """List absent inputs without opening any partial capture.

    This probe is deliberately limited to path presence.  Existing artifacts
    still go through the strict validators below, so a malformed partial file
    is never converted into a successful or synthetic hardware result.
    """

    folder = HERE / stage
    if not folder.is_dir():
        return [stage]
    required = [
        folder / "source-manifest.json",
        folder / "binary-normal.json",
        folder / "build-normal.receipt.json",
    ]
    for job in expected_jobs(plan):
        required.extend(
            folder / f"{job['name']}{suffix}"
            for suffix in (".receipt.json", ".csv", ".json", ".stdout", ".stderr")
        )
    # Hardware rows also bind to the primary native row in the same stage.
    # Treat that dependency as pending while the native lane is still being
    # captured rather than manufacturing a hardware/native identity.
    primary = plan["primary"]
    for repeat in range(1, int(primary["repeats"]) + 1):
        for shape in primary["shapes"]:
            name = f"native-r{repeat}-primary-{shape}"
            required.extend(
                folder / f"{name}{suffix}"
                for suffix in (".receipt.json", ".json", ".stdout", ".stderr", ".rss.json")
            )
    return [relative(path) for path in required if not path.exists()]


def _pending_document(stage_selection: str, plan: dict[str, Any],
                      missing: dict[str, list[str]], reason: str | None = None) -> dict[str, Any]:
    selected = ("baseline",) if stage_selection == "baseline" else \
        ("candidate",) if stage_selection == "candidate" else ("baseline", "candidate")
    stages = {
        stage: {
            "status": "pending",
            "missing_artifacts": missing.get(stage, []),
        }
        for stage in selected
    }
    document: dict[str, Any] = {
        "schema": "xlsx_shared_traversal_whole_child_hardware_analysis_v1",
        "status": "pending",
        "stage_selection": list(selected),
        "plan": relative(PLAN_PATH),
        "plan_sha256": sha(PLAN_PATH),
        "run_script": relative(HERE / "run.py"),
        "run_script_sha256": sha(HERE / "run.py"),
        "numeric_analyzer": {
            "path": relative(NUMERIC_PATH),
            "sha256": sha(NUMERIC_PATH),
        },
        "stages": stages,
        "scope": plan["hardware"]["scope"],
        "limitations": [
            "Hardware captures are incomplete; no counters or derived metrics are fabricated.",
            "The hardware lane covers a whole fresh child, including setup, corpus generation, publication and output oracles.",
            "No latency gate, operation-local attribution or isolated planning counter follows.",
        ],
    }
    if reason is not None:
        document["reason"] = reason
    return document


def analyze(stage_selection: str = "both") -> dict[str, Any]:
    require(stage_selection in ("baseline", "candidate", "both"),
            f"unknown hardware stage selection: {stage_selection}")
    plan = plan_data()
    selected = ("baseline",) if stage_selection == "baseline" else \
        ("candidate",) if stage_selection == "candidate" else ("baseline", "candidate")
    missing = {stage: _missing_stage_artifacts(stage, plan) for stage in selected}
    if any(missing[stage] for stage in selected):
        return _pending_document(stage_selection, plan, missing)
    if stage_selection == "both" and not (HERE / "comparison.json").is_file():
        return _pending_document(stage_selection, plan,
                                  {stage: [] for stage in selected},
                                  "comparison.json is missing; hardware is conditional on native admission")
    stages = {stage: analyze_stage(stage, plan) for stage in selected}
    document: dict[str, Any] = {
        "schema": "xlsx_shared_traversal_whole_child_hardware_analysis_v1",
        "status": "pass",
        "stage_selection": list(selected),
        "plan": relative(PLAN_PATH),
        "plan_sha256": sha(PLAN_PATH),
        "run_script": relative(HERE / "run.py"),
        "run_script_sha256": sha(HERE / "run.py"),
        "numeric_analyzer": {
            "path": relative(NUMERIC_PATH),
            "sha256": sha(NUMERIC_PATH),
        },
        "stages": stages,
        "scope": plan["hardware"]["scope"],
        "limitations": [
            "Counters cover the whole fresh child and include setup, corpus generation, publication and output oracles.",
            "IPC and branch-miss percentage are reported only for complete grouped counter coverage.",
            "No latency gate, operation-local attribution, isolated planning counter, cold-cache, range, scaling or native-producer claim follows.",
        ],
    }
    if set(stages) == {"baseline", "candidate"}:
        document["native_admission"] = native_admission()
        document["comparison"] = compare_stages(stages["baseline"], stages["candidate"])
    return document


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--stage", choices=("baseline", "candidate", "both"), default="both")
    parser.add_argument("--write", action="store_true",
                        help="write the validated JSON and Markdown artifacts")
    parser.add_argument("--output", type=Path,
                        help="JSON output path; defaults to hardware-analysis.json")
    parser.add_argument("--markdown-output", type=Path,
                        help="Markdown output path; defaults beside JSON output")
    args = parser.parse_args()
    output = args.output or HERE / "hardware-analysis.json"
    markdown = args.markdown_output or output.with_suffix(".md")
    try:
        document = analyze(args.stage)
        if args.write:
            output.parent.mkdir(parents=True, exist_ok=True)
            output.write_text(json.dumps(document, indent=2, sort_keys=True) + "\n",
                              encoding="utf-8")
            markdown.parent.mkdir(parents=True, exist_ok=True)
            markdown.write_text(render_markdown(document), encoding="utf-8")
    except (EvidenceError, OSError, json.JSONDecodeError, ValueError, KeyError) as error:
        print(f"analyze_hardware.py: error: {error}", file=sys.stderr)
        return 2
    if args.write:
        print(f"0546 whole-child hardware diagnostic verified: {output}")
    else:
        print(json.dumps(document, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
