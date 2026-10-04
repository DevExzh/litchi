"""Final independent offline replay and custody review for the 0900 phase packet."""
from __future__ import annotations

import hashlib
import json
import math
import pathlib
import statistics
from typing import Any

P = pathlib.Path(__file__).resolve().parent
ROOT = P.parents[3]
PHASES = ("open_ns", "plan_ns", "commit_ns", "publication_ns")
ALL_PHASES = PHASES + ("reopen_ns",)


def load(path: pathlib.Path) -> Any:
    return json.loads(path.read_text())


def digest(path: pathlib.Path) -> dict[str, Any]:
    data = path.read_bytes()
    return {
        "path": str(path.resolve()),
        "bytes": len(data),
        "sha256": hashlib.sha256(data).hexdigest(),
    }


def close(left: float, right: float) -> bool:
    return math.isclose(left, right, rel_tol=1e-12, abs_tol=1e-9)


def midpoint(values: list[int]) -> int:
    ordered = sorted(values)
    left = ordered[(len(ordered) - 1) // 2]
    right = ordered[len(ordered) // 2]
    return left // 2 + right // 2 + (left % 2 + right % 2) // 2


def nearest_rank(values: list[int], percentile: int) -> int:
    index = ((percentile * len(values) + 99) // 100) - 1
    return sorted(values)[min(index, len(values) - 1)]


def stats(values: list[int]) -> dict[str, int | float]:
    ordered = sorted(values)
    return {
        "p50": midpoint(values),
        "p95": nearest_rank(values, 95),
        "p99": nearest_rank(values, 99),
        "mean": statistics.mean(values),
        "min": ordered[0],
        "max": ordered[-1],
    }


def interval(values: list[int]) -> dict[str, float]:
    mean = statistics.mean(values)
    deviation = statistics.stdev(values) if len(values) > 1 else 0.0
    degrees = len(values) - 1
    z = 1.959963984540054
    df = float(degrees)
    critical = (z + (z**3 + z) / (4 * df)
                + (5 * z**5 + 16 * z**3 + 3 * z) / (96 * df**2)
                + (3 * z**7 + 19 * z**5 + 17 * z**3 - 15 * z) / (384 * df**3))
    margin = critical * deviation / math.sqrt(len(values))
    return {
        "lower": max(0.0, mean - margin),
        "upper": mean + margin,
        "standard_deviation": deviation,
    }


def command_for(plan: dict[str, Any], folder: pathlib.Path, shape: str, lane: str) -> list[str]:
    samples = plan["qualification_samples"] if lane == "qualification" else plan["samples"]
    warmup = plan["qualification_warmup"] if lane == "qualification" else plan["warmup"]
    return [
        "/usr/bin/taskset", "-c", str(plan["cpu"]), plan["binary"]["path"],
        "--warmup", str(warmup), "--samples", str(samples), "--case", plan["case"],
        "--xlsx-cell-crud-shape", shape, "--json", str(folder / "report.json"),
    ]


def schedule(plan: dict[str, Any], lane: str) -> list[tuple[int, str]]:
    if lane == "qualification":
        return [(0, shape) for shape in plan["shapes"]]
    return [
        (block, shape)
        for block in range(plan["native_repeats"])
        for shape in (plan["shapes"] if block % 2 == 0 else list(reversed(plan["shapes"])))
    ]


def row_review(plan: dict[str, Any], lane: str, block: int, shape: str) -> dict[str, Any]:
    folder = P / lane / f"b{block:02}-{shape}"
    command = load(folder / "command.json")
    start = load(folder / "start.json")
    report = load(folder / "report.json")
    validation = load(folder / "validation.json")
    samples = plan["qualification_samples"] if lane == "qualification" else plan["samples"]
    warmup = plan["qualification_warmup"] if lane == "qualification" else plan["warmup"]
    expected_argv = command_for(plan, folder, shape, lane)
    expected_helper = P / "run.py" if lane == "qualification" and block == 0 and shape == "medium" else P / "run_v2.py"
    required_artifacts = [folder / name for name in ("report.json", "start.json", "stderr.txt", "stdout.txt")]
    errors: list[str] = []

    def need(condition: bool, message: str) -> None:
        if not condition:
            errors.append(message)

    need(command["exit_code"] == 0, "command exit code")
    need(command["argv"] == expected_argv, "command argv")
    need(command["cwd"] == str(ROOT), "command cwd")
    need(command["environment"] == plan["environment"], "command environment")
    need(command["plan"] == digest(P / "plan.json"), "command plan descriptor")
    need(command["helper"] == digest(expected_helper), "command helper descriptor")
    need(command["binary"] == plan["binary"], "command binary descriptor")
    need(start["argv"] == expected_argv and start["cwd"] == str(ROOT), "start invocation")
    need(start["environment"] == plan["environment"], "start environment")
    need(start["plan"] == digest(P / "plan.json"), "start plan descriptor")
    need(start["helper"] == digest(expected_helper), "start helper descriptor")
    need(start["binary"] == plan["binary"], "start binary descriptor")
    need(report["schema_version"] == 1, "report schema")
    need(report["binary_identity"]["binary_sha256"] == plan["binary"]["sha256"], "report binary")
    need(report["tool"]["profile"] == "release" and report["tool"]["instrumentation"] == "none", "report tool")
    need(report["environment"]["cpu_affinity"] == str(plan["cpu"]), "report affinity")
    need(report["environment"]["rustc_version"] is None, "report rustc metadata")
    need(report["environment"]["git_revision"] is None, "report git revision metadata")
    need(report["environment"]["git_worktree_dirty"] is None, "report git dirty metadata")
    config = report["configuration"]
    need(config["cases"] == [plan["case"]], "report case")
    need(config["xlsx_cell_crud_shapes"] == [shape], "report shape")
    need(config["samples_per_case"] == samples, "report sample count")
    need(config["warmup_iterations_per_case"] == warmup, "report warmup count")
    need(len(report["results"]) == 1 and report["results"][0]["case"] == plan["case"], "report result")
    result = report["results"][0]
    summary = result["source"]["xlsx_cell_values"]
    vectors = {phase: summary[phase] for phase in ALL_PHASES}
    need(all(len(values) == samples for values in vectors.values()), "phase vector lengths")
    need(all(isinstance(value, int) and value >= 0 for values in vectors.values() for value in values), "phase vector values")
    phase_sums = [sum(vectors[phase][index] for phase in PHASES) for index in range(samples)]
    elapsed = result["elapsed_ns"]
    order = elapsed["sample_order"]
    need(sorted(order) == list(range(samples)), "sample order permutation")
    need(order == sorted(range(samples), key=lambda index: (phase_sums[index], index)), "phase sum order")
    need([phase_sums[index] for index in order] == elapsed["samples"], "phase sum alignment")
    measured = stats(elapsed["samples"])
    for key, value in measured.items():
        need(elapsed[key] == value if isinstance(value, int) else close(float(elapsed[key]), float(value)), f"elapsed {key}")
    ci = interval(elapsed["samples"])
    need(close(float(elapsed["standard_deviation"]), ci["standard_deviation"]), "elapsed standard deviation")
    need(close(float(elapsed["confidence_interval_95"]["lower"]), ci["lower"]), "elapsed confidence lower")
    need(close(float(elapsed["confidence_interval_95"]["upper"]), ci["upper"]), "elapsed confidence upper")
    phase_statistics = {phase: stats(vectors[phase]) for phase in ALL_PHASES}
    phase_shares = {phase: sum(vectors[phase]) / sum(elapsed["samples"]) for phase in PHASES}
    analysis_row = None
    # The canonical analysis is independently checked by the caller after all rows are collected.
    for path in required_artifacts:
        need(path.is_file(), f"missing artifact {path.name}")
    recorded = command["artifacts"]
    need(len(recorded) == len(required_artifacts), "command artifact count")
    need({item["path"] for item in recorded} == {str(path.resolve()) for path in required_artifacts}, "command artifact set")
    artifact_descriptors = []
    for path in required_artifacts:
        actual = digest(path)
        matches = [item for item in recorded if item["path"] == actual["path"]]
        need(len(matches) == 1 and matches[0] == actual, f"artifact custody {path.name}")
        artifact_descriptors.append(actual)
    review = {
        "lane": lane,
        "block": block,
        "shape": shape,
        "samples": samples,
        "warmup": warmup,
        "command": digest(folder / "command.json"),
        "start": digest(folder / "start.json"),
        "report": digest(folder / "report.json"),
        "validation": digest(folder / "validation.json"),
        "helper": digest(expected_helper),
        "artifact_descriptors": artifact_descriptors,
        "runtime_metadata": {
            "rustc_version": report["environment"]["rustc_version"],
            "git_revision": report["environment"]["git_revision"],
            "git_worktree_dirty": report["environment"]["git_worktree_dirty"],
        },
        "phase_sum_ns": sum(phase_sums),
        "elapsed_sum_ns": sum(elapsed["samples"]),
        "phase_sums_reconciled": not errors or all("phase sum" not in error and "phase vector" not in error for error in errors),
        "sample_order_reconciled": not any("sample order" in error or "phase sum order" in error or "phase sum alignment" in error for error in errors),
        "elapsed_statistics": measured,
        "phase_statistics": phase_statistics,
        "phase_share": phase_shares,
        "errors": errors,
    }
    return review


def main() -> None:
    plan = load(P / "plan.json")
    rows = []
    for lane in ("qualification", "native"):
        capture = load(P / lane / "capture.json")
        expected_children = 2 if lane == "qualification" else plan["native_repeats"] * len(plan["shapes"])
        if capture["status"] != "pass" or capture["children"] != expected_children:
            raise SystemExit(f"capture manifest mismatch: {lane}")
        for block, shape in schedule(plan, lane):
            rows.append(row_review(plan, lane, block, shape))
    errors = [f"{row['lane']}/{row['block']}/{row['shape']}: {error}" for row in rows for error in row["errors"]]
    analysis = load(P / "analysis.json")
    analysis_matches = True
    analysis_errors: list[str] = []
    if analysis.get("status") != "pass" or len(analysis.get("rows", [])) != len(rows):
        analysis_matches = False
        analysis_errors.append("analysis status/row count")
    else:
        for independent, canonical in zip(rows, analysis["rows"]):
            for key in ("lane", "block", "shape", "samples", "elapsed_statistics", "phase_statistics", "phase_share"):
                canonical_key = {"elapsed_statistics": "total", "phase_statistics": "phases", "phase_share": "phase_shares"}.get(key, key)
                if independent[key] != canonical[canonical_key]:
                    analysis_matches = False
                    analysis_errors.append(f"analysis row {independent['lane']}/{independent['block']}/{independent['shape']} {key}")
    native_summaries: dict[str, Any] = {}
    for shape in plan["shapes"]:
        selected = [row for row in rows if row["lane"] == "native" and row["shape"] == shape]
        p50_values = [row["elapsed_statistics"]["p50"] for row in selected]
        native_summaries[shape] = {
            "children": len(selected),
            "p50_ns_range": [min(p50_values), max(p50_values)],
            "p50_ns_median": statistics.median(p50_values),
            "p50_max_min_ratio": max(p50_values) / min(p50_values),
            "phase_share_ranges": {
                phase: [min(row["phase_share"][phase] for row in selected), max(row["phase_share"][phase] for row in selected)]
                for phase in PHASES
            },
        }
        native_summaries[shape]["variation_flag"] = native_summaries[shape]["p50_max_min_ratio"] > plan["variation_flag"]["threshold"]
    inputs = {
        "plan": digest(P / "plan.json"),
        "original_helper": digest(P / "run.py"),
        "corrected_helper": digest(P / "run_v2.py"),
        "metadata_addendum": digest(P / "metadata-addendum.json"),
        "scope_clarification": digest(P / "scope-clarification.json"),
        "source_scope": digest(P / "source-scope.md"),
        "methodology_review": digest(P / "methodology-review.md"),
        "review_helper": digest(P / "independent_review_v2.py"),
        "analysis": digest(P / "analysis.json"),
        "qualification_capture": digest(P / "qualification" / "capture.json"),
        "native_capture": digest(P / "native" / "capture.json"),
        "source_manifest": digest(pathlib.Path(plan["source_manifest"]["path"])),
        "build_receipt": digest(pathlib.Path(plan["build"]["path"])),
        "binary": digest(pathlib.Path(plan["binary"]["path"])),
        "semantic_validator": digest(pathlib.Path(plan["prior_semantic_validator"]["path"])),
        "timer_source": digest(ROOT / "tools/perf-baseline/src/lib.rs"),
    }
    output = {
        "schema": "litchi.xlsx-phase-independent-review-0900.v1",
        "status": "pass" if not errors and analysis_matches else "fail",
        "reviewer": "independent_phase_review_0900",
        "scope": {
            "mode": "offline-only retained-artifact review",
            "evaluation_head": plan["head"],
            "baseline_source_revision": plan["baseline_source_revision"],
            "iwork_excluded": True,
            "new_workload_build_or_profiler": False,
            "performance_claim": None,
            "adoption_decision": None,
            "qualification_excluded_from_native_summary": True,
            "timing_interpretation": "open, selector planning, staged sets plus commit, publication including returned publication snapshot drop; reopen and remaining drops are excluded",
        },
        "inputs": inputs,
        "custody": {
            "source_files_verified": plan["source_files_verified"],
            "source_manifest_matches_plan": inputs["source_manifest"] == plan["source_manifest"],
            "binary_matches_plan": inputs["binary"] == plan["binary"],
            "build_receipt_matches_plan": inputs["build_receipt"] == plan["build"],
            "command_artifacts_checked": sum(len(row["artifact_descriptors"]) for row in rows),
            "command_artifact_descriptor_errors": [],
            "runtime_metadata_expected_null": True,
            "runtime_metadata_all_null": all(not any(row["runtime_metadata"].values()) for row in rows),
        },
        "populations": {
            "qualification_children": sum(row["lane"] == "qualification" for row in rows),
            "qualification_samples": sum(row["samples"] for row in rows if row["lane"] == "qualification"),
            "native_children": sum(row["lane"] == "native" for row in rows),
            "native_samples": sum(row["samples"] for row in rows if row["lane"] == "native"),
            "raw_reports_replayed": len(rows),
            "native_schedule": [(row["block"], row["shape"]) for row in rows if row["lane"] == "native"],
        },
        "numeric_replay": {
            "phase_names": list(PHASES),
            "all_phase_names": list(ALL_PHASES),
            "phase_sum_reconciliation_pass": not any("phase sum" in error or "phase vector" in error for error in errors),
            "sample_order_reconciliation_pass": not any("sample order" in error or "phase sum order" in error or "phase sum alignment" in error for error in errors),
            "elapsed_statistics_recomputed": not any("elapsed " in error for error in errors),
            "analysis_rows_reconciled": analysis_matches,
            "analysis_errors": analysis_errors,
            "native_summaries": native_summaries,
            "qualification_phase_shares": {
                f"{row['shape']}": row["phase_share"] for row in rows if row["lane"] == "qualification"
            },
        },
        "rows": rows,
        "findings": {
            "blocking": errors + analysis_errors,
            "limitations": [
                "Qualification and native timing populations are separate; qualification is used for semantic admission only.",
                "The retained binary reports null runtime Git fields because its isolated compile-time manifest tree was removed; rustc_version is null because the frozen runtime PATH omits rustc. External source, build, and binary custody are authoritative.",
                "Phase shares describe this current-source diagnostic and do not establish a before/after speedup or production eligibility.",
            ],
        },
    }
    with (P / "review-final.json").open("x") as handle:
        handle.write(json.dumps(output, indent=2, sort_keys=True) + "\n")
    if errors or not analysis_matches:
        raise SystemExit("independent review failed: " + repr(errors + analysis_errors))


if __name__ == "__main__":
    main()
