#!/usr/bin/env python3
"""Validate and compare the frozen 0531 final native ABBA capture.

The final checkout was rebuilt after test-only corrections and therefore has a
different release binary from the measured candidate.  This analyzer binds a
fresh 44-child matrix to ``final-native-plan.json`` and ``final_capture.py``.
It validates the exact receipt command, source-manifest custody, binary hash,
artifact inventory, raw result identity/oracles, and global ABBA order before
performing the same native admission arithmetic as ``analyze.py``.

The ``final`` stage is the candidate side of the arithmetic, while all output
keeps its true ``final`` stage, source-manifest, and binary labels.  A missing
child produces a pending envelope without synthetic metrics.  No capture,
build, profile, or source operation is performed here.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import importlib.util
import json
import os
import random
import sys
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
NATIVE_PATH = HERE / "analyze.py"
PLAN_PATH = HERE / "plan.json"
FINAL_PLAN_PATH = HERE / "final-native-plan.json"
FROZEN_INPUTS_PATH = HERE / "final-native-frozen-inputs.json"
FINAL_CAPTURE_PATH = HERE / "final_capture.py"
RUN_PATH = HERE / "run.py"
SOURCE_BINDING_PATH = HERE / "final-source-binding.json"
QUALITY_SUMMARY_PATH = HERE / "final-quality-summary.json"

_spec = importlib.util.spec_from_file_location("mce_native_0531_final", NATIVE_PATH)
if _spec is None or _spec.loader is None:
    raise ImportError(f"cannot load native analyzer: {NATIVE_PATH}")
NATIVE = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(NATIVE)
BASE = NATIVE.BASE

PRIMARY_CASE = "xlsx_source_backed_cell_values_one_percent_edit_save"
TIMING_STATS = ("p50", "p95", "p99", "mean")
TIME_FORMAT = '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}'
EXPECTED_ORDER = [("baseline", 1), ("final", 1), ("final", 2), ("baseline", 2)]
EXPECTED_CHILDREN = 44


def require(condition: bool, message: str) -> None:
    if not condition:
        raise BASE.EvidenceError(message)


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read_json(path: Path, label: str | None = None) -> Any:
    require(path.is_file() and not path.is_symlink(),
            f"{label or path.name} is missing or not a regular file")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as error:
        raise BASE.EvidenceError(f"{label or path.name} is not valid JSON: {error}") from error


def utc(value: Any, label: str) -> _datetime.datetime:
    require(isinstance(value, str) and value, f"{label} timestamp is missing")
    try:
        parsed = _datetime.datetime.fromisoformat(value)
    except ValueError as error:
        raise BASE.EvidenceError(f"{label} timestamp is malformed") from error
    require(parsed.tzinfo is not None, f"{label} timestamp has no timezone")
    return parsed


def load_frozen() -> tuple[dict[str, Any], dict[str, Any], dict[str, Any], dict[str, Any]]:
    plan = read_json(PLAN_PATH, "primary plan")
    final_plan = read_json(FINAL_PLAN_PATH, "final native plan")
    frozen = read_json(FROZEN_INPUTS_PATH, "final native frozen inputs")
    source_binding = read_json(SOURCE_BINDING_PATH, "final source binding")
    require(isinstance(plan, dict), "primary plan is not an object")
    require(isinstance(final_plan, dict), "final native plan is not an object")
    require(isinstance(frozen, dict), "final native frozen inputs are not an object")
    require(isinstance(source_binding, dict), "final source binding is not an object")
    require(final_plan.get("schema") == "litchi-0531-final-native-plan-v1",
            "final native plan schema differs")
    require(frozen.get("final-native-plan.json") == sha(FINAL_PLAN_PATH),
            "final native plan frozen-input hash differs")
    require(frozen.get("final_capture.py") == sha(FINAL_CAPTURE_PATH),
            "final capture frozen-input hash differs")
    require(final_plan.get("primary_plan_sha256") == sha(PLAN_PATH),
            "final native plan is not bound to the primary plan")
    require(final_plan.get("gates") == plan.get("gates"),
            "final native gates differ from the primary plan")
    require(final_plan.get("order") == [list(item) for item in EXPECTED_ORDER],
            "final native ABBA order differs from the frozen order")
    require(final_plan.get("baseline_binary_sha256"),
            "final native baseline binary hash is missing")
    require(final_plan.get("final_binary_sha256"),
            "final native binary hash is missing")
    require(final_plan.get("baseline_binary_sha256") != final_plan.get("final_binary_sha256"),
            "final native baseline and final binary hashes unexpectedly match")
    require(final_plan.get("final_source_manifest_sha256") == source_binding.get("final_manifest_sha256"),
            "final source binding and final native plan manifest hashes differ")
    require(plan.get("primary", {}).get("case") == PRIMARY_CASE,
            "primary plan case differs from the frozen XLSX case")
    NATIVE._validated_gates(plan)
    jobs = NATIVE._expected_native_jobs(plan)
    require(len(jobs) == 22, f"primary native job count differs from 22: {len(jobs)}")
    require(sum(1 for job in jobs if job["kind"] == "primary") == 4,
            "primary native matrix does not contain four rows")
    require(sum(1 for job in jobs if job["kind"] == "guard") == 18,
            "guard native matrix does not contain eighteen rows")
    require(len({job["guard"] for job in jobs if job["kind"] == "guard"}) == len(plan["guards"]),
            "guard id inventory differs from the primary plan")
    require(final_plan.get("matrix", "").find("44") >= 0,
            "final native plan does not state the complete 44-child matrix")
    return plan, final_plan, frozen, source_binding


def native_jobs(plan: dict[str, Any], repeat: int) -> list[dict[str, Any]]:
    result = []
    for original in NATIVE._expected_native_jobs(plan):
        if original["repeat"] != repeat:
            continue
        job = dict(original)
        job["name"] = "final-" + original["name"]
        result.append(job)
    require(result, f"no final native jobs for repeat {repeat}")
    return result


def stage_jobs(plan: dict[str, Any], final_plan: dict[str, Any], stage: str) -> list[dict[str, Any]]:
    repeats = [repeat for planned_stage, repeat in EXPECTED_ORDER if planned_stage == stage]
    require(repeats, f"final native order has no {stage} repeat")
    jobs: list[dict[str, Any]] = []
    for repeat in repeats:
        jobs.extend(native_jobs(plan, repeat))
    return jobs


def expected_capture_command(stage: str, job: dict[str, Any], plan: dict[str, Any],
                             binary: dict[str, Any], folder: Path) -> list[str]:
    scratch = Path(plan["owned_paths"][0])
    return [
        "taskset", "-c", str(plan["cpu"]), "/usr/bin/time", "-f", TIME_FORMAT,
        "-o", str(folder / f"{job['name']}.rss.json"),
        str(scratch / f"{stage}-normal"),
        "--warmup", str(job["warmup"]), "--samples", str(job["samples"]),
        "--case", job["case"], "--xlsx-cell-crud-shape", job["shape"],
        "--json", str(folder / f"{job['name']}.json"),
    ]


def expected_build_command(plan: dict[str, Any]) -> list[str]:
    return [
        "env", "CARGO_BUILD_JOBS=2", "CARGO_INCREMENTAL=0", "cargo", "build",
        "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml",
        "--bin", "litchi-perf-baseline", "--target-dir", str(plan["owned_paths"][1]),
    ]


def check_artifacts(folder: Path, name: str, receipt: dict[str, Any],
                   expected_names: set[str]) -> None:
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict) and set(artifacts) == expected_names,
            f"{name} artifact inventory differs")
    for filename, digest in artifacts.items():
        artifact = folder / filename
        require(artifact.is_file() and not artifact.is_symlink(),
                f"{name} artifact is missing or not regular: {filename}")
        require(isinstance(digest, str) and digest == sha(artifact),
                f"{name} artifact digest differs: {filename}")


def check_build(stage: str, plan: dict[str, Any], final_plan: dict[str, Any],
                manifest_sha: str) -> dict[str, Any]:
    folder = HERE / stage
    identity = read_json(folder / "binary-normal.json", f"{stage} binary identity")
    require(isinstance(identity, dict), f"{stage} binary identity is not an object")
    expected_hash = (final_plan["baseline_binary_sha256"]
                     if stage == "baseline" else final_plan["final_binary_sha256"])
    expected_path = Path(plan["owned_paths"][0]) / f"{stage}-normal"
    require(identity.get("path") == str(expected_path),
            f"{stage} binary path differs from final capture binding")
    require(identity.get("sha256") == expected_hash,
            f"{stage} binary hash differs from final native plan")
    require(isinstance(identity.get("bytes"), int) and identity["bytes"] > 0,
            f"{stage} binary size is invalid")
    if expected_path.is_file() and not expected_path.is_symlink():
        require(sha(expected_path) == expected_hash
                and expected_path.stat().st_size == identity["bytes"],
                f"{stage} retained binary custody hash or size differs")
    else:
        # The final native matrix is retained after cleanup.  A post-cleanup
        # replay may therefore lack the two owned scratch trees, but it must
        # bind their removal to the primary plan and prove that both paths
        # are absent.  Do not expose this live-custody branch in the report:
        # the report remains a function of the frozen receipts and metrics.
        cleanup_path = HERE / "cleanup.json"
        cleanup = read_json(cleanup_path, "cleanup receipt")
        owned = plan.get("owned_paths")
        require(isinstance(owned, list) and all(isinstance(item, str) and item for item in owned),
                "plan.owned_paths is not a non-empty string list")
        require(cleanup.get("plan_sha256") == sha(PLAN_PATH),
                "cleanup plan hash differs")
        require(cleanup.get("removed") == owned
                and cleanup.get("owned_paths_absent") is True,
                "cleanup removed paths or absence marker differs")
        require(all(not os.path.lexists(path) for path in owned),
                "cleanup-owned path remains present")
    require(identity.get("source_manifest_sha256") == manifest_sha,
            f"{stage} binary source-manifest hash differs")

    build_receipt_path = folder / "build-normal.receipt.json"
    build_receipt = read_json(build_receipt_path, f"{stage} build receipt")
    require(isinstance(build_receipt, dict), f"{stage} build receipt is not an object")
    require(identity.get("build_receipt_sha256") == sha(build_receipt_path),
            f"{stage} binary build receipt hash differs")
    require(build_receipt.get("exit_code") == 0 and build_receipt.get("binary_sha256") is None,
            f"{stage} build receipt exit or binary field differs")
    require(build_receipt.get("command") == expected_build_command(plan),
            f"{stage} build command differs")
    require(build_receipt.get("plan_sha256") == sha(PLAN_PATH),
            f"{stage} build plan hash differs")
    require(build_receipt.get("script_sha256") == sha(RUN_PATH),
            f"{stage} build driver hash differs")
    require(build_receipt.get("source_manifest_sha256") == manifest_sha
            and build_receipt.get("working_source_manifest_sha256") == manifest_sha,
            f"{stage} build source custody differs")
    check_artifacts(folder, "build-normal", build_receipt,
                    {"build-normal.stdout", "build-normal.stderr"})
    return {
        "stage": stage,
        "label": "baseline" if stage == "baseline" else "final",
        **identity,
        "build_receipt_sha256": sha(build_receipt_path),
        "manifest_sha256": manifest_sha,
    }


def missing_artifacts(stage: str, jobs: list[dict[str, Any]]) -> list[str]:
    folder = HERE / stage
    missing: list[str] = []
    for filename in ("source-manifest.json", "binary-normal.json",
                     "build-normal.receipt.json"):
        if not (folder / filename).is_file():
            missing.append(f"{stage}/{filename}")
    for job in jobs:
        for suffix in (".host.json", ".json", ".rss.json", ".stdout", ".stderr", ".receipt.json"):
            if not (folder / (job["name"] + suffix)).is_file():
                missing.append(f"{stage}/{job['name']}{suffix}")
    return missing


def check_host(path: Path, name: str) -> dict[str, Any]:
    value = read_json(path, name)
    require(isinstance(value, dict), f"{name} host observation is not an object")
    require(isinstance(value.get("observed_utc"), str) and value["observed_utc"],
            f"{name} host observation timestamp is missing")
    require(isinstance(value.get("processes"), list),
            f"{name} host process list is malformed")
    require(value.get("scope") == "Pre-child observation; not proof of an idle host.",
            f"{name} host observation scope differs")
    return value


def check_receipt(stage: str, job: dict[str, Any], plan: dict[str, Any],
                  final_plan: dict[str, Any], binary: dict[str, Any],
                  manifest_sha: str, final_manifest_sha: str) -> tuple[dict[str, Any], dict[str, Any]]:
    folder = HERE / stage
    name = job["name"]
    receipt_path = folder / f"{name}.receipt.json"
    receipt = read_json(receipt_path, f"{stage}/{name} receipt")
    require(isinstance(receipt, dict), f"{name} receipt is not an object")
    require(receipt.get("exit_code") == 0, f"{name} did not exit successfully")
    require(receipt.get("command") == expected_capture_command(stage, job, plan, binary, folder),
            f"{name} command differs from final capture")
    require(receipt.get("binary_sha256") == binary["sha256"],
            f"{name} binary hash differs")
    require(receipt.get("plan_sha256") == sha(PLAN_PATH),
            f"{name} plan hash differs")
    require(receipt.get("script_sha256") == sha(RUN_PATH),
            f"{name} run driver hash differs")
    require(receipt.get("source_manifest_sha256") == manifest_sha,
            f"{name} stage source-manifest hash differs")
    # final_capture.py intentionally checks the final checkout for baseline
    # children while retaining the baseline stage manifest in the receipt.
    require(receipt.get("working_source_manifest_sha256") == final_manifest_sha,
            f"{name} working source-manifest custody differs")
    start = utc(receipt.get("start_utc"), f"{name} start")
    end = utc(receipt.get("end_utc"), f"{name} end")
    require(end >= start, f"{name} receipt ends before it starts")
    check_artifacts(
        folder, name, receipt,
        {f"{name}.host.json", f"{name}.json", f"{name}.rss.json",
         f"{name}.stdout", f"{name}.stderr"},
    )
    host = check_host(folder / f"{name}.host.json", f"{stage}/{name}")
    raw = read_json(folder / f"{name}.json", f"{stage}/{name} raw result")
    row = NATIVE._validate_result(raw, plan, job, binary, False)
    rss = BASE.validate_rss(folder / f"{name}.rss.json")
    row["stage"] = stage
    row["true_binary_label"] = binary["label"]
    row["true_source_manifest_sha256"] = receipt["working_source_manifest_sha256"]
    row["host_observation"] = {
        "path": str(folder / f"{name}.host.json"),
        "observed_utc": host["observed_utc"],
        "process_count": len(host["processes"]),
        "scope": host["scope"],
    }
    row["rss"] = {"scope": "whole_child_process", **rss}
    return receipt, row


def check_stage(stage: str, plan: dict[str, Any], final_plan: dict[str, Any],
                final_manifest_sha: str) -> tuple[dict[str, Any], list[tuple[str, dict[str, Any]]]]:
    folder = HERE / stage
    require(folder.is_dir(), f"{stage} stage directory is missing")
    manifest_path = folder / "source-manifest.json"
    manifest = read_json(manifest_path, f"{stage} source manifest")
    require(isinstance(manifest, dict) and manifest, f"{stage} source manifest is empty")
    manifest_sha = sha(manifest_path)
    if stage == "final":
        require(manifest_sha == final_manifest_sha,
                "final source manifest differs from final native plan")
    jobs = stage_jobs(plan, final_plan, stage)
    expected_names = {job["name"] for job in jobs}
    actual_names = {
        path.name[:-len(".receipt.json")]
        for path in folder.glob("final-native-*.receipt.json")
    }
    require(actual_names == expected_names,
            f"{stage} final-native receipt set differs: {sorted(actual_names ^ expected_names)}")
    binary = check_build(stage, plan, final_plan, manifest_sha)
    rows: list[dict[str, Any]] = []
    receipts: list[tuple[str, dict[str, Any]]] = []
    for job in jobs:
        receipt, row = check_receipt(stage, job, plan, final_plan, binary,
                                     manifest_sha, final_manifest_sha)
        rows.append(row)
        receipts.append((job["name"], receipt))
    actual_keys = {(row["kind"], row["guard"], row["repeat"], row["case"], row["shape"])
                  for row in rows}
    expected_keys = {(job["kind"], job["guard"], job["repeat"], job["case"], job["shape"])
                     for job in jobs}
    require(actual_keys == expected_keys and len(rows) == len(jobs),
            f"{stage} validated row matrix differs")
    rows.sort(key=lambda row: (row["kind"], -1 if row["guard"] is None else row["guard"],
                               row["repeat"], row["case"], row["shape"]))
    return {
        "stage": stage,
        "label": "baseline" if stage == "baseline" else "final",
        "manifest_sha256": manifest_sha,
        "binary_identity": binary,
        "native": {
            "rows": rows,
            "row_count": len(rows),
            "total_samples": sum(row["samples"] for row in rows),
        },
        "custody": {
            "receipt_count": len(receipts),
            "source_manifest_entries": len(manifest),
            "receipts_non_overlapping": True,
        },
    }, receipts


def validate_global_order(receipts: list[tuple[str, str, dict[str, Any]]],
                          plan: dict[str, Any]) -> list[dict[str, Any]]:
    ordered = sorted(receipts, key=lambda item: utc(item[2]["start_utc"], item[0]))
    expected: list[tuple[str, str]] = []
    capture_order: list[dict[str, Any]] = []
    ordinal = 0
    for stage, repeat in EXPECTED_ORDER:
        jobs = native_jobs(plan, repeat)
        names = [job["name"] for job in jobs]
        capture_order.append({"ordinal": ordinal + 1, "stage": stage,
                              "repeat": repeat, "job_names": names})
        expected.extend((stage, name) for name in names)
        ordinal += 1
    actual = [(stage, name) for name, stage, _ in ordered]
    require(actual == expected, "final native receipt order differs from frozen ABBA order")
    for (_, _, left), (_, _, right) in zip(ordered, ordered[1:]):
        require(utc(left["end_utc"], "receipt end") <= utc(right["start_utc"], "receipt start"),
                "final native child receipts overlap")
    return capture_order


def add_stage_labels(records: list[dict[str, Any]], repeat_stage: bool = False) -> None:
    for record in records:
        if repeat_stage:
            record["repeat_stage"] = record["stage"]
            record["baseline_stage"] = record["stage"]
            record["candidate_stage"] = record["stage"]
        else:
            record["baseline_stage"] = "baseline"
            record["candidate_stage"] = "final"


def compare_native(baseline: dict[str, Any], final: dict[str, Any]) -> dict[str, Any]:
    comparisons: list[dict[str, Any]] = []
    adverse: list[dict[str, Any]] = []
    drift: list[dict[str, Any]] = []
    rng = random.Random(BASE.BOOTSTRAP_SEED)
    NATIVE._compare_timing_lane(
        baseline, final, "native", rng, comparisons, adverse, drift,
    )
    add_stage_labels(comparisons)
    add_stage_labels(adverse)
    add_stage_labels(drift, repeat_stage=True)
    for record in comparisons:
        record["candidate_label"] = "final"
    for record in adverse:
        record["candidate_label"] = "final"
    return {
        "timing_comparisons": comparisons,
        "adverse_flags_over_five_percent": adverse,
        "same_build_drift_over_five_percent": drift,
        "bootstrap": {
            "iterations": BASE.BOOTSTRAP_ITERATIONS,
            "seed": BASE.BOOTSTRAP_SEED,
            "scope": "matched-child within-stage resampling; final median / baseline median",
        },
    }


def pending_report(plan: dict[str, Any], final_plan: dict[str, Any],
                   missing: list[str]) -> dict[str, Any]:
    return {
        "status": "pending",
        "stage": "final-native-compare",
        "plan_sha256": sha(PLAN_PATH),
        "final_plan_sha256": sha(FINAL_PLAN_PATH),
        "final_capture_sha256": sha(FINAL_CAPTURE_PATH),
        "structured_gates": dict(plan["gates"]),
        "admission_status": "pending",
        "missing_artifacts": missing,
        "matrix_validation": {
            "expected_children": EXPECTED_CHILDREN,
            "expected_rows_per_stage": 22,
            "status": "pending",
        },
        "scope": "Final native ABBA matrix; no synthetic metrics are produced while captures are incomplete.",
    }


def analyze() -> dict[str, Any]:
    plan, final_plan, frozen, source_binding = load_frozen()
    final_manifest_sha = final_plan["final_source_manifest_sha256"]
    all_jobs = {stage: stage_jobs(plan, final_plan, stage)
                for stage in ("baseline", "final")}
    missing = [item for stage in ("baseline", "final")
               for item in missing_artifacts(stage, all_jobs[stage])]
    if missing:
        return pending_report(plan, final_plan, missing)

    stage_evidence: dict[str, dict[str, Any]] = {}
    all_receipts: list[tuple[str, str, dict[str, Any]]] = []
    stage_receipts: dict[str, list[tuple[str, dict[str, Any]]]] = {}
    for stage in ("baseline", "final"):
        evidence, receipts = check_stage(stage, plan, final_plan, final_manifest_sha)
        stage_evidence[stage] = evidence
        stage_receipts[stage] = receipts
        all_receipts.extend((name, stage, receipt) for name, receipt in receipts)
    capture_order = validate_global_order(all_receipts, plan)

    baseline = stage_evidence["baseline"]
    final = stage_evidence["final"]
    require(baseline["native"]["row_count"] == 22 and final["native"]["row_count"] == 22,
            "final native stage row counts are not 22")
    require(baseline["native"]["total_samples"] == 1340
            and final["native"]["total_samples"] == 1340,
            "final native stage sample totals are not 1340")
    require(baseline["binary_identity"]["sha256"] == final_plan["baseline_binary_sha256"],
            "baseline true binary label does not match final plan")
    require(final["binary_identity"]["sha256"] == final_plan["final_binary_sha256"],
            "final true binary label does not match final plan")
    require(baseline["binary_identity"]["sha256"] != final["binary_identity"]["sha256"],
            "baseline and final binary identities unexpectedly match")

    comparison = compare_native(baseline, final)
    native_admission = NATIVE._native_admission(baseline, final, plan)
    native_admission["baseline_stage"] = "baseline"
    native_admission["candidate_stage"] = "final"
    for row in native_admission["rows"]:
        row["baseline_stage"] = "baseline"
        row["candidate_stage"] = "final"
    return {
        "status": "pass",
        "stage": "final-native-compare",
        # Keep the original analyzer's top-level plan binding compatible with
        # downstream adapters.  The final plan has an explicit separate hash.
        "plan_sha256": sha(PLAN_PATH),
        "final_plan_sha256": sha(FINAL_PLAN_PATH),
        "final_capture_sha256": sha(FINAL_CAPTURE_PATH),
        "frozen_inputs_sha256": sha(FROZEN_INPUTS_PATH),
        "source_binding_sha256": sha(SOURCE_BINDING_PATH),
        "structured_gates": dict(final_plan["gates"]),
        "baseline_stage": "baseline",
        "candidate_stage": "final",
        "candidate_label": "final",
        "baseline": baseline,
        # The final stage is exposed through the compatible candidate slot;
        # candidate_stage and true_stage_labels preserve its actual identity.
        "candidate": final,
        "true_stage_labels": {
            "baseline": {
                "stage": "baseline",
                "binary_sha256": baseline["binary_identity"]["sha256"],
                "source_manifest_sha256": baseline["manifest_sha256"],
            },
            "final": {
                "stage": "final",
                "binary_sha256": final["binary_identity"]["sha256"],
                "source_manifest_sha256": final["manifest_sha256"],
            },
        },
        "matrix_validation": {
            "status": "pass",
            "expected_children": EXPECTED_CHILDREN,
            "validated_children": len(all_receipts),
            "expected_rows_per_stage": 22,
            "validated_rows_per_stage": {
                "baseline": baseline["native"]["row_count"],
                "final": final["native"]["row_count"],
            },
            "primary_rows_per_stage": 4,
            "guard_rows_per_stage": 18,
            "guard_shape_jobs_per_repeat": 9,
            "guard_ids": len(plan["guards"]),
            "capture_order": capture_order,
        },
        "comparison": comparison,
        "native_admission": native_admission,
        "admission_status": native_admission["decision"],
        "allocation_diagnostics": {
            "status": "not-captured",
            "used_for_gate": False,
            "scope": "Final native matrix has no separate allocator lane; native admission is independent.",
        },
        "conditional_lanes": {
            "profile": {
                "status": "unmeasured",
                "required_reduction_percent": float(final_plan["gates"]["planning_ir_reduction_percent"]),
                "used_for_gate": False,
                "reason": "Final native evidence is complete; planning Ir still requires a separately bound profile.",
            },
        },
        "numerical_verifier": {
            "path": str(NATIVE_PATH.relative_to(HERE.parent)),
            "sha256": sha(NATIVE_PATH),
        },
        "scope": (
            "Final native XLSX primary plus shared OOXML guard ABBA matrix. "
            "Final is the candidate side for arithmetic; true baseline/final "
            "binary and source labels are retained. Publication is diagnostic; "
            "no speedup claim is made by this report."
        ),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path,
                        default=HERE / "final-native-comparison.json")
    args = parser.parse_args()
    try:
        result = analyze()
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n",
                               encoding="utf-8")
    except BASE.EvidenceError as error:
        print(f"final native evidence check failed: {error}", file=sys.stderr)
        return 1
    print(f"final native {result['stage']} evidence {result['status']}: {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
