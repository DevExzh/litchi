#!/usr/bin/env python3
"""Bounded native tail guard for the two 0519 elapsed p99 outliers.

The plan is deliberately frozen before capture.  Run ``--write-plan`` once,
then run ``--capture`` after the quality lane is terminal.  ``--analyze``
replays only the retained tail-guard CSVs.  The baseline executable is bound
to its retained build receipt and is never compared with the working-tree
baseline source.  Candidate captures bind both current-source snapshots to
the retained candidate build manifest.

The eight fresh processes run A1/B1/B2/A2 over both cases.  Each case has
two independent before/after pairs:

* A1: retained baseline, file-batch then owned-batch
* B1: retained candidate, file-batch then owned-batch
* B2: retained candidate, owned-batch then file-batch
* A2: retained baseline, owned-batch then file-batch

The first four-child draft was corrected before any raw tail-guard artifact
was created; this final plan therefore has two A/B pairs per case.

This is bounded diagnostic evidence.  It keeps the original comparison flags
and does not dismiss a flag as noise based on this follow-up.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import hashlib
import json
import math
import os
import random
import re
import subprocess
import sys
import time
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
SCRATCH = Path("/tmp/litchi-goal-0519")
PLAN_PATH = HERE / "tail-guard-plan.json"
OUTPUT_DIR = HERE / "tail-guard"
REPORT_JSON = HERE / "tail-guard.json"
REPORT_MD = HERE / "tail-guard.md"
ORIGINAL_COMPARISON = HERE / "candidate-comparison.json"

# Import the existing row validator through the same source-bound capture
# module used by the native campaigns.  Nothing is captured during import.
sys.path.insert(0, str(HERE))
import capture as capture_module  # noqa: E402
from run import sources  # noqa: E402


PLAN_SCHEMA = "managed_paragraph_tail_guard_plan_0519_v2"
CAPTURE_SCHEMA = "managed_paragraph_tail_guard_capture_0519_v2"
REPORT_SCHEMA = "managed_paragraph_tail_guard_0519_v2"
CPU = 2
SAMPLES = 200
WARMUPS = 10
REPEATS = 1
BOOTSTRAP_ITERATIONS = 4000
DEFAULT_SEED = 0x0519A11
CASES = {
    "1": "p128-k1-file-batch",
    "2": "p128-k1-owned-batch",
}
ORDER = (
    ("A1", "baseline", "1"),
    ("A1", "baseline", "2"),
    ("B1", "candidate", "1"),
    ("B1", "candidate", "2"),
    ("B2", "candidate", "2"),
    ("B2", "candidate", "1"),
    ("A2", "baseline", "2"),
    ("A2", "baseline", "1"),
)
PAIR_ORDER = (("A1", "B1"), ("A2", "B2"))
PHASES = ("elapsed_ns", "open_ns", "edit_ns", "commit_ns", "publish_ns", "drop_ns")
SHORT_PHASES = ("open_ns", "edit_ns", "commit_ns", "drop_ns")
STATS = ("p50", "mean", "p95", "p99")
THRESHOLDS = {"p50": 5.0, "mean": 5.0, "p95": 10.0, "p99": 15.0, "rss": 5.0}
RSS_RE = re.compile(r"Maximum resident set size \(kbytes\):\s*(\d+)")
INDEX_FIELDS = ("repeat", "warmup", "ordinal")
WORK_FIELDS = ("budget_before_work", "budget_live_work", "budget_after_work")
SOURCE_COUNTER_FIELDS = (
    "source_read_calls",
    "source_requested_bytes",
    "source_returned_bytes",
    "source_zero_length_calls",
)


class TailGuardError(RuntimeError):
    """Fail-closed tail-guard error."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise TailGuardError(message)


def file_sha256(path: Path) -> str:
    require(path.is_file(), f"missing file: {path}")
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def mapping_sha256(mapping: dict[str, str]) -> str:
    payload = (json.dumps(mapping, indent=2) + "\n").encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def read_json(path: Path) -> dict[str, Any]:
    require(path.is_file(), f"missing JSON: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise TailGuardError(f"cannot read JSON {path}: {error}") from error
    require(isinstance(value, dict), f"JSON object required: {path}")
    return value


def write_json(path: Path, value: dict[str, Any], exclusive: bool = False) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    mode = "x" if exclusive else "w"
    with path.open(mode, encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write("\n")


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def build_binding(variant: str, require_binary: bool = True) -> dict[str, Any]:
    build_dir = HERE / variant
    receipt_path = build_dir / "build-receipt.json"
    manifest_path = build_dir / "source-manifest.json"
    receipt = read_json(receipt_path)
    require(receipt.get("exit_code") == 0, f"{variant}: retained build did not pass")
    binary_text = receipt.get("binary")
    require(isinstance(binary_text, str) and binary_text, f"{variant}: binary path missing")
    binary = Path(binary_text)
    expected_binary_sha = receipt.get("binary_sha256")
    require(isinstance(expected_binary_sha, str), f"{variant}: binary digest missing")
    if binary.exists():
        actual_binary_sha = file_sha256(binary)
        require(actual_binary_sha == expected_binary_sha,
                f"{variant}: retained binary hash differs from build receipt")
    else:
        require(not require_binary,
                f"{variant}: retained binary is required for capture: {binary}")
    expected_manifest_sha = receipt.get("source_manifest_sha256")
    require(isinstance(expected_manifest_sha, str), f"{variant}: source manifest digest missing")
    require(file_sha256(manifest_path) == expected_manifest_sha,
            f"{variant}: retained source manifest hash differs from build receipt")
    command = receipt.get("command")
    require(isinstance(command, list) and command, f"{variant}: build command missing")
    return {
        "variant": variant,
        "build_receipt_path": str(receipt_path),
        "build_receipt_sha256": file_sha256(receipt_path),
        "build_manifest_path": str(manifest_path),
        "source_manifest_sha256": expected_manifest_sha,
        "binary_path": str(binary),
        "binary_sha256": expected_binary_sha,
    }


def current_source_sha() -> str:
    return mapping_sha256(sources())


def count_original_flags(report: dict[str, Any]) -> int:
    total = 0
    comparisons = report.get("comparisons")
    require(isinstance(comparisons, dict), "candidate comparison has no comparisons")
    for records in comparisons.values():
        require(isinstance(records, list), "candidate comparison records are malformed")
        for record in records:
            require(isinstance(record, dict), "candidate comparison record is malformed")
            flags = record.get("adverse_flags", [])
            require(isinstance(flags, list), "candidate comparison flags are malformed")
            total += len(flags)
    return total


def make_plan() -> dict[str, Any]:
    baseline = build_binding("baseline")
    candidate = build_binding("candidate")
    current_candidate_sha = current_source_sha()
    require(current_candidate_sha == candidate["source_manifest_sha256"],
            "current source is not the retained candidate build source")
    comparison = read_json(ORIGINAL_COMPARISON)
    original_count = count_original_flags(comparison)
    return {
        "schema": PLAN_SCHEMA,
        "frozen_utc": utc_now(),
        "protocol": {
            "cases": CASES,
            "order": [
                {"slot": slot, "variant": variant, "case_id": case_id,
                 "case": CASES[case_id]}
                for slot, variant, case_id in ORDER
            ],
            "samples": SAMPLES,
            "warmups": WARMUPS,
            "repeats": REPEATS,
            "cpu": CPU,
            "fresh_process_per_case": True,
            "phase_metrics": list(PHASES),
            "rss": "whole-child Maximum resident set size from /usr/bin/time -v stderr",
            "statistics": "nearest-rank p50/p95/p99 and arithmetic mean",
            "thresholds_percent": THRESHOLDS,
            "bootstrap": {
                "iterations": BOOTSTRAP_ITERATIONS,
                "seed": DEFAULT_SEED,
                "unit": "unpaired measured rows within this single internal repeat; candidate median / baseline median",
                "interpretation": "descriptive only",
            },
            "correction": "The uncaptured four-child draft conflated case and replication. Corrected before capture to eight children: each A1/B1/B2/A2 slot runs both cases, yielding two independent A/B pairs per case. No raw tail-guard artifacts existed before this correction.",
        },
        "base_revision": "45d71cb6f0cd5d003c544b61c748552a513b0245",
        "scope": "Native DOCX tail guard for r2 p128-k1-file-batch and p128-k1-owned-batch; profile and hardware lanes excluded",
        "builds": {"baseline": baseline, "candidate": candidate},
        "candidate_source_manifest_current_sha256": current_candidate_sha,
        "original_comparison": {
            "path": str(ORIGINAL_COMPARISON),
            "sha256": file_sha256(ORIGINAL_COMPARISON),
            "adverse_flag_count": original_count,
            "retention": "all original flags remain retained regardless of tail-guard outcome",
        },
    }


def load_plan(require_binaries: bool = False) -> dict[str, Any]:
    plan = read_json(PLAN_PATH)
    require(plan.get("schema") == PLAN_SCHEMA, "tail-guard plan schema differs")
    require(plan.get("protocol", {}).get("cases") == CASES, "tail-guard case plan differs")
    require(plan.get("protocol", {}).get("order") == [
        {"slot": slot, "variant": variant, "case_id": case_id, "case": CASES[case_id]}
        for slot, variant, case_id in ORDER
    ], "tail-guard order differs")
    require(plan["protocol"]["samples"] == SAMPLES
            and plan["protocol"]["warmups"] == WARMUPS
            and plan["protocol"]["repeats"] == REPEATS
            and plan["protocol"]["cpu"] == CPU,
            "tail-guard sample protocol differs")
    for variant in ("baseline", "candidate"):
        retained = build_binding(variant, require_binary=require_binaries)
        expected = plan.get("builds", {}).get(variant)
        require(isinstance(expected, dict), f"plan has no {variant} binding")
        for field in ("build_receipt_sha256", "source_manifest_sha256", "binary_sha256", "binary_path"):
            require(expected.get(field) == retained.get(field),
                    f"{variant} retained build binding differs from frozen plan: {field}")
    require(file_sha256(ORIGINAL_COMPARISON) == plan["original_comparison"]["sha256"],
            "original candidate comparison changed after plan freeze")
    return plan


def artifact_path(slot: str, case: str, suffix: str) -> Path:
    return OUTPUT_DIR / f"{slot}-{case}.{suffix}"


def capture_one(plan: dict[str, Any], slot: str, variant: str, case_id: str) -> dict[str, Any]:
    case = CASES[case_id]
    paragraphs, replacements, source, mode = capture_module.prior.parse_name(case)
    binding = build_binding(variant, require_binary=True)
    plan_binding = plan["builds"][variant]
    require(binding["binary_sha256"] == plan_binding["binary_sha256"],
            f"{slot}: binary changed after plan freeze")
    before_source_sha = current_source_sha()
    candidate_source_sha = plan["builds"]["candidate"]["source_manifest_sha256"]
    if variant == "candidate":
        require(before_source_sha == candidate_source_sha,
                f"{slot}: current candidate source differs before capture")
    binary_before_sha = file_sha256(Path(binding["binary_path"]))

    csv_path = artifact_path(slot, case, "csv")
    stdout_path = artifact_path(slot, case, "stdout")
    stderr_path = artifact_path(slot, case, "stderr")
    receipt_path = artifact_path(slot, case, "json")
    for path in (csv_path, stdout_path, stderr_path, receipt_path):
        require(not path.exists(), f"tail-guard artifact already exists: {path}")
    corpus = SCRATCH / "tail-guard-corpora" / f"{slot}-{case}"
    require(not corpus.exists(), f"tail-guard corpus path already exists: {corpus}")
    corpus.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    command = [
        "/usr/bin/time", "-v", "taskset", "-c", str(CPU), binding["binary_path"],
        "--paragraphs", str(paragraphs), "--replacements", str(replacements),
        "--source", source, "--mode", mode,
        "--samples", str(SAMPLES), "--warmups", str(WARMUPS), "--repeats", str(REPEATS),
        "--artifact-dir", str(corpus), "--output", str(csv_path),
    ]
    started = utc_now()
    tick = time.monotonic()
    with stdout_path.open("x", encoding="utf-8") as stdout, stderr_path.open("x", encoding="utf-8") as stderr:
        result = subprocess.run(command, cwd=REPO, stdout=stdout, stderr=stderr)
    elapsed_seconds = time.monotonic() - tick
    after_source_sha = current_source_sha()
    binary_after_sha = file_sha256(Path(binding["binary_path"]))
    artifacts = {
        path.name: file_sha256(path)
        for path in (csv_path, stdout_path, stderr_path)
        if path.exists()
    }
    receipt = {
        "schema": CAPTURE_SCHEMA,
        "slot": slot,
        "case_id": case_id,
        "case": case,
        "variant": variant,
        "command": command,
        "started_utc": started,
        "elapsed_seconds": elapsed_seconds,
        "exit_code": result.returncode,
        "cpu": CPU,
        "fresh_process_per_case": True,
        "samples": SAMPLES,
        "warmups": WARMUPS,
        "repeats": REPEATS,
        "source_manifest_before_sha256": before_source_sha,
        "source_manifest_after_sha256": after_source_sha,
        "source_unchanged": before_source_sha == after_source_sha,
        "candidate_source_manifest_sha256": candidate_source_sha,
        "candidate_source_before_matches_build": before_source_sha == candidate_source_sha,
        "candidate_source_after_matches_build": after_source_sha == candidate_source_sha,
        "working_tree_baseline_binding": "not_required_for_retained_baseline_binary" if variant == "baseline" else "candidate_manifest_required_before_and_after",
        "build_receipt_path": binding["build_receipt_path"],
        "build_receipt_sha256": binding["build_receipt_sha256"],
        "build_manifest_path": binding["build_manifest_path"],
        "build_source_manifest_sha256": binding["source_manifest_sha256"],
        "binary_path": binding["binary_path"],
        "binary_sha256_before": binary_before_sha,
        "binary_sha256_after": binary_after_sha,
        "binary_sha256": binding["binary_sha256"],
        "plan_sha256": file_sha256(PLAN_PATH),
        "cleanup_verified": not corpus.exists(),
        "scope": plan["scope"],
        "artifacts": artifacts,
    }
    write_json(receipt_path, receipt, exclusive=True)
    require(result.returncode == 0, f"{slot}: benchmark exited {result.returncode}")
    require(receipt["source_unchanged"], f"{slot}: source changed during capture")
    require(binary_before_sha == binding["binary_sha256"] == binary_after_sha,
            f"{slot}: retained binary changed during capture")
    require(receipt["cleanup_verified"], f"{slot}: benchmark left corpus behind")
    rows = read_rows(csv_path)
    validate_rows(case, rows)
    return receipt


def read_rows(path: Path) -> list[dict[str, str]]:
    require(path.is_file(), f"missing tail-guard CSV: {path}")
    with path.open(newline="", encoding="utf-8") as stream:
        return list(csv.DictReader(stream))


def validate_rows(case: str, rows: list[dict[str, str]]) -> None:
    try:
        capture_module.validate(case, rows, SAMPLES, WARMUPS, REPEATS)
    except (AssertionError, KeyError, ValueError) as error:
        raise TailGuardError(f"{case}: capture.validate failed: {error}") from error
    require(len(rows) == SAMPLES + WARMUPS,
            f"{case}: expected {SAMPLES + WARMUPS} rows, found {len(rows)}")
    for row in rows:
        require(row.get("budget_managed") == "true", f"{case}: unmanaged budget row")
        for field in capture_module.prior.BOOLEAN_KEYS:
            require(row.get(field) == "true", f"{case}: {field} is false")


def rss_from_stderr(path: Path) -> int:
    matches = RSS_RE.findall(path.read_text(encoding="utf-8"))
    require(len(matches) == 1, f"{path.name}: expected one whole-child RSS value")
    return int(matches[0])


def nearest_rank(values: list[float], quantile: float) -> float:
    ordered = sorted(values)
    index = max(0, math.ceil(quantile * len(ordered)) - 1)
    return ordered[index]


def phase_stats(rows: list[dict[str, str]]) -> dict[str, dict[str, float]]:
    measured = [row for row in rows if row.get("warmup") == "false"]
    require(len(measured) == SAMPLES, f"expected {SAMPLES} measured rows")
    result: dict[str, dict[str, float]] = {}
    for field in PHASES:
        values = [float(row[field]) for row in measured]
        result[field] = {
            "p50": nearest_rank(values, 0.50),
            "mean": sum(values) / len(values),
            "p95": nearest_rank(values, 0.95),
            "p99": nearest_rank(values, 0.99),
        }
    return result


def row_key(row: dict[str, str]) -> tuple[int, str, int]:
    return int(row["repeat"]), row["warmup"], int(row["ordinal"])


def row_summary(rows: list[dict[str, str]]) -> dict[str, Any]:
    identity_keys = capture_module.prior.IDENTITY_KEYS
    identities = {tuple(row[key] for key in identity_keys) for row in rows}
    require(len(identities) == 1, "tail-guard output identity changes within a run")
    return {
        "total": len(rows),
        "warmups": sum(row["warmup"] == "true" for row in rows),
        "measured": sum(row["warmup"] == "false" for row in rows),
        "identity": {key: next(iter(identities))[index] for index, key in enumerate(identity_keys)},
        "guards_all_true": all(
            row.get("budget_managed") == "true"
            and all(row.get(key) == "true" for key in capture_module.prior.BOOLEAN_KEYS)
            for row in rows
        ),
    }


def counter_values(rows: list[dict[str, str]], fields: tuple[str, ...]) -> dict[str, list[int]]:
    return {field: [int(row[field]) for row in rows] for field in fields}


def parity_for_fields(
    baseline: list[dict[str, str]], candidate: list[dict[str, str]], fields: tuple[str, ...]
) -> dict[str, Any]:
    left = {row_key(row): row for row in baseline}
    right = {row_key(row): row for row in candidate}
    require(set(left) == set(right), "baseline/candidate row identities differ")
    differences: list[dict[str, Any]] = []
    for key in sorted(left):
        for field in fields:
            if left[key].get(field) != right[key].get(field):
                differences.append({
                    "row": list(key), "field": field,
                    "baseline": left[key].get(field), "candidate": right[key].get(field),
                })
    return {"fields": list(fields), "equal": not differences, "differences": differences}


def parity_summary(baseline: list[dict[str, str]], candidate: list[dict[str, str]]) -> dict[str, Any]:
    output = parity_for_fields(baseline, candidate, capture_module.prior.IDENTITY_KEYS)
    oracle = parity_for_fields(baseline, candidate, capture_module.prior.BOOLEAN_KEYS)
    work = parity_for_fields(baseline, candidate, WORK_FIELDS)
    source = parity_for_fields(baseline, candidate, SOURCE_COUNTER_FIELDS)
    all_true = all(
        row.get("budget_managed") == "true"
        and all(row.get(field) == "true" for field in capture_module.prior.BOOLEAN_KEYS)
        for row in baseline + candidate
    )
    return {
        "row_count_each": len(baseline),
        "output_identity_parity": output,
        "oracle_parity": oracle,
        "work_parity": work,
        "source_counter_parity": source,
        "all_oracle_and_release_guards_true": all_true,
    }


def bootstrap_p50_ratio(
    baseline: list[float], candidate: list[float], seed: int
) -> dict[str, Any]:
    rng = random.Random(seed)
    ratios: list[float] = []
    for _ in range(BOOTSTRAP_ITERATIONS):
        left = [baseline[rng.randrange(len(baseline))] for _ in baseline]
        right = [candidate[rng.randrange(len(candidate))] for _ in candidate]
        base_median = nearest_rank(left, 0.50)
        candidate_median = nearest_rank(right, 0.50)
        ratios.append(candidate_median / base_median if base_median else math.inf)
    return {
        "low": nearest_rank(ratios, 0.025),
        "high": nearest_rank(ratios, 0.975),
        "iterations": BOOTSTRAP_ITERATIONS,
        "seed": seed,
        "unit": "unpaired measured rows within this single internal repeat; candidate median / baseline median",
        "interpretation": "descriptive only",
    }


def comparison_for_metric(
    metric: str,
    stat: str,
    baseline_value: float,
    candidate_value: float,
    seed: int | None = None,
) -> dict[str, Any]:
    delta = ((candidate_value / baseline_value) - 1.0) * 100.0 if baseline_value else None
    value: dict[str, Any] = {
        "metric": metric,
        "stat": stat,
        "baseline": baseline_value,
        "candidate": candidate_value,
        "ratio_candidate_over_baseline": candidate_value / baseline_value if baseline_value else None,
        "delta_pct": delta,
        "threshold_pct": THRESHOLDS[stat],
    }
    if metric != "rss_kib" and stat == "p50" and seed is not None:
        value["bootstrap95_ci_ratio_of_medians"] = None
    return value


def original_flags(plan: dict[str, Any]) -> dict[str, Any]:
    report = read_json(ORIGINAL_COMPARISON)
    require(file_sha256(ORIGINAL_COMPARISON) == plan["original_comparison"]["sha256"],
            "original comparison changed after plan freeze")
    all_flags: list[dict[str, Any]] = []
    comparisons = report["comparisons"]
    for pair, records in comparisons.items():
        for record in records:
            for flag in record.get("adverse_flags", []):
                values = record["metrics"][flag["metric"]][flag["stat"]]
                complete_flag = dict(flag, baseline=values["baseline"], candidate=values["candidate"])
                all_flags.append({"pair": pair, "case": record["case"], "flag": complete_flag})
    require(len(all_flags) == plan["original_comparison"]["adverse_flag_count"],
            "original comparison flag count changed after plan freeze")
    selected = [item for item in all_flags if item["case"] in CASES.values()]
    return {
        "report_path": str(ORIGINAL_COMPARISON),
        "report_sha256": file_sha256(ORIGINAL_COMPARISON),
        "total_adverse_flags": len(all_flags),
        "all_flags": all_flags,
        "selected_case_flags": selected,
        "retention": "all original flags are retained; tail-guard results do not dismiss them",
    }


def load_capture(plan: dict[str, Any], slot: str, variant: str, case_id: str) -> dict[str, Any]:
    case = CASES[case_id]
    binding = build_binding(variant, require_binary=False)
    receipt_path = artifact_path(slot, case, "json")
    receipt = read_json(receipt_path)
    require(receipt.get("schema") == CAPTURE_SCHEMA, f"{slot}: capture schema differs")
    for field, expected in (("slot", slot), ("variant", variant), ("case_id", case_id), ("case", case)):
        require(receipt.get(field) == expected, f"{slot}: receipt {field} differs")
    require(receipt.get("plan_sha256") == file_sha256(PLAN_PATH), f"{slot}: plan binding differs")
    require(receipt.get("build_receipt_sha256") == binding["build_receipt_sha256"],
            f"{slot}: build receipt binding differs")
    require(receipt.get("build_source_manifest_sha256") == binding["source_manifest_sha256"],
            f"{slot}: build source manifest binding differs")
    require(receipt.get("binary_sha256") == binding["binary_sha256"], f"{slot}: binary receipt binding differs")
    require(receipt.get("binary_sha256_before") == binding["binary_sha256"]
            and receipt.get("binary_sha256_after") == binding["binary_sha256"],
            f"{slot}: binary changed during capture")
    require(receipt.get("source_unchanged") is True, f"{slot}: source was changed during capture")
    if variant == "candidate":
        candidate_sha = plan["builds"]["candidate"]["source_manifest_sha256"]
        require(receipt.get("source_manifest_before_sha256") == candidate_sha
                and receipt.get("source_manifest_after_sha256") == candidate_sha,
                f"{slot}: candidate source binding before/after differs")
        require(receipt.get("candidate_source_before_matches_build") is True
                and receipt.get("candidate_source_after_matches_build") is True,
                f"{slot}: candidate source binding flags are false")
    for path in (artifact_path(slot, case, "csv"), artifact_path(slot, case, "stdout"), artifact_path(slot, case, "stderr")):
        expected_sha = receipt.get("artifacts", {}).get(path.name)
        require(expected_sha == file_sha256(path), f"{slot}: artifact hash differs for {path.name}")
    command = receipt.get("command")
    require(isinstance(command, list), f"{slot}: command is malformed")
    command_text = [str(token) for token in command]
    require("valgrind" not in command_text and "perf" not in command_text,
            f"{slot}: non-native tool entered tail guard")
    for value in ("taskset", "-c", str(CPU), "--samples", str(SAMPLES), "--warmups", str(WARMUPS), "--repeats", str(REPEATS)):
        require(value in command_text, f"{slot}: command missing {value}")
    rows = read_rows(artifact_path(slot, case, "csv"))
    validate_rows(case, rows)
    rss = rss_from_stderr(artifact_path(slot, case, "stderr"))
    return {
        "slot": slot,
        "variant": variant,
        "case": case,
        "receipt": receipt,
        "rows": rows,
        "row_summary": row_summary(rows),
        "phase_stats": phase_stats(rows),
        "rss_kib": rss,
    }


def compare_case(
    case: str, baseline: dict[str, Any], candidate: dict[str, Any]
) -> dict[str, Any]:
    metrics: dict[str, dict[str, Any]] = {}
    flags: list[dict[str, Any]] = []
    for metric in PHASES:
        metrics[metric] = {}
        base_values = [float(row[metric]) for row in baseline["rows"] if row["warmup"] == "false"]
        candidate_values = [float(row[metric]) for row in candidate["rows"] if row["warmup"] == "false"]
        for stat in STATS:
            seed_material = f"{DEFAULT_SEED}:{case}:{metric}:{stat}".encode("utf-8")
            seed = int.from_bytes(hashlib.sha256(seed_material).digest()[:8], "big")
            value = comparison_for_metric(
                metric,
                stat,
                baseline["phase_stats"][metric][stat],
                candidate["phase_stats"][metric][stat],
                seed,
            )
            value["bootstrap95_ci_ratio_of_medians"] = bootstrap_p50_ratio(
                base_values, candidate_values, seed
            ) if stat == "p50" else None
            metrics[metric][stat] = value
            if value["delta_pct"] is not None and value["delta_pct"] > THRESHOLDS[stat]:
                flags.append({
                    "metric": metric,
                    "stat": stat,
                    "baseline": value["baseline"],
                    "candidate": value["candidate"],
                    "delta_pct": value["delta_pct"],
                    "threshold_pct": THRESHOLDS[stat],
                })
    rss_value = comparison_for_metric(
        "rss_kib", "rss", baseline["rss_kib"], candidate["rss_kib"]
    )
    rss_value["bootstrap95_ci_ratio_of_medians"] = None
    metrics["rss_kib"] = {"rss": rss_value}
    if rss_value["delta_pct"] is not None and rss_value["delta_pct"] > THRESHOLDS["rss"]:
        flags.append({
            "metric": "rss_kib", "stat": "rss",
            "baseline": rss_value["baseline"], "candidate": rss_value["candidate"],
            "delta_pct": rss_value["delta_pct"], "threshold_pct": THRESHOLDS["rss"],
        })
    return {
        "baseline_slot": baseline["slot"],
        "candidate_slot": candidate["slot"],
        "metrics": metrics,
        "adverse_flags": flags,
        "flag_counts": {
            "all": len(flags),
            "tail_p95_p99": sum(item["stat"] in ("p95", "p99") for item in flags),
            "short_phase": sum(item["metric"] in SHORT_PHASES for item in flags),
            "rss": sum(item["metric"] == "rss_kib" for item in flags),
        },
        "parity": parity_summary(baseline["rows"], candidate["rows"]),
    }


def analyze(plan: dict[str, Any]) -> dict[str, Any]:
    loaded: dict[tuple[str, str], dict[str, Any]] = {}
    for slot, variant, case_id in ORDER:
        loaded[(slot, case_id)] = load_capture(plan, slot, variant, case_id)
    cases: dict[str, Any] = {}
    pair_comparisons: list[dict[str, Any]] = []
    for case_id, case in CASES.items():
        pairs: dict[str, Any] = {}
        for baseline_slot, candidate_slot in PAIR_ORDER:
            baseline = loaded[(baseline_slot, case_id)]
            candidate = loaded[(candidate_slot, case_id)]
            comparison = compare_case(case, baseline, candidate)
            pair_id = f"{baseline_slot}-{candidate_slot}"
            pairs[pair_id] = {
                "baseline": {
                    "slot": baseline_slot,
                    "row_summary": baseline["row_summary"],
                    "phase_stats": baseline["phase_stats"],
                    "rss_kib": baseline["rss_kib"],
                },
                "candidate": {
                    "slot": candidate_slot,
                    "row_summary": candidate["row_summary"],
                    "phase_stats": candidate["phase_stats"],
                    "rss_kib": candidate["rss_kib"],
                },
                "comparison": comparison,
            }
            pair_comparisons.append({"case": case, "pair": pair_id, "comparison": comparison})
        cases[case] = {"pairs": pairs}
    original = original_flags(plan)
    report = {
        "schema": REPORT_SCHEMA,
        "plan_sha256": file_sha256(PLAN_PATH),
        "scope": plan["scope"],
        "protocol": plan["protocol"],
        "bindings": plan["builds"],
        "runs": [
            {
                "slot": loaded[(slot, case_id)]["slot"],
                "variant": loaded[(slot, case_id)]["variant"],
                "case": loaded[(slot, case_id)]["case"],
                "receipt": loaded[(slot, case_id)]["receipt"],
                "row_summary": loaded[(slot, case_id)]["row_summary"],
                "phase_stats": loaded[(slot, case_id)]["phase_stats"],
                "rss_kib": loaded[(slot, case_id)]["rss_kib"],
            }
            for slot, _, case_id in ORDER
        ],
        "cases": cases,
        "pair_comparisons": pair_comparisons,
        "original_comparison": original,
        "summary": {
            "cases": len(cases),
            "runs": len(loaded),
            "paired_comparisons": len(pair_comparisons),
            "new_adverse_flags": sum(len(item["comparison"]["adverse_flags"]) for item in pair_comparisons),
            "new_tail_p95_p99_flags": sum(item["comparison"]["flag_counts"]["tail_p95_p99"] for item in pair_comparisons),
            "new_short_phase_flags": sum(item["comparison"]["flag_counts"]["short_phase"] for item in pair_comparisons),
            "new_rss_flags": sum(item["comparison"]["flag_counts"]["rss"] for item in pair_comparisons),
            "original_adverse_flags_retained": original["total_adverse_flags"],
            "output_identity_parity": all(
                item["comparison"]["parity"]["output_identity_parity"]["equal"] for item in pair_comparisons
            ),
            "oracle_and_release_guards_true": all(
                item["comparison"]["parity"]["all_oracle_and_release_guards_true"] for item in pair_comparisons
            ),
            "work_parity": all(item["comparison"]["parity"]["work_parity"]["equal"] for item in pair_comparisons),
            "source_counter_parity": all(item["comparison"]["parity"]["source_counter_parity"]["equal"] for item in pair_comparisons),
        },
    }
    write_json(REPORT_JSON, report)
    REPORT_MD.write_text(render_markdown(report), encoding="utf-8")
    return report


def fmt(value: Any) -> str:
    if isinstance(value, float):
        return f"{value:.3f}"
    return str(value)


def render_markdown(report: dict[str, Any]) -> str:
    lines = [
        "# 0519 bounded native tail guard",
        "",
        "This diagnostic captures the two r2 elapsed p99 outliers with retained baseline and candidate binaries in eight fresh processes on CPU 2. The fixed order is A1/B1/B2/A2 over both cases: A1 baseline file+owned, B1 candidate file+owned, B2 candidate owned+file, A2 baseline owned+file. Each case therefore has two independent A/B pairs.",
        "",
        "The protocol uses 200 measured samples, 10 warmups, and one internal repeat per process. Phase values are nearest-rank p50/p95/p99 plus arithmetic mean in nanoseconds; RSS is whole-child GNU time maximum in KiB. p50 bootstrap intervals are unpaired and descriptive only.",
        "",
        "All original candidate-comparison flags are retained below. This bounded follow-up does not dismiss any original flag as noise or change the release decision by itself.",
        "",
        "## Summary",
        "",
        f"- New adverse flags: {report['summary']['new_adverse_flags']} (tail p95/p99: {report['summary']['new_tail_p95_p99_flags']}; short phases: {report['summary']['new_short_phase_flags']}; RSS: {report['summary']['new_rss_flags']}).",
        f"- Original adverse flags retained: {report['summary']['original_adverse_flags_retained']}.",
        f"- Output identity parity: `{report['summary']['output_identity_parity']}`; oracle/release guards true: `{report['summary']['oracle_and_release_guards_true']}`; Work parity: `{report['summary']['work_parity']}`; source-counter parity: `{report['summary']['source_counter_parity']}`.",
        "",
        "## Case comparison",
        "",
        "| Case | Pair | Elapsed p50 (baseline→candidate ns) | Publish p50 (baseline→candidate ns) | RSS (baseline→candidate KiB) | New flags |",
        "| --- | --- | ---: | ---: | ---: | ---: |",
    ]
    for case, case_value in report["cases"].items():
        for pair, value in case_value["pairs"].items():
            base = value["comparison"]["metrics"]
            lines.append(
                f"| {case} | {pair} | {fmt(base['elapsed_ns']['p50']['baseline'])}→{fmt(base['elapsed_ns']['p50']['candidate'])} ({fmt(base['elapsed_ns']['p50']['delta_pct'])}%) | {fmt(base['publish_ns']['p50']['baseline'])}→{fmt(base['publish_ns']['p50']['candidate'])} ({fmt(base['publish_ns']['p50']['delta_pct'])}%) | {fmt(base['rss_kib']['rss']['baseline'])}→{fmt(base['rss_kib']['rss']['candidate'])} ({fmt(base['rss_kib']['rss']['delta_pct'])}%) | {len(value['comparison']['adverse_flags'])} |"
            )
    lines += ["", "## All phase, tail, and RSS statistics", ""]
    for case, case_value in report["cases"].items():
        for pair, value in case_value["pairs"].items():
            lines += [f"### {case} — {pair}", "", "| Metric | Stat | Baseline | Candidate | Ratio | Δ% | Flag |", "| --- | --- | ---: | ---: | ---: | ---: | --- |"]
            flags = {(item["metric"], item["stat"]) for item in value["comparison"]["adverse_flags"]}
            for metric in (*PHASES, "rss_kib"):
                for stat in STATS if metric != "rss_kib" else ("rss",):
                    item = value["comparison"]["metrics"][metric][stat]
                    lines.append(
                        f"| {metric} | {stat} | {fmt(item['baseline'])} | {fmt(item['candidate'])} | {fmt(item['ratio_candidate_over_baseline'])} | {fmt(item['delta_pct'])} | {'ADVERSE' if (metric, stat) in flags else ''} |"
                    )
            lines.append("")
            lines.append("Bootstrap intervals are attached to each p50 metric in JSON; all phase values above remain the authoritative point statistics.")
            lines.append("")
            lines.append("Flags with absolute values:")
            if value["comparison"]["adverse_flags"]:
                for item in value["comparison"]["adverse_flags"]:
                    lines.append(
                        f"- `{item['metric']}.{item['stat']}` {item['delta_pct']:+.3f}% ({fmt(item['baseline'])}→{fmt(item['candidate'])}; threshold {item['threshold_pct']}%)."
                    )
            else:
                lines.append("- none")
            lines.append("")
            parity = value["comparison"]["parity"]
            lines += [
                "Parity:",
                f"- Output identity: `{parity['output_identity_parity']['equal']}`; oracle fields: `{parity['oracle_parity']['equal']}`; Work fields: `{parity['work_parity']['equal']}`; source counters: `{parity['source_counter_parity']['equal']}`.",
                "",
            ]
    lines += [
        "## Capture bindings",
        "",
        "| Slot | Variant | Case | Binary SHA256 | Source before | Source after | RSS KiB |",
        "| --- | --- | --- | --- | --- | --- | ---: |",
    ]
    for run in report["runs"]:
        receipt = run["receipt"]
        lines.append(
            f"| {run['slot']} | {run['variant']} | {run['case']} | `{receipt['binary_sha256']}` | `{receipt['source_manifest_before_sha256']}` | `{receipt['source_manifest_after_sha256']}` | {run['rss_kib']} |"
        )
    lines += [
        "",
        "Baseline rows are bound to the retained baseline binary/build manifest without requiring the current working tree to equal the old baseline source. Candidate rows bind both current source snapshots to the retained candidate manifest.",
        "",
        "## Original comparison flags retained",
        "",
        "| Pair | Case | Metric | Stat | Δ% | Baseline | Candidate | Threshold |",
        "| --- | --- | --- | --- | ---: | ---: | ---: | ---: |",
    ]
    for item in report["original_comparison"]["all_flags"]:
        flag = item["flag"]
        lines.append(
            f"| {item['pair']} | {item['case']} | {flag.get('metric', '')} | {flag.get('stat', '')} | {fmt(flag.get('delta_pct'))} | {fmt(flag.get('baseline', ''))} | {fmt(flag.get('candidate', ''))} | {fmt(flag.get('threshold_pct', ''))} |"
        )
    lines += ["", "Full rows, receipts, source/binary bindings, guards, parity differences, and bootstrap metadata are retained in `tail-guard.json`.", ""]
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    modes = parser.add_mutually_exclusive_group(required=True)
    modes.add_argument("--write-plan", action="store_true", help="freeze bindings and protocol before capture")
    modes.add_argument("--capture", action="store_true", help="run the eight fresh-process tail-guard captures and analyze them")
    modes.add_argument("--analyze", action="store_true", help="analyze already retained tail-guard captures")
    args = parser.parse_args()
    try:
        if args.write_plan:
            require(not PLAN_PATH.exists(), f"frozen plan already exists: {PLAN_PATH}")
            plan = make_plan()
            write_json(PLAN_PATH, plan, exclusive=True)
            print(json.dumps({"plan": str(PLAN_PATH), "sha256": file_sha256(PLAN_PATH), "adverse_flags_retained": plan["original_comparison"]["adverse_flag_count"]}, sort_keys=True))
        elif args.capture:
            plan = load_plan(require_binaries=True)
            for slot, variant, case_id in ORDER:
                receipt = capture_one(plan, slot, variant, case_id)
                print(slot, CASES[case_id], "rows passed", flush=True)
            report = analyze(plan)
            print(json.dumps({"json": str(REPORT_JSON), "markdown": str(REPORT_MD), "new_adverse_flags": report["summary"]["new_adverse_flags"]}, sort_keys=True))
        else:
            plan = load_plan()
            report = analyze(plan)
            print(json.dumps({"json": str(REPORT_JSON), "markdown": str(REPORT_MD), "new_adverse_flags": report["summary"]["new_adverse_flags"]}, sort_keys=True))
        return 0
    except (TailGuardError, OSError, subprocess.SubprocessError, KeyError, ValueError) as error:
        print(f"tail guard failed: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
