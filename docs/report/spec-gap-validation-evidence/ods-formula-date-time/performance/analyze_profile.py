#!/usr/bin/env python3
"""Validate and summarize a date/time performance capture.

The capture runner stores one JSON object per timed child in each
``measurements.jsonl`` file.  This analyzer treats those rows and their raw
stdout/time receipts as the evidence source: it checks the complete sample
set, output checksums, accounting invariants, and matched baseline/candidate
groups before producing descriptive deltas.  A five percent change is always
reported; it is not classified as noise or assigned a cause here.
"""

from __future__ import annotations

import argparse
from collections import defaultdict
import hashlib
import json
import math
from pathlib import Path
import random
import re
import statistics
from typing import Any


HERE = Path(__file__).resolve().parent
DEFAULT_RESULTS = HERE / "results"
MATRIX_PATH = HERE / "case-matrix.json"
PHASES = ("evaluate", "parse-evaluate")
THRESHOLD = 0.05
TIME_RE = re.compile(r"Maximum resident set size \(kbytes\):\s*(\d+)")
SHAPE_RE = re.compile(r"^(?P<rows>[1-9][0-9]*)x(?P<columns>[1-9][0-9]*)$")
BOOTSTRAP_SEED = 20260920
BOOTSTRAP_REPLICATES = 10_000
BOOTSTRAP_CONFIDENCE = 0.95

SAMPLE_COUNTERS = (
    "elapsed_ns",
    "alloc_calls",
    "dealloc_calls",
    "requested_bytes",
    "released_bytes",
    "live_before",
    "live_after",
    "peak_live_delta",
    "work",
    "memory_retained",
    "reference_reads",
    "output_bytes",
    "checksum",
)

ROW_MEDIANS = (
    "output_bytes_p50",
    "bytes_per_repeat_p50",
    "elapsed_ns_p50",
    "elapsed_ns_mean",
    "elapsed_ns_p95",
    "elapsed_ns_p99",
    "allocator_calls_p50",
    "requested_bytes_p50",
    "released_bytes_p50",
    "peak_live_delta_p50",
    "memory_retained_p50",
    "work_p50",
    "reference_reads_p50",
    "checksum_p50",
    "live_before_p50",
    "live_after_p50",
)


class CaptureError(RuntimeError):
    """A fail-closed capture evidence error."""


def load(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as error:
        raise CaptureError(f"missing JSON file: {path}") from error
    except json.JSONDecodeError as error:
        raise CaptureError(f"invalid JSON in {path}: {error}") from error


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(chunk)
    except OSError as error:
        raise CaptureError(f"cannot hash {path}: {error}") from error
    return digest.hexdigest()


def matrix_declarations() -> dict[str, Any]:
    value = load(MATRIX_PATH)
    if not isinstance(value, dict):
        raise CaptureError("case-matrix.json is not an object")
    controls = value.get("matched_controls")
    date_cases = value.get("candidate_cases")
    phases = value.get("phases")
    counts = value.get("counts")
    if not isinstance(controls, list) or not isinstance(date_cases, list):
        raise CaptureError("case-matrix.json has malformed case lists")
    if not isinstance(phases, list) or phases != list(PHASES):
        raise CaptureError(f"case-matrix.json phases differ from {list(PHASES)!r}")
    if not isinstance(counts, dict):
        raise CaptureError("case-matrix.json has no counts object")
    all_rows = controls + date_cases
    if any(not isinstance(row, dict) for row in all_rows):
        raise CaptureError("case-matrix.json contains a non-object case row")
    control_names = [row.get("name") for row in controls]
    date_names = [row.get("name") for row in date_cases]
    if any(not isinstance(name, str) for name in control_names + date_names):
        raise CaptureError("case-matrix.json contains a case without a string name")
    if len(set(control_names + date_names)) != len(control_names) + len(date_names):
        raise CaptureError("case-matrix.json contains duplicate case names")
    expected_counts = {
        "matched_controls": len(control_names),
        "candidate_cases": len(date_names),
        "baseline_timed_groups": len(control_names) * len(PHASES),
        "candidate_timed_groups": (len(control_names) + len(date_names)) * len(PHASES),
    }
    for key, expected in expected_counts.items():
        equal(f"case matrix count {key}", counts.get(key), expected)
    warmups = nonnegative_int(value.get("warmups"), "case matrix warmups")
    samples = nonnegative_int(value.get("samples"), "case matrix samples")
    if warmups <= 0 or samples <= 0:
        raise CaptureError("case-matrix.json warmups and samples must be positive")
    case_rows: dict[str, dict[str, Any]] = {}
    for row in all_rows:
        name = row["name"]
        for field in ("source", "evaluation_path", "expected", "shape"):
            if field not in row:
                raise CaptureError(f"case matrix row {name}: missing {field}")
        reference_reads = nonnegative_int(
            row.get("reference_reads"), f"case matrix row {name} reference_reads"
        )
        case_rows[name] = {
            "row": row,
            "reference_reads": reference_reads,
            "geometry": shape_geometry(row["shape"], f"case matrix row {name} shape"),
        }
    return {
        "contract_sha256": value.get("contract_sha256"),
        "oracle_vectors_sha256": value.get("oracle_vectors_sha256"),
        "baseline_commit": value.get("baseline_commit"),
        "controls": control_names,
        "date_cases": date_names,
        "phases": list(PHASES),
        "warmups": warmups,
        "samples": samples,
        "rows": case_rows,
    }


def equal(label: str, actual: Any, expected: Any) -> None:
    if actual != expected:
        raise CaptureError(f"{label}: expected {expected!r}, observed {actual!r}")


def finite_number(value: Any, label: str) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError) as error:
        raise CaptureError(f"{label}: expected a number, observed {value!r}") from error
    if not math.isfinite(number):
        raise CaptureError(f"{label}: expected a finite number, observed {value!r}")
    return number


def nonnegative_int(value: Any, label: str) -> int:
    if isinstance(value, bool):
        raise CaptureError(f"{label}: expected a nonnegative integer")
    try:
        number = int(value)
    except (TypeError, ValueError) as error:
        raise CaptureError(f"{label}: expected a nonnegative integer") from error
    if number < 0 or str(value) not in {str(number), f"{number}.0"}:
        raise CaptureError(f"{label}: expected a nonnegative integer, observed {value!r}")
    return number


def shape_geometry(value: Any, label: str) -> dict[str, Any]:
    """Parse the matrix's execution geometry, independently of its payload."""
    if value == "scalar":
        return {"kind": "scalar", "rows": 0, "columns": 0, "elements": 1}
    if not isinstance(value, str):
        raise CaptureError(f"{label}: expected scalar or NxM shape, observed {value!r}")
    match = SHAPE_RE.fullmatch(value)
    if match is None:
        raise CaptureError(f"{label}: malformed execution shape {value!r}")
    rows = int(match.group("rows"))
    columns = int(match.group("columns"))
    return {"kind": "array", "rows": rows, "columns": columns, "elements": rows * columns}


def same_expected_outcome(observed: Any, expected: Any) -> bool:
    """Compare typed preflight outcomes with the harness's numeric tolerance."""
    if isinstance(expected, bool):
        return isinstance(observed, bool) and observed == expected
    if isinstance(expected, (int, float)):
        return (
            isinstance(observed, (int, float))
            and not isinstance(observed, bool)
            and math.isfinite(float(observed))
            and math.isfinite(float(expected))
            and abs(float(observed) - float(expected)) <= max(abs(float(expected)), 1.0) * 1.0e-10
        )
    if isinstance(expected, list):
        return (
            isinstance(observed, list)
            and len(observed) == len(expected)
            and all(same_expected_outcome(actual, wanted) for actual, wanted in zip(observed, expected))
        )
    if isinstance(expected, dict):
        return (
            isinstance(observed, dict)
            and observed.keys() == expected.keys()
            and all(same_expected_outcome(observed[key], value) for key, value in expected.items())
        )
    return type(observed) is type(expected) and observed == expected


def rows(directory: Path) -> list[dict[str, Any]]:
    path = directory / "measurements.jsonl"
    if not path.is_file():
        raise CaptureError(f"missing measurements: {path}")
    result: list[dict[str, Any]] = []
    for line_number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
        if not line.strip():
            continue
        try:
            value = json.loads(line)
        except json.JSONDecodeError as error:
            raise CaptureError(f"{path}:{line_number}: invalid JSON: {error}") from error
        if not isinstance(value, dict):
            raise CaptureError(f"{path}:{line_number}: measurement is not an object")
        result.append(value)
    return result


def raw_child(directory: Path, row: dict[str, Any], label: str) -> tuple[dict[str, Any], int]:
    raw_stdout = row.get("raw_stdout")
    raw_time = row.get("raw_time")
    if not isinstance(raw_stdout, str) or not raw_stdout:
        raise CaptureError(f"{label}: raw_stdout is missing")
    if not isinstance(raw_time, str) or not raw_time:
        raise CaptureError(f"{label}: raw_time is missing")
    stdout_path = directory / raw_stdout
    time_path = directory / raw_time
    if not stdout_path.is_file():
        raise CaptureError(f"{label}: missing raw stdout {stdout_path}")
    if not time_path.is_file():
        raise CaptureError(f"{label}: missing raw time receipt {time_path}")
    lines = [line for line in stdout_path.read_text(encoding="utf-8").splitlines() if line.strip()]
    if len(lines) != 1:
        raise CaptureError(f"{label}: expected one child JSON line, observed {len(lines)}")
    try:
        child = json.loads(lines[0])
    except json.JSONDecodeError as error:
        raise CaptureError(f"{label}: invalid child JSON: {error}") from error
    if not isinstance(child, dict):
        raise CaptureError(f"{label}: child JSON is not an object")
    match = TIME_RE.search(time_path.read_text(encoding="utf-8"))
    if match is None:
        raise CaptureError(f"{label}: missing maximum RSS in {time_path}")
    return child, int(match.group(1))


def validate_sample(sample: dict[str, Any], label: str) -> None:
    for field in SAMPLE_COUNTERS:
        nonnegative_int(sample.get(field), f"{label} sample {field}")
    if sample["live_before"] != sample["live_after"]:
        raise CaptureError(f"{label}: allocator live bytes do not balance")
    if sample["released_bytes"] > sample["requested_bytes"]:
        raise CaptureError(f"{label}: released bytes exceed requested bytes")


def validate_row(
    row: dict[str, Any],
    directory: Path,
    label: str,
    warmups: int,
    contract: dict[str, Any] | None = None,
) -> dict[str, Any]:
    required = {
        "case", "phase", "sample_index", "supported", "warmups", "iterations", "repeat",
        "input_bytes", "output_bytes_p50", "bytes_per_repeat_p50", "elapsed_ns_p50",
        "elapsed_ns_mean", "elapsed_ns_p95", "elapsed_ns_p99", "elapsed_ns_per_repeat",
        "allocator_calls_p50", "requested_bytes_p50", "released_bytes_p50",
        "peak_live_delta_p50", "memory_retained_p50", "work_p50", "work_per_repeat",
        "reference_reads_p50", "reference_reads_per_repeat", "checksum_p50", "live_before_p50",
        "live_after_p50", "samples", "rss_kib", "binary_sha256", "source_git_head",
        "raw_stdout", "raw_time",
    }
    missing = sorted(required - row.keys())
    if missing:
        raise CaptureError(f"{label}: missing row fields {missing}")
    if row["supported"] is not True:
        raise CaptureError(f"{label}: unsupported timed row")
    equal(f"{label} warmups", row["warmups"], warmups)
    equal(f"{label} iterations", row["iterations"], 1)
    repeat = nonnegative_int(row["repeat"], f"{label} repeat")
    if repeat == 0:
        raise CaptureError(f"{label}: repeat must be positive")
    for field in ROW_MEDIANS:
        nonnegative_int(row[field], f"{label} {field}")
    nonnegative_int(row["input_bytes"], f"{label} input_bytes")
    rss = nonnegative_int(row["rss_kib"], f"{label} rss_kib")
    if rss == 0:
        raise CaptureError(f"{label}: RSS receipt is zero")
    raw_samples = row["samples"]
    if not isinstance(raw_samples, list) or len(raw_samples) != 1 or not isinstance(raw_samples[0], dict):
        raise CaptureError(f"{label}: expected exactly one raw child sample")
    sample = raw_samples[0]
    validate_sample(sample, label)
    child, receipt_rss = raw_child(directory, row, label)
    equal(f"{label} raw RSS", rss, receipt_rss)
    equal(f"{label} child case", child.get("case"), row["case"])
    equal(f"{label} child phase", child.get("phase"), row["phase"])
    equal(f"{label} child sample", child.get("samples"), raw_samples)
    equal(f"{label} checksum p50", row["checksum_p50"], sample["checksum"])
    equal(f"{label} output p50", row["output_bytes_p50"], sample["output_bytes"])
    equal(f"{label} elapsed p50", row["elapsed_ns_p50"], sample["elapsed_ns"])
    equal(f"{label} elapsed mean", row["elapsed_ns_mean"], sample["elapsed_ns"])
    equal(f"{label} elapsed p95", row["elapsed_ns_p95"], sample["elapsed_ns"])
    equal(f"{label} elapsed p99", row["elapsed_ns_p99"], sample["elapsed_ns"])
    equal(f"{label} allocator calls p50", row["allocator_calls_p50"], sample["alloc_calls"])
    equal(f"{label} requested bytes p50", row["requested_bytes_p50"], sample["requested_bytes"])
    equal(f"{label} released bytes p50", row["released_bytes_p50"], sample["released_bytes"])
    equal(f"{label} peak live p50", row["peak_live_delta_p50"], sample["peak_live_delta"])
    equal(f"{label} retained p50", row["memory_retained_p50"], sample["memory_retained"])
    equal(f"{label} work p50", row["work_p50"], sample["work"])
    equal(f"{label} reads p50", row["reference_reads_p50"], sample["reference_reads"])
    equal(f"{label} live before p50", row["live_before_p50"], sample["live_before"])
    equal(f"{label} live after p50", row["live_after_p50"], sample["live_after"])
    expected_bytes = int(row["input_bytes"]) + int(sample["output_bytes"]) // repeat
    equal(f"{label} bytes per repeat", row["bytes_per_repeat_p50"], expected_bytes)
    expected_elapsed = int(sample["elapsed_ns"]) / repeat
    observed_elapsed = finite_number(row.get("elapsed_ns_per_repeat"), f"{label} elapsed per repeat")
    if abs(observed_elapsed - expected_elapsed) > max(1.0e-9, expected_elapsed * 1.0e-9):
        raise CaptureError(f"{label}: elapsed per-repeat normalization is inconsistent")
    expected_work = int(sample["work"]) / repeat
    observed_work = finite_number(row.get("work_per_repeat"), f"{label} work per repeat")
    if abs(observed_work - expected_work) > max(1.0e-9, expected_work * 1.0e-9):
        raise CaptureError(f"{label}: work per-repeat normalization is inconsistent")
    expected_reads = int(sample["reference_reads"]) / repeat
    observed_reads = finite_number(row.get("reference_reads_per_repeat"), f"{label} reads per repeat")
    if abs(observed_reads - expected_reads) > max(1.0e-9, expected_reads * 1.0e-9):
        raise CaptureError(f"{label}: reads per-repeat normalization is inconsistent")
    if contract is not None:
        geometry = contract["geometry"]
        expected_shape = geometry["kind"]
        equal(f"{label} execution shape kind", row.get("shape"), expected_shape)
        equal(f"{label} execution rows", row.get("rows"), geometry["rows"])
        equal(f"{label} execution columns", row.get("columns"), geometry["columns"])
        equal(f"{label} execution elements", row.get("elements"), geometry["elements"])
        expected_total_reads = contract["reference_reads"]
        if contract["row"].get("failure") == "Cancelled":
            # Cancellation is sticky for the child: the first resolver read
            # cancels the remaining repeats, so the total is one rather than
            # the ordinary per-repeat contract multiplied by repeat.
            expected_total_reads = 1
        else:
            expected_total_reads *= repeat
        equal(f"{label} timed reference reads", sample["reference_reads"], expected_total_reads)
    return {
        "case": str(row["case"]),
        "phase": str(row["phase"]),
        "sample_index": nonnegative_int(row["sample_index"], f"{label} sample_index"),
        "row": row,
        "sample": sample,
        "rss_kib": rss,
    }


def validate_preflight(
    directory: Path,
    expected_cases: list[str],
    matrix_rows: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    receipt = load(directory / "preflight.json")
    equal(f"{directory} preflight status", receipt.get("status"), "ok")
    equal(f"{directory} preflight cases", receipt.get("cases"), expected_cases)
    equal(f"{directory} preflight phase", receipt.get("phase"), "evaluate")
    equal(f"{directory} preflight warmups", receipt.get("warmups"), 0)
    equal(f"{directory} preflight iterations", receipt.get("iterations"), 1)
    observed_reads = receipt.get("reference_reads")
    descriptions = receipt.get("descriptions")
    if not isinstance(observed_reads, dict) or not isinstance(descriptions, dict):
        raise CaptureError(f"{directory} preflight is missing read bounds or descriptions")
    equal(f"{directory} preflight read keys", set(observed_reads), set(expected_cases))
    equal(f"{directory} preflight description keys", set(descriptions), set(expected_cases))
    for case in expected_cases:
        contract = matrix_rows[case]
        equal(
            f"{directory} preflight reads {case}",
            nonnegative_int(observed_reads.get(case), f"{directory} preflight reads {case}"),
            contract["reference_reads"],
        )
        description = descriptions[case]
        if not isinstance(description, dict):
            raise CaptureError(f"{directory} preflight description {case} is not an object")
        matrix_row = contract["row"]
        for field in ("source", "evaluation_path", "shape"):
            equal(f"{directory} preflight {case} {field}", description.get(field), matrix_row[field])
        if "expected" not in description or not same_expected_outcome(
            description["expected"], matrix_row["expected"]
        ):
            raise CaptureError(
                f"{directory} preflight expected outcome mismatch for {case}: "
                f"expected {matrix_row['expected']!r}, observed {description.get('expected')!r}"
            )
    return receipt


def validate_source_manifest(directory: Path) -> dict[str, Any]:
    manifest = load(directory / "source-manifest.json")
    before = manifest.get("before")
    after = manifest.get("after")
    if not isinstance(before, dict) or not isinstance(after, dict):
        raise CaptureError(f"{directory} source manifest lacks before/after snapshots")
    for field in (
        "git_head",
        "dirty_status_sha256_excluding_profile",
        "dirty_status_bytes_excluding_profile",
        "source_sha256",
        "workspace_source_sha256",
        "workspace_lock_sha256",
        "harness_sha256",
        "profile_input_sha256",
    ):
        if field not in before or field not in after:
            raise CaptureError(f"{directory} source manifest lacks {field}")
        equal(f"{directory} source manifest {field}", after[field], before[field])
    for field in (
        "source_sha256_unchanged",
        "workspace_source_sha256_unchanged",
        "profile_input_sha256_unchanged",
    ):
        equal(f"{directory} source manifest {field}", manifest.get(field), True)
    if not isinstance(manifest.get("binary_sha256"), str) or not manifest["binary_sha256"]:
        raise CaptureError(f"{directory} source manifest lacks binary identity")
    return manifest


def validate_freeze_receipt(summary: dict[str, Any]) -> dict[str, Any]:
    receipt = summary.get("candidate_freeze")
    if not isinstance(receipt, dict):
        raise CaptureError("capture summary lacks candidate freeze receipt")
    path = receipt.get("path")
    digest = receipt.get("sha256")
    manifest = receipt.get("manifest")
    if not isinstance(path, str) or not path:
        raise CaptureError("candidate freeze receipt lacks path")
    if not isinstance(digest, str) or re.fullmatch(r"[0-9a-f]{64}", digest) is None:
        raise CaptureError("candidate freeze receipt has malformed SHA-256")
    if not isinstance(manifest, dict):
        raise CaptureError("candidate freeze receipt lacks manifest")
    selected = manifest.get("selected_files")
    if not isinstance(selected, dict) or not selected:
        raise CaptureError("candidate freeze receipt has no selected file map")
    for relative, expected in selected.items():
        if not isinstance(relative, str) or not relative:
            raise CaptureError("candidate freeze selected file has no path")
        if not isinstance(expected, str) or re.fullmatch(r"[0-9a-f]{64}", expected) is None:
            raise CaptureError(f"candidate freeze selected file {relative!r} has malformed SHA-256")
    freeze_path = Path(path)
    if not freeze_path.is_file():
        freeze_path = HERE.parent / "gates" / "freeze.json"
    if not freeze_path.is_file():
        raise CaptureError("candidate freeze file unavailable for digest verification")
    equal("candidate freeze receipt SHA-256", sha256_file(freeze_path), digest)
    on_disk = load(freeze_path)
    equal("candidate freeze receipt manifest", on_disk, manifest)
    captures = summary.get("captures")
    if not isinstance(captures, list):
        raise CaptureError("capture summary has no capture list for freeze binding")
    candidate_capture = next(
        (capture for capture in captures if isinstance(capture, dict) and capture.get("label") == "candidate-final"),
        None,
    )
    if not isinstance(candidate_capture, dict):
        raise CaptureError("capture summary has no candidate-final freeze binding")
    equal("candidate capture freeze SHA-256", candidate_capture.get("freeze_sha256"), digest)
    equal("candidate capture freeze path", candidate_capture.get("freeze_path"), path)
    if "freeze_base_commit" in candidate_capture:
        equal("candidate capture freeze base", candidate_capture["freeze_base_commit"], manifest.get("base_commit"))
    return {"path": path, "sha256": digest, "selected_files": len(selected), "verified_path": str(freeze_path)}


def validate_capture(
    directory: Path,
    expected_cases: list[str],
    phases: list[str],
    warmups: int,
    samples: int,
    matrix_rows: dict[str, dict[str, Any]] | None = None,
) -> dict[str, Any]:
    environment = load(directory / "environment.json")
    equal(f"{directory} environment cases", environment.get("cases"), expected_cases)
    equal(f"{directory} environment phases", environment.get("phases"), phases)
    equal(f"{directory} environment warmups", environment.get("warmups_per_child"), warmups)
    equal(f"{directory} environment samples", environment.get("samples_per_group"), samples)
    cleanup = load(directory / "target-cleanup.json")
    equal(f"{directory} target cleanup", cleanup.get("removed"), True)
    source_manifest = validate_source_manifest(directory)
    if matrix_rows is None:
        matrix_rows = {}
    preflight = validate_preflight(directory, expected_cases, matrix_rows) if matrix_rows else None
    rows_seen = rows(directory)
    expected_count = len(expected_cases) * len(phases) * samples
    equal(f"{directory} row count", len(rows_seen), expected_count)
    groups: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for row in rows_seen:
        key = (row.get("case"), row.get("phase"))
        if key[0] not in expected_cases or key[1] not in phases:
            raise CaptureError(f"{directory}: unexpected group {key}")
        contract = matrix_rows.get(key[0]) if matrix_rows else None
        if matrix_rows and contract is None:
            raise CaptureError(f"{directory}: no matrix contract for {key[0]}")
        checked = validate_row(
            row,
            directory,
            f"{directory} {key} sample {row.get('sample_index')}",
            warmups,
            contract,
        )
        groups[key].append(checked)
    expected_groups = {(case, phase) for case in expected_cases for phase in phases}
    equal(f"{directory} groups", set(groups), expected_groups)
    normalized: dict[str, Any] = {}
    for key, group in groups.items():
        indices = sorted(item["sample_index"] for item in group)
        equal(f"{directory} {key} sample indices", indices, list(range(1, samples + 1)))
        invariants = ("checksum", "expected", "shape", "rows", "columns", "elements", "input_bytes", "output_bytes_p50", "reference_reads_p50", "work_p50")
        first = group[0]["row"]
        for field in invariants:
            values = {item["row"].get(field) for item in group}
            if len(values) != 1:
                raise CaptureError(f"{directory} {key}: {field} changed across samples: {values}")
        binaries = {item["row"].get("binary_sha256") for item in group}
        heads = {item["row"].get("source_git_head") for item in group}
        equal(f"{directory} {key} binary", len(binaries), 1)
        equal(f"{directory} {key} binary manifest", binaries.pop(), source_manifest["binary_sha256"])
        equal(f"{directory} {key} source head", len(heads), 1)
        normalized["%s/%s" % key] = group_summary(group)
    return {
        "environment": environment,
        "source_manifest": source_manifest,
        "preflight": preflight,
        "groups": normalized,
        "rows": len(rows_seen),
    }


def percentile(values: list[float], fraction: float) -> float:
    if not values:
        raise CaptureError("cannot calculate a percentile of an empty sample")
    ordered = sorted(values)
    index = (len(ordered) - 1) * fraction
    low = math.floor(index)
    high = math.ceil(index)
    if low == high:
        return ordered[low]
    return ordered[low] + (ordered[high] - ordered[low]) * (index - low)


def metric_values(group: list[dict[str, Any]]) -> dict[str, list[float]]:
    result: dict[str, list[float]] = defaultdict(list)
    for item in group:
        sample = item["sample"]
        repeat = nonnegative_int(item["row"]["repeat"], "group repeat")
        result["elapsed_ns_per_repeat"].append(sample["elapsed_ns"] / repeat)
        result["rss_kib"].append(float(item["rss_kib"]))
        result["output_bytes_per_repeat"].append(sample["output_bytes"] / repeat)
        result["alloc_calls_per_repeat"].append(sample["alloc_calls"] / repeat)
        result["requested_bytes_per_repeat"].append(sample["requested_bytes"] / repeat)
        result["released_bytes_per_repeat"].append(sample["released_bytes"] / repeat)
        result["peak_live_delta"].append(float(sample["peak_live_delta"]))
        result["memory_retained"].append(float(sample["memory_retained"]))
        result["work_per_repeat"].append(sample["work"] / repeat)
        result["reference_reads_per_repeat"].append(sample["reference_reads"] / repeat)
    return result


def group_summary(group: list[dict[str, Any]]) -> dict[str, Any]:
    values = metric_values(group)
    result: dict[str, Any] = {"samples": len(group), "metrics": {}}
    for name, sequence in values.items():
        result["metrics"][name] = {
            "min": min(sequence),
            "median": statistics.median(sequence),
            "p95": percentile(sequence, 0.95),
            "p99": percentile(sequence, 0.99),
            "max": max(sequence),
            "values": sequence,
        }
    row = group[0]["row"]
    # These fields describe the produced value and the contract-required
    # reference-read count. Work, output bytes, and allocator/resource
    # counters remain accounting metrics so an implementation can change its
    # amount of work without being reported as a semantic mismatch.
    result["correctness"] = {
        "checksum": row["checksum_p50"],
        "expected": row.get("expected"),
        "shape": row.get("shape"),
        "rows": row.get("rows"),
        "columns": row.get("columns"),
        "elements": row.get("elements"),
        "input_bytes": row.get("input_bytes"),
        "reference_reads": row.get("reference_reads_p50"),
    }
    return result


def delta(candidate: float, baseline: float) -> dict[str, Any]:
    difference = candidate - baseline
    relative = None if baseline == 0 else difference / baseline
    return {
        "baseline": baseline,
        "candidate": candidate,
        "difference": difference,
        "relative": relative,
        "absolute_threshold": abs(relative) >= THRESHOLD if relative is not None else candidate != baseline,
        "positive_regression": relative >= THRESHOLD if relative is not None else candidate > baseline,
    }


def bootstrap_seed(group: str, metric: str, quantile: str) -> int:
    seed = BOOTSTRAP_SEED
    for byte in f"{group}|{metric}|{quantile}".encode("utf-8"):
        seed = ((seed ^ byte) * 1_099_511_628_211) & ((1 << 63) - 1)
    return seed


def bootstrap_delta_interval(
    baseline_values: list[float],
    candidate_values: list[float],
    fraction: float,
    seed: int,
) -> dict[str, Any]:
    if not baseline_values or not candidate_values:
        raise CaptureError("cannot bootstrap an empty comparison sample")
    rng = random.Random(seed)
    differences: list[float] = []
    relatives: list[float] = []
    for _ in range(BOOTSTRAP_REPLICATES):
        baseline_resample = [baseline_values[rng.randrange(len(baseline_values))] for _ in baseline_values]
        candidate_resample = [candidate_values[rng.randrange(len(candidate_values))] for _ in candidate_values]
        baseline_statistic = percentile(baseline_resample, fraction)
        candidate_statistic = percentile(candidate_resample, fraction)
        difference = candidate_statistic - baseline_statistic
        differences.append(difference)
        if baseline_statistic != 0:
            relatives.append(difference / baseline_statistic)
    lower = (1.0 - BOOTSTRAP_CONFIDENCE) / 2.0
    upper = 1.0 - lower
    return {
        "seed": seed,
        "replicates": BOOTSTRAP_REPLICATES,
        "confidence": BOOTSTRAP_CONFIDENCE,
        "difference_ci95": {
            "low": percentile(differences, lower),
            "high": percentile(differences, upper),
        },
        "relative_ci95": (
            {
                "low": percentile(relatives, lower),
                "high": percentile(relatives, upper),
            }
            if relatives
            else None
        ),
    }


def compare_groups(baseline: dict[str, Any], candidate: dict[str, Any]) -> dict[str, Any]:
    base_groups = baseline["groups"]
    cand_groups = candidate["groups"]
    shared = sorted(set(base_groups) & set(cand_groups))
    correctness_mismatches: list[dict[str, Any]] = []
    comparisons: dict[str, Any] = {}
    flags: list[dict[str, Any]] = []
    metric_names = (
        "elapsed_ns_per_repeat", "rss_kib", "output_bytes_per_repeat", "alloc_calls_per_repeat",
        "requested_bytes_per_repeat", "released_bytes_per_repeat", "peak_live_delta",
        "memory_retained", "work_per_repeat", "reference_reads_per_repeat",
    )
    quantiles = ("median", "p95", "p99")
    for key in shared:
        base_correct = base_groups[key]["correctness"]
        cand_correct = cand_groups[key]["correctness"]
        if base_correct != cand_correct:
            correctness_mismatches.append({"group": key, "baseline": base_correct, "candidate": cand_correct})
        group_comparison: dict[str, Any] = {}
        for metric in metric_names:
            group_comparison[metric] = {}
            for quantile in quantiles:
                base_value = base_groups[key]["metrics"][metric][quantile]
                cand_value = cand_groups[key]["metrics"][metric][quantile]
                measured = delta(cand_value, base_value)
                group_comparison[metric][quantile] = measured
                if measured["absolute_threshold"]:
                    flag = {
                        "group": key,
                        "metric": metric,
                        "quantile": quantile,
                        "sample_count": {
                            "baseline": len(base_groups[key]["metrics"][metric]["values"]),
                            "candidate": len(cand_groups[key]["metrics"][metric]["values"]),
                        },
                        **measured,
                    }
                    flag["bootstrap"] = bootstrap_delta_interval(
                        base_groups[key]["metrics"][metric]["values"],
                        cand_groups[key]["metrics"][metric]["values"],
                        {"median": 0.5, "p95": 0.95, "p99": 0.99}[quantile],
                        bootstrap_seed(key, metric, quantile),
                    )
                    flags.append(flag)
        comparisons[key] = group_comparison
    if correctness_mismatches:
        raise CaptureError("matched controls changed correctness/checksum:\n" + json.dumps(correctness_mismatches, indent=2, sort_keys=True))
    return {
        "threshold": THRESHOLD,
        "bootstrap": {
            "seed": BOOTSTRAP_SEED,
            "replicates": BOOTSTRAP_REPLICATES,
            "confidence": BOOTSTRAP_CONFIDENCE,
            "method": "independent resampling of normalized baseline and candidate samples",
        },
        "matched_groups": len(shared),
        "comparisons": comparisons,
        "flags": flags,
        "positive_regressions": [flag for flag in flags if flag["positive_regression"]],
        "candidate_only_groups": sorted(set(cand_groups) - set(base_groups)),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--results", type=Path, default=DEFAULT_RESULTS)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--expected-warmups", type=int)
    parser.add_argument("--expected-samples", type=int)
    args = parser.parse_args()
    results = args.results.resolve()
    matrix = matrix_declarations()
    summary = load(results / "capture-summary.json")
    freeze_receipt = validate_freeze_receipt(summary)
    captures = summary.get("captures")
    if not isinstance(captures, list) or len(captures) != 2:
        raise CaptureError("capture-summary.json must contain baseline and candidate captures")
    controls = summary.get("controls")
    date_cases = summary.get("date_cases")
    if not isinstance(controls, list) or not isinstance(date_cases, list):
        raise CaptureError("capture summary has malformed case declarations")
    equal("summary controls", controls, matrix["controls"])
    equal("summary date cases", date_cases, matrix["date_cases"])
    equal("summary phases", summary.get("phases"), matrix["phases"])
    equal("summary baseline commit", summary.get("baseline_commit"), matrix["baseline_commit"])
    equal("summary contract hash", summary.get("contract_sha256"), matrix["contract_sha256"])
    equal("summary oracle hash", summary.get("oracle_sha256"), matrix["oracle_vectors_sha256"])
    warmups = nonnegative_int(summary.get("warmups_per_child"), "summary warmups")
    samples = nonnegative_int(summary.get("samples_per_group"), "summary samples")
    expected_warmups = args.expected_warmups if args.expected_warmups is not None else matrix["warmups"]
    expected_samples = args.expected_samples if args.expected_samples is not None else matrix["samples"]
    equal("summary warmups", warmups, expected_warmups)
    equal("summary samples", samples, expected_samples)
    profile_before = summary.get("profile_input_sha256_before")
    profile_after = summary.get("profile_input_sha256_after")
    if not isinstance(profile_before, dict) or not isinstance(profile_after, dict):
        raise CaptureError("capture summary lacks profile input snapshots")
    equal("summary profile input snapshot", profile_after, profile_before)
    phases = matrix["phases"]
    baseline_dir = results / f"baseline-{summary['baseline_commit']}"
    candidate_dir = results / "candidate-final"
    baseline = validate_capture(
        baseline_dir, controls, phases, warmups, samples, matrix["rows"]
    )
    candidate = validate_capture(
        candidate_dir, list(controls) + list(date_cases), phases, warmups, samples, matrix["rows"]
    )
    equal(
        "baseline profile input snapshot",
        baseline["source_manifest"]["before"]["profile_input_sha256"],
        profile_before,
    )
    equal(
        "candidate profile input snapshot",
        candidate["source_manifest"]["before"]["profile_input_sha256"],
        profile_before,
    )
    comparison = compare_groups(baseline, candidate)
    output = {
        "status": "ok",
        "baseline_commit": summary.get("baseline_commit"),
        "contract_sha256": summary.get("contract_sha256"),
        "oracle_sha256": summary.get("oracle_sha256"),
        "candidate_freeze": freeze_receipt,
        "matrix": {
            "path": str(MATRIX_PATH),
            "warmups": matrix["warmups"],
            "samples": matrix["samples"],
            "controls": len(matrix["controls"]),
            "date_cases": len(matrix["date_cases"]),
        },
        "warmups_per_child": warmups,
        "samples_per_group": samples,
        "baseline_rows": baseline["rows"],
        "candidate_rows": candidate["rows"],
        "baseline_groups": len(baseline["groups"]),
        "candidate_groups": len(candidate["groups"]),
        "comparison": comparison,
        "captures": {
            "baseline": {"environment": baseline["environment"]},
            "candidate": {"environment": candidate["environment"]},
        },
    }
    destination = args.output.resolve() if args.output else results / "profile-analysis.json"
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(json.dumps(output, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({
        "status": output["status"],
        "baseline_rows": output["baseline_rows"],
        "candidate_rows": output["candidate_rows"],
        "matched_groups": comparison["matched_groups"],
        "threshold_flags": len(comparison["flags"]),
        "positive_regressions": len(comparison["positive_regressions"]),
        "output": str(destination),
    }, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except CaptureError as error:
        raise SystemExit(f"analyze_profile.py: {error}") from error
