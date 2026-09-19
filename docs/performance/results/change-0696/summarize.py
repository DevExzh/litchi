#!/usr/bin/env python3
"""Summarize every retained native leg for the 0696 workflow matrix.

The raw TSV files remain the authoritative per-leg evidence.  This script
adds descriptive statistics, explicit A/A drift pairs, candidate-versus-
baseline pairs, and semantic metadata checks without pooling legs.
"""

import hashlib
import json
import math
import random
import statistics
from pathlib import Path


P = Path(__file__).resolve().parent
EXPECTED_WORKFLOWS = {
    "one": {"real", "control", "generated", "notes-poi", "notes-lo"},
    "noop": {"real", "control", "generated", "notes-poi", "notes-lo"},
    "two": {"real", "control", "generated"},
}
EXPECTED_LEGS = {"a0", "a1", "a2", "a3", "b0", "b1"}
STAT_FIELDS = ("p50_ns", "mean_ns", "p95_ns", "p99_ns")
rng = random.Random(693)


def quantile(values, fraction):
    if not values:
        raise AssertionError("quantile of an empty sample")
    return sorted(values)[max(0, math.ceil(len(values) * fraction) - 1)]


def bootstrap_median(values):
    count = len(values)
    samples = [statistics.median(rng.choices(values, k=count)) for _ in range(2000)]
    return [quantile(samples, 0.025), quantile(samples, 0.975)]


def manifest_paths(prefix):
    combined = P / f"{prefix}-runs.json"
    if combined.exists():
        return [combined]
    paths = [P / f"{prefix}-runs-{suffix}.json" for suffix in ("baseline", "compare")]
    return [path for path in paths if path.exists()]


def load_runs(prefix):
    paths = manifest_paths(prefix)
    if not paths:
        raise SystemExit(f"missing {prefix} run manifest")
    records = []
    for path in paths:
        records.extend(json.loads(path.read_text()))
    return records


def output_path(value):
    path = Path(value)
    return path if path.is_absolute() else P / path


def parse(path, expected_samples=None):
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
    if expected_samples is not None and declared != expected_samples:
        raise AssertionError(f"{path}: expected {expected_samples} samples, got {declared}")
    if len(rows) != declared:
        raise AssertionError(f"{path}: declared {declared} samples, found {len(rows)}")
    if [row["sample"] for row in rows] != list(range(declared)):
        raise AssertionError(f"{path}: sample indices are not contiguous")
    metadata = {}
    for line in lines:
        fields = line.split("\t")
        if fields and not fields[0].isdigit() and not fields[0].startswith("sample"):
            metadata[fields[0]] = fields[1:]
    return header[1:], rows, metadata


def metadata_value(metadata, key):
    values = metadata.get(key)
    if values is None:
        return None
    return values[0] if len(values) == 1 else values


def target_coordinates(metadata, key):
    values = metadata.get(key)
    if values is None or len(values) != 2:
        raise AssertionError(f"missing or malformed {key} metadata")
    if not values[0].startswith("slide:") or not values[1].startswith("shape:"):
        raise AssertionError(f"malformed {key} metadata: {values}")
    return {
        "slide": int(values[0].removeprefix("slide:")),
        "shape": int(values[1].removeprefix("shape:")),
    }


def normalized_metadata(metadata, workflow):
    normalized = {key: list(values) for key, values in metadata.items()}
    normalized.setdefault("workflow", [workflow])
    if normalized["workflow"] != [workflow]:
        raise AssertionError(f"workflow metadata mismatch: {normalized['workflow']} vs {workflow}")
    if workflow == "one":
        target = target_coordinates(normalized, "target")
        if "target1" in normalized or "target2" in normalized:
            raise AssertionError("one-edit output unexpectedly has multi-target metadata")
        required = {
            "before_revision_sha256",
            "after_revision_sha256",
            "before_semantic_sha256",
            "after_semantic_sha256",
            "candidate_archive_sha256",
            "reopened_target_text_sha256",
            "correctness_target_text",
        }
        missing = required - set(normalized)
        if missing:
            raise AssertionError(f"one-edit output lacks metadata: {sorted(missing)}")
        if normalized["before_revision_sha256"] == normalized["after_revision_sha256"]:
            raise AssertionError("one-edit revision did not change")
        if normalized["before_semantic_sha256"] == normalized["after_semantic_sha256"]:
            raise AssertionError("one-edit semantic digest did not change")
        normalized["target_coordinates"] = [target]
    elif workflow == "noop":
        target = target_coordinates(normalized, "target1")
        required = {
            "commit_is_changed": ["false"],
            "revision_identical": ["true"],
            "output_identical": ["true"],
        }
        for key, expected in required.items():
            if normalized.get(key) != expected:
                raise AssertionError(f"no-op {key} is {normalized.get(key)!r}, expected {expected}")
        if normalized.get("before_revision_sha256") != normalized.get("after_revision_sha256"):
            raise AssertionError("no-op revision metadata differs")
        if normalized.get("before_archive_sha256") != normalized.get("after_archive_sha256"):
            raise AssertionError("no-op archive metadata differs")
        if normalized.get("before_semantic_sha256") != normalized.get("after_semantic_sha256"):
            raise AssertionError("no-op semantic metadata differs")
        required_metadata = {
            "before_revision_sha256",
            "after_revision_sha256",
            "before_archive_sha256",
            "after_archive_sha256",
            "before_semantic_sha256",
            "after_semantic_sha256",
            "correctness_target_text_sha256",
        }
        missing = required_metadata - set(normalized)
        if missing:
            raise AssertionError(f"no-op output lacks metadata: {sorted(missing)}")
        normalized["target_coordinates"] = [target]
    elif workflow == "two":
        first = target_coordinates(normalized, "target1")
        second = target_coordinates(normalized, "target2")
        if first["slide"] == second["slide"]:
            raise AssertionError("two-edit targets are not on distinct slides")
        if "target" in normalized:
            raise AssertionError("two-edit output unexpectedly has one-edit target metadata")
        required = {
            "before_revision_sha256",
            "after_revision_sha256",
            "before_semantic_sha256",
            "after_semantic_sha256",
            "candidate_archive_sha256",
            "reopened_target1_text_sha256",
            "reopened_target2_text_sha256",
            "correctness_target_text",
        }
        missing = required - set(normalized)
        if missing:
            raise AssertionError(f"two-edit output lacks metadata: {sorted(missing)}")
        if normalized["before_revision_sha256"] == normalized["after_revision_sha256"]:
            raise AssertionError("two-edit revision did not change")
        if normalized["before_semantic_sha256"] == normalized["after_semantic_sha256"]:
            raise AssertionError("two-edit semantic digest did not change")
        if normalized["reopened_target1_text_sha256"] != normalized["reopened_target2_text_sha256"]:
            raise AssertionError("two-edit target markers do not agree")
        normalized["target_coordinates"] = [first, second]
    else:
        raise AssertionError(f"unknown workflow {workflow!r}")
    return normalized


def stable_metadata(metadata):
    # Keep every correctness and provenance field.  The explicit sample and
    # warmup counts are run-shape metadata and are checked independently.
    return {
        key: value
        for key, value in metadata.items()
        if key not in {"samples", "warmups"}
    }


def stats(values):
    return {
        "samples": len(values),
        "p50_ns": statistics.median(values),
        "mean_ns": statistics.mean(values),
        "p95_ns": quantile(values, 0.95),
        "p99_ns": quantile(values, 0.99),
        "min_ns": min(values),
        "max_ns": max(values),
        "median_bootstrap_95ci_ns": bootstrap_median(values),
    }


def percent_delta(value, baseline):
    if baseline == 0:
        return None
    return (value - baseline) * 100.0 / baseline


def comparison(case, phase, kind, candidate_leg, baseline_leg, by_key):
    candidate = by_key[(case, candidate_leg, phase)]
    baseline = by_key[(case, baseline_leg, phase)]
    delta = {field: candidate[field] - baseline[field] for field in STAT_FIELDS}
    delta_pct = {
        field: percent_delta(candidate[field], baseline[field]) for field in STAT_FIELDS
    }
    return {
        "case": case,
        "phase": phase,
        "kind": kind,
        "pair": f"{candidate_leg}/{baseline_leg}",
        "candidate_leg": candidate_leg,
        "baseline_leg": baseline_leg,
        "candidate": {field: candidate[field] for field in STAT_FIELDS},
        "baseline": {field: baseline[field] for field in STAT_FIELDS},
        "delta_ns": delta,
        "delta_pct": delta_pct,
    }


def expected_cases():
    cases_path = P / "cases.json"
    if cases_path.exists():
        return {
            item["case"]: item["workflow"] for item in json.loads(cases_path.read_text())
        }
    return {
        f"{workflow}-{name}": workflow
        for workflow, names in EXPECTED_WORKFLOWS.items()
        for name in sorted(names)
    }


def main():
    cases = expected_cases()
    runs = load_runs("native")
    if len({(run["case"], run["leg"]) for run in runs}) != len(runs):
        raise AssertionError("duplicate native case/leg manifest entry")
    summary = []
    bindings = {}
    leg_metadata = {}
    by_key = {}
    present_legs = set()
    for run in runs:
        if run.get("exit_code") != 0:
            raise AssertionError(f"native run failed: {run}")
        case = run["case"]
        if case not in cases:
            raise AssertionError(f"unexpected native case {case}")
        workflow = cases[case]
        leg = run["leg"]
        present_legs.add(leg)
        path = output_path(run["output"])
        if hashlib.sha256(path.read_bytes()).hexdigest() != run["output_sha256"]:
            raise AssertionError(f"native output hash mismatch: {path}")
        phases, rows, metadata = parse(path, expected_samples=100)
        normalized = normalized_metadata(metadata, workflow)
        stable = stable_metadata(normalized)
        case_legs = leg_metadata.setdefault(case, {})
        if case_legs and stable != next(iter(case_legs.values())):
            raise AssertionError(f"semantic/provenance metadata drift in {case} leg {leg}")
        case_legs[leg] = stable
        bindings[case] = stable
        for phase in phases:
            if not phase.endswith("_ns"):
                raise AssertionError(f"unexpected native timing column {phase}")
            values = [row[phase] for row in rows]
            record = dict(case=case, workflow=workflow, leg=leg, phase=phase)
            record.update(stats(values))
            summary.append(record)
            by_key[(case, leg, phase)] = record
    if set(bindings) != set(cases):
        raise AssertionError(f"case coverage mismatch: {set(cases) ^ set(bindings)}")
    if not present_legs.issubset(EXPECTED_LEGS):
        raise AssertionError(f"unexpected native legs: {present_legs - EXPECTED_LEGS}")
    if "b0" in present_legs or "b1" in present_legs:
        required_legs = {"a0", "a1", "a2", "a3", "b0", "b1"}
    else:
        required_legs = {"a0", "a1"}
    if present_legs != required_legs:
        raise AssertionError(f"native leg coverage is {present_legs}, expected {required_legs}")
    for case in cases:
        case_legs = set(leg_metadata[case])
        if case_legs != present_legs:
            raise AssertionError(f"{case} has incomplete leg coverage: {case_legs}")

    comparisons = []
    for case in cases:
        phases = sorted({record["phase"] for record in summary if record["case"] == case})
        for phase in phases:
            comparisons.append(comparison(case, phase, "aa", "a1", "a0", by_key))
            if {"a2", "a3", "b0", "b1"}.issubset(present_legs):
                comparisons.append(comparison(case, phase, "aa", "a3", "a2", by_key))
                comparisons.append(comparison(case, phase, "candidate", "b0", "a2", by_key))
                comparisons.append(comparison(case, phase, "candidate", "b1", "a3", by_key))

    (P / "native-summary.json").write_text(json.dumps(summary, indent=2) + "\n")
    (P / "semantic-bindings.json").write_text(json.dumps(bindings, indent=2) + "\n")
    (P / "native-leg-metadata.json").write_text(json.dumps(leg_metadata, indent=2) + "\n")
    (P / "native-comparisons.json").write_text(json.dumps(comparisons, indent=2) + "\n")
    print(f"Native summary: {len(summary)} leg/phase records across {len(bindings)} cases")
    print(f"Native comparisons: {len(comparisons)} explicit A/A or candidate pairs")
    if "b0" in present_legs:
        print("Candidate comparisons are descriptive b0/a2 and b1/a3 pairs; raw legs remain separate.")
    else:
        print("Baseline-only run: emitted a0/a1 drift pairs; candidate comparisons await compare legs.")


if __name__ == "__main__":
    main()
